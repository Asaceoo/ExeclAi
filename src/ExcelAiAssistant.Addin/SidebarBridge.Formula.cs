using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Skills;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// SidebarBridge 分域 partial（阶段 A 机械拆分，见 docs/AI交接文档.md）。
    /// 公共 API 即 WebView2 前端契约（hostObjects.excelAiBridge.*），签名不得改动。
    /// </summary>
    public sealed partial class SidebarBridge : IDisposable
    {
        public Task<string> GenerateFormula(string dataAreas, string targetAreas, string requirement)
        {
            return ExecuteAsync(async () =>
            {
                var description = (requirement ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(description))
                {
                    throw new ArgumentException("请描述你想让公式做什么，例如「计算每行数值占总计的百分比」。", nameof(requirement));
                }

                var dataAreaList = NormalizeAreas(dataAreas);
                var targetAreaList = NormalizeAreas(targetAreas);
                if (targetAreaList.Count == 0)
                {
                    targetAreaList.Add(CaptureCurrentSelectionAddress());
                }

                // A1 构造器优先链路：先尝试 FormulaBuilder 确定性构造（零模型调用、内化表头钳制与引用卫生），
                // 未命中模板/构造失败 → 落入下方模型自由生成兜底（builder=false）。
                try
                {
                    var headers = TryReadHeaderNames(dataAreaList.FirstOrDefault());
                    var sourceWidth = TryReadSourceWidth(dataAreaList.FirstOrDefault());
                    var intent = FormulaBuilder.ParseIntent(description, headers, sourceWidth);
                    if (intent != null)
                    {
                        var builtFormula = FormulaBuilder.Build(new FormulaBuilder.BuildOptions
                        {
                            Template = intent.Template,
                            SourceAddress = dataAreaList.Count > 0 ? dataAreaList[0] : targetAreaList[0],
                            ValueColumn = intent.ValueColumn,
                            CriteriaColumn = intent.CriteriaColumn,
                            Criteria = intent.Criteria,
                            WeightColumn = intent.WeightColumn,
                            HeaderRows = intent.HeaderRows,
                            RankOrder = intent.RankOrder
                        });
                        DiagnosticsLog.Write("formula-builder hit: " + intent.Template + " <- " + description);
                        return JsonConvert.SerializeObject(new
                        {
                            ok = true,
                            formula = builtFormula,
                            explanation = "（模板构造）命中「" + (intent.MatchedKeyword ?? intent.Template) + "」模板，由公式构造器确定性生成：表头行已自动跳过、引用已规范化。如需调整条件或列，直接修改需求再生成即可。",
                            dataAreas = dataAreaList,
                            targetAreas = targetAreaList,
                            builder = true
                        });
                    }
                }
                catch (ArgumentException builderEx)
                {
                    // 构造器参数不完整（如条件值缺失）→ 落模型兜底，把构造器提示并入需求上下文
                    DiagnosticsLog.Write("formula-builder miss: " + builderEx.Message);
                }
                catch (Exception builderEx2)
                {
                    DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs builder-priority  " + builderEx2.GetType().Name + ": " + builderEx2.Message);
                }

                var profile = RequireProfile();
                using var client = new OpenAiCompatibleClient(profile);
                using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(90));

                var prompt = new StringBuilder();
                if (dataAreaList.Count > 0)
                {
                    prompt.AppendLine("\u6570\u636e\u9009\u533a\uff08\u53ef\u53c2\u8003\u7684\u6e90\u6570\u636e\uff0c\u53ea\u8bfb\uff0c\u4e0d\u4f1a\u6539\u52a8\uff09\uff1a");
                    foreach (var area in dataAreaList.Take(6))
                    {
                        prompt.Append("- ").AppendLine(area);
                        var preview = TryReadRangePreview(area);
                        if (preview != null)
                        {
                            prompt.AppendLine("  \u9884\u89c8\uff1a" + preview);
                        }
                    }
                }
                else
                {
                    prompt.AppendLine("\u6570\u636e\u9009\u533a\uff1a\uff08\u672a\u6307\u5b9a\uff09");
                }

                prompt.AppendLine("\u76ee\u6807\u9009\u533a\uff08\u516c\u5f0f\u8981\u586b\u5165\u7684\u4f4d\u7f6e\uff0c\u8bf7\u6309\u9996\u683c\u7528\u76f8\u5bf9/\u7edd\u5bf9\u5f15\u7528\u914d\u5408\u5411\u4e0b\u6216\u5411\u53f3\u586b\u5145\uff09\uff1a");
                foreach (var area in targetAreaList)
                {
                    prompt.Append("- ").AppendLine(area);
                }

                prompt.AppendLine("\u9700\u6c42\uff1a" + description);
                prompt.AppendLine();
                prompt.AppendLine(
                    "请生成一条 Excel 公式。你的输出只能包含「说明」与「公式」两部分，不要多余文字、不要代码围栏：");
                prompt.AppendLine("说明：请简洁、有结构，严格按下面 4 行固定格式输出，每行只有一句话，让新手一眼看懂——");
                prompt.AppendLine("调用函数：<这条公式调用哪些函数，如 SUM、IF、VLOOKUP、COUNTIF 等>");
                prompt.AppendLine("括号内：<每个参数各代表什么、分别在算什么，多项用“；”分隔>");
                prompt.AppendLine("括号外：<= 与函数名本身分别表示什么>");
                prompt.AppendLine("执行顺序：<有嵌套时按“内层先算 → 外层再拿结果算”写出；无嵌套则写“直接依次计算”>");
                prompt.AppendLine("公式：（以 = 开头的单条公式）");

                var input = new List<ChatMessage>
                {
                    ChatMessage.System(
                        "你是 Excel 公式专家。根据用户给的数据选区与目标选区、需求，生成一条正确、能填充的公式，并用简洁、有结构的方式解释，避免一大段文字。\n输出格式绝对按照：\n1. 「说明：」后严格换行输出 4 行，每行一句话，行首标签固定为：调用函数：、括号内：、括号外：、执行顺序：。其中括号外要讲清开头的 = 与函数名本身的含义；执行顺序若有嵌套要按“内层先算 → 外层再拿结果算”写出。\n2. 最后以「公式：」开头的单条 Excel 公式。\n不要任何其他内容、不要 markdown 代码围栏。"),
                    ChatMessage.User(prompt.ToString())
                };

                var result = await client.CompleteAsync(input, timeoutSource.Token).ConfigureAwait(false);
                var (formula, explanation) = SplitGeneratedFormula(result?.Content ?? string.Empty);
                if (string.IsNullOrWhiteSpace(formula))
                {
                    throw new InvalidOperationException("\u6a21\u578b\u672a\u80fd\u751f\u6210\u6709\u6548\u516c\u5f0f\uff0c\u8bf7\u91cd\u8bd5\u6216\u8c03\u6574\u63cf\u8ff0\u3002");
                }

                return JsonConvert.SerializeObject(new
                {
                    ok = true,
                    formula,
                    explanation,
                    dataAreas = dataAreaList,
                    targetAreas = targetAreaList
                });
            });
        }

        /// <summary>
        /// 公式助手：把已生成的公式插入到目标选区首格（可多选），写入为相对公式并记录撤销快照，
        /// 方便用户随后向下/向右填充；插入前捕获目标首格旧公式（D2 公式历史 diff），
        /// 返回更新后的可回滚历史。
        /// </summary>
        public Task<string> InsertFormula(string targetAreas, string formula)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(formula))
                {
                    throw new ArgumentException("没有可插入的公式。请先生成公式。", nameof(formula));
                }

                var targetAreaList = NormalizeAreas(targetAreas);
                if (targetAreaList.Count == 0)
                {
                    targetAreaList.Add(CaptureCurrentSelectionAddress());
                }

                // D2 公式历史 diff：插入前捕获每个目标首格的旧公式（读不到/非公式记 null）
                var beforeFormulas = new List<string>();
                foreach (var area in targetAreaList)
                {
                    beforeFormulas.Add(TryReadCellFormula(area));
                }

                var host = CreateHost(CurrentTurnLabel());
                var operationIds = new List<string>();
                var inserted = new List<string>();
                foreach (var area in targetAreaList)
                {
                    var (sheetName, rangeAddress) = SplitSheetAndRange(area);
                    var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["address"] = rangeAddress,
                        ["formula"] = formula
                    };
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        args["sheet"] = sheetName;
                    }

                    var result = host.InsertFormulaToRange(args);
                    if (!result.Ok)
                    {
                        var reason = result.Data != null ? Convert.ToString(result.Data) : result.ErrorCode;
                        throw new InvalidOperationException($"插入到 {area} 失败：{reason ?? "未知错误"}");
                    }

                    var operationId = Convert.ToString(ExcelToolHostHelpers.GetOperationId(result));
                    if (!string.IsNullOrEmpty(operationId))
                    {
                        operationIds.Add(operationId);
                    }

                    inserted.Add(area);
                }

                // D2 公式历史 diff：全部插入成功后落盘历史（失败不落，避免半截记录）
                try
                {
                    var change = FormulaDiff.DescribeChange(
                        beforeFormulas.Count > 0 ? beforeFormulas[0] : null, formula);
                    AppendFormulaHistory(new FormulaHistoryEntry
                    {
                        Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        Target = targetAreaList[0],
                        OldFormula = beforeFormulas.Count > 0 ? beforeFormulas[0] : null,
                        NewFormula = formula.Trim(),
                        Change = change
                    });
                }
                catch (Exception historyEx)
                {
                    // 历史落盘失败不阻断插入主流程
                    DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs formula-history  " + historyEx.GetType().Name + ": " + historyEx.Message);
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    inserted,
                    operation_ids = operationIds.ToArray(),
                    change = FormulaDiff.DescribeChange(
                        beforeFormulas.Count > 0 ? beforeFormulas[0] : null, formula),
                    history = GetHistoryRecords()
                }));
            });
        }

        /// <summary>D2 公式历史 diff：读目标首格当前公式（无工作簿/非公式/COM 失败返 null）。</summary>
        private string TryReadCellFormula(string address)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(address)) return null;
                var (sheetName, rangeAddress) = SplitSheetAndRange(address);
                dynamic workbook = excelApplication.ActiveWorkbook;
                if (workbook == null) return null;
                dynamic sheet = sheetName == null ? (object)excelApplication.ActiveSheet : workbook.Worksheets[sheetName];
                dynamic cell = sheet.Range[rangeAddress].Cells[1, 1];
                var raw = Convert.ToString(cell.Formula);
                if (string.IsNullOrWhiteSpace(raw) || !raw.StartsWith("=", StringComparison.Ordinal)) return null;
                return raw;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // ---------- D2 公式历史 diff：落盘与读取（RootFolder\formula-history.json，≤50 条） ----------

        private sealed class FormulaHistoryEntry
        {
            public string Time { get; set; }
            public string Target { get; set; }
            public string OldFormula { get; set; }
            public string NewFormula { get; set; }
            public string Change { get; set; }
        }

        private static readonly object FormulaHistoryLock = new object();

        private static string FormulaHistoryPath =>
            System.IO.Path.Combine(RootFolder, "formula-history.json");

        private static void AppendFormulaHistory(FormulaHistoryEntry entry)
        {
            lock (FormulaHistoryLock)
            {
                var entries = LoadFormulaHistoryUnsafe();
                entries.Insert(0, entry);
                if (entries.Count > 50) entries.RemoveRange(50, entries.Count - 50);
                System.IO.File.WriteAllText(
                    FormulaHistoryPath,
                    JsonConvert.SerializeObject(entries, Formatting.Indented),
                    new UTF8Encoding(false));
            }
        }

        private static List<FormulaHistoryEntry> LoadFormulaHistoryUnsafe()
        {
            try
            {
                if (!System.IO.File.Exists(FormulaHistoryPath)) return new List<FormulaHistoryEntry>();
                return JsonConvert.DeserializeObject<List<FormulaHistoryEntry>>(
                    System.IO.File.ReadAllText(FormulaHistoryPath)) ?? new List<FormulaHistoryEntry>();
            }
            catch
            {
                return new List<FormulaHistoryEntry>(); // 文件损坏 → 从空开始，不拖死主流程
            }
        }

        /// <summary>D2 公式历史 diff：前端拉取最近公式变更（≤50 条）。</summary>
        public Task<string> GetFormulaHistory()
        {
            return ExecuteAsync(() =>
            {
                List<FormulaHistoryEntry> entries;
                lock (FormulaHistoryLock)
                {
                    entries = LoadFormulaHistoryUnsafe();
                }
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    entries = entries.ConvertAll(e => new
                    {
                        time = e.Time,
                        target = e.Target,
                        old_formula = e.OldFormula,
                        new_formula = e.NewFormula,
                        change = e.Change
                    })
                }));
            });
        }

        /// <summary>
        /// 公式助手：用 Excel Evaluate 实算目标首格的公式结果作为「结果预览」，不改动工作表。
        /// 相对引用按首格所在位置解析（模型生成时已按目标首格写好具体引用，Evaluate 即得该格应显示的值）。
        /// </summary>
        public Task<string> PreviewFormula(string formula, string targetAreas)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(formula))
                {
                    throw new ArgumentException("没有可预览的公式。请先生成公式。", nameof(formula));
                }

                var formulaText = formula.Trim();
                if (!formulaText.StartsWith("="))
                {
                    formulaText = "=" + formulaText;
                }

                var areaList = NormalizeAreas(targetAreas);
                if (areaList.Count == 0)
                {
                    areaList.Add(CaptureCurrentSelectionAddress());
                }

                var first = areaList[0];
                var (sheetName, rangeAddress) = SplitSheetAndRange(first);
                dynamic workbook = excelApplication.ActiveWorkbook;
                dynamic worksheet = string.IsNullOrEmpty(sheetName)
                    ? (object)excelApplication.ActiveSheet
                    : workbook.Sheets[sheetName];
                dynamic topLeft = worksheet.Range[rangeAddress].Cells[1, 1];

                string cellAddress;
                try { cellAddress = Convert.ToString(topLeft.Address); }
                catch { cellAddress = first; }

                // 安全护栏：与 eval_formula 工具同口径——CALL/REGISTER 系禁止试算（审查者视角：两通道同一防线）
                var blocked = FormulaBuilder.FindBlockedFunction(formulaText);
                if (blocked != null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = false,
                        error = "预览被拦截：公式包含被禁函数 " + blocked + "（可执行外部代码，禁止试算）。"
                    }));
                }

                object value;
                string error = null;
                try
                {
                    // Evaluate 在此工作表上下文内按具体引用实算，返回该格公式应显示的值，不做任何写入。
                    value = worksheet.Evaluate(formulaText);
                }
                catch (Exception exception)
                {
DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs:223  {exception.GetType().Name}: {exception.Message}");
                    value = null;
                    error = exception.Message;
                }

                if (error != null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = false,
                        error = $"预览失败：{error}"
                    }));
                }

                if (value == null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        value = ""
                    }));
                }

                if (value is string or int or long or double or float or bool or DateTime)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        value = FormatPreviewValue(value),
                        cell = cellAddress
                    }));
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    value = Convert.ToString(value),
                    cell = cellAddress
                }));
            });
        }

        /// <summary>把 Evaluate 返回的原始值格式化为可供显示/preview 的字符串（时间、数值、文本）。</summary>
        private static string FormatPreviewValue(object raw)
        {
            string FormatNumeric(double d)
            {
                return Math.Abs(d - Math.Round(d)) < 1e-9 ? d.ToString("G", System.Globalization.CultureInfo.InvariantCulture) : d.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
            }

            switch (raw)
            {
                case null:
                    return string.Empty;
                case bool flag:
                    return flag ? "TRUE" : "FALSE";
                case DateTime time:
                    return time.ToString("yyyy-MM-dd HH:mm:ss");
                case double numeric:
                    return FormatNumeric(numeric);
                case float single:
                    return FormatNumeric((double)single);
                case int integer:
                    return integer.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case long integer64:
                    return integer64.ToString(System.Globalization.CultureInfo.InvariantCulture);
                default:
                    var asString = Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture);
                    return string.IsNullOrEmpty(asString) ? string.Empty : asString.Trim();
            }
        }

        private static List<string> NormalizeAreas(string raw)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return list;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var part in Regex.Split(raw, "[\r\n,，;；]+"))
            {
                var trimmed = part.Trim();
                if (trimmed.Length == 0 || !seen.Add(trimmed))
                {
                    continue;
                }

                list.Add(trimmed);
            }

            return list;
        }

        private static (string Formula, string Explanation) SplitGeneratedFormula(string raw)
        {
            var text = (raw ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(text))
            {
                return (string.Empty, string.Empty);
            }

            const string noteLabel = "\u8bf4\u660e\uff1a";
            const string formulaLabel = "\u516c\u5f0f\uff1a";
            var noteIndex = text.IndexOf(noteLabel, StringComparison.Ordinal);
            var formulaIndex = text.LastIndexOf(formulaLabel, StringComparison.Ordinal);
            var explanation = string.Empty;
            var formula = string.Empty;
            if (formulaIndex >= 0)
            {
                formula = CleanFormulaPart(text.Substring(formulaIndex + formulaLabel.Length));
            }

            if (noteIndex >= 0)
            {
                var end = formulaIndex >= 0 ? formulaIndex : text.Length;
                explanation = text.Substring(noteIndex + noteLabel.Length, Math.Max(0, end - noteIndex - noteLabel.Length)).Trim().Trim('\r', '\n');
            }

            if (formulaIndex < 0)
            {
                formula = CleanGeneratedFormula(text);
            }

            return (formula, explanation);
        }

        private static string CleanFormulaPart(string candidate)
        {
            candidate = (candidate ?? string.Empty).Trim();
            candidate = Regex.Replace(candidate, "(?i)^```[A-Za-z]*", string.Empty).Trim();
            candidate = Regex.Replace(candidate, "```$", string.Empty).Trim();
            foreach (var line in Regex.Split(candidate, "\r?\n"))
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("="))
                {
                    return trimmed;
                }
            }

            return candidate;
        }

        private static string CleanGeneratedFormula(string raw)
        {
            var text = (raw ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            text = Regex.Replace(text, "(?im)^```[A-Za-z]*$", string.Empty).Trim();
            foreach (var line in Regex.Split(text, "\r?\n"))
            {
                var candidate = line.Trim();
                if (candidate.StartsWith("="))
                {
                    return candidate;
                }
            }

            var index = text.IndexOf('=');
            return index >= 0 ? text.Substring(index).Trim() : text;
        }

        private string CaptureCurrentSelectionAddress()
        {
            try
            {
                dynamic selection = excelApplication.Selection;
                return Convert.ToString(selection?.Address ?? "A1") ?? "A1";
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs:398  silent-catch covered (no exception variable)");
                return "A1";
            }
        }

        /// <summary>A1 构造器优先链路：读数据选区表头行列名（≤12 列），失败返回 null。</summary>
        private string[] TryReadHeaderNames(string address)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(address)) return null;
                var (sheetName, rangeAddress) = SplitSheetAndRange(address);
                dynamic workbook = excelApplication.ActiveWorkbook;
                if (workbook == null) return null;
                dynamic sheet = sheetName == null ? (object)excelApplication.ActiveSheet : workbook.Worksheets[sheetName];
                dynamic range = sheet.Range[rangeAddress];
                var cols = Math.Min(Convert.ToInt32(range.Columns.Count), 12);
                var headers = new string[cols];
                for (var c = 0; c < cols; c++)
                {
                    headers[c] = Convert.ToString(range.Cells[1, c + 1].Value2) ?? string.Empty;
                }
                if (range != null && Marshal.IsComObject(range)) Marshal.FinalReleaseComObject(range);
                return headers;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>A1 构造器优先链路：读数据选区列宽，失败返回 0。</summary>
        private int TryReadSourceWidth(string address)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(address)) return 0;
                var (sheetName, rangeAddress) = SplitSheetAndRange(address);
                dynamic workbook = excelApplication.ActiveWorkbook;
                if (workbook == null) return 0;
                dynamic sheet = sheetName == null ? (object)excelApplication.ActiveSheet : workbook.Worksheets[sheetName];
                dynamic range = sheet.Range[rangeAddress];
                var width = Convert.ToInt32(range.Columns.Count);
                if (range != null && Marshal.IsComObject(range)) Marshal.FinalReleaseComObject(range);
                return width;
            }
            catch (Exception)
            {
                return 0;
            }
        }


        private string TryReadRangePreview(string address)
        {
            try
            {
                var (sheetName, rangeAddress) = SplitSheetAndRange(address);
                dynamic workbook = excelApplication.ActiveWorkbook;
                if (workbook == null)
                {
                    return null;
                }

                dynamic sheet = sheetName == null ? workbook.ActiveSheet : workbook.Worksheets[sheetName];
                dynamic range = sheet.Range(rangeAddress);
                var value = range.Value2;
                var lines = new List<string>();
                const int maxRows = 8;
                const int maxCols = 6;
                if (value is object[,] matrix)
                {
                    var rows = Math.Min(matrix.GetLength(0), maxRows);
                    var cols = Math.Min(matrix.GetLength(1), maxCols);
                    for (var row = 0; row < rows; row++)
                    {
                        var cells = new List<string>();
                        for (var column = 0; column < cols; column++)
                        {
                            cells.Add(Convert.ToString(matrix[row, column]) ?? string.Empty);
                        }

                        lines.Add("\u7b2c" + (row + 1) + "\u884c: " + string.Join(" | ", cells));
                    }
                }
                else if (value != null)
                {
                    lines.Add("\u533a\u57df\u4ec5\u4e00\u683c: " + Convert.ToString(value));
                }

                if (range != null && Marshal.IsComObject(range))
                {
                    Marshal.FinalReleaseComObject(range);
                }

                return lines.Count == 0 ? null : string.Join("\n", lines);
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs:448  silent-catch covered (no exception variable)");
                return null;
            }
        }

        private object ComputeSendScope()
        {
            var contextJson = JsonConvert.SerializeObject(GetCurrentContextObject());
            var contextTokens = TokenEstimator.Estimate(contextJson);
            var components = new List<object>
            {
                new { name = "current_context", label = "当前选区与上下文", tokens = contextTokens }
            };
            var referenceTokens = 0;
            foreach (var reference in references)
            {
                var tokens = TokenEstimator.EstimateJson(ToReferenceObject(reference));
                referenceTokens += tokens;
                components.Add(new
                {
                    name = "reference_" + reference.Id,
                    label = "引用 " + reference.Address,
                    tokens,
                    rows = reference.RowCount,
                    columns = reference.ColumnCount
                });
            }

            var conversationTokens = TokenEstimator.Estimate(conversation);
            var scopeText = contextJson + " " + string.Join(" ", references.Select(ToReferenceObject));
            var detections = SensitivityScanner.Scan(scopeText);
            var total = contextTokens + referenceTokens + conversationTokens + TokenEstimator.Estimate("系统指令");
            return new
            {
                ok = true,
                total_tokens = total,
                context_tokens = contextTokens,
                reference_tokens = referenceTokens,
                conversation_tokens = conversationTokens,
                components = components.Cast<object>().ToArray(),
                sensitive = detections.Select(detection => new
                {
                    detection.Kind,
                    detection.Label,
                    detection.Category,
                    detection.Count,
                    detection.Sample
                }).Cast<object>().ToArray(),
                has_sensitive = detections.Count > 0
            };
        }

        /// <summary>
        /// 公式助手：插入后自动巡检——对目标选区做错误值/悬空引用等浅层体检（复用 ExcelFormulaService.DiagnoseFormula）。
        /// 返回 { ok, total_issues, summary }，前端据此 toast 提示。
        /// </summary>
        /// <summary>
        /// 公式助手：读数据选区表头列名（C6 选区显示列名）。返回 { ok, headers: [...] }。
        /// </summary>
        public Task<string> GetRangeHeaders(string address)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(address))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "缺少地址。" }));
                }
                var headers = TryReadHeaderNames(address.Trim());
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    headers = headers ?? new string[0]
                }));
            });
        }


        public Task<string> DiagnoseFormulaRange(string address)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(address))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "缺少诊断地址。" }));
                }
                var (sheetName, rangeAddress) = SplitSheetAndRange(address.Trim());
                var service = new ExcelFormulaService(excelApplication);
                var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["address"] = rangeAddress
                };
                if (!string.IsNullOrEmpty(sheetName))
                {
                    args["sheet"] = sheetName;
                }
                var result = service.DiagnoseFormula(args);
                if (!result.Ok)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = Convert.ToString(result.Data) ?? result.ErrorCode }));
                }
                var data = result.Data as Newtonsoft.Json.Linq.JObject
                    ?? (result.Data != null ? Newtonsoft.Json.Linq.JObject.FromObject(result.Data) : new Newtonsoft.Json.Linq.JObject());
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    total_issues = data.Value<int?>("total_issues") ?? 0,
                    summary = data["summary"]
                }));
            });
        }

        /// <summary>
        /// D1 多公式批量生成：需求每行一条（≤8 条），逐条走「构造器优先 → 模型兜底」同链路。
        /// 返回 { ok, items: [{ requirement, formula, explanation, builder, ok, error }] }。
        /// 单条失败不阻断其余（AI 使用方视角：批量场景下一条模型抖动不应废掉整批）。
        /// </summary>
        public Task<string> GenerateFormulasBatch(string dataAreas, string targetAreas, string requirements)
        {
            return ExecuteAsync(async () =>
            {
                var lines = NormalizeRequirementLines(requirements);
                if (lines.Count == 0)
                {
                    throw new ArgumentException("请先在需求框里写需求（每行一条，最多 8 条）。", nameof(requirements));
                }
                if (lines.Count > 8)
                {
                    throw new ArgumentException("批量生成一次最多 8 条，当前 " + lines.Count + " 条。请分批。", nameof(requirements));
                }

                var dataAreaList = NormalizeAreas(dataAreas);
                var targetAreaList = NormalizeAreas(targetAreas);
                if (targetAreaList.Count == 0)
                {
                    targetAreaList.Add(CaptureCurrentSelectionAddress());
                }

                var headers = TryReadHeaderNames(dataAreaList.FirstOrDefault());
                var sourceWidth = TryReadSourceWidth(dataAreaList.FirstOrDefault());
                var items = new List<object>();
                foreach (var line in lines)
                {
                    items.Add(await GenerateSingleFormula(line, dataAreaList, targetAreaList, headers, sourceWidth).ConfigureAwait(false));
                }

                return JsonConvert.SerializeObject(new { ok = true, items });
            });
        }

        /// <summary>D1：需求行归一化（去空行/去重，保序）。</summary>
        private static List<string> NormalizeRequirementLines(string raw)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(raw)) return list;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in Regex.Split(raw, "\r?\n"))
            {
                var trimmed = line.Trim();
                if (trimmed.Length == 0 || trimmed.StartsWith("#") || !seen.Add(trimmed))
                {
                    continue; // 空行与 # 注释行跳过
                }
                list.Add(trimmed);
            }
            return list;
        }

        /// <summary>
        /// D1：单条公式生成（构造器优先 → 模型兜底），与 GenerateFormula 同链路但独立容错——
        /// 单条抛异常转为 { ok=false, error } 条目，不阻断批量。
        /// </summary>
        private async Task<object> GenerateSingleFormula(
            string description, List<string> dataAreaList, List<string> targetAreaList,
            string[] headers, int sourceWidth)
        {
            // 第一优先：构造器确定性生成（零模型调用）
            try
            {
                var intent = FormulaBuilder.ParseIntent(description, headers, sourceWidth);
                if (intent != null)
                {
                    var builtFormula = FormulaBuilder.Build(new FormulaBuilder.BuildOptions
                    {
                        Template = intent.Template,
                        SourceAddress = dataAreaList.Count > 0 ? dataAreaList[0] : targetAreaList[0],
                        ValueColumn = intent.ValueColumn,
                        CriteriaColumn = intent.CriteriaColumn,
                        Criteria = intent.Criteria,
                        WeightColumn = intent.WeightColumn,
                        HeaderRows = intent.HeaderRows,
                        RankOrder = intent.RankOrder
                    });
                    DiagnosticsLog.Write("formula-batch builder hit: " + intent.Template + " <- " + description);
                    return new
                    {
                        requirement = description,
                        formula = builtFormula,
                        explanation = "（模板构造）命中「" + (intent.MatchedKeyword ?? intent.Template) + "」模板，由公式构造器确定性生成。",
                        builder = true,
                        ok = true,
                        error = (string)null
                    };
                }
            }
            catch (ArgumentException builderEx)
            {
                DiagnosticsLog.Write("formula-batch builder miss: " + builderEx.Message);
                // 落模型兜底
            }
            catch (Exception builderEx2)
            {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs batch-builder  " + builderEx2.GetType().Name + ": " + builderEx2.Message);
                // 落模型兜底
            }

            // 第二优先：模型兜底（与 GenerateFormula 同 prompt 契约）
            try
            {
                var profile = RequireProfile();
                using var client = new OpenAiCompatibleClient(profile);
                using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(90));

                var prompt = new StringBuilder();
                if (dataAreaList.Count > 0)
                {
                    prompt.AppendLine("数据选区（可参考的源数据，只读，不会改动）：");
                    foreach (var area in dataAreaList.Take(6))
                    {
                        prompt.Append("- ").AppendLine(area);
                        var preview = TryReadRangePreview(area);
                        if (preview != null)
                        {
                            prompt.AppendLine("  预览：" + preview);
                        }
                    }
                }
                else
                {
                    prompt.AppendLine("数据选区：（未指定）");
                }

                prompt.AppendLine("目标选区（公式要填入的位置，请按首格用相对/绝对引用配合向下或向右填充）：");
                foreach (var area in targetAreaList)
                {
                    prompt.Append("- ").AppendLine(area);
                }

                prompt.AppendLine("需求：" + description);
                prompt.AppendLine();
                prompt.AppendLine("请生成一条 Excel 公式。你的输出只能包含「说明」与「公式」两部分，不要多余文字、不要代码围栏：");
                prompt.AppendLine("说明：请简洁、有结构，严格按下面 4 行固定格式输出，每行只有一句话，让新手一眼看懂——");
                prompt.AppendLine("调用函数：<这条公式调用哪些函数，如 SUM、IF、VLOOKUP、COUNTIF 等>");
                prompt.AppendLine("括号内：<每个参数各代表什么、分别在算什么，多项用“；”分隔>");
                prompt.AppendLine("括号外：<= 与函数名本身分别表示什么>");
                prompt.AppendLine("执行顺序：<有嵌套时按“内层先算 → 外层再拿结果算”写出；无嵌套则写“直接依次计算”>");
                prompt.AppendLine("公式：（以 = 开头的单条公式）");

                var input = new List<ChatMessage>
                {
                    ChatMessage.System(
                        "你是 Excel 公式专家。根据用户给的数据选区与目标选区、需求，生成一条正确、能填充的公式，并用简洁、有结构的方式解释，避免一大段文字。\n输出格式绝对按照：\n1. 「说明：」后严格换行输出 4 行，每行一句话，行首标签固定为：调用函数：、括号内：、括号外：、执行顺序：。其中括号外要讲清开头的 = 与函数名本身的含义；执行顺序若有嵌套要按“内层先算 → 外层再拿结果算”写出。\n2. 最后以「公式：」开头的单条 Excel 公式。\n不要任何其他内容、不要 markdown 代码围栏。"),
                    ChatMessage.User(prompt.ToString())
                };

                var result = await client.CompleteAsync(input, timeoutSource.Token).ConfigureAwait(false);
                var (formula, explanation) = SplitGeneratedFormula(result?.Content ?? string.Empty);
                if (string.IsNullOrWhiteSpace(formula))
                {
                    return new
                    {
                        requirement = description,
                        formula = (string)null,
                        explanation = (string)null,
                        builder = false,
                        ok = false,
                        error = "模型未能生成有效公式，请调整描述后重试。"
                    };
                }

                return new
                {
                    requirement = description,
                    formula,
                    explanation,
                    builder = false,
                    ok = true,
                    error = (string)null
                };
            }
            catch (Exception modelEx)
            {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.Formula.cs batch-model  " + modelEx.GetType().Name + ": " + modelEx.Message);
                return new
                {
                    requirement = description,
                    formula = (string)null,
                    explanation = (string)null,
                    builder = false,
                    ok = false,
                    error = modelEx.Message
                };
            }
        }

    }
}
