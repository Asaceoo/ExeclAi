using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 公式构造与验证闭环（服务层薄壳）：
    /// build_formula（纯逻辑） / eval_formula（Application.Evaluate 试算，写前验证） /
    /// diagnose_formula（区域公式静态体检五类根因）/ rewrite_formulas（≤500 格批改：绝对相对/IFERROR）。
    /// </summary>
    public sealed class ExcelFormulaService
    {
        private const int MaxRewriteCells = 500;
        private const int MaxDiagnoseCells = 200;

        private readonly dynamic application;

        public ExcelFormulaService(object application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        // ---------- build_formula（纯逻辑，无 COM） ----------

        public ToolResult BuildFormula(Dictionary<string, object> args)
        {
            try
            {
                var o = new FormulaBuilder.BuildOptions
                {
                    Template = GetString(args, "template"),
                    SourceAddress = GetString(args, "source_address"),
                    ValueColumn = GetInt(args, "value_column", 1),
                    SheetName = GetOptional(args, "sheet_name"),
                    HeaderRows = GetInt(args, "header_rows", 1),
                    CriteriaColumn = args.TryGetValue("criteria_column", out var cc) && cc != null ? Convert.ToInt32(cc) : (int?)null,
                    Criteria = GetOptional(args, "criteria"),
                    WeightColumn = args.TryGetValue("weight_column", out var wc) && wc != null ? Convert.ToInt32(wc) : (int?)null,
                };
                var formula = FormulaBuilder.Build(o);
                AppLog.Info("formula", "build_formula template=" + o.Template + " -> " + formula);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        formula,
                        template = o.Template,
                        source_address = o.SourceAddress,
                        header_rows = o.HeaderRows,
                        basis = "公式由 FormulaBuilder 结构化生成：表头行已钳制（数据区从第 " + (o.HeaderRows + 1) + " 行起）、同表裸地址/跨表规范限定。写入前建议用 eval_formula 试算验证。"
                    }
                };
            }
            catch (ArgumentException aex)
            {
                return Invalid("INVALID_ARGS", aex.Message);
            }
        }

        // ---------- eval_formula（Application.Evaluate，写前验证） ----------

        public ToolResult EvalFormula(Dictionary<string, object> args)
        {
            var formula = GetString(args, "formula");
            if (string.IsNullOrWhiteSpace(formula))
            {
                return Invalid("INVALID_ARGS", "缺少 formula。");
            }
            var expr = formula.TrimStart().TrimStart('=');
            // 安全护栏：CALL/REGISTER 系可执行外部代码——两通道（工具/UI 预览）共用 FormulaBuilder 黑名单
            var blocked = FormulaBuilder.FindBlockedFunction(expr);
            if (blocked != null)
            {
                return Invalid("BLOCKED_FUNCTION", "公式包含被禁函数 " + blocked + "（可执行外部代码，禁止试算）。");
            }
            try
            {
                var result = application.Evaluate(expr);
                var described = DescribeValue(result);
                var displayText = described.Item1; var kind = described.Item2; var errorCode = described.Item3;
                AppLog.Info("formula", "eval_formula -> " + kind + " " + displayText);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        formula = "=" + expr,
                        kind,
                        value = displayText,
                        error_code = errorCode,
                        note = kind == "error"
                            ? "公式当前会产生该错误值——写入前必须修正（常见：引用名不存在/函数名拼写/区域形态）。"
                            : (expr.ToUpperInvariant().Contains("NOW()") || expr.ToUpperInvariant().Contains("TODAY()") || expr.ToUpperInvariant().Contains("RAND()") ? "试算通过。注意：公式含易变函数（NOW/TODAY/RAND），每次重算结果会变化。" : "试算通过。真实写入后结果可能随引用区域数据变化。")
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("EVAL_ERROR", "试算失败：" + ex.Message + "（常见：函数在本宿主版本不支持/引用不存在）。");
            }
        }

        private static Tuple<string, string, string> DescribeValue(object v)
        {
            if (v == null) return Tuple.Create("(空)", "empty", (string)null);
            if (v is int err && err >= 2000 && err <= 2042)
            {
                return Tuple.Create(ErrName(err), "error", ErrName(err));
            }
            if (v is double d) return Tuple.Create(d.ToString("R", CultureInfo.InvariantCulture), "number", (string)null);
            if (v is bool b) return Tuple.Create(b ? "TRUE" : "FALSE", "boolean", (string)null);
            if (v is string str) return Tuple.Create(str, "text", (string)null);
            if (v is object[,] arr)
            {
                var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
                var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
                var sample = arr.GetValue(r1, c1);
                return Tuple.Create("数组 " + (r2 - r1 + 1) + "×" + (c2 - c1 + 1) + "，首格=" + (sample != null ? sample.ToString() : "空"), "array", (string)null);
            }
            return Tuple.Create(v.ToString(), "other", (string)null);
        }

        private static string ErrName(int code)
        {
            switch (code)
            {
                case 2000: return "#NULL!";
                case 2007: return "#DIV/0!";
                case 2015: return "#VALUE!";
                case 2023: return "#REF!";
                case 2029: return "#NAME?";
                case 2036: return "#NUM!";
                case 2042: return "#N/A";
                default: return "#" + code;
            }
        }

        // ---------- diagnose_formula（区域公式静态体检） ----------

        public ToolResult DiagnoseFormula(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null) return Invalid("NO_WORKBOOK", "当前没有打开的工作簿。");
                var sheetName = GetOptional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (rows * cols > MaxDiagnoseCells)
                {
                    return Invalid("TOO_MANY_CELLS", "诊断区域过大（" + rows + "×" + cols + "），上限 " + MaxDiagnoseCells + " 格。请缩小范围或分段。");
                }

                // 工作表名集合（判悬空引用）
                var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var ws in workbook.Worksheets)
                {
                    sheetNames.Add(Convert.ToString(ws.Name));
                }

                var items = new List<object>();
                var counts = new Dictionary<string, int>
                {
                    { "ok", 0 }, { "has_error_value", 0 }, { "dangling_sheet_ref", 0 },
                    { "external_ref", 0 }, { "circular_self", 0 }, { "refcode_in_text", 0 }, { "not_formula", 0 }
                };

                foreach (var cell in range.Cells)
                {
                    var addr = Convert.ToString(cell.Address, CultureInfo.InvariantCulture);
                    var formulaText = Convert.ToString(cell.Formula);
                    if (string.IsNullOrWhiteSpace(formulaText) || !formulaText.StartsWith("=", StringComparison.Ordinal))
                    {
                        counts["not_formula"]++;
                        continue;
                    }
                    var issues = new List<string>();
                    // 错误值检测用 Value 类型（Int32 2000~2042 = XlCVErr），不用 Text——
                    // 列宽不足时 Text 显示 "####" 同样以 # 开头，会造成假阳性（审查者视角 R 修正）。
                    var rawValue = cell.Value;
                    if (rawValue is int errCode && errCode >= 2000 && errCode <= 2042)
                    {
                        issues.Add("HAS_ERROR_VALUE:" + ErrName(errCode));
                        counts["has_error_value"]++;
                    }
                    if (formulaText.IndexOf("#REF!", StringComparison.Ordinal) >= 0)
                    {
                        issues.Add("REFCODE_IN_TEXT:#REF!");
                        counts["refcode_in_text"]++;
                    }
                    if (Regex.IsMatch(formulaText, @"\[\d+\]"))
                    {
                        issues.Add("EXTERNAL_REF:公式引用了外部工作簿（[n]形态）");
                        counts["external_ref"]++;
                    }
                    // 自引用（本格地址出现在自身公式内）
                    var ownLocal = addr.Contains("$") ? addr.Replace("$", "") : addr;
                    var ownShort = addr.Split('!').Last();
                    if (formulaText.IndexOf(ownShort, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        issues.Add("CIRCULAR_SELF:公式引用了自身");
                        counts["circular_self"]++;
                    }
                    // 悬空 sheet 引用
                    foreach (Match m in Regex.Matches(formulaText, "(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.\\u4e00-\\u9fa5]*))!"))
                    {
                        var name = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                        if (string.IsNullOrEmpty(name) || !sheetNames.Contains(name))
                        {
                            issues.Add("DANGLING_SHEET_REF:工作表「" + name + "」不存在（引用悬空）");
                            counts["dangling_sheet_ref"]++;
                        }
                    }

                    if (issues.Count == 0)
                    {
                        counts["ok"]++;
                        items.Add(new { cell = addr, verdict = "ok", issues = Array.Empty<string>(), formula = Trunc(formulaText) });
                    }
                    else
                    {
                        items.Add(new { cell = addr, verdict = "issues", issues = issues.Take(5).ToArray(), formula = Trunc(formulaText) });
                    }
                }

                var totalIssues = counts["has_error_value"] + counts["dangling_sheet_ref"] + counts["external_ref"] + counts["circular_self"] + counts["refcode_in_text"];
                var advice = totalIssues == 0
                    ? "区域公式未发现结构性问题。"
                    : "建议：#REF! 与悬空表引用删除或改指存在的表；错误值先用 eval_formula 试算修正；自引用改引用其他单元格。";
                AppLog.Info("formula", "diagnose_formula " + address + " issues=" + totalIssues);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(range.Address),
                        scanned = rows * cols,
                        summary = counts,
                        total_issues = totalIssues,
                        items,
                        advice
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException cex)
            {
                return Invalid("EXCEL_COM_ERROR", "诊断失败（HRESULT=0x" + cex.HResult.ToString("X8") + "）：" + cex.Message);
            }
            finally
            {
                Release(range); Release(sheet); Release(workbook);
            }
        }

        // ---------- rewrite_formulas（区域批改，≤500 格） ----------

        public ToolResult RewriteFormulas(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            var op = (GetString(args, "operation") ?? string.Empty).Trim().ToLowerInvariant();
            if (op != "to_absolute" && op != "to_relative" && op != "wrap_iferror")
            {
                return Invalid("INVALID_ARGS", "operation 仅支持 to_absolute / to_relative / wrap_iferror。");
            }
            var fallback = GetOptional(args, "iferror_fallback") ?? "0";

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null) return Invalid("NO_WORKBOOK", "当前没有打开的工作簿。");
                var sheetName = GetOptional(args, "sheet");
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                range = sheet.Range[address];

                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (rows * cols > MaxRewriteCells)
                {
                    return Invalid("TOO_MANY_CELLS", "批改区域过大（" + rows + "×" + cols + "），上限 " + MaxRewriteCells + " 格。请缩小范围或分批。");
                }

                var changed = 0;
                var skipped = 0;
                var samples = new List<string>();
                foreach (var cell in range.Cells)
                {
                    var f = Convert.ToString(cell.Formula);
                    if (string.IsNullOrWhiteSpace(f) || !f.StartsWith("=", StringComparison.Ordinal))
                    {
                        skipped++;
                        continue;
                    }
                    string nf;
                    if (op == "to_absolute") nf = "=" + FormulaBuilder.ToAbsolute(f.Substring(1));
                    else if (op == "to_relative") nf = "=" + FormulaBuilder.ToRelative(f.Substring(1));
                    else nf = FormulaBuilder.WrapIfError(f, fallback);
                    if (nf != f)
                    {
                        cell.Formula = nf;
                        changed++;
                        if (samples.Count < 5) samples.Add(Convert.ToString(cell.Address) + " → " + nf);
                    }
                }

                AppLog.Warn("formula", "rewrite_formulas op=" + op + " changed=" + changed + " skipped=" + skipped + " range=" + Convert.ToString(range.Address));
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = new List<string> { Convert.ToString(sheet.Name) + "!" + Convert.ToString(range.Address) },
                    Data = new
                    {
                        operation = op,
                        changed,
                        skipped_non_formula = skipped,
                        samples,
                        note = "批改完成（≤500 格护栏内，直写不走撤销快照）。回滚口径：用 diagnose_formula 复核 + 以 samples 中改写前公式手工还原，或改用 eval_formula 先行试算避免误改。"
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException cex)
            {
                return Invalid("EXCEL_COM_ERROR", "批改失败（HRESULT=0x" + cex.HResult.ToString("X8") + "）：" + cex.Message);
            }
            finally
            {
                Release(range); Release(sheet); Release(workbook);
            }
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value); } catch { }
            }
        }

        private static string Trunc(string text)
        {
            return text != null && text.Length > 200 ? text.Substring(0, 200) + "…" : text;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                return Convert.ToString(v);
            }
            return string.Empty;
        }

        private static string GetOptional(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                var t = Convert.ToString(v);
                return string.IsNullOrWhiteSpace(t) ? null : t;
            }
            return null;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                if (int.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out var n)) return n;
            }
            return fallback;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}
