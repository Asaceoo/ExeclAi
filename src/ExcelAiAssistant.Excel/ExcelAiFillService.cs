using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 数据智能填充（P6-C'）：两段式。
    ///   ai_fill_blank_cells（只读）：读选区（含表头+数据）→ 定位空白格 → 分组并附样本，供模型语义推理建议值，不写入。
    ///   apply_fill_blank_cells（写入·受控）：把模型确认的 {地址→值} 就地回填空白格，非空格保持不变；整块写回经 writer 记录撤销。
    /// 仅空白格视为待填；占位文本非空格一律不动。写前受控执行总开关校验。
    /// </summary>
    public sealed class ExcelAiFillService
    {
        private const int DefaultMaxBlanks = 5000;
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        private readonly Func<bool> isControlledEnabled;

        public ExcelAiFillService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null,
            Func<bool> isControlledEnabled = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
            this.isControlledEnabled = isControlledEnabled ?? (() => true);
        }

        // ===================== 分析（只读，L0） =====================
        public ToolResult AnalyzeBlanks(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var hasHeader = PBool(args, "has_header", true);
            // D4：默认 same_sheet，仅用户显式选 whole_workbook 才整簿引用
            var referenceScope = Optional(args, "reference_scope") ?? "same_sheet";
            if (!IsValidReferenceScope(referenceScope))
            {
                return new ToolResult { Ok = false, ErrorCode = "INVALID_SCOPE", Data = $"reference_scope 仅支持 none/same_sheet/whole_workbook，当前为 {referenceScope}。" };
            }
            var maxBlanks = PInt(args, "max_blanks", DefaultMaxBlanks);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                (workbook, sheet) = ResolveSource(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                // dynamic 传染：显式还原静态类型，避免下方 GroupByColumn 动态分派
                object[,] block = ReadMatrix(range, rows, cols);

                var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader);
                var truncated = blanks.Count > maxBlanks;
                var page = truncated ? blanks.GetRange(0, maxBlanks) : blanks;
                var groups = BlankFillAnalyzer.GroupByColumn(block, page, hasHeader, sampleLimit: 5);
                var rowContexts = BlankFillAnalyzer.BuildRowContext(block, page, hasHeader);

                var samplePayload = new List<object>();
                foreach (var kv in groups)
                {
                    var info = kv.Value;
                    var blankAddrs = new List<string>();
                    foreach (var b in info.Blanks) blankAddrs.Add(b.CellAddress);
                    samplePayload.Add(new
                    {
                        column = info.Column,
                        header = info.ColumnName,
                        blank_cells = blankAddrs,
                        sample_values = info.Samples
                    });
                }

                var rowPayload = new List<object>();
                foreach (var rc in rowContexts)
                {
                    var cells = new List<object>();
                    foreach (var c in rc.Cells)
                    {
                        cells.Add(new
                        {
                            column = c.Column,
                            header = c.ColumnName,
                            value = c.IsBlank ? null : c.Value,
                            is_blank = c.IsBlank
                        });
                    }

                    rowPayload.Add(new { row = rc.Row, cells = cells });
                }

                var references = BuildReferenceSamples(sheetName, referenceScope, sheet, workbook);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        source = $"{sheetName ?? Convert.ToString(sheet.Name)} {address}",
                        has_header = hasHeader,
                        blank_count = blanks.Count,
                        shown_count = page.Count,
                        has_more = truncated,
                        truncated_prompt = truncated ? $"选区空缺较多（共 {blanks.Count} 个），仅处理前 {maxBlanks} 个空格，其余忽略。请在本提示中原样转达用户，避免“填了没生效”的误解。" : null,
                        reference_scope = referenceScope,
                        header = hasHeader ? ReadHeaders(block) : null,
                        blanks_by_column = samplePayload,
                        rows_neighboring = rowPayload,
                        reference_samples = references,
                        fill_contract = new
                        {
                            rule = "对每个空格给出 fill 与 reason。找不到语义依据的数值空格填字符串「待补」；reason 为空也必须填「待补」；不填 0、不填均值、不覆盖非空白格。"
                        }
                    }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 写入（确认门 + 受控，落盘） =====================
        public ToolResult ApplyFills(Dictionary<string, object> args)
        {
            if (isControlledEnabled != null && !isControlledEnabled())
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "CONTROLLED_EXECUTION_OFF",
                    Data = "受控执行总开关未开启，已拒绝写入。请在设置中开启「受控执行」后再试。"
                };
            }

            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var fills = ParseFills(args);
            if (fills.Count == 0)
            {
                // fills 是必填参数：缺失/为空=调用方错误，不得当成“没有可填入的空白格”假成功
                return new ToolResult { Ok = false, ErrorCode = "INVALID_FILLS", Data = "缺少必填参数 fills：请提供至少一个 目标单元格地址 → 填充值 的映射。" };
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                (workbook, sheet) = ResolveSource(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                object[,] block = ReadMatrix(range, rows, cols);

                // 只允许回填本区域内的空白格：非空格即刻拒绝，绝不覆盖
                var applied = 0;
                var rejected = new List<string>();
                var failed = 0;
                foreach (var entry in fills)
                {
                    if (!TryParseAddress(entry.Key, out var row, out var col, out var addrNote))
                    {
                        rejected.Add(addrNote);
                        continue;
                    }

                    if (row < 1 || row > rows || col < 1 || col > cols)
                    {
                        rejected.Add(entry.Key + "（超出选区范围）");
                        continue;
                    }

                    // 原值非空白（含占位文本）→ 拒绝覆盖
                    if (!BlankFillAnalyzer.IsLikelyBlank(block[row - 1, col - 1]))
                    {
                        rejected.Add(entry.Key + "（非空白格，已跳过）");
                        continue;
                    }

                    block[row - 1, col - 1] = entry.Value ?? string.Empty;
                    applied++;
                }

                if (applied == 0)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { applied = 0, rejected = rejected, message = "没有可填入的空白格，未做任何修改。" }
                    };
                }

                // 整块写回：非空格保持原值、空白格被填充；writer 内部记录撤销快照
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var writeArgs = new Dictionary<string, object> { ["address"] = address, ["sheet"] = sheetName };
                object[,] outMatrix = NormalizeBlock(block);
                var res = writer != null
                    ? writer(writeArgs, outMatrix, "value", "apply_fill_blank_cells")
                    : WriteDirect(sheet, address, outMatrix);
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { applied = applied, rejected = rejected, failed = failed, cells = EditableAddresses(fills) }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ===================== 内部 =====================
        private (dynamic Workbook, dynamic Sheet) ResolveSource(string sheetName)
        {
            var workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
            var sheet = sheetName == null ? (dynamic)application.ActiveSheet : (dynamic)workbook.Worksheets.Item[sheetName];
            return (workbook, sheet);
        }

        private object[] ReadHeaders(object[,] block)
        {
            var cols = block.GetLength(1);
            var result = new object[cols];
            for (var c = 0; c < cols; c++)
            {
                result[c] = block[0, c] is string s ? s : block[0, c];
            }

            return result;
        }

        /// <summary>按 reference_scope 收集模型可参考的其他数据样本（有限样本，防 token 爆炸）。</summary>
        private List<object> BuildReferenceSamples(string sourceSheet, string scope, dynamic sourceSheetObj, dynamic workbook)
        {
            var result = new List<object>();
            if (string.Equals(scope, "none", StringComparison.OrdinalIgnoreCase)) return result;
            try
            {
                if (sourceSheetObj == null) return result;
                if (string.Equals(scope, "same_sheet", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(scope) || string.Equals(scope, "whole_workbook", StringComparison.OrdinalIgnoreCase))
                {
                    // 本表前几行样本（若与选区不同）
                    if (string.Equals(sourceSheetObj.Name, sourceSheet ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                    {
                        // 直接取活动表 UsedRange 首行，仅作为参考上下文
                        dynamic used = null;
                        try
                        {
                            used = sourceSheetObj.UsedRange;
                            var usedRows = Convert.ToInt32(used.Rows.Count);
                            if (usedRows > 1)
                            {
                                var m = ReadMatrix(used, 1, Convert.ToInt32(used.Columns.Count));
                                result.Add(new { source = "sheet_first_row", sample = ToList(m) });
                            }
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            Release((object)used);
                        }
                    }
                }

                // whole_workbook：列出工作表名，供模型决定是否值得跨表引用（不展开全表内容防 token 爆炸）
                if (string.Equals(scope, "whole_workbook", StringComparison.OrdinalIgnoreCase) && workbook != null)
                {
                    var names = new List<string>();
                    var count = Convert.ToInt32(workbook.Worksheets.Count);
                    for (var i = 1; i <= count; i++)
                    {
                        dynamic ws = null;
                        try
                        {
                            ws = workbook.Worksheets.Item[i];
                            names.Add(Convert.ToString(ws.Name));
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            Release((object)ws);
                        }
                    }

                    result.Add(new { source = "workbook_sheets", names = names });
                }
            }
            catch (Exception)
            {
                // 参考样本仅用于辅助，失败不影响主流程
            }

            return result;
        }

        private static bool IsValidReferenceScope(string scope)
        {
            return string.Equals(scope, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scope, "same_sheet", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scope, "whole_workbook", StringComparison.OrdinalIgnoreCase);
        }

        private static List<object> ToList(object[,] m)
        {
            var list = new List<object>();
            if (m == null) return list;
            var cols = m.GetLength(1);
            for (var c = 0; c < cols; c++)
            {
                var v = m[0, c];
                list.Add(v is string s ? s : v);
            }

            return list;
        }

        private static Dictionary<string, object> ParseFills(Dictionary<string, object> args)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (!args.TryGetValue("fills", out var raw)) return result;
            if (raw is IDictionary<string, object> dict)
            {
                foreach (var kv in dict)
                {
                    var v = kv.Value;
                    result[kv.Key] = v == null ? null : ConvertThing(v);
                }

                return result;
            }

            if (raw is Newtonsoft.Json.Linq.JObject jobj)
            {
                foreach (var kv in jobj)
                {
                    var v = kv.Value;
                    result[kv.Key] = v == null || v.Type == Newtonsoft.Json.Linq.JTokenType.Null ? null : ConvertThing(v);
                }

                return result;
            }

            if (raw is Newtonsoft.Json.Linq.JArray jarr)
            {
                foreach (var item in jarr)
                {
                    var obj = item as Newtonsoft.Json.Linq.JObject;
                    if (obj == null) continue;
                    var addr = obj["address"] ?? obj["cell"];
                    var val = obj["value"];
                    if (addr != null)
                    {
                        result[Convert.ToString(addr)] = val == null || val.Type == Newtonsoft.Json.Linq.JTokenType.Null
                            ? null
                            : ConvertThing(val);
                    }
                }
            }

            return result;
        }

        /// <summary>把建议值规整为可写回 Excel 的标量：保留数值/布尔类型，字符串原样，其余转文本。</summary>
        private static object ConvertThing(object v)
        {
            if (v is Newtonsoft.Json.Linq.JValue jv)
            {
                switch (jv.Type)
                {
                    case Newtonsoft.Json.Linq.JTokenType.String: return (string)jv;
                    case Newtonsoft.Json.Linq.JTokenType.Integer: return (long)jv;
                    case Newtonsoft.Json.Linq.JTokenType.Float: return (double)jv;
                    case Newtonsoft.Json.Linq.JTokenType.Boolean: return (bool)jv;
                    default: return jv.ToString();
                }
            }

            if (v is string s) return s;
            if (v is bool b) return b;
            if (v is double || v is float || v is decimal) return v;
            if (v is int || v is long || v is short || v is byte) return v;
            return v == null ? null : v.ToString();
        }

        private static List<string> EditableAddresses(Dictionary<string, object> fills)
        {
            var list = new List<string>();
            foreach (var k in fills.Keys) list.Add(k);
            return list;
        }

        private static bool TryParseAddress(string addr, out int row, out int col, out string note)
        {
            row = 0;
            col = 0;
            note = null;
            if (string.IsNullOrWhiteSpace(addr))
            {
                note = "（空地址）";
                return false;
            }

            var key = addr.Trim().ToUpperInvariant();
            int i = 0;
            var colPart = string.Empty;
            while (i < key.Length && char.IsLetter(key[i])) { colPart += key[i]; i++; }
            if (colPart.Length == 0 || colPart.Length > 3)
            {
                note = addr + "（无法解析）";
                return false;
            }

            var rowPart = key.Substring(i);
            if (rowPart.Length == 0 || !int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row < 1)
            {
                note = addr + "（无法解析）";
                return false;
            }

            col = ColumnLettersToNumber(colPart);
            if (col < 1)
            {
                note = addr + "（无法解析）";
                return false;
            }

            return true;
        }

        private static int ColumnLettersToNumber(string letters)
        {
            int result = 0;
            foreach (var c in letters)
            {
                if (c < 'A' || c > 'Z') return 0;
                result = result * 26 + (c - 'A' + 1);
            }

            return result;
        }

        private static object[,] NormalizeBlock(object[,] block)
        {
            var rows = block.GetLength(0);
            var cols = block.GetLength(1);
            var result = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var v = block[r, c];
                    result[r, c] = v ?? DBNull.Value;
                }
            }

            return result;
        }

        private ToolResult WriteDirect(dynamic sheet, string address, object[,] values)
        {
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            dynamic range = null;
            try
            {
                // 兜底写入：从 address 顶格写满 rows x cols
                var resized = ExcelRangeHelper.Resize(address, rows, cols);
                if (string.IsNullOrEmpty(resized)) return new ToolResult { Ok = false, ErrorCode = "INVALID_ADDRESS", Data = "地址无效：" + address };
                range = sheet.Range[resized];
                range.Value2 = values;
                return new ToolResult { Ok = true };
            }
            finally
            {
                Release((object)range);
            }
        }

        private static object[,] ReadMatrix(dynamic range, int rows, int cols)
        {
            var result = new object[rows, cols];
            object value = null;
            try
            {
                value = range.Value2;
                if (value is object[,] array)
                {
                    for (var r = 0; r < rows; r++)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            result[r, c] = array[r + array.GetLowerBound(0), c + array.GetLowerBound(1)];
                        }
                    }

                    return result;
                }

                if (rows == 1 && cols == 1)
                {
                    result[0, 0] = value;
                }

                return result;
            }
            finally
            {
                Release((object)value);
            }
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static bool PBool(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null && bool.TryParse(Convert.ToString(v), out var b))
            {
                return b;
            }

            return fallback;
        }

        private static int PInt(Dictionary<string, object> args, string key, int fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null &&
                int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
            {
                return i;
            }

            return fallback;
        }

        private static void Release(object value)
        {
            if (value == null) return;
            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(value))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}