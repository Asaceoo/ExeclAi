using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第二期单元格/结构增强工具：合并拆分、单元格类型重置、公式批量转换、引用切换、公式引用追踪、下钻、可见行复制。
    /// </summary>
    public sealed class ExcelAdvancedEditService
    {
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        // 逐 cell/逐行 COM 迭代的资源护栏；防止模型传整列/整表规模导致百万次 COM 调用失控。
        private const int MaxResetCells = 100000;
        private const int MaxCopyCells = 100000;

        public ExcelAdvancedEditService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
        }

        public ToolResult ManageMergeSplit(Dictionary<string, object> args)
        {
            var action = Required(args, "action");
            var address = Required(args, "address");
            var direction = Optional(args, "direction");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic value = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                switch (action)
                {
                    case "merge":
                        range.Merge();
                        break;
                    case "unmerge":
                        range.UnMerge();
                        break;
                    case "fill_merged":
                        FillMerged(sheet, range);
                        break;
                    case "merge_similar":
                        value = range.Value2;
                        var matrix = FlattenValues(value, Convert.ToInt32(range.Rows.Count), Convert.ToInt32(range.Columns.Count));
                        if (string.Equals(direction, "row", StringComparison.OrdinalIgnoreCase))
                        {
                            MergeSimilarRowWise(sheet, range, matrix);
                        }
                        else
                        {
                            MergeSimilarColumnWise(sheet, range, matrix);
                        }

                        break;
                    default:
                        return Invalid("BAD_ACTION", $"不支持的操作：{action}。");
                }

                return new ToolResult { Ok = true, Data = new { action, direction = direction, merged = true } };
            }
            finally
            {
                Release((object)value);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private void MergeSimilarColumnWise(dynamic sheet, dynamic range, object[,] matrix)
        {
            var rows = matrix.GetLength(0);
            var columns = matrix.GetLength(1);
            for (var column = 0; column < columns; column++)
            {
                var runStart = 0;
                for (var row = 1; row <= rows; row++)
                {
                    var changed = row == rows ||
                                  !object.Equals(NormalizeComparable(matrix[row, column]), NormalizeComparable(matrix[row - 1, column]));
                    if (!changed)
                    {
                        continue;
                    }

                    if (row - runStart > 1)
                    {
                        MergeVertical(sheet, range, runStart + 1, row, column + 1);
                    }

                    runStart = row;
                }
            }
        }

        private void MergeVertical(dynamic sheet, dynamic range, int topRow, int bottomRow, int excelColumn)
        {
            var columnName = ExcelRangeHelper.ColumnName(excelColumn);
            dynamic run = null;
            try
            {
                run = sheet.Range[columnName + topRow + ":" + columnName + bottomRow];
                run.Merge();
            }
            finally
            {
                Release((object)run);
            }
        }

        private void MergeSimilarRowWise(dynamic sheet, dynamic range, object[,] matrix)
        {
            var rows = matrix.GetLength(0);
            var columns = matrix.GetLength(1);
            for (var row = 0; row < rows; row++)
            {
                var runStart = -1;
                object runValue = null;
                for (var column = 0; column <= columns; column++)
                {
                    var current = column < columns ? matrix[row, column] : null;
                    if (column < columns && object.Equals(current, runValue) && runStart >= 0)
                    {
                        continue;
                    }

                    if (runStart >= 0 && column - runStart > 1)
                    {
                        MergeHorizontal(sheet, range, row + 1, runStart + 1, column);
                    }

                    runStart = column < columns ? column : -1;
                    runValue = current;
                }
            }
        }

        private void MergeHorizontal(dynamic sheet, dynamic range, int excelRow, int leftColumn, int rightColumn)
        {
            dynamic run = null;
            try
            {
                run = sheet.Range[ExcelRangeHelper.CellAddress(excelRow, leftColumn) + ":" + ExcelRangeHelper.CellAddress(excelRow, rightColumn)];
                run.Merge();
            }
            finally
            {
                Release((object)run);
            }
        }

        private void FillMerged(dynamic sheet, dynamic range)
        {
            var merges = range.MergeCells;
            if (merges == null || Convert.ToBoolean(merges) == false)
            {
                return;
            }

            dynamic areas = null;
            try
            {
                areas = range.MergeArea;
                if (areas == null)
                {
                    return;
                }

                var topLeft = areas.Cells[1, 1];
                var anchor = topLeft.Value2;
                dynamic fillRange = null;
                try
                {
                    fillRange = areas.MergeCells ? areas.UnMerge() : null;
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:FillMerged 静默兜底触发: " + xaiEx.Message);
                    // 无合并区域时忽略。
                }

                if (fillRange == null)
                {
                    // UnMerge() 不返回区域，改为逐单元格填充。
                    areas.Value2 = matrixOf(areas.Rows.Count, areas.Columns.Count, anchor);
                }

                Release((object)topLeft);
            }
            finally
            {
                Release((object)areas);
            }
        }

        private static object[,] matrixOf(int rows, int columns, object value)
        {
            var result = new object[rows, columns];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < columns; c++)
                {
                    result[r, c] = value;
                }
            }

            return result;
        }

        private static object NormalizeComparable(object value)
        {
            if (value == null)
            {
                return null;
            }

            return value is Newtonsoft.Json.Linq.JValue j ? j.Value : value;
        }

        public ToolResult ResetValueType(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var targetType = Required(args, "target_type").ToLowerInvariant();
            if (targetType != "text" && targetType != "number" && targetType != "date")
            {
                return Invalid("INVALID_TYPE", $"target_type 仅支持 text/number/date，当前为 {targetType}。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic value = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                if (rows * (long)columns > MaxResetCells)
                {
                    return Invalid("TOO_MANY_CELLS", $"区域过大（{rows} 行 × {columns} 列），请缩小范围或分批处理。");
                }

                value = range.Value2;
                var matrix = FlattenValues(value, rows, columns);

                var converted = new object[rows, columns];
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < columns; c++)
                    {
                        converted[r, c] = ConvertCell(matrix[r, c], targetType);
                    }
                }

                ApplyCellWise(range, matrix, converted, targetType);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(range.Address),
                        target_type = targetType,
                        cells = rows * columns
                    }
                };
            }
            finally
            {
                Release((object)value);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private static object ConvertCell(object value, string targetType)
        {
            var text = value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            switch (targetType)
            {
                case "text":
                    return string.IsNullOrWhiteSpace(text) ? null : text;
                case "number":
                    if (value is double || value is int || value is long || value is decimal)
                    {
                        return value;
                    }

                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        return number;
                    }

                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out number))
                    {
                        return number;
                    }

                    return null;
                case "date":
                    if (value is DateTime)
                    {
                        return value;
                    }

                    if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) ||
                        DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                    {
                        return date;
                    }

                    return null;
                default:
                    return value;
            }
        }

        private void ApplyCellWise(dynamic range, object[,] original, object[,] converted, string targetType)
        {
            var rows = converted.GetLength(0);
            var columns = converted.GetLength(1);
            if (targetType == "number" || targetType == "date")
            {
                // 先写值，再统一设置数字格式。
                range.Value2 = converted;
                range.NumberFormat = targetType == "date" ? "yyyy/m/d" : "General";
            }
            else
            {
                // 文本类型：先设文本格式避免自动转换，再逐格写入字符串。
                range.NumberFormat = "@";
                var top = range.Cells[1, 1];
                var last = range.Cells[rows, columns];
                dynamic target = null;
                dynamic cell = null;
                try
                {
                    target = range.Worksheet.Range[top, last];
                    for (var r = 0; r < rows; r++)
                    {
                        for (var c = 0; c < columns; c++)
                        {
                            cell = target.Cells[r + 1, c + 1];
                            cell.Value2 = converted[r, c] as string;
                            Release((object)cell);
                            cell = null;
                        }
                    }
                }
                finally
                {
                    Release((object)cell);
                    Release((object)target);
                    Release((object)top);
                    Release((object)last);
                }
            }
        }

        public ToolResult TransformFormulas(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            // actions 经 JSON 反序列化为 JArray/List<object> 等形态，`as IEnumerable<object>` 对 JArray 为 null，
            // 且元素是 JObject 而非 Dictionary<string,object>。按非泛型 IEnumerable + JObject(ToDict) 多形态读取。
            var actions = new List<Dictionary<string, object>>();
            if (args.TryGetValue("actions", out var raw) && raw != null)
            {
                try
                {
                    if (raw is System.Collections.IEnumerable en && !(raw is string))
                    {
                        foreach (var item in en)
                        {
                            var dict = ToDict(item);
                            if (dict != null)
                            {
                                actions.Add(dict);
                            }
                        }
                    }
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:TransformFormulas 静默兜底触发: " + xaiEx.Message);
                }
            }

            if (actions.Count == 0)
            {
                return Invalid("MISSING_ACTIONS", "缺少必填参数：actions。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic formula = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                formula = range.Formula;
                var original = FlattenValues(formula, rows, columns);
                var transformed = new object[rows, columns];
                var changed = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < columns; c++)
                    {
                        var cell = CellText(original[r, c]);
                        if (!cell.StartsWith("=", StringComparison.Ordinal))
                        {
                            transformed[r, c] = original[r, c];
                            continue;
                        }

                        var next = cell;
                        foreach (var item in actions)
                        {
                            next = ApplyAction(next, item);
                        }

                        transformed[r, c] = next;
                        if (!string.Equals(next, cell, StringComparison.Ordinal))
                        {
                            changed++;
                        }
                    }
                }

                var sheetName = Convert.ToString(sheet.Name);
                var rangeAddress = Convert.ToString(range.Address);
                WriteFormula(sheet, range, transformed, rows, columns);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = sheetName,
                        address = rangeAddress,
                        changed_cells = changed,
                        total_cells = rows * columns
                    }
                };
            }
            finally
            {
                Release((object)formula);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private static string ApplyAction(string formula, Dictionary<string, object> action)
        {
            var type = Convert.ToString(action.TryGetValue("type", out var t) ? t : string.Empty);
            var inner = formula.Substring(1); // 去掉 = 前缀
            switch (type)
            {
                case "iferror":
                    var fallback = ExcelCells(action, "fallback", "\"\"");
                    return "=IFERROR(" + inner + ", " + fallback + ")";
                case "ifna":
                    var naFallback = ExcelCells(action, "fallback", "\"\"");
                    return "=IFNA(" + inner + ", " + naFallback + ")";
                case "round":
                    var digits = Convert.ToInt32(action.TryGetValue("digits", out var d) ? d : 2);
                    return "=ROUND(" + inner + ", " + digits + ")";
                case "custom":
                    var func = Convert.ToString(action.TryGetValue("function", out var f) ? f : string.Empty);
                    if (string.IsNullOrWhiteSpace(func))
                    {
                        return formula;
                    }

                    var tail = ExcelCells(action, "args", string.Empty);
                    return "=" + func + "(" + inner + (string.IsNullOrEmpty(tail) ? string.Empty : ", " + tail) + ")";
                default:
                    return formula;
            }
        }

        private static Dictionary<string, object> ToDict(object item)
        {
            if (item is Dictionary<string, object> d)
            {
                return d;
            }

            if (item is Newtonsoft.Json.Linq.JObject jo)
            {
                var result = new Dictionary<string, object>();
                foreach (var property in jo.Properties())
                {
                    result[property.Name] = UnwrapToken(property.Value);
                }

                return result;
            }

            return null;
        }

        private static object UnwrapToken(Newtonsoft.Json.Linq.JToken token)
        {
            if (token is Newtonsoft.Json.Linq.JValue value)
            {
                return value.Value;
            }

            if (token is Newtonsoft.Json.Linq.JObject obj)
            {
                return ToDict(obj);
            }

            if (token is Newtonsoft.Json.Linq.JArray array)
            {
                var list = new List<object>();
                foreach (var item in array)
                {
                    list.Add(UnwrapToken(item));
                }

                return list;
            }

            return token?.ToString();
        }

        private static string ExcelCells(Dictionary<string, object> map, string key, string defaultValue)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            var text = Convert.ToString(value);
            if (text.Length > 0 && char.IsDigit(text[0]))
            {
                return text;
            }

            // 非纯数字按字面量转义，防止公式注入用数字函数。
            if (text.All(ch => char.IsLetterOrDigit(ch) || ch == '_'))
            {
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }

            return text;
        }

        public ToolResult ConvertFormulaReferences(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var mode = Optional(args, "mode") ?? "to_absolute";
            if (!string.Equals(mode, "to_absolute", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(mode, "to_relative", StringComparison.OrdinalIgnoreCase))
            {
                // fail-closed（M16 全量审查）：schema 有 enum 约束，代码侧同样复核，
                // 防止拼写错误被静默按"相对引用"处理。
                return Invalid("INVALID_ARGS", $"mode 仅支持 to_absolute / to_relative，收到：{mode}");
            }
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic formula = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                formula = range.Formula;
                var original = FlattenValues(formula, rows, columns);
                var converted = new object[rows, columns];
                var changed = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < columns; c++)
                    {
                        var cell = CellText(original[r, c]);
                        if (!cell.StartsWith("=", StringComparison.Ordinal))
                        {
                            converted[r, c] = original[r, c];
                            continue;
                        }

                        var next = ConvertRefsInFormula(cell, mode);
                        converted[r, c] = next;
                        if (!string.Equals(next, cell, StringComparison.Ordinal))
                        {
                            changed++;
                        }
                    }
                }

                var sheetName = Convert.ToString(sheet.Name);
                var rangeAddress = Convert.ToString(range.Address);
                WriteFormula(sheet, range, converted, rows, columns);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = sheetName,
                        address = rangeAddress,
                        mode = mode,
                        changed_cells = changed
                    }
                };
            }
            finally
            {
                Release((object)formula);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        internal static string ConvertRefsInFormula(string formula, string mode)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return formula;
            }

            // 先标记引号内文本（"…" 文本字面量、'…' 工作表名），这些区域里的 "abc123" 绝不是单元格引用，
            // 绝不能改写；否则 to_absolute 会把 ="abc123" 变成 =$abc$123 直接损坏公式。
            var quoted = new bool[formula.Length];
            var inDouble = false;
            var inSingle = false;
            for (var i = 0; i < formula.Length; i++)
            {
                var c = formula[i];
                if (inDouble)
                {
                    quoted[i] = true;
                    if (c == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            quoted[i + 1] = true;
                            i++;
                        }
                        else
                        {
                            inDouble = false;
                        }
                    }

                    continue;
                }

                if (inSingle)
                {
                    quoted[i] = true;
                    if (c == '\'')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '\'')
                        {
                            quoted[i + 1] = true;
                            i++;
                        }
                        else
                        {
                            inSingle = false;
                        }
                    }

                    continue;
                }

                if (c == '"')
                {
                    inDouble = true;
                    quoted[i] = true;
                }
                else if (c == '\'')
                {
                    inSingle = true;
                    quoted[i] = true;
                }
            }

            // 只匹配 A1/$A$1/a1/$a$1 形式的引用；边界断言挡住紧随词元（命名范围/长列名）的误配；
            // 尾随 `(` 的函数名（LOG10、BIN2DEC、HEX2DEC 等含数字函数名）不得改写，否则成 $LOG$10 损坏公式。
            var pattern = new Regex("(?<![\\w$])\\$?([A-Za-z]{1,3})\\$?(\\d+)(?![\\w(])");
            var output = new StringBuilder();
            var last = 0;
            foreach (Match m in pattern.Matches(formula))
            {
                if (m.Index < last)
                {
                    continue;
                }

                var inQuote = false;
                for (var k = m.Index; k < m.Index + m.Length; k++)
                {
                    if (quoted[k])
                    {
                        inQuote = true;
                        break;
                    }
                }

                if (inQuote)
                {
                    output.Append(formula, last, m.Index + m.Length - last);
                    last = m.Index + m.Length;
                    continue;
                }

                output.Append(formula, last, m.Index - last);
                var column = m.Groups[1].Value;
                var row = m.Groups[2].Value;
                output.Append(mode == "to_absolute"
                    ? "$" + column + "$" + row
                    : column + row);
                last = m.Index + m.Length;
            }

            output.Append(formula, last, formula.Length - last);
            return output.ToString();
        }

        public ToolResult TraceFormulaReferences(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic precedents = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                string formula = null;
                try
                {
                    formula = CellText(range.Formula);
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:TraceFormulaReferences 静默兜底触发: " + xaiEx.Message);
                    formula = null;
                }

                if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("=", StringComparison.Ordinal))
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { sheet = Convert.ToString(sheet.Name), address = Convert.ToString(range.Address), has_formula = false, references = new object[0] }
                    };
                }

                var refs = new List<object>();
                precedents = range.Precedents;
                if (precedents != null)
                {
                    foreach (dynamic area in EnumerateAreas(precedents))
                    {
                        var refSheetName = Convert.ToString(area.Parent.Name);
                        var refAddress = Convert.ToString(area.Address);
                        var refValue = SafeReadValue(area);
                        refs.Add(new { sheet = refSheetName, address = refAddress, value = refValue });
                        Release((object)area);
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(range.Address),
                        has_formula = true,
                        formula = formula,
                        reference_count = refs.Count,
                        references = refs
                    }
                };
            }
            finally
            {
                Release((object)precedents);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private static IEnumerable<object> EnumerateAreas(dynamic range)
        {
            var list = new List<object>();
            dynamic areas = null;
            try
            {
                areas = range.Areas;
                var count = Convert.ToInt32(areas.Count);
                for (var i = 1; i <= count; i++)
                {
                    list.Add(areas.Item[i]);
                }
            }
            finally
            {
                // Areas 集合 RCW 及时释放（M16 全量审查：原实现泄漏，每次 trace 一个）。
                ReleaseCom(areas);
            }

            return list;
        }

        /// <summary>Range/集合类自有 RCW 的安全释放（失败不阻断收尾）。</summary>
        private static void ReleaseCom(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(comObject))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception xaiEx) {
                AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:ReleaseCom 静默兜底触发: " + xaiEx.Message);
            }
        }

        public ToolResult DrillExcelRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var trace = TraceFormulaReferences(new Dictionary<string, object>(args, StringComparer.OrdinalIgnoreCase));
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        display = CellText(SafeReadValue(range)),
                        formula = CellText(range.Formula),
                        trace = trace.Ok ? trace.Data : null
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

        public ToolResult CopyVisibleRows(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            dynamic workbook = null;
            dynamic sourceSheet = null;
            dynamic sourceRange = null;
            dynamic rows2d = null;
            try
            {
                Resolve(ref workbook, ref sourceSheet, Optional(args, "source_sheet"));
                sourceRange = sourceSheet.Range[sourceAddress];
                var rows = Convert.ToInt32(sourceRange.Rows.Count);
                var columns = Convert.ToInt32(sourceRange.Columns.Count);
                if (rows * (long)columns > MaxCopyCells)
                {
                    return Invalid("TOO_MANY_CELLS", $"源区域过大（{rows} 行 × {columns} 列），请缩小范围或分批处理。");
                }

                var all = FlattenValues(sourceRange.Value2, rows, columns);
                var visibleRows = new List<object[]>();
                for (var r = 0; r < rows; r++)
                {
                    if (IsRowHidden(sourceRange, r))
                    {
                        continue;
                    }

                    var rowBuffer = new object[columns];
                    for (var c = 0; c < columns; c++)
                    {
                        rowBuffer[c] = all[r, c];
                    }

                    visibleRows.Add(rowBuffer);
                }

                var output = new object[visibleRows.Count, columns];
                for (var i = 0; i < visibleRows.Count; i++)
                {
                    for (var c = 0; c < columns; c++)
                    {
                        output[i, c] = visibleRows[i][c];
                    }
                }

                if (writer != null)
                {
                    var targetSheet = Optional(args, "target_sheet");
                    var targetAddress = Optional(args, "target_address");
                    if (targetSheet == null && targetAddress == null)
                    {
                        // fail-closed（M16 全量审查 P0）：缺省写回源区域会把可见行覆盖到源区顶部、
                        // 隐藏/被筛选行数据被静默破坏。必须显式指定目标。
                        return Invalid("INVALID_ARGS", "必须提供 target_address 或 target_sheet：缺省写回源区域会用可见行覆盖隐藏行数据。");
                    }

                    var writeArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["sheet"] = targetSheet ?? Convert.ToString(sourceSheet.Name),
                        ["address"] = targetAddress ?? Convert.ToString(sourceRange.Address)
                    };
                    var writeResult = writer(writeArgs, output, "auto", "copy_visible_rows_to_sheet");
                    if (!writeResult.Ok)
                    {
                        return writeResult;
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        source = Convert.ToString(sourceRange.Address),
                        source_rows = rows,
                        copied_rows = visibleRows.Count,
                        columns = columns
                    }
                };
            }
            finally
            {
                Release((object)rows2d);
                Release((object)sourceRange);
                Release((object)sourceSheet);
                Release((object)workbook);
            }
        }

        private static bool IsRowHidden(dynamic sourceRange, int offset)
        {
            dynamic row = null;
            try
            {
                row = sourceRange.Rows.Item[offset + 1];
                return Convert.ToBoolean(row.Hidden);
            }
            catch (Exception xaiEx) {
                AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:IsRowHidden 静默兜底触发: " + xaiEx.Message);
                return false;
            }
            finally
            {
                Release((object)row);
            }
        }

        private void WriteFormula(dynamic sheet, dynamic range, object[,] formulas, int rows, int columns)
        {
            if (writer != null)
            {
                var writeArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = Convert.ToString(sheet.Name),
                    ["address"] = Convert.ToString(range.Address)
                };
                var result = writer(writeArgs, formulas, "formula", "transform_formula");
                if (!result.Ok)
                {
                    throw new InvalidOperationException(CellText(result.Data));
                }

                return;
            }

            range.Formula = formulas;
        }

        private static object SafeReadValue(dynamic source)
        {
            try
            {
                return source.Value2;
            }
            catch (Exception xaiEx) {
                AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:SafeReadValue 静默兜底触发: " + xaiEx.Message);
                return null;
            }
        }

        private static string CellText(object value)
        {
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
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

        private static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        // CVErr 错误单元格归一为错误文本，防止错误值被改写成数字（M16 全量审查）。
                        result[row, column] = ExcelErrorValues.Normalize(array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)]);
                    }
                }

                return result;
            }

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = ExcelErrorValues.Normalize(source);
            }

            return result;
        }

        private static void Release(object value)
        {
            if (value == null)
            {
                return;
            }

            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(value))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
                }
            }
            catch (Exception xaiEx) {
                AppLog.Warn("adv-edit", "ExcelAdvancedEditService.cs:Release 静默兜底触发: " + xaiEx.Message);
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}