using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 只读/导航级工具：工作表列表、区域预览、读写值与公式、当前选区、选中区域、查找单元格、设置/清除筛选。
    /// </summary>
    public sealed class ExcelReadService
    {
        private readonly dynamic application;

        public ExcelReadService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult GetSheetList(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheets = new List<object>();
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    dynamic used = null;
                    try
                    {
                        used = worksheet.UsedRange;
                        // XlSheetVisibility：-1=可见、0=隐藏、2=深度隐藏（M16 修复：
                        // 原 (Visible != 0) 判定把 VeryHidden 误标为可见）。
                        var rawVisibility = Convert.ToInt32(worksheet.Visible);
                        var visible = rawVisibility == -1;
                        sheets.Add(new
                        {
                            name = Convert.ToString(worksheet.Name),
                            index = Convert.ToInt32(worksheet.Index),
                            visible,
                            hidden_state = rawVisibility == -1 ? "visible" : rawVisibility == 2 ? "very_hidden" : "hidden",
                            used_range = Convert.ToString(used?.Address ?? "None"),
                            used_rows = Convert.ToInt32(used?.Rows.Count ?? 0),
                            used_columns = Convert.ToInt32(used?.Columns.Count ?? 0)
                        });
                    }
                    finally
                    {
                        Release((object)used);
                        // 枚举变量 RCW 释放（M16 全量审查：原实现泄漏，每表一个）。
                        Release((object)worksheet);
                    }
                }

                return new ToolResult { Ok = true, Data = new { sheet_count = sheets.Count, sheets } };
            }
            finally
            {
                Release((object)workbook);
            }
        }

        public ToolResult PreviewSheet(Dictionary<string, object> args)
        {
            var previewRows = GetInt(args, "rows", 8);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic used = null;
            dynamic values = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                used = sheet.UsedRange;
                var totalRows = Convert.ToInt32(used.Rows.Count);
                var totalColumns = Convert.ToInt32(used.Columns.Count);
                if (totalRows == 0 || totalColumns == 0)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { sheet = Convert.ToString(sheet.Name), total_rows = 0, total_columns = 0, empty = true }
                    };
                }

                var topCount = Math.Min(previewRows, totalRows);
                var top = ReadValues(used, 1, 1, topCount, totalColumns);
                object[,] bottom = null;
                if (totalRows > topCount)
                {
                    bottom = ReadValues(used, totalRows - Math.Min(previewRows, totalRows - topCount) + 1, 1, Math.Min(previewRows, totalRows - topCount), totalColumns);
                }

                object[,] sample;
                if (totalRows * totalColumns > 50000)
                {
                    // 超大表：仅用首尾抽样推断列类型，避免整块读取海量数组。
                    sample = MergeSample(top, bottom, totalColumns);
                }
                else
                {
                    values = used.Value2;
                    sample = FlattenValues(values, totalRows, totalColumns);
                }

                var columnTypes = InferColumnTypes(sample, totalColumns);

                return new ToolResult
                {
                    Ok = true,
                    Truncated = totalRows > topCount,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(used.Address),
                        total_rows = totalRows,
                        total_columns = totalColumns,
                        top_rows = top,
                        bottom_rows = bottom,
                        column_types = columnTypes
                    }
                };
            }
            finally
            {
                Release((object)values);
                Release((object)used);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult ReadRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            // 上限兜底：用户传超大 max_cells 不能无限放大单次读取；clamp 到 [1, 200000]
            var rawMax = GetInt(args, "max_cells", 2000);
            var maxCells = rawMax < 1 ? 1 : (rawMax > 200000 ? 200000 : rawMax);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic values = null;
            dynamic formula = null;
            dynamic numberFormats = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                long cellCount = rows * (long)columns;
                if (cellCount > maxCells)
                {
                    // 大区域：分段抽样读取首尾行，避免一次性海量数组导致超时/占用。
                    var sampleRows = Math.Min(3, rows);
                    var top = ReadValues(range, 1, 1, sampleRows, columns);
                    var bottomCount = Math.Min(2, rows - sampleRows);
                    var bottom = bottomCount > 0
                        ? ReadValues(range, rows - bottomCount + 1, 1, bottomCount, columns)
                        : null;
                    return new ToolResult
                    {
                        Ok = true,
                        Truncated = true,
                        Data = new
                        {
                            sheet = Convert.ToString(sheet.Name),
                            address = Convert.ToString(range.Address),
                            rows,
                            columns,
                            cell_count = cellCount,
                            required_count = maxCells,
                            summary = true,
                            top_rows = top,
                            bottom_rows = bottom,
                            number_format_sample = SampleNumberFormat(range)
                        }
                    };
                }

                values = range.Value2;
                formula = range.Formula;
                numberFormats = range.NumberFormat;
                var valueData = FlattenValues(values, rows, columns);
                var formulaData = FlattenValues(formula, rows, columns);
                var formatData = FlattenValues(numberFormats, rows, columns);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(range.Address),
                        rows,
                        columns,
                        cell_count = cellCount,
                        values = valueData,
                        formulas = formulaData,
                        number_formats = formatData
                    }
                };
            }
            finally
            {
                Release((object)numberFormats);
                Release((object)formula);
                Release((object)values);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult ReadSelection(Dictionary<string, object> args)
        {
            var rawMax = GetInt(args, "max_cells", 2000);
            var maxCells = rawMax < 1 ? 1 : (rawMax > 200000 ? 200000 : rawMax);
            dynamic selection = null;
            dynamic values = null;
            dynamic formula = null;
            dynamic numberFormats = null;
            try
            {
                selection = application.Selection;
                if (selection == null)
                {
                    return Invalid("NO_SELECTION", "当前没有活动选区。");
                }

                var rows = Convert.ToInt32(selection.Rows.Count);
                var columns = Convert.ToInt32(selection.Columns.Count);
                long cellCount = rows * (long)columns;

                if (cellCount > maxCells)
                {
                    // 大区域选区：分段抽样读取首尾行，避免一次性海量数组。
                    var sampleRows = Math.Min(3, rows);
                    var top = ReadValues(selection, 1, 1, sampleRows, columns);
                    var bottomCount = Math.Min(2, rows - sampleRows);
                    var bottom = bottomCount > 0
                        ? ReadValues(selection, rows - bottomCount + 1, 1, bottomCount, columns)
                        : null;
                    return new ToolResult
                    {
                        Ok = true,
                        Truncated = true,
                        Data = new
                        {
                            sheet = Convert.ToString(application.ActiveSheet?.Name ?? "?"),
                            address = Convert.ToString(selection.Address),
                            rows,
                            columns,
                            cell_count = cellCount,
                            summary = true,
                            top_rows = top,
                            bottom_rows = bottom,
                            values = (object[,])null,
                            formulas = (object[,])null,
                            number_formats = (object[,])null
                        }
                    };
                }

                values = selection.Value2;
                formula = selection.Formula;
                numberFormats = selection.NumberFormat;

                var valueData = FlattenValues(values, rows, columns);
                var formulaData = FlattenValues(formula, rows, columns);
                var formatData = FlattenValues(numberFormats, rows, columns);

                return new ToolResult
                {
                    Ok = true,
                    Truncated = false,
                    Data = new
                    {
                        sheet = Convert.ToString(application.ActiveSheet?.Name ?? "?"),
                        address = Convert.ToString(selection.Address),
                        rows,
                        columns,
                        cell_count = cellCount,
                        summary = false,
                        top_rows = (object[,])null,
                        bottom_rows = (object[,])null,
                        values = valueData,
                        formulas = formulaData,
                        number_formats = formatData
                    }
                };
            }
            finally
            {
                Release((object)numberFormats);
                Release((object)formula);
                Release((object)values);
                Release((object)selection);
            }
        }

        public ToolResult SelectRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var activate = GetBool(args, "activate", true);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = GetSheet(workbook, Optional(args, "sheet"));
                if (activate)
                {
                    workbook.Activate();
                    sheet.Activate();
                }

                range = sheet.Range[address];
                range.Select();
                range.Activate();
                return new ToolResult
                {
                    Ok = true,
                    Data = new { sheet = Convert.ToString(sheet.Name), selected = Convert.ToString(range.Address) }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult FindCells(Dictionary<string, object> args)
        {
            var what = Required(args, "what");
            var mode = (Optional(args, "mode") ?? "value").ToLowerInvariant();
            if (mode != "value" && mode != "formula")
            {
                return Invalid("INVALID_MODE", "mode 仅支持 value 或 formula。");
            }

            var matchCase = GetBool(args, "match_case", false);
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic used = null;
            dynamic range = null;
            dynamic data = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var address = Optional(args, "address");
                range = string.IsNullOrWhiteSpace(address)
                    ? sheet.UsedRange
                    : sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                var originRow = Convert.ToInt32(range.Row);
                var originColumn = Convert.ToInt32(range.Column);
                // 全量扫描必须整体拉取才不漏匹配；超大区域（如不带 address 的整张 UsedRange）禁止无界读取，
                // 让用户缩小范围而不是静默超时/爆内存。
                if (rows * (long)columns > 200000)
                {
                    return Invalid("RANGE_TOO_LARGE",
                        $"扫描区域过大（{rows} 行 × {columns} 列）。请提供 address 缩小查找范围后再试。");
                }

                data = mode == "formula" ? range.Formula : range.Value2;
                var matrix = FlattenValues(data, rows, columns);

                var matches = new List<string>();
                for (var row = 0; row < rows; row++)
                {
                    for (var column = 0; column < columns; column++)
                    {
                        var cell = matrix[row, column];
                        if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                        {
                            continue;
                        }

                        if (Convert.ToString(cell).IndexOf(what, comparison) >= 0)
                        {
                            matches.Add(ExcelRangeHelper.CellAddress(originRow + row, originColumn + column));
                        }
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Truncated = matches.Count > 200,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        scanned_range = Convert.ToString(range.Address),
                        mode,
                        count = matches.Count,
                        matches = matches.Take(200).ToArray()
                    }
                };
            }
            finally
            {
                Release((object)data);
                Release((object)range);
                Release((object)used);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult ManageFilter(Dictionary<string, object> args)
        {
            var action = Required(args, "action");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic values = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                if (string.Equals(action, "clear", StringComparison.OrdinalIgnoreCase))
                {
                    range = sheet.UsedRange;
                    range.EntireRow.Hidden = false;
                    try
                    {
                        sheet.AutoFilterMode = false;
                    }
                    catch (Exception)
                    {
                        // 某些环境没有开启自动筛选时置空会抛错，忽略。
                    }

                    return new ToolResult { Ok = true, Data = new { cleared = true } };
                }

                var address = Required(args, "address");
                var column = GetInt(args, "column", 1);
                var @operator = Required(args, "operator");
                var value = Optional(args, "value");
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                if (column < 1 || column > columns)
                {
                    return Invalid("INVALID_COLUMN", $"列号必须介于 1 到 {columns}。");
                }

                values = range.Value2;
                var matrix = FlattenValues(values, rows, columns);
                var window = range.Rows;
                try
                {
                    for (var row = 0; row < rows; row++)
                    {
                        if (row == 0)
                        {
                            continue;
                        }

                        var keep = DataTransformService.Matches(matrix[row, column - 1], @operator, value);
                        dynamic rowRange = null;
                        try
                        {
                            rowRange = window.Item[row + 1];
                            rowRange.EntireRow.Hidden = !keep;
                        }
                        finally
                        {
                            Release((object)rowRange);
                        }
                    }
                }
                finally
                {
                    Release((object)window);
                }

                var visible = 0;
                for (var row = 0; row < rows; row++)
                {
                    if (row == 0)
                    {
                        continue;
                    }

                    if (DataTransformService.Matches(matrix[row, column - 1], @operator, value))
                    {
                        visible++;
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{Convert.ToString(sheet.Name)}!{Convert.ToString(range.Address)}" },
                    Data = new { filtered = true, matched_rows = visible, total_rows = Math.Max(rows - 1, 0) }
                };
            }
            finally
            {
                Release((object)values);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public static string[] InferColumnTypes(object[,] sample, int columnCount)
        {
            var result = new string[columnCount];
            var rows = sample.GetLength(0);
            for (var column = 0; column < columnCount; column++)
            {
                var hasValue = false;
                var allNumbers = true;
                var allDates = true;
                var allText = true;
                var someEmpty = false;
                // 首行视为表头，不参与类型推断。
                for (var row = 1; row < rows; row++)
                {
                    var cell = sample[row, column];
                    if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                    {
                        someEmpty = true;
                        continue;
                    }

                    hasValue = true;
                    var text = Convert.ToString(cell).Trim();
                    var isNumber = double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
                    var isDate = DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    if (!isNumber)
                    {
                        allNumbers = false;
                    }

                    if (!isDate)
                    {
                        allDates = false;
                    }

                    allText = allText && !isNumber && !isDate;
                }

                if (!hasValue)
                {
                    result[column] = "blank";
                }
                else if (allNumbers && !someEmpty)
                {
                    result[column] = "number";
                }
                else if (allDates && allNumbers)
                {
                    result[column] = DateTime.TryParse(Convert.ToString(sample[0, column]), CultureInfo.InvariantCulture, DateTimeStyles.None, out _)
                        ? "date"
                        : "number";
                }
                else if ((allNumbers || allDates) && !allText && !someEmpty)
                {
                    result[column] = ((allNumbers && allDates) ? "number" : (allNumbers ? "number" : "date"));
                }
                else if (hasValue && !(allNumbers || allDates) && !allText)
                {
                    result[column] = "mixed";
                }
                else
                {
                    result[column] = "text";
                }
            }

            return result;
        }

        private object[,] ReadValues(dynamic sourceRange, int startRow, int startColumn, int rowCount, int columnCount)
        {
            if (rowCount <= 0 || columnCount <= 0)
            {
                return new object[0, 0];
            }

            dynamic first = null;
            dynamic last = null;
            dynamic value = null;
            try
            {
                first = sourceRange.Cells[startRow, startColumn];
                last = sourceRange.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
                value = sourceRange.Range[first, last].Value2;
                return FlattenValues(value, rowCount, columnCount);
            }
            finally
            {
                Release((object)value);
                Release((object)last);
                Release((object)first);
            }
        }

        private static object[,] Slice(object[,] matrix, int startRow, int count)
        {
            var result = new object[count, matrix.GetLength(1)];
            for (var row = 0; row < count && startRow + row < matrix.GetLength(0); row++)
            {
                for (var column = 0; column < matrix.GetLength(1); column++)
                {
                    result[row, column] = matrix[startRow + row, column];
                }
            }

            return result;
        }

        private static object[,] MergeSample(object[,] top, object[,] bottom, int columns)
        {
            var topRows = top?.GetLength(0) ?? 0;
            var bottomRows = bottom?.GetLength(0) ?? 0;
            var result = new object[topRows + bottomRows, columns];
            for (var row = 0; row < topRows; row++)
            {
                for (var column = 0; column < columns; column++)
                {
                    result[row, column] = top[row, column];
                }
            }

            for (var row = 0; row < bottomRows; row++)
            {
                for (var column = 0; column < columns; column++)
                {
                    result[topRows + row, column] = bottom[row, column];
                }
            }

            return result;
        }

        private static string SampleNumberFormat(dynamic range)
        {
            dynamic first = null;
            try
            {
                first = range.Cells[1, 1];
                return Convert.ToString(first.NumberFormat);
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                Release((object)first);
            }
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = GetSheet(workbook, sheetName);
        }

        private dynamic GetSheet(dynamic workbook, string sheetName)
        {
            return string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source == null || rows == 0 || columns == 0)
            {
                return result;
            }

            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }

                return result;
            }

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToBoolean(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        }
}