using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// JYYJtool 借鉴工具：③ 文本语义 / 数字 / 行列表格 / 工作表 / 复制粘贴增强。
    /// 带有 undo 友好的写回调；结构级改动通过 recorder 记录。
    /// </summary>
    public sealed class ExcelJyyjtoolService
    {
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        private readonly OperationRecorder recorder;

        public ExcelJyyjtoolService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null,
            OperationRecorder recorder = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
            this.recorder = recorder;
        }

        // ======================== 文本语义 ========================

        public ToolResult ExtractTextCharacters(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var kind = Optional(args, "kind") ?? "number";
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var text = ToText(matrix, r, c);
                        matrix[r, c] = string.IsNullOrEmpty(text)
                            ? string.Empty
                            : JyyjtoolAlgorithms.ExtractCharacters(text, kind);
                    }
                }

                return WriteGrid(args, sheetName, target, matrix, rows, cols);
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult InsertTextIntoCell(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var insertion = Optional(args, "insertion") ?? string.Empty;
            var when = (Optional(args, "when") ?? "char").ToLowerInvariant();
            var where = (Optional(args, "where") ?? "before").ToLowerInvariant();
            var index = GetInt(args, "index", when == "char" ? 0 : 1);
            var separator = Optional(args, "separator") ?? string.Empty;
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var text = ToText(matrix, 0, 0);
                matrix[0, 0] = JyyjtoolAlgorithms.InsertSubstring(text, insertion, when, where, index, separator);
                return WriteGrid(args, sheetName, target, matrix, 1, 1);
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult DeleteTextMatches(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var pattern = Required(args, "pattern");
            if (pattern.Length > 256)
            {
                return Invalid("INVALID_PATTERN", "正则表达式过长（超过 256 字符），请简化后重试。");
            }

            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                var changed = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var text = ToText(matrix, r, c);
                        if (string.IsNullOrEmpty(text))
                        {
                            continue;
                        }

                        var cleaned = JyyjtoolAlgorithms.DeleteMatches(text, pattern);
                        if (cleaned != text)
                        {
                            changed++;
                        }

                        matrix[r, c] = cleaned;
                    }
                }

                var result = WriteGrid(args, sheetName, target, matrix, rows, cols);
                if (result.Ok)
                {
                    result.Data = new { changed_cells = changed };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult ConvertTextToPinyin(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var style = Optional(args, "style") ?? "full";
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        matrix[r, c] = JyyjtoolAlgorithms.ToPinyin(ToText(matrix, r, c), style);
                    }
                }

                return WriteGrid(args, sheetName, target, matrix, rows, cols);
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult ConvertChineseVariant(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var direction = Optional(args, "direction") ?? "simplified_to_traditional";
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        matrix[r, c] = JyyjtoolAlgorithms.ConvertChineseVariant(ToText(matrix, r, c), direction);
                    }
                }

                return WriteGrid(args, sheetName, target, matrix, rows, cols);
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult AmountToUppercase(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var target = Optional(args, "target_address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var output = new object[rows, 1];
                var converted = 0;
                var unparsed = 0;
                for (var r = 0; r < rows; r++)
                {
                    var text = ToText(matrix, r, 0);
                    decimal amount;
                    if (decimal.TryParse(text, NumberStyles.Number | NumberStyles.AllowCurrencySymbol, CultureInfo.CurrentCulture, out amount)
                        || decimal.TryParse(text, NumberStyles.Number | NumberStyles.AllowCurrencySymbol, CultureInfo.InvariantCulture, out amount))
                    {
                        output[r, 0] = JyyjtoolAlgorithms.NumberToUppercaseAmount(amount);
                        converted++;
                    }
                    else
                    {
                        // 非金额文本不破坏原值，原样保留并提示未转换行数。
                        output[r, 0] = text;
                        unparsed++;
                    }
                }

                if (string.IsNullOrWhiteSpace(target))
                {
                    ResolveAddress((object)sheet, address, out var originRow, out var originCol);
                    // 缺省写到源区域右侧相邻列、起始行与源对齐（M16 修复：原 +1 行偏移会写到 B2 起而非 B1）。
                    target = ExcelRangeHelper.ColumnName(originCol + 1) + originRow;
                }

                var result = WriteGrid(args, sheetName, target, output, rows, 1);
                if (result.Ok)
                {
                    result.Data = new { sheet = sheetName, target_address = target, rows, converted, unparsed };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        // ======================== 数字 ========================

        public ToolResult BankersRoundValues(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var decimals = GetInt(args, "decimals", 2);
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                var changed = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        if (matrix[r, c] == null || string.IsNullOrWhiteSpace(ToText(matrix, r, c)))
                        {
                            continue;
                        }

                        if (double.TryParse(ToText(matrix, r, c), NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                        {
                            var rounded = JyyjtoolAlgorithms.BankersRound((decimal)num, decimals);
                            if ((decimal)num != rounded)
                            {
                                changed++;
                            }

                            matrix[r, c] = rounded;
                        }
                    }
                }

                var result = WriteGrid(args, sheetName, target, matrix, rows, cols);
                if (result.Ok)
                {
                    result.Data = new { changed_cells = changed, decimals };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult PickNumbersToSum(Dictionary<string, object> args)
        {
            var candidatesAddress = Required(args, "candidates_address");
            var target = GetDouble(args, "target", 0);
            var tolerance = GetDouble(args, "tolerance", 0.01);
            var outputAddress = Optional(args, "output_address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, candidatesAddress, out var sheetName);
                var candidates = new List<double>();
                foreach (var cell in matrix)
                {
                    if (cell != null && double.TryParse(Convert.ToString(cell), NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                    {
                        candidates.Add(num);
                    }
                }

                if (candidates.Count == 0)
                {
                    return Invalid("NO_CANDIDATES", "候选区域没有可用的数字。");
                }

                var chosen = JyyjtoolAlgorithms.FindSubsetSum(candidates.ToArray(), target, tolerance);
                var chosenValues = chosen.Select(i => candidates[i]).ToList();
                var sum = chosenValues.Sum();

                ToolResult result;
                if (!string.IsNullOrWhiteSpace(outputAddress))
                {
                    var values = new object[Math.Max(chosenValues.Count, 1), 1];
                    for (var i = 0; i < chosenValues.Count; i++)
                    {
                        values[i, 0] = chosenValues[i];
                    }

                    result = WriteGrid(args, sheetName, outputAddress, values, values.GetLength(0), 1);
                }
                else
                {
                    result = new ToolResult { Ok = true };
                }

                if (result.Ok)
                {
                    result.Data = new
                    {
                        found = chosenValues.Count > 0,
                        target,
                        sum,
                        within_tolerance = Math.Abs(sum - target) <= tolerance,
                        chosen_indices = chosen.ToArray(),
                        chosen_values = chosenValues.ToArray(),
                        output_address = string.IsNullOrWhiteSpace(outputAddress) ? null : outputAddress
                    };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult FillRandomNumbers(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var min = GetDouble(args, "min", 0);
            var max = GetDouble(args, "max", 100);
            var decimals = GetInt(args, "decimals", 0);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var random = new Random();
                var output = new object[rows, cols];
                var scale = Math.Pow(10, Math.Max(0, Math.Min(decimals, 6)));
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var raw = min + (max - min) * random.NextDouble();
                        output[r, c] = Math.Round(raw * scale) / scale;
                    }
                }

                var result = WriteGrid(args, Convert.ToString(sheet.Name), address, output, rows, cols);
                if (result.Ok)
                {
                    result.Data = new { rows, columns = cols, min, max, decimals };
                }

                return result;
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult FillNumberSequence(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var step = GetDouble(args, "step", 1);
            var count = Math.Max(1, Math.Min(GetInt(args, "count", 10), 10000));
            var direction = (Optional(args, "direction") ?? "down").ToLowerInvariant();
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var writeDown = direction != "right";
                var start = args.TryGetValue("start", out var startValue) && startValue != null
                    ? ConvertNum(startValue)
                    // schema：address 为「序列起始单元格」，start 缺省用当前单元格值。
                    // range.Value2 对多格选区返回数组，直接 ConvertNum 会得 0；Flatten 后取左上角，
                    // 使选区多格时以首行首列为起点，符合功能区「选中数据列后向下填充」的常见姿势。
                    : ConvertNum(FlattenValues(range.Value2, Convert.ToInt32(range.Rows.Count), Convert.ToInt32(range.Columns.Count))[0, 0]);

                var length = count;
                var target = writeDown
                    ? ExcelRangeHelper.Resize(address, length, 1)
                    : ExcelRangeHelper.Resize(address, 1, length);
                if (string.IsNullOrEmpty(target))
                {
                    return Invalid("INVALID_ADDRESS", $"地址无效：{address}");
                }

                var output = new object[writeDown ? length : 1, writeDown ? 1 : length];
                for (var i = 0; i < length; i++)
                {
                    var v = start + step * i;
                    if (writeDown)
                    {
                        output[i, 0] = v;
                    }
                    else
                    {
                        output[0, i] = v;
                    }
                }

                var result = WriteGrid(args, Convert.ToString(sheet.Name), target, output, output.GetLength(0), output.GetLength(1));
                if (result.Ok)
                {
                    result.Data = new { start, step, count = length, direction };
                }

                return result;
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult GeneratePermutationList(Dictionary<string, object> args)
        {
            var n = GetInt(args, "n", 3);
            var outputAddress = Required(args, "output_address");
            var combination = GetBool(args, "combination", false);
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                n = Math.Max(1, Math.Min(n, 8));
                List<string> lines;
                if (combination)
                {
                    lines = new List<string>();
                    for (var a = 1; a <= n; a++)
                    {
                        for (var b = a + 1; b <= n; b++)
                        {
                            lines.Add($"{a},{b}");
                        }
                    }
                }
                else
                {
                    lines = JyyjtoolAlgorithms.GeneratePermutations(n, 5000).ToList();
                }

                var values = new object[lines.Count, 1];
                for (var i = 0; i < lines.Count; i++)
                {
                    values[i, 0] = lines[i];
                }

                var sheetName = Convert.ToString(sheet.Name);
                var result = WriteGrid(args, sheetName, outputAddress, values, values.GetLength(0), 1);
                if (result.Ok)
                {
                    result.Data = new { n, combination, count = values.GetLength(0), output_address = outputAddress };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        // ======================== 行列表格 ========================

        public ToolResult InsertBlankRowsEveryN(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var interval = Math.Max(1, GetInt(args, "interval", 1));
            var count = Math.Max(1, GetInt(args, "count", 1));
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var totalRows = Convert.ToInt32(range.Rows.Count);
                var sheetName = Convert.ToString(sheet.Name);

                // 每隔 interval 行插入 count 行空行：在第 k*interval 行后插 count 行。
                var boundaries = new List<int>();
                for (var b = 1; b * interval < totalRows; b++)
                {
                    boundaries.Add(b * interval);
                }

                const int xlShiftDown = -4121;
                var inserted = 0;
                for (var i = boundaries.Count - 1; i >= 0; i--)
                {
                    var b = boundaries[i];
                    var from = Math.Min(b + 1, totalRows);
                    var to = Math.Min(b + count, totalRows);
                    dynamic rowsRange = null;
                    try
                    {
                        rowsRange = sheet.Rows[$"{from}:{to}"];
                        rowsRange.Insert(xlShiftDown);
                        inserted += count;
                    }
                    finally
                    {
                        Release((object)rowsRange);
                    }
                }

                RaiseStructural(args, "dimension-insert", new { sheet = sheetName, dimension = "row", count = inserted });
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!{address}" },
                    Data = new { inserted_rows = inserted, interval, per_step = count }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult DeleteBlankRowsRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var onlyFullyBlank = GetBool(args, "only_fully_blank", false);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var sheetName = Convert.ToString(sheet.Name);
                var values = (object[,])FlattenValues(range.Value2, rows, cols);
                var deleted = 0;
                for (var r = rows - 1; r >= 0; r--)
                {
                    var blank = false;
                    if (onlyFullyBlank)
                    {
                        blank = true;
                        for (var c = 0; c < cols; c++)
                        {
                            if (!IsBlank(values[r, c]))
                            {
                                blank = false;
                                break;
                            }
                        }
                    }
                    else
                    {
                        blank = IsBlank(values[r, 0]);
                    }

                    if (!blank)
                    {
                        continue;
                    }

                    dynamic rowRange = null;
                    try
                    {
                        rowRange = range.Rows[r + 1];
                        rowRange.EntireRow.Delete();
                    }
                    finally
                    {
                        Release((object)rowRange);
                    }

                    deleted++;
                }

                RaiseStructural(args, "dimension-insert", new { sheet = sheetName, dimension = "row", count = deleted, undo = "delete" });
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!{address}" },
                    Data = new { deleted_rows = deleted, only_fully_blank = onlyFullyBlank }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult MakePayslipRows(Dictionary<string, object> args)
        {
            var headerAddress = Required(args, "header_address");
            var dataAddress = Required(args, "data_address");
            var outputAddress = Optional(args, "output_address") ?? dataAddress;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var header = ReadRange((object)sheet, headerAddress, out _);
                var data = ReadRange((object)sheet, dataAddress, out _);
                var headerRows = header.GetLength(0);
                var dataRows = data.GetLength(0);
                if (headerRows == 0)
                {
                    return Invalid("EMPTY_HEADER", "表头区域为空。");
                }

                var cols = data.GetLength(1);
                var headerCols = header.GetLength(1);
                var width = Math.Max(cols, headerCols);
                var output = new object[dataRows * (headerRows + 1), width];
                for (var i = 0; i < dataRows; i++)
                {
                    for (var h = 0; h < Math.Min(headerRows, header.GetLength(1) > 0 ? headerRows : 1); h++)
                    {
                        for (var c = 0; c < width; c++)
                        {
                            if (h < headerRows)
                            {
                                var hv = (header.GetLength(0) > h && header.GetLength(1) > c) ? header[h, c] : null;
                                output[i * (headerRows + 1) + h, c] = hv;
                            }
                        }
                    }

                    for (var c = 0; c < width; c++)
                    {
                        output[i * (headerRows + 1) + (headerRows), c] = (data.GetLength(1) > c) ? data[i, c] : null;
                    }
                }

                var result = WriteGrid(args, sheetName, outputAddress, output, output.GetLength(0), width);
                if (result.Ok)
                {
                    result.Data = new { sheet = sheetName, output_rows = output.GetLength(0), output_address = outputAddress };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult FillBlankCellsDown(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var copyValue = GetBool(args, "copy_value", true);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic formula = null;
            dynamic values = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var sheetName = Convert.ToString(sheet.Name);
                var current = copyValue ? range.Value2 : range.Formula;
                var matrix = (object[,])FlattenValues(current, rows, cols);
                int filled = 0;
                for (var c = 0; c < cols; c++)
                {
                    object carry = null;
                    for (var r = 0; r < rows; r++)
                    {
                        var blank = IsBlank(matrix[r, c]);
                        if (blank)
                        {
                            if (carry != null)
                            {
                                matrix[r, c] = carry;
                                filled++;
                            }
                        }
                        else
                        {
                            carry = matrix[r, c];
                        }
                    }
                }

                var result = WriteGrid(args, sheetName, address, matrix, rows, cols);
                if (result.Ok)
                {
                    result.Data = new { filled_cells = filled, copy_value = copyValue };
                }

                return result;
            }
            finally
            {
                Release((object)values);
                Release((object)formula);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult MergeDuplicateCells(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var firstOnly = GetBool(args, "first_only", true);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var sheetName = Convert.ToString(sheet.Name);
                var values = (object[,])FlattenValues(range.Value2, rows, cols);
                // 按列对相邻相同单元格垂直合并
                for (var c = 0; c < cols; c++)
                {
                    var runStart = 0;
                    for (var r = 1; r <= rows; r++)
                    {
                        var changed = r == rows || !object.Equals(Normalize(values[r, c]), Normalize(values[r - 1, c]));
                        if (!changed)
                        {
                            continue;
                        }

                        if (r - runStart > 1)
                        {
                            dynamic top = null;
                            dynamic bottom = null;
                            dynamic mergeRange = null;
                            try
                            {
                                top = range.Cells[runStart + 1, c + 1];
                                bottom = range.Cells[r, c + 1];
                                mergeRange = sheet.Range[top, bottom];
                                if (firstOnly)
                                {
                                    mergeRange.Value2 = values[runStart, c];
                                }

                                mergeRange.Merge();
                            }
                            finally
                            {
                                Release((object)mergeRange);
                                Release((object)bottom);
                                Release((object)top);
                            }
                        }

                        runStart = r;
                    }
                }

                RaiseStructural(args, "merge-cells", new { sheet = sheetName, address, first_only = firstOnly });
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!{address}" },
                    Data = new { merged = true, address, first_only = firstOnly }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult FuzzyMatchColumns(Dictionary<string, object> args)
        {
            var firstAddress = Required(args, "first_address");
            var secondAddress = Required(args, "second_address");
            // 工具分类审查第十批（参数 Clamp 家族）：threshold>1 恒无匹配、<0 恒全匹配（静默无意义）——Clamp 0~1。
            var threshold = Math.Min(Math.Max(GetDouble(args, "threshold", 0.6), 0.0), 1.0);
            var outputAddress = Optional(args, "output_address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var first = ReadRange((object)sheet, firstAddress, out _);
                var second = ReadRange((object)sheet, secondAddress, out _);
                var left = SingleColumn(first);
                var right = SingleColumn(second);
                var matches = new List<object[]>();
                for (var i = 0; i < left.Count; i++)
                {
                    if (IsBlank(left[i]))
                    {
                        continue;
                    }

                    var best = -1;
                    var bestSim = 0.0;
                    for (var j = 0; j < right.Count; j++)
                    {
                        if (IsBlank(right[j]))
                        {
                            continue;
                        }

                        var sim = Similarity(ToText2(left[i]), ToText2(right[j]));
                        if (sim > bestSim)
                        {
                            bestSim = sim;
                            best = j;
                        }
                    }

                    matches.Add(new object[] { ToText2(left[i]), best >= 0 ? ToText2(right[best]) : null, bestSim, bestSim >= threshold });
                }

                ToolResult result;
                if (!string.IsNullOrWhiteSpace(outputAddress))
                {
                    var m = Math.Max(matches.Count, 1);
                    var values = new object[m, 3];
                    for (var i = 0; i < m; i++)
                    {
                        if (i < matches.Count)
                        {
                            values[i, 0] = matches[i][0];
                            values[i, 1] = matches[i][1];
                            values[i, 2] = matches[i][3];
                        }
                        else
                        {
                            values[i, 2] = false;
                        }
                    }

                    result = WriteGrid(args, sheetName, outputAddress, values, m, 3);
                }
                else
                {
                    result = new ToolResult { Ok = true };
                }

                if (result.Ok)
                {
                    result.Data = new
                    {
                        threshold,
                        matched = matches.Count(x => (bool)x[3]),
                        total = matches.Count,
                        output_address = string.IsNullOrWhiteSpace(outputAddress) ? null : outputAddress
                    };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult WrapFormulasWithIfError(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var fallback = Optional(args, "fallback") ?? string.Empty;
            var target = Optional(args, "target_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic formulas = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var sheetName = Convert.ToString(sheet.Name);
                formulas = range.Formula;
                var matrix = (object[,])FlattenValues(formulas, rows, cols);
                var wrapped = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var text = matrix[r, c] == null ? string.Empty : Convert.ToString(matrix[r, c]);
                        if (text.StartsWith("=", StringComparison.Ordinal) &&
                            !text.StartsWith("=IFERROR(", StringComparison.OrdinalIgnoreCase))
                        {
                            matrix[r, c] = $"=IFERROR({text.Substring(1)},{fallback})";
                            wrapped++;
                        }
                    }
                }

                var result = WriteGrid(args, sheetName, target, matrix, rows, cols);
                if (result.Ok)
                {
                    result.Data = new { wrapped_cells = wrapped, fallback };
                }

                return result;
            }
            finally
            {
                Release((object)formulas);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult Convert2dTo1d(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var keyColumns = GetInt(args, "key_columns", 1);
            var outputAddress = Optional(args, "output_address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var matrix = ReadRange((object)sheet, address, out _);
                var unpivoted = JyyjtoolAlgorithms.Unpivot(matrix, keyColumns);
                if (unpivoted.Count == 0)
                {
                    return Invalid("EMPTY", "区域没有可转置的数据。");
                }

                var width = unpivoted.Max(r => r.Length);
                var values = new object[unpivoted.Count, width];
                for (var i = 0; i < unpivoted.Count; i++)
                {
                    for (var c = 0; c < width; c++)
                    {
                        values[i, c] = unpivoted[i].Length > c ? unpivoted[i][c] : null;
                    }
                }

                string dest;
                if (string.IsNullOrWhiteSpace(outputAddress))
                {
                    outputAddress = CellBelow((object)sheet, address, values.GetLength(0) + 4);
                    dest = outputAddress;
                }
                else
                {
                    dest = outputAddress;
                }

                var result = WriteGrid(args, sheetName, dest, values, values.GetLength(0), width);
                if (result.Ok)
                {
                    result.Data = new { output_rows = values.GetLength(0), output_columns = width, key_columns = keyColumns, output_address = dest };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult Convert1dTo2d(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var columnsPerRow = Math.Max(1, GetInt(args, "columns_per_row", 3));
            var outputAddress = Optional(args, "output_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var matrix = ReadRange((object)sheet, address, out _);
                var all = new List<object>();
                for (var r = 0; r < matrix.GetLength(0); r++)
                {
                    for (var c = 0; c < matrix.GetLength(1); c++)
                    {
                        if (!IsBlank(matrix[r, c]))
                        {
                            all.Add(matrix[r, c]);
                        }
                    }
                }

                var rowCount = all.Count == 0 ? 0 : (int)Math.Ceiling(all.Count * 1.0 / columnsPerRow);
                var values = new object[rowCount, columnsPerRow];
                for (var i = 0; i < all.Count; i++)
                {
                    values[i / columnsPerRow, i % columnsPerRow] = all[i];
                }

                var result = WriteGrid(args, sheetName, outputAddress, values, rowCount, columnsPerRow);
                if (result.Ok)
                {
                    result.Data = new { output_rows = rowCount, output_columns = columnsPerRow, output_address = outputAddress };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult CreatePivotTable(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var rowField = Math.Max(1, GetInt(args, "row_field", 1));
            var valueField = Math.Max(1, GetInt(args, "value_field", 2));
            var aggregation = (Optional(args, "aggregation") ?? "sum").ToLowerInvariant();
            if (aggregation != "sum" && aggregation != "count" && aggregation != "average")
            {
                return Invalid("INVALID_AGGREGATION",
                    $"aggregation 仅支持 sum/count/average，当前为 {aggregation}。");
            }

            var outputAddress = Optional(args, "output_address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var matrix = ReadRange((object)sheet, address, out _);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                if (rows < 2)
                {
                    return Invalid("EMPTY", "明细区域至少需要表头与一行数据。");
                }

                if (rowField > cols)
                {
                    return Invalid("INVALID_COLUMN", $"row_field={rowField} 超出明细区域列数 {cols}。");
                }

                if (valueField > cols)
                {
                    return Invalid("INVALID_COLUMN", $"value_field={valueField} 超出明细区域列数 {cols}。");
                }

                var rf = rowField - 1;
                var vf = valueField - 1;
                var groups = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
                var groupCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var header = matrix[0, rf];
                var valueHeader = matrix[0, vf];
                for (var r = 1; r < rows; r++)
                {
                    var key = matrix[r, rf] == null ? string.Empty : Convert.ToString(matrix[r, rf]);
                    if (!groups.TryGetValue(key, out var list))
                    {
                        list = new List<double>();
                        groups[key] = list;
                    }

                    if (!groupCounts.TryGetValue(key, out var cnt))
                    {
                        groupCounts[key] = 1;
                    }
                    else
                    {
                        groupCounts[key] = cnt + 1;
                    }

                    if (double.TryParse(Convert.ToString(matrix[r, vf]), NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                    {
                        list.Add(num);
                    }
                }

                var outputRows = groups.Count + 1;
                var values = new object[outputRows, 4];
                values[0, 0] = header;
                values[0, 1] = valueHeader;
                values[0, 2] = "聚合";
                values[0, 3] = "取值";
                var idx = 1;
                foreach (var pair in groups)
                {
                    var aggValue = aggregation == "count"
                        ? (object)(
                            groupCounts.TryGetValue(pair.Key, out var cnt) ? cnt : 0)
                        : aggregation == "average"
                            ? (pair.Value.Count == 0 ? 0D : pair.Value.Average())
                            : (object)pair.Value.Sum();
                    values[idx, 0] = pair.Key;
                    values[idx, 1] = aggValue;
                    values[idx, 2] = aggregation;
                    values[idx, 3] = "";
                    idx++;
                }

                string dest;
                if (string.IsNullOrWhiteSpace(outputAddress))
                {
                    outputAddress = CellBelow((object)sheet, address, values.GetLength(0) + 4);
                    dest = outputAddress;
                }
                else
                {
                    dest = outputAddress;
                }

                var result = WriteGrid(args, sheetName, dest, values, values.GetLength(0), 4);
                if (result.Ok)
                {
                    result.Data = new { groups = groups.Count, aggregation, row_field = rowField, value_field = valueField, output_address = dest };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        // ======================== 工作表 ========================

        public ToolResult SplitSheetByField(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var fieldColumn = Math.Max(1, GetInt(args, "field_column", 1));
            var prefix = Optional(args, "sheet_prefix");
            var keepHeader = GetBool(args, "keep_header", true);
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                var sheetName = Convert.ToString(sheet.Name);
                var matrix = ReadRange((object)sheet, address, out _);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                if (rows < 1)
                {
                    return Invalid("EMPTY", "数据区域为空。");
                }

                if (fieldColumn > cols)
                {
                    return Invalid("INVALID_COLUMN", $"field_column={fieldColumn} 超出数据区域列数 {cols}。");
                }

                var fc = fieldColumn - 1;
                var groups = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
                var headerRow = new List<object[]>();
                for (var r = 0; r < rows; r++)
                {
                    var row = new object[cols];
                    for (var c = 0; c < cols; c++)
                    {
                        row[c] = matrix[r, c];
                    }

                    if (r == 0)
                    {
                        headerRow.Add(row);
                        continue;
                    }

                    var key = matrix[r, fc] == null ? string.Empty : Convert.ToString(matrix[r, fc]);
                    if (!groups.TryGetValue(key, out var list))
                    {
                        list = new List<object[]>();
                        groups[key] = list;
                    }

                    list.Add(row);
                }

                var created = new List<string>();
                foreach (var pair in groups)
                {
                    var baseName = string.IsNullOrWhiteSpace(prefix)
                        ? SanitizeSheetName(pair.Key)
                        : SanitizeSheetName(prefix + pair.Key);
                    var finalName = UniqueSheetName(workbook, baseName);
                    dynamic anchor = null;
                    dynamic newSheet = null;
                    try
                    {
                        anchor = workbook.Worksheets[workbook.Worksheets.Count];
                        newSheet = workbook.Worksheets.Add(After: anchor);
                        newSheet.Name = finalName;
                    }
                    finally
                    {
                        // 锚点/新建表均为工作表级共享 RCW：不做强制释放（防"RCW 分离"），
                        // 置空即断开本方法引用，交由 GC 回收（M16 真机 WPS 实证）。
                        anchor = null;
                        newSheet = null;
                    }
                    var valueList = pair.Value;
                    var outRows = valueList.Count + (keepHeader ? 1 : 0);
                    var outMatrix = new object[outRows, cols];
                    var rowIdx = 0;
                    if (keepHeader && headerRow.Count > 0)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            outMatrix[0, c] = headerRow[0][c];
                        }

                        rowIdx = 1;
                    }

                    for (var i = 0; i < valueList.Count; i++, rowIdx++)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            outMatrix[rowIdx, c] = valueList[i][c];
                        }
                    }

                    WriteGrid(args, finalName, "A1", outMatrix, outRows, cols);
                    created.Add(finalName);
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { source_sheet = sheetName, field_column = fieldColumn, sheets = created.ToArray() }
                };
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        public ToolResult CreateSheetByValue(Dictionary<string, object> args)
        {
            var nameCell = Required(args, "name_cell");
            var baseSheet = Optional(args, "base_sheet");
            var position = (Optional(args, "position") ?? "end").ToLowerInvariant();
            var reference = Optional(args, "reference");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var newSheetName = Convert.ToString(ReadScalar((object)sheet, nameCell));
                if (string.IsNullOrWhiteSpace(newSheetName))
                {
                    return Invalid("EMPTY_NAME", "名称单元格为空。");
                }

                newSheetName = SanitizeSheetName(newSheetName);
                newSheetName = UniqueSheetName(workbook, newSheetName);

                dynamic newSheet = null;
                try
                {
                    dynamic anchor = null;
                    try
                    {
                        anchor = workbook.Worksheets[workbook.Worksheets.Count];
                        newSheet = workbook.Worksheets.Add(After: anchor);
                    }
                    finally
                    {
                        // 锚点为工作表级共享 RCW，遵循项目策略不做强制释放（防"RCW 分离"）。
                        anchor = null;
                    }

                    // position=before/after 时移动到参考表旁；end 时保持追加在末尾
                    if (string.Equals(position, "before", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(position, "after", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrWhiteSpace(reference))
                        {
                            try
                            {
                                dynamic refSheet = workbook.Worksheets.Item[reference];
                                if (string.Equals(position, "before", StringComparison.OrdinalIgnoreCase))
                                {
                                    newSheet.Move(Before: refSheet);
                                }
                                else
                                {
                                    newSheet.Move(After: refSheet);
                                }

                                refSheet = null; // 工作表级共享 RCW 不强制释放
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }

                    newSheet.Name = newSheetName;
                    if (!string.IsNullOrWhiteSpace(baseSheet))
                    {
                        dynamic baseRange = null;
                        try
                        {
                            dynamic src = workbook.Worksheets.Item[baseSheet];
                            baseRange = src.UsedRange;
                            dynamic dst = newSheet.Range[ExcelRangeHelper.Resize("A1", Convert.ToInt32(baseRange.Rows.Count), Convert.ToInt32(baseRange.Columns.Count))];
                            dst.Value2 = baseRange.Value2;
                            Release((object)dst);
                            Release((object)src);
                        }
                        finally
                        {
                            Release((object)baseRange);
                        }
                    }

                    RaiseStructural(args, "sheet-insert", new { sheets = new[] { newSheetName } });
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { sheet = newSheetName, created_from = nameCell }
                    };
                }
                finally
                {
                    // 新建表的使用全部结束后再释放（M16 全量审查：原实现从不释放）。
                    Release((object)newSheet);
                }
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult CompareTwoSheets(Dictionary<string, object> args)
        {
            var sheetA = Optional(args, "sheet_a");
            var addressA = Optional(args, "address_a");
            var sheetB = Optional(args, "sheet_b");
            var addressB = Optional(args, "address_b");
            var outputAddress = Optional(args, "output_address");
            dynamic workbook = null;
            dynamic sheetAObj = null;
            dynamic sheetBObj = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheetAObj = string.IsNullOrWhiteSpace(sheetA) ? application.ActiveSheet : workbook.Worksheets.Item[sheetA];
                sheetBObj = string.IsNullOrWhiteSpace(sheetB)
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[sheetB];
                var nameA = string.IsNullOrWhiteSpace(sheetA) ? Convert.ToString(application.ActiveSheet.Name) : sheetA;
                var nameB = string.IsNullOrWhiteSpace(sheetB) ? nameA : sheetB;

                // 工具分类审查第十批修复（静默假成功）：sheet 与 address 全缺省时 A=B=活动表 A1 起——
                // 同区域自比恒「无差异」（假成功，AI/用户误以为数据一致）。人话拒绝并给出正确用法。
                if (string.IsNullOrWhiteSpace(sheetA) && string.IsNullOrWhiteSpace(sheetB) &&
                    string.IsNullOrWhiteSpace(addressA) && string.IsNullOrWhiteSpace(addressB))
                {
                    return Invalid("SAME_SOURCE",
                        "sheet_a/address_a 与 sheet_b/address_b 均未指定，两方指向同一区域，对比无意义。请至少指定 sheet_b 或 address_b（如对比同表两块区域传不同的 address_a/address_b）。");
                }

                var ma = ReadRange((object)sheetAObj, string.IsNullOrWhiteSpace(addressA) ? "A1" : addressA, out _);
                var mb = ReadRange((object)sheetBObj, string.IsNullOrWhiteSpace(addressB) ? "A1" : addressB, out _);
                var maxRows = Math.Max(ma.GetLength(0), mb.GetLength(0));
                var maxCols = Math.Max(ma.GetLength(1), mb.GetLength(1));

                var diff = new List<object[]>();
                for (var r = 0; r < maxRows; r++)
                {
                    for (var c = 0; c < maxCols; c++)
                    {
                        var va = r < ma.GetLength(0) && c < ma.GetLength(1) ? ToText2(ma[r, c]) : null;
                        var vb = r < mb.GetLength(0) && c < mb.GetLength(1) ? ToText2(mb[r, c]) : null;
                        if (SafeEqual(va, vb))
                        {
                            continue;
                        }

                        var cell = r + 1 + "," + ExcelRangeHelper.ColumnName(c + 1);
                        diff.Add(new object[] { cell, va, vb, va == null ? "added" : vb == null ? "removed" : "changed" });
                    }
                }

                ToolResult result;
                if (!string.IsNullOrWhiteSpace(outputAddress))
                {
                    var n = diff.Count == 0 ? 1 : diff.Count;
                    var values = new object[n, 4];
                    if (diff.Count == 0)
                    {
                        values[0, 0] = "无差异";
                    }
                    else
                    {
                        for (var i = 0; i < diff.Count; i++)
                        {
                            values[i, 0] = diff[i][0];
                            values[i, 1] = diff[i][3];
                            values[i, 2] = diff[i][1];
                            values[i, 3] = diff[i][2];
                        }
                    }

                    result = WriteGrid(args, Convert.ToString(application.ActiveSheet.Name), outputAddress, values, n, 4);
                }
                else
                {
                    result = new ToolResult { Ok = true };
                }

                if (result.Ok)
                {
                    result.Data = new
                    {
                        sheet_a = nameA,
                        sheet_b = nameB,
                        equal = diff.Count == 0,
                        difference_count = diff.Count,
                        output_address = string.IsNullOrWhiteSpace(outputAddress) ? null : outputAddress
                    };
                }

                return result;
            }
            finally
            {
                Release((object)sheetBObj);
                Release((object)sheetAObj);
                Release((object)workbook);
            }
        }

        public ToolResult ProtectSheetWatermark(Dictionary<string, object> args)
        {
            var action = (Optional(args, "action") ?? "protect").ToLowerInvariant();
            var password = Optional(args, "password");
            var watermark = Optional(args, "watermark");
            var watermarkCell = Optional(args, "watermark_cell") ?? "A1";
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                switch (action)
                {
                    case "protect":
                        sheet.Protect(password ?? string.Empty);
                        break;
                    case "unprotect":
                        sheet.Unprotect(password ?? string.Empty);
                        break;
                    case "watermark":
                        if (string.IsNullOrEmpty(watermark))
                        {
                            return Invalid("EMPTY_WATERMARK", "水印文本为空。");
                        }

                        dynamic cellRange = null;
                        try
                        {
                            cellRange = sheet.Range[watermarkCell];
                            cellRange.Value2 = watermark;
                        }
                        finally
                        {
                            Release((object)cellRange);
                        }

                        break;
                    default:
                        return Invalid("BAD_ACTION", $"不支持的操作：{action}");
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!{watermarkCell}" },
                    Data = new { action, sheet = sheetName, watermark = watermark }
                };
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        // ======================== 复制粘贴增强 ========================

        public ToolResult CopyRangeAsImage(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var appearance = (Optional(args, "appearance") ?? "screen").ToLowerInvariant();
            var anchor = Optional(args, "anchor") ?? "A1";
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                // WPS 兼容（全量审查实测）：隐藏实例上 CopyPicture/Select/Paste 任一步都可能
                // 抛裸 E_FAIL。分步容错给出明确语义错误码，而不是让裸 COM 错误穿透。
                try
                {
                    range.CopyPicture(
                        Appearance: appearance == "printer" ? 2 : 1,
                        Format: -4147);
                }
                catch (System.Runtime.InteropServices.COMException ce)
                {
                    return Invalid("IMAGE_COPY_FAILED",
                        $"复制区域为图片失败（WPS 隐藏实例可能不支持 CopyPicture，请让宿主可见后重试）：{ce.Message}");
                }
                System.Threading.Thread.Sleep(120); // FIX（M26 遗留 1）：剪贴板渲染稳定等待（WPS 隐藏实例 Paste 竞态主因之一）

                dynamic anchorRange = null;
                try
                {
                    anchorRange = sheet.Range[anchor];
                    try
                    {
                        anchorRange.Select();
                    }
                    catch (Exception)
                    {
                        // 隐藏实例 Select 可能失败：降级 Activate，再不行直接 Paste（贴到活动格）
                        try { anchorRange.Activate(); } catch { }
                    }

                    // FIX（M26 遗留 1）：WPS 隐藏实例 sheet.Paste() OLE 剪贴板受限（0x800A03EC）。
                    // 三级降级：①sheet.Paste() ②sheet.Pictures().Paste()（图片对象通道）
                    // ③宿主瞬时可见一闪再贴（隐藏实例为 harness/MCP 自建，闪现安全）
                    bool pasted = false;
                    Exception lastPasteError = null;
                    try { sheet.Paste(); pasted = true; }
                    catch (Exception pe1) { lastPasteError = pe1; }
                    if (!pasted)
                    {
                        try
                        {
                            object picObj = sheet.Pictures().Paste();
                            dynamic pic = (object)picObj != null ? (dynamic)picObj : null;
                            pasted = pic != null;
                            if (pic != null) { try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)pic); } catch { } }
                        }
                        catch (Exception pe2) { lastPasteError = pe2; }
                    }
                    if (!pasted)
                    {
                        try
                        {
                            workbook.Application.Visible = true;
                            System.Threading.Thread.Sleep(150);
                            sheet.Paste();
                            pasted = true;
                        }
                        catch (Exception pe3) { lastPasteError = pe3; }
                        finally { try { workbook.Application.Visible = false; } catch { } }
                    }
                    if (!pasted)
                    {
                        var msg = lastPasteError is System.Runtime.InteropServices.COMException ? "（三级通道 CopyPicture→Pictures.Paste→瞬时可见均被 WPS 拒绝）" : "";
                        return Invalid("IMAGE_PASTE_FAILED", $"粘贴图片失败{msg}：{lastPasteError?.Message}");
                    }
                }
                finally
                {
                    Release((object)anchorRange);
                }
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{Convert.ToString(sheet.Name)}!{address}" },
                    Data = new { sheet = Convert.ToString(sheet.Name), source = address, copied = true }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult CopyVisibleRangeOnly(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var target = Optional(args, "target_address");
            if (string.IsNullOrWhiteSpace(target))
            {
                // fail-closed（M16 全量审查）：缺省写回源区会把可见行覆盖到源区顶部，
                // 隐藏/被筛选行的数据被静默破坏。必须显式给 target_address。
                return Invalid("INVALID_ARGS", "必须提供 target_address：缺省写回源区域会用可见行覆盖隐藏行数据。请把可见行写到其它位置。");
            }
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic values = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var sheetName = Convert.ToString(sheet.Name);
                values = range.Value2;
                var matrix = (object[,])FlattenValues(values, rows, cols);

                var copied = new List<object[]>();
                for (var r = 0; r < rows; r++)
                {
                    if (Convert.ToBoolean(range.Rows[r + 1].Hidden)) // 被筛选隐藏则跳过
                    {
                        continue;
                    }

                    var line = new object[cols];
                    for (var c = 0; c < cols; c++)
                    {
                        line[c] = matrix[r, c];
                    }

                    copied.Add(line);
                }

                if (copied.Count == 0)
                {
                    return new ToolResult { Ok = true, Data = new { copied_rows = 0, hide = true } };
                }

                var values2 = new object[copied.Count, cols];
                for (var i = 0; i < copied.Count; i++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        values2[i, c] = copied[i][c];
                    }
                }

                var result = WriteGrid(args, sheetName, target, values2, values2.GetLength(0), cols);
                if (result.Ok)
                {
                    result.Data = new { copied_rows = copied.Count, target_address = target, only_visible = true };
                }

                return result;
            }
            finally
            {
                Release((object)values);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult SmartPasteDelimiter(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var delimiterKey = (Optional(args, "delimiter") ?? "auto").ToLowerInvariant();
            var custom = Optional(args, "custom");
            var outputAddress = Optional(args, "output_address") ?? address;
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var sheetName = Convert.ToString(sheet.Name);
                var matrix = ReadRange((object)sheet, address, out _);

                char? delimiter = null;
                string customDelimiter = null;
                if (delimiterKey == "comma") delimiter = ',';
                else if (delimiterKey == "semicolon") delimiter = ';';
                else if (delimiterKey == "tab") delimiter = '\t';
                else if (delimiterKey == "pipe") delimiter = '|';
                else if (delimiterKey == "space") delimiter = ' ';
                else if (delimiterKey == "custom")
                {
                    if (string.IsNullOrWhiteSpace(custom))
                    {
                        return Invalid("INVALID_DELIMITER", "delimiter=custom 时必须提供非空 custom 分隔符。");
                    }

                    customDelimiter = custom;
                }

                var rows = matrix.GetLength(0);
                var output = new List<string[]>();
                for (var r = 0; r < rows; r++)
                {
                    var text = ToText2(matrix[r, 0]);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        output.Add(new[] { string.Empty });
                        continue;
                    }

                    string[] parts;
                    if (customDelimiter != null)
                    {
                        parts = JyyjtoolAlgorithms.SplitSmart(text, customDelimiter);
                    }
                    else if (delimiter == null)
                    {
                        var detected = JyyjtoolAlgorithms.DetectDelimiter(text);
                        parts = JyyjtoolAlgorithms.SplitSmart(text, detected);
                    }
                    else
                    {
                        parts = JyyjtoolAlgorithms.SplitSmart(text, delimiter.Value);
                    }

                    output.Add(parts);
                }

                var width = output.Count == 0 ? 1 : output.Max(p => p.Length);
                var values = new object[output.Count, width];
                for (var i = 0; i < output.Count; i++)
                {
                    for (var c = 0; c < width; c++)
                    {
                        values[i, c] = output[i].Length > c ? output[i][c] : string.Empty;
                    }
                }

                var result = WriteGrid(args, sheetName, outputAddress, values, values.GetLength(0), width);
                if (result.Ok)
                {
                    result.Data = new { output_rows = values.GetLength(0), output_columns = width, delimiter = delimiterKey, output_address = outputAddress };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        // ======================== 辅助 ========================

        /// <summary>跨区键值匹配回填（VLOOKUP 免公式版）：以源区域 key_column 逐行去 lookup_address 匹配，
        /// 命中则把查找区域 lookup_value_column 的值写入源区域 value_column；未命中保持原格不变。</summary>
        public ToolResult CrossReferenceColumns(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var keyColumn = GetInt(args, "key_column", 1);
            var valueColumn = GetInt(args, "value_column", 1);
            var lookupAddress = Required(args, "lookup_address");
            var lookupKeyColumn = GetInt(args, "lookup_key_column", 1);
            var lookupValueColumn = GetInt(args, "lookup_value_column", 1);
            var lastMatch = GetBool(args, "last_match", false);
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var matrix = ReadRange((object)sheet, address, out var sheetName);
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                if (keyColumn < 1 || keyColumn > cols || valueColumn < 1 || valueColumn > cols)
                {
                    return Invalid("INVALID_COLUMN", $"key_column/value_column 超出源区域列范围 1..{cols}。");
                }

                var lookupMatrix = ReadRange((object)sheet, lookupAddress, out _);
                if (lookupKeyColumn < 1 || lookupKeyColumn > lookupMatrix.GetLength(1)
                    || lookupValueColumn < 1 || lookupValueColumn > lookupMatrix.GetLength(1))
                {
                    return Invalid("INVALID_COLUMN", "lookup_key_column/lookup_value_column 超出查找区域列范围。");
                }

                var matched = JyyjtoolAlgorithms.CrossReferenceValues(
                    matrix, keyColumn - 1, lookupMatrix, lookupKeyColumn - 1, lookupValueColumn - 1, lastMatch);
                var filled = 0;
                for (var r = 0; r < rows; r++)
                {
                    if (matched[r] != null)
                    {
                        matrix[r, valueColumn - 1] = matched[r];
                        filled++;
                    }
                }

                var result = WriteGrid(args, sheetName, address, matrix, rows, cols);
                if (result.Ok)
                {
                    result.Data = new
                    {
                        sheet = sheetName,
                        address,
                        filled,
                        total = rows,
                        lookup_address = lookupAddress,
                        last_match = lastMatch
                    };
                }

                return result;
            }
            finally
            {
                Release((object)workbook);
                Release((object)sheet);
            }
        }

        private ToolResult WriteGrid(Dictionary<string, object> args, string sheetName, string address, object[,] values, int rows, int cols)
        {
            if (rows == 0)
            {
                rows = values.GetLength(0);
            }

            if (cols == 0)
            {
                cols = values.GetLength(1);
            }

            if (writer != null)
            {
                var writeArgs = new Dictionary<string, object>(args, StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = sheetName,
                    ["address"] = address
                };
                return writer(writeArgs, values, "auto", "jr_tool");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, sheetName);
                var resized = ExcelRangeHelper.Resize(address, rows, cols);
                if (string.IsNullOrEmpty(resized))
                {
                    return Invalid("INVALID_ADDRESS", $"地址无效：{address}");
                }

                range = sheet.Range[resized];
                range.Value2 = values;
                return new ToolResult { Ok = true, AffectedRanges = { $"{sheetName}!{resized}" } };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private string CellBelow(dynamic sheet, string address, int howManyRowsBelow)
        {
            dynamic range = null;
            try
            {
                range = sheet.Range[address];
                return ExcelRangeHelper.CellAddress(
                    Convert.ToInt32(range.Row) + howManyRowsBelow,
                    Convert.ToInt32(range.Column));
            }
            finally
            {
                Release((object)range);
            }
        }

        private object[,] ReadRange(dynamic sheet, string address, out string sheetName)
        {
            dynamic range = null;
            dynamic values = null;
            try
            {
                range = sheet.Range[address];
                sheetName = Convert.ToString(sheet.Name);
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (rows == 0 || cols == 0)
                {
                    return new object[0, 0];
                }

                values = range.Value2;
                return FlattenValues(values, rows, cols);
            }
            finally
            {
                Release((object)values);
                Release((object)range);
            }
        }

        private dynamic ResolveAddress(dynamic sheet, string address, out int originRow, out int originCol)
        {
            dynamic range = null;
            try
            {
                range = sheet.Range[address];
                originRow = Convert.ToInt32(range.Row);
                originCol = Convert.ToInt32(range.Column);
                return range;
            }
            finally
            {
                Release((object)range);
            }
        }

        private object ReadScalar(dynamic sheet, string address)
        {
            dynamic range = null;
            try
            {
                range = sheet.Range[address];
                return range.Value2;
            }
            finally
            {
                Release((object)range);
            }
        }

        private List<object> SingleColumn(object[,] matrix)
        {
            var result = new List<object>();
            for (var r = 0; r < matrix.GetLength(0); r++)
            {
                for (var c = 0; c < matrix.GetLength(1); c++)
                {
                    result.Add(matrix[r, c]);
                }
            }

            return result;
        }

        private static bool SafeEqual(string a, string b)
        {
            return (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) ||
                   string.Equals(a ?? string.Empty, b ?? string.Empty, StringComparison.Ordinal);
        }

        private static double Similarity(string a, string b)
        {
            if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b))
            {
                return 1.0;
            }

            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return 0.0;
            }

            return 1.0 - (double)Levenshtein(a, b) / Math.Max(a.Length, b.Length);
        }

        private static int Levenshtein(string a, string b)
        {
            var dp = new int[a.Length + 1, b.Length + 1];
            for (var i = 0; i <= a.Length; i++)
            {
                dp[i, 0] = i;
            }

            for (var j = 0; j <= b.Length; j++)
            {
                dp[0, j] = j;
            }

            for (var i = 1; i <= a.Length; i++)
            {
                for (var j = 1; j <= b.Length; j++)
                {
                    var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    dp[i, j] = Math.Min(Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1), dp[i - 1, j - 1] + cost);
                }
            }

            return dp[a.Length, b.Length];
        }

        private static object Normalize(object value)
        {
            // CVErr 错误单元格读回 Int32：先归一为错误文本，防止被 double.TryParse 当数字、
            // 导致不同错误格被判相等而合并（M16 全量审查）。
            value = ExcelErrorValues.Normalize(value);
            if (value is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                return d;
            }

            return value ?? string.Empty;
        }

        private static IEnumerable<int> ColumnRange(int from, int to)
        {
            for (var i = from; i <= to; i++)
            {
                yield return i;
            }
        }

        private static bool IsBlank(object value)
        {
            // CVErr 错误格文本为 "#N/A" 等，非空白——不会被误判为空、也不会被当作数字填充（M16）。
            return value == null || string.IsNullOrWhiteSpace(Convert.ToString(ExcelErrorValues.Normalize(value)));
        }

        private static string ToText(object[,] matrix, int row, int col)
        {
            if (row < 0 || row >= matrix.GetLength(0) || col < 0 || col >= matrix.GetLength(1))
            {
                return string.Empty;
            }

            return ToText2(matrix[row, col]);
        }

        private static string ToText2(object value)
        {
            // CVErr 错误单元格归一为错误文本（如 "#N/A"），防止 "-2146826246" 进入金额大写/
            // 银行家舍入/凑数/透视聚合等数值解析路径（M16 全量审查）。
            return value == null ? string.Empty : Convert.ToString(ExcelErrorValues.Normalize(value));
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

        private double ConvertNum(object value)
        {
            if (double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
            {
                return num;
            }

            return 0;
        }

        private void RaiseStructural(Dictionary<string, object> args, string undoKind, object payload)
        {
            if (recorder == null)
            {
                return;
            }

            recorder("jr_tool", args, undoKind, null, payload);
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "工作表";
            }

            var sb = new StringBuilder();
            foreach (var c in name.Trim())
            {
                if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']')
                {
                    sb.Append('_');
                }
                else
                {
                    sb.Append(c);
                }
            }

            var result = sb.ToString();
            return result.Length > 31 ? result.Substring(0, 31) : result;
        }

        private static string UniqueSheetName(dynamic workbook, string baseName)
        {
            var name = baseName;
            var n = 2;
            while (SheetExists(workbook, name))
            {
                name = $"{baseName}_{n++}";
            }

            return name;
        }

        private static bool SheetExists(dynamic workbook, string name)
        {
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                if (string.Equals(Convert.ToString(worksheet.Name), name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = application.ActiveSheet;
            }
            else
            {
                // fail-closed（M16 全量审查）：表名不存在时不得静默回退到活动表——
                // 那会把写入打到错误的工作表上。显式报 INVALID_SHEET。
                sheet = null;
                try
                {
                    sheet = workbook.Worksheets.Item[sheetName];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    sheet = null;
                }
                catch (InvalidCastException)
                {
                    // WPS 对不存在表名可能抛类型转换类异常，同按"不存在"处理。
                    sheet = null;
                }

                if (sheet == null)
                {
                    throw new KeyNotFoundException("工作表不存在：" + sheetName);
                }
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

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            try
            {
                var raw = value is Newtonsoft.Json.Linq.JValue j ? j.Value : value;
                return Convert.ToInt32(raw);
            }
            catch (Exception)
            {
                // 与 GetDouble 对齐：非法整数输入回退默认值，而非抛异常成为 500。
                return defaultValue;
            }
        }

        private static double GetDouble(Dictionary<string, object> args, string key, double defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            var raw = value is Newtonsoft.Json.Linq.JValue j ? j.Value : value;
            return double.TryParse(Convert.ToString(raw), NumberStyles.Any, CultureInfo.InvariantCulture, out var num)
                ? num
                : defaultValue;
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToBoolean(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(object value)
        {
            if (value == null || !System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                return;
            }

            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
            catch (System.Runtime.InteropServices.InvalidComObjectException)
            {
                // 嵌套 writer（WriteGenericRange→WriteRangeCore）已把同源 workbook/sheet 的 RCW 分离，
                // 局部 finally 再次清理属冗余，忽略，避免把成功写入的调用打成 EXCEL_COM_ERROR。
            }
        }
    }
}