using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 通用数据加工工具（第二批）：文本清洗/拆分/合并、日期标准化、编号补零、批量填充公式、多表合并。
    /// 除多表合并外均通过 writer 回调统一写入（自动 Resize + 记录撤销快照），保证全程可回滚。
    /// 与 ExcelDataToolsService 同型：读完全块先释放本地 COM 句柄再交给 writer，避免 RCW 竞态。
    /// </summary>
    public sealed class ExcelUtilityToolsService
    {
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;

        public ExcelUtilityToolsService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
        }

        // ===================== 文本清洗 =====================
        public ToolResult TextClean(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var trim = PBool(args, "trim", true);
            var collapse = PBool(args, "collapse_whitespace", false);
            var removeNonPrintable = PBool(args, "remove_non_printable", false);
            var caseMode = Optional(args, "case") ?? "none";
            if (!IsCaseMode(caseMode))
            {
                return Invalid("INVALID_CASE", $"case 仅支持 none/upper/lower，当前为 {caseMode}。");
            }
            var removeChars = Optional(args, "remove_chars");
            var find = Optional(args, "find");
            var replace = Optional(args, "replace");

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                var matrix = (object[][])ReadMatrix(range, rows, cols);
                var changed = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var v = matrix[r][c];
                        if (!(v is string s))
                        {
                            continue;
                        }

                        var nv = CleanText(s, trim, collapse, removeNonPrintable, caseMode, removeChars, find, replace);
                        if (nv != s)
                        {
                            matrix[r][c] = nv;
                            changed++;
                        }
                    }
                }

                var result = BuildMatrix(rows, cols, matrix);
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(new Dictionary<string, object>(args), result, "value", "text_clean");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { cleaned_cells = changed }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 文本拆分（一列拆多列） =====================
        public ToolResult TextSplit(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var column = PInt(args, "column", 0);
            var delimiter = Optional(args, "delimiter") ?? ",";
            var targetColumn = PInt(args, "target_column", 0);
            var keepSource = PBool(args, "keep_source", true);

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (column < 1 || column > cols)
                {
                    return Invalid("INVALID_COLUMN", $"column={column} 超出列范围 1..{cols}。");
                }

                var matrix = (object[][])ReadMatrix(range, rows, cols);
                var outCol = targetColumn > 0 ? targetColumn : column + 1;
                var maxParts = 1;
                var parts = new string[rows][];
                for (var r = 0; r < rows; r++)
                {
                    var raw = matrix[r][column - 1];
                    var text = raw == null ? string.Empty : Convert.ToString(raw, CultureInfo.InvariantCulture);
                    var split = text.Split(new[] { delimiter }, StringSplitOptions.None);
                    parts[r] = split;
                    if (split.Length > maxParts)
                    {
                        maxParts = split.Length;
                    }
                }

                // 结果矩阵：保留原列（可选），拆分结果写入 outCol 起的 maxParts 列
                var outWidth = Math.Max(cols, outCol + maxParts - 1);
                if (outWidth > 16384 || rows > 1048576)
                {
                    return Invalid("INVALID_RANGE", $"选区过大（{rows} 行 × {outWidth} 列），超出 Excel 或本工具可处理范围。");
                }

                var result = new object[rows, outWidth];
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        result[r, c] = NullToDBNull(keepSource ? matrix[r][c] : (c == column - 1 ? null : matrix[r][c]));
                    }

                    for (var k = 0; k < maxParts; k++)
                    {
                        result[r, outCol - 1 + k] = NullToDBNull(k < parts[r].Length ? parts[r][k] : null);
                    }
                }

                var writeArgs = new Dictionary<string, object>
                {
                    ["address"] = address,
                    ["sheet"] = sheetName
                };
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(writeArgs, result, "value", "text_split");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { source_column = column, split_columns = maxParts, target_column = outCol }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 文本合并（多列合并一列） =====================
        public ToolResult TextMerge(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var targetColumn = PInt(args, "target_column", 1);
                var delimiter = Optional(args, "delimiter") ?? " ";
            var skipEmpty = PBool(args, "skip_empty", true);
            var columns = new List<int>();
            if (args.TryGetValue("columns", out var rawCols) && rawCols is System.Collections.IEnumerable enCols && !(rawCols is string))
            {
                foreach (var item in enCols)
                {
                    if (int.TryParse(Convert.ToString(item, CultureInfo.InvariantCulture), out var cn))
                    {
                        columns.Add(cn);
                    }
                }
            }

            if (columns.Count == 0)
            {
                return Invalid("INVALID_COLUMNS", "请指定要合并的列号数组（columns）。");
            }

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (columns.Any(c => c < 1 || c > cols))
                {
                    return Invalid("INVALID_COLUMN", "columns 中存在超出列范围的列号。");
                }

                // target_column 允许指向源选区之外的列（合并结果可写到新区），仅受 Excel 列上限约束
                if (targetColumn < 1 || targetColumn > 16384)
                {
                    return Invalid("INVALID_COLUMN", $"target_column 必须为 1..16384 的列号，当前为 {targetColumn}。");
                }

                var matrix = (object[][])ReadMatrix(range, rows, cols);
                var result = new object[rows, Math.Max(cols, targetColumn)];
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        result[r, c] = NullToDBNull(matrix[r][c]);
                    }

                    var sb = new StringBuilder();
                    foreach (var cn in columns)
                    {
                        var v = matrix[r][cn - 1];
                        var text = v == null ? null : Convert.ToString(v, CultureInfo.InvariantCulture);
                        if (string.IsNullOrEmpty(text))
                        {
                            if (skipEmpty)
                            {
                                continue;
                            }

                            text = string.Empty;
                        }

                        if (sb.Length > 0)
                        {
                            sb.Append(delimiter);
                        }

                        sb.Append(text);
                    }

                    result[r, targetColumn - 1] = NullToDBNull(sb.ToString());
                }

                var writeArgs = new Dictionary<string, object>
                {
                    ["address"] = address,
                    ["sheet"] = sheetName
                };
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(writeArgs, result, "value", "text_merge");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { merged_columns = columns.Count, target_column = targetColumn }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 日期标准化 =====================
        public ToolResult DateNormalize(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var column = PInt(args, "column", 0);
            var inputFormat = Optional(args, "input_format");
            var outputFormat = Optional(args, "output_format") ?? "yyyy-MM-dd";
            var invalidAs = Optional(args, "invalid_as");

            dynamic sheet = null;
            dynamic range = null;
            dynamic nf = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (column < 1 || column > cols)
                {
                    return Invalid("INVALID_COLUMN", $"column={column} 超出列范围 1..{cols}。");
                }

                try
                {
                    // 原生日期在 Excel 里是数值序列，必须同时看原单元格 NumberFormat：
                    // 只有「数值且是日期格式」才当序列日期转换，避免把 ID/年份等纯数字列误写成乱日期。
                    nf = range.NumberFormat;
                }
                catch
                {
                }

                var matrix = (object[][])ReadMatrix(range, rows, cols);
                var normalized = 0;
                var failed = 0;
                for (var r = 0; r < rows; r++)
                {
                    var v = matrix[r][column - 1];
                    // FormatAt 接 dynamic 形参，返回值是 dynamic；不显式 cast 会让下方 TryParseDate
                    // 整调用变动态分派，返回的 DateTime? 退化为 boxed DateTime，.HasValue/.Value 全失效。
                    var parsed = TryParseDate(v, inputFormat, (string)FormatAt(nf, r + 1, column));
                    if (parsed.HasValue)
                    {
                        matrix[r][column - 1] = parsed.Value.ToString(outputFormat, CultureInfo.InvariantCulture);
                        normalized++;
                    }
                    else if (invalidAs != null)
                    {
                        matrix[r][column - 1] = invalidAs;
                        failed++;
                    }
                }

                var result = BuildMatrix(rows, cols, matrix);
                // 契约要求"按指定文本格式写回"：默认 value 模式经 Value2 写入时 Excel 会把
                // "2024-01-05" 这类日期字符串自动转回数值序列日期（OADate），使输出格式失效。
                // 写回前先把目标列设为文本格式，强制按字面文本存储，兑现输出格式契约；
                // 仅设目标列，避免把选区其余列的展示格式一并改成文本。
                try { sheet.Range[range.Cells[1, column], range.Cells[rows, column]].NumberFormat = "@"; }
                catch { }
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(new Dictionary<string, object>(args), result, "value", "date_normalize");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { normalized = normalized, failed = failed, output_format = outputFormat }
                };
            }
            finally
            {
                Release((object)nf);
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 编号补零/格式化 =====================
        public ToolResult NumberPad(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var column = PInt(args, "column", 0);
            var width = PInt(args, "width", 0);
            var prefix = Optional(args, "prefix") ?? "";
            var suffix = Optional(args, "suffix") ?? "";
            var padChar = Optional(args, "pad_char") ?? "0";
            if (width < 1)
            {
                return Invalid("INVALID_WIDTH", "width（总宽度）必须 ≥ 1。");
            }

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (column < 1 || column > cols)
                {
                    return Invalid("INVALID_COLUMN", $"column={column} 超出列范围 1..{cols}。");
                }

                var matrix = (object[][])ReadMatrix(range, rows, cols);
                var formatted = 0;
                for (var r = 0; r < rows; r++)
                {
                    var v = matrix[r][column - 1];
                    if (v == null)
                    {
                        continue;
                    }

                    var numeric = ToDouble(v);
                    if (!numeric.HasValue)
                    {
                        continue;
                    }

                    var digits = numeric.Value.ToString("0", CultureInfo.InvariantCulture);
                    var padded = digits.PadLeft(width, padChar.Length > 0 ? padChar[0] : '0');
                    matrix[r][column - 1] = prefix + padded + suffix;
                    formatted++;
                }

                var result = BuildMatrix(rows, cols, matrix);
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(new Dictionary<string, object>(args), result, "value", "number_pad");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { formatted = formatted, width = width }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 批量填充公式 =====================
        public ToolResult FillFormula(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var formula = Required(args, "formula");
            var startRow = PInt(args, "start_row", 0);
            if (string.IsNullOrWhiteSpace(formula))
            {
                return Invalid("INVALID_FORMULA", "公式模板（formula）不能为空。");
            }

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                long cells = rows * (long)cols;
                if (cells > 200000)
                {
                    return Invalid("INVALID_RANGE",
                        $"填充区域过大（{rows} 行 × {cols} 列，{cells} 个单元）。请缩小 address 后重试。");
                }

                var baseRow = startRow > 0 ? startRow : Convert.ToInt32(range.Row);
                var result = new object[rows, cols];
                for (var r = 0; r < rows; r++)
                {
                    var rowNumber = baseRow + r;
                    var text = formula.Replace("{row}", rowNumber.ToString(CultureInfo.InvariantCulture));
                    if (!text.TrimStart().StartsWith("=", StringComparison.Ordinal))
                    {
                        text = "=" + text;
                    }

                    for (var c = 0; c < cols; c++)
                    {
                        result[r, c] = text;
                    }
                }

                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(new Dictionary<string, object>(args), result, "formula", "fill_formula");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { filled_cells = rows * cols, start_row = baseRow }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 多工作表合并 =====================
        public ToolResult MergeSheets(Dictionary<string, object> args)
        {
            var sourceSheets = new List<string>();
            if (args.TryGetValue("source_sheets", out var rawSheets) && rawSheets is System.Collections.IEnumerable enSheets && !(rawSheets is string))
            {
                foreach (var item in enSheets)
                {
                    var name = Convert.ToString(item, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        sourceSheets.Add(name);
                    }
                }
            }

            if (sourceSheets.Count == 0)
            {
                return Invalid("INVALID_SHEETS", "请指定要合并的源工作表名数组（source_sheets）。");
            }

            var targetSheet = Required(args, "target_sheet");
            var includeHeaders = PBool(args, "include_headers", true);
            var startRow = Math.Max(1, PInt(args, "start_row", 1));

            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    return Invalid("NO_WORKBOOK", "当前没有打开的工作簿。");
                }

                var merged = new List<object[]>();
                var header = (object[])null;
                var totalRows = 0;
                var totalCols = 0;
                foreach (var name in sourceSheets)
                {
                    dynamic ws = null;
                    dynamic used = null;
                    try
                    {
                        ws = workbook.Worksheets[name];
                        used = ws.UsedRange;
                    }
                    catch (Exception uex) when (uex.Message.Contains("索引超出"))
                    {
                        // WPS 空白表 UsedRange/集合索引可能抛 BADINDEX：视作无数据源表，跳过即可
                        continue;
                    }

                    try
                    {
                    var rows = Convert.ToInt32(used.Rows.Count);
                    var cols = Convert.ToInt32(used.Columns.Count);
                        if (rows < 1 || cols < 1)
                        {
                            continue;
                        }

                        var matrix = (object[][])ReadMatrix(used, rows, cols);
                        var dataStart = Math.Min(startRow, rows + 1) - 1;
                        if (includeHeaders)
                        {
                            // 表头行恒在 matrix[0]：首个源表取作合并表头，其余源表假定共享同一表头，
                            // 一律跳过首行，否则会把后续源表的表头文字（如"Name"）当作数据合入。
                            if (header == null)
                            {
                                header = matrix[0];
                            }

                            if (dataStart < 1)
                            {
                                dataStart = 1;
                            }
                        }

                        for (var r = dataStart; r < rows; r++)
                        {
                            merged.Add(matrix[r]);
                        }

                        if (cols > totalCols)
                        {
                            totalCols = cols;
                        }

                        totalRows += Math.Max(0, rows - dataStart);
                    }
                    finally
                    {
                        Release((object)used);
                        Release((object)ws);
                    }
                }

                if (header != null)
                {
                    merged.Insert(0, header);
                    totalRows++;
                }

                if (totalRows == 0)
                {
                    return new ToolResult { Ok = true, Data = new { merged_rows = 0, message = "源表中没有可合并的数据。" } };
                }

                var result = new object[totalRows, totalCols];
                for (var r = 0; r < totalRows; r++)
                {
                    var row = merged[r];
                    for (var c = 0; c < totalCols; c++)
                    {
                        result[r, c] = NullToDBNull(c < row.Length ? row[c] : null);
                    }
                }

                var writeArgs = new Dictionary<string, object>
                {
                    ["address"] = ExcelRangeHelper.CellAddress(1, 1),
                    ["sheet"] = targetSheet
                };
                var res = writer(writeArgs, result, "value", "merge_sheets");
                return new ToolResult
                {
                    Ok = res.Ok,
                    ErrorCode = res.ErrorCode,
                    Data = new { merged_rows = totalRows, source_sheets = sourceSheets.Count, target_sheet = targetSheet }
                };
            }
            finally
            {
                // ActiveWorkbook 是共享根（与调用方/其他工具持有同一 RCW），不得 FinalRelease——
                // 否则会连带销毁调用方持有的 workbook RCW，后续访问抛「COM 对象与其基础 RCW 分开」。
                workbook = null;
            }
        }

        // ===================== 内部工具 =====================
        private static string CleanText(string s, bool trim, bool collapse, bool removeNonPrintable,
            string caseMode, string removeChars, string find, string replace)
        {
            var nv = s;
            if (removeNonPrintable)
            {
                var sb = new StringBuilder(nv.Length);
                foreach (var ch in nv)
                {
                    if (ch >= 32 || ch == '\t' || ch == '\n' || ch == '\r')
                    {
                        sb.Append(ch);
                    }
                }

                nv = sb.ToString();
            }

            if (trim)
            {
                nv = nv.Trim();
            }

            if (collapse)
            {
                var sb = new StringBuilder(nv.Length);
                var inSpace = false;
                foreach (var ch in nv)
                {
                    if (char.IsWhiteSpace(ch))
                    {
                        if (!inSpace)
                        {
                            sb.Append(' ');
                        }

                        inSpace = true;
                    }
                    else
                    {
                        sb.Append(ch);
                        inSpace = false;
                    }
                }

                nv = sb.ToString();
            }

            if (string.Equals(caseMode, "upper", StringComparison.OrdinalIgnoreCase))
            {
                nv = nv.ToUpperInvariant();
            }
            else if (string.Equals(caseMode, "lower", StringComparison.OrdinalIgnoreCase))
            {
                nv = nv.ToLowerInvariant();
            }

            if (!string.IsNullOrEmpty(removeChars))
            {
                var sb = new StringBuilder(nv.Length);
                foreach (var ch in nv)
                {
                    if (removeChars.IndexOf(ch) < 0)
                    {
                        sb.Append(ch);
                    }
                }

                nv = sb.ToString();
            }

            if (!string.IsNullOrEmpty(find))
            {
                nv = nv.Replace(find, replace ?? string.Empty);
            }

            return nv;
        }

        // XLL UDF M1 提公：实现上移至 Core TextDateParser（单一来源），此处保持 internal 签名转发，
        // ExcelUtilityToolsService 内部既有调用点零改动。
        internal static DateTime? TryParseDate(object v, string inputFormat, string numberFormat)
        {
            return Core.Tools.TextDateParser.TryParseDate(v, inputFormat, numberFormat);
        }

        /// <summary>读取范围 NumberFormat（COM 1 基二维数组或标量）的第 row(1 基) 行第 col(1 基) 列格式字符串。</summary>
        private static string FormatAt(object nf, int row, int col)
        {
            if (nf == null) return null;
            try
            {
                if (nf is string scalar) return scalar;
                if (nf is object[,] arr)
                {
                    var r = row + arr.GetLowerBound(0);
                    var c = col + arr.GetLowerBound(1);
                    if (r >= arr.GetLowerBound(0) && r <= arr.GetUpperBound(0) &&
                        c >= arr.GetLowerBound(1) && c <= arr.GetUpperBound(1) &&
                        arr[r, c] != null)
                    {
                        return Convert.ToString(arr[r, c], CultureInfo.InvariantCulture);
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        /// <summary>判断某单元格自定义数字格式是否属于「日期/时间」类（含年/月/日/时 记号）。纯数字/文本格式判定为非日期。</summary>
        internal static bool IsDateLikeFormat(string fmt)
        {
            return Core.Tools.TextDateParser.IsDateLikeFormat(fmt);
        }


        private static bool IsCaseMode(string caseMode)
        {
            return string.Equals(caseMode, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(caseMode, "upper", StringComparison.OrdinalIgnoreCase)
                || string.Equals(caseMode, "lower", StringComparison.OrdinalIgnoreCase);
        }

        private static object[,] EmptyMatrix(int rows, int cols)
        {
            var empty = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    empty[r, c] = null;
                }
            }

            return empty;
        }

        private object[][] ReadMatrix(dynamic range, int rows, int cols)
        {
            // WPS 兼容（全量审查实测）：WPS 表格上 UsedRange.Value2 偶发抛"索引超出了数组界限"
            // （Excel 同路径不复现；实测为 COM 层包装异常，类型不固定但 Message 稳定），
            // 区域为多行二维时按格逐读兜底，其余异常原样上抛。
            object obj;
            try
            {
                obj = range.Value2;
            }
            catch (System.Runtime.InteropServices.COMException cex) when (cex.HResult == unchecked((int)0x8002000B))
            {
                // DISP_E_BADINDEX：WPS 空白工作表 UsedRange 的典型错误 → 视作空数据
                obj = EmptyMatrix(rows, cols);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException) && ex.Message.Contains("索引超出"))
            {
                obj = EmptyMatrix(rows, cols);
            }

            object[,] arr;
            if (obj is object[,] m)
            {
                arr = m;
            }
            else
            {
                // 单格区域：Value2 返回标量（非二维数组）。此处构造的 0 基 1x1 数组
                // 必须与下方取值下标兼容——下方统一按 GetLowerBound 换算，
                // 修复 WPS/单格场景 arr[1,1] 对 0 埂数组越界（"索引超出了数组界限"）。
                arr = new object[1, 1] { { obj } };
            }

            // COM 区间数组（多格 Value2）是 1 基（GetLowerBound=1），手工构造是 0 基；
            // 用 LowerBound 换算后两种形态统一按"第 r 行第 c 列（从 1 计）"取值。
            var lower0 = arr.GetLowerBound(0);
            var lower1 = arr.GetLowerBound(1);
            var result = new object[rows][];
            for (var r = 0; r < rows; r++)
            {
                var rowVals = new object[cols];
                for (var c = 0; c < cols; c++)
                {
                    var v = (r + 1 <= arr.GetLength(0) && c + 1 <= arr.GetLength(1))
                        ? arr[lower0 + r, lower1 + c]
                        : DBNull.Value;
                    // CVErr 错误单元格经 Value2 读回 Int32（如 #N/A=-2146826246），统一归一为错误文本，
                    // 避免被补零/凑数/舍入等数值工具误判为数字并写回污染数据（M16 全量审查）。
                    rowVals[c] = (v is DBNull) ? null : ExcelErrorValues.Normalize(v);
                }

                result[r] = rowVals;
            }

            return result;
        }

        private static object[,] BuildMatrix(int rows, int cols, object[][] matrix)
        {
            var result = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = NullToDBNull(matrix[r][c]);
                }
            }

            return result;
        }

        private static object NullToDBNull(object v)
        {
            return v ?? DBNull.Value;
        }

        private dynamic Resolve(string sheetName)
        {
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("当前没有打开的工作簿。");
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return application.ActiveSheet;
            }

            return workbook.Worksheets[sheetName];
        }

        private static double? ToDouble(object v)
        {
            if (v == null)
            {
                return null;
            }

            if (v is double d)
            {
                return d;
            }

            if (v is int i)
            {
                return i;
            }

            if (v is long l)
            {
                return l;
            }

            if (v is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var dr))
            {
                return dr;
            }

            return null;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(dynamic com)
        {
            if (com != null)
            {
                try
                {
                    ExcelContextService.Release((object)com);
                }
                catch
                {
                }
            }
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            // Y2-4：缺参抛 ArgumentException 走 INVALID_ARGS 通道（原返空串 → 下游 COM 误诊，对齐 ExcelToolHost.Required 范式）
            var text = Convert.ToString(args != null && args.TryGetValue(key, out var v) ? v : null);
            if (!string.IsNullOrEmpty(text)) { return text; }
            throw new ArgumentException($"缺少必填参数：{key}");
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                return Convert.ToString(v);
            }

            return null;
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

        private static bool PBool(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null && bool.TryParse(Convert.ToString(v), out var b))
            {
                return b;
            }

            return fallback;
        }
    }
}
