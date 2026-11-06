using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// Markdown 表格双向转换：选区导出为 Markdown 表格、Markdown 表格导入为 Excel 区域。
    /// </summary>
    public sealed class ExcelMarkdownService
    {
        private readonly dynamic application;
        // 全量读入护栏：导出选区上限，防止整张表被拉成 17 亿元素矩阵 OOM。
        private const long MaxExportCells = 100000;

        public ExcelMarkdownService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult ExportToMarkdown(Dictionary<string, object> args)
        {
            var requested = Optional(args, "address");
            var includeHeaders = GetBool(args, "include_headers", true);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic value = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = string.IsNullOrWhiteSpace(requested) ? application.Selection : sheet.Range[requested];
                int rows = Convert.ToInt32(range.Rows.Count);
                int columns = Convert.ToInt32(range.Columns.Count);
                if (rows == 0 || columns == 0)
                {
                    return Invalid("EMPTY_RANGE", "区域为空，无法导出。");
                }

                // 全量读入护栏：rows×cols 用 long 防 int 溢出，超限显式拒绝而非拉全表。
                var cellCount = rows * (long)columns;
                if (cellCount > MaxExportCells)
                {
                    return Invalid("TOO_MANY_CELLS", $"选区过大（{rows} 行 × {columns} 列，{cellCount} 个单元），超过导出上限 {MaxExportCells}，请缩小范围。");
                }

                value = range.Value2;
                var matrix = FlattenValues(value, rows, columns);

                // 统一转成字符串矩阵，供 Markdown 排版用。
                var text = new string[rows, columns];
                for (var row = 0; row < rows; row++)
                {
                    for (var column = 0; column < columns; column++)
                    {
                        text[row, column] = CellText(matrix[row, column]);
                    }
                }

                // include_headers=true：首行作表头，正文从第 2 行起（不能再把首行重复进正文）；
                // false：选区全是数据，无表头，合成占位表头「列1..列n」、正文含全部行。
                var bodyStart = 0;
                string[] header;
                if (includeHeaders && rows >= 1)
                {
                    header = Enumerable.Range(0, columns).Select(i => text[0, i]).ToArray();
                    bodyStart = 1;
                }
                else
                {
                    header = Enumerable.Range(0, columns).Select(i => "列" + (i + 1)).ToArray();
                    bodyStart = 0;
                }

                var cells = new List<string[]>();
                for (var row = bodyStart; row < rows; row++)
                {
                    cells.Add(Enumerable.Range(0, columns).Select(i => text[row, i]).ToArray());
                }

                var markdown = BuildTable(header, cells, columns);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = Convert.ToString(range.Address),
                        rows = rows,
                        columns = columns,
                        include_headers = includeHeaders,
                        markdown = markdown
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

        public ToolResult ParseMarkdownToMatrix(Dictionary<string, object> args)
        {
            var markdown = Optional(args, "markdown");
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return Invalid("MISSING_MARKDOWN", "缺少必填参数：markdown。");
            }

            var lines = markdown
                .Replace("\r\n", "\n")
                .Split('\n')
                .Select(line => line.Trim())
                .Where(line => line.StartsWith("|", StringComparison.Ordinal) && !IsDivider(line))
                .ToList();
            if (lines.Count == 0)
            {
                return Invalid("PARSE_ERROR", "没有解析到 Markdown 表格行。");
            }

            var rows = new List<string[]>();
            foreach (var line in lines)
            {
                var row = line.Trim('|')
                    // M16 全量审查：导出侧把单元格内的 | 转义为 \|，导入侧必须只按未转义的 | 拆分
                    // 并反转义，否则含管道符的单元格在"导出→再导入"往返中被错误拆列。
                    .SplitUnescapedPipe()
                    .Select(cell => cell.Trim().Replace("\\|", "|"))
                    .ToArray();
                rows.Add(row);
            }

            var columnCount = rows.Max(row => row.Length);
            var matrix = new object[rows.Count, columnCount];
            for (var r = 0; r < rows.Count; r++)
            {
                for (var c = 0; c < rows[r].Length; c++)
                {
                    matrix[r, c] = rows[r][c];
                }
            }

            return new ToolResult
            {
                Ok = true,
                Data = new { rows = rows.Count, columns = columnCount, values = matrix }
            };
        }

        public ToolResult ImportToExcel(
            Dictionary<string, object> args,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer)
        {
            var parsed = ParseMarkdownToMatrix(args);
            if (!parsed.Ok)
            {
                return parsed;
            }

            var data = parsed.Data;
            var values = (object[,])((dynamic)data).values;
            var valueMode = GetString(args, "value_mode", "auto");
            return writer(args, values, valueMode, "import_markdown_to_excel");
        }

        private static string GetString(Dictionary<string, object> args, string key, string defaultValue)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value) ? defaultValue : value;
        }

        private static bool IsDivider(string line)
        {
            return line.Trim('|')
                .Split('|')
                .All(cell => cell.Trim().All(ch => ch == ':' || ch == '-' || ch == ' '));
        }

        private static string BuildTable(string[] header, List<string[]> cells, int columns)
        {
            var widths = new int[columns];
            for (var c = 0; c < columns; c++)
            {
                widths[c] = header[c].Length;
            }

            foreach (var row in cells)
            {
                for (var c = 0; c < columns; c++)
                {
                    widths[c] = Math.Max(widths[c], row[c].Length);
                }
            }

            var builder = new StringBuilder();
            builder.AppendLine(FormatRow(header, widths));
            builder.AppendLine(string.Join("|", widths.Select(w => "-".PadLeft(Math.Max(w, 3), '-'))) + "|");
            foreach (var row in cells)
            {
                builder.AppendLine(FormatRow(row, widths));
            }

            return builder.ToString();
        }

        private static string FormatRow(string[] row, int[] widths)
        {
            var parts = new string[widths.Length];
            for (var c = 0; c < widths.Length; c++)
            {
                var value = c < row.Length ? row[c] : string.Empty;
                parts[c] = value.PadRight(widths[c]);
            }

            return "| " + string.Join(" | ", parts) + " |";
        }

        private static string CellText(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var text = Convert.ToString(value);
            text = text.Replace("\r\n", " ").Replace("\n", " ").Replace("|", "\\|");
            return text;
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
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
            catch (Exception)
            {
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }

/// <summary>
/// M16 全量审查：按"未被反斜杠转义"的竖线拆分 Markdown 表格行——
/// 与导出侧 CellText 的 | → \| 转义配套，保证导出→导入往返无损。
/// </summary>
internal static class MarkdownPipeSplitExtensions
{
    public static string[] SplitUnescapedPipe(this string line)
    {
        var parts = new List<string>();
        var current = new System.Text.StringBuilder();
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '\\' && i + 1 < line.Length && line[i + 1] == '|')
            {
                current.Append("\\|"); // 保留转义序列，由上层统一反转义
                i++;
                continue;
            }

            if (ch == '|')
            {
                parts.Add(current.ToString());
                current.Clear();
                continue;
            }

            current.Append(ch);
        }

        parts.Add(current.ToString());
        return parts.ToArray();
    }
}
}
