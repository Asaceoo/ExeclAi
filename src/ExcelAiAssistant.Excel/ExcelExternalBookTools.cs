using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// P5-C1 跨工作簿只读引用工具：open_excel_external_book / read_external_range / close_external_book。
    /// 仅读外部句柄，不改动活动工作簿；写操作仍限当前活动工作簿（安全边界）。
    /// </summary>
    public static class ExcelExternalBookTools
    {
        /// <summary>单次外部读取的单元格总数上限：防止极端超宽/超高 range 触发巨型编组与序列化（OOM）。</summary>
        private const int MaxExternalCells = 2_000_000;

        public static void Register(ToolRegistry registry, ExternalBookRegistry books)
        {
            registry.Register(new ToolDescriptor(
                "open_excel_external_book",
                "以只读句柄打开指定路径的工作簿，返回句柄与工作表清单，供后续 read_external_range 引用。只读，不会改动该工作簿。",
                // 高危：path 由模型自主给出，ReadOnly 会自动放行，等于静默任意文件读取；须弹确认让用户看清被打开的文件路径。
                RiskLevel.Dangerous,
                args => OpenTool(books, args),
                new Dictionary<string, object>
                {
                    ["type"] = "object",
                    ["properties"] = new Dictionary<string, object>
                    {
                        ["path"] = new Dictionary<string, object> { ["name"] = "path", ["type"] = "string", ["description"] = "待打开工作簿的完整路径(.xlsx/.xls)" }
                    },
                    ["required"] = new[] { "path" }
                }));

            registry.Register(new ToolDescriptor(
                "read_external_range",
                "读取外部只读句柄中指定工作簿、工作表、区域的值，作为上下文引用。",
                RiskLevel.ReadOnly,
                args => ReadTool(books, args),
                new Dictionary<string, object>
                {
                    ["type"] = "object",
                    ["properties"] = new Dictionary<string, object>
                    {
                        ["handle"] = new Dictionary<string, object> { ["name"] = "handle", ["type"] = "string", ["description"] = "open_excel_external_book 返回的句柄" },
                        ["sheet"] = new Dictionary<string, object> { ["name"] = "sheet", ["type"] = "string", ["description"] = "工作表名" },
                        ["range"] = new Dictionary<string, object> { ["name"] = "range", ["type"] = "string", ["description"] = "区域地址，如 A1:C9" },
                        ["max_rows"] = new Dictionary<string, object> { ["name"] = "max_rows", ["type"] = "integer", ["description"] = "可选，最多读取行数" }
                    },
                    ["required"] = new[] { "handle", "sheet", "range" }
                }));

            registry.Register(new ToolDescriptor(
                "close_external_book",
                "显式关闭外部只读句柄并释放资源。",
                RiskLevel.ReadOnly,
                args => CloseTool(books, args),
                new Dictionary<string, object>
                {
                    ["type"] = "object",
                    ["properties"] = new Dictionary<string, object>
                    {
                        ["handle"] = new Dictionary<string, object> { ["name"] = "handle", ["type"] = "string", ["description"] = "要关闭的句柄" }
                    },
                    ["required"] = new[] { "handle" }
                }));
        }

        private static ToolResult OpenTool(ExternalBookRegistry books, Dictionary<string, object> args)
        {
            var path = AsString(args, "path");
            if (string.IsNullOrWhiteSpace(path))
            {
                return Fail("MISSING_PATH", "缺少 path 参数（工作簿完整路径）。");
            }

            try
            {
                var info = books.Open(path);
                return Succeed(JsonConvert.SerializeObject(new
                {
                    handle = info.Token,
                    name = info.Name,
                    path = info.Path,
                    sheets = info.Sheets
                }));
            }
            catch (Exception ex)
            {
                return Fail("OPEN_FAILED", "打开外部工作簿失败：" + ex.Message);
            }
        }

        private static ToolResult ReadTool(ExternalBookRegistry books, Dictionary<string, object> args)
        {
            var handle = AsString(args, "handle");
            var sheet = AsString(args, "sheet");
            var range = AsString(args, "range");
            if (string.IsNullOrWhiteSpace(handle) || string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(range))
            {
                return Fail("MISSING_PARAM", "需要 handle、sheet、range 三个参数。");
            }

            if (!books.TryGet(handle, out var workbook) || workbook == null)
            {
                return Fail("HANDLE_CLOSED", "句柄不存在或已关闭：" + handle);
            }

            dynamic book = workbook;

            // max_rows 必须有限且为正，防止负值/超大值触发整列无限读取与巨型序列化
            var rawMax = ParseInt(args, "max_rows", 200);
            var maxRows = rawMax < 1 ? 200 : (rawMax > 20000 ? 20000 : rawMax);
            try
            {
                var rows = ReadValues(book, sheet, range, maxRows);
                return Succeed(SerializeTable(rows));
            }
            catch (DataTooLargeException ex)
            {
                return Fail("TOO_MANY_CELLS", ex.Message);
            }
            catch (Exception ex)
            {
                return Fail("READ_FAILED", "读取外部区域失败：" + ex.Message);
            }
        }

        private static ToolResult CloseTool(ExternalBookRegistry books, Dictionary<string, object> args)
        {
            var handle = AsString(args, "handle");
            if (string.IsNullOrWhiteSpace(handle))
            {
                return Fail("MISSING_PARAM", "缺少 handle 参数。");
            }

            var ok = books.Close(handle);
            return Succeed(JsonConvert.SerializeObject(new { ok = ok, handle = handle, closed = ok }));
        }

        private static ToolResult Succeed(object data)
        {
            return new ToolResult { Ok = true, Data = data };
        }

        private static ToolResult Fail(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        /// <summary>读取区域为 List[List[object]]（按行），标量为单行单列。</summary>
        private static List<List<object>> ReadValues(dynamic book, string sheet, string range, int maxRows)
        {
            var result = new List<List<object>>();
            dynamic ws = null;
            dynamic rng = null;
            try
            {
                try
                {
                    ws = book.Worksheets[sheet];
                }
                catch
                {
                    throw new KeyNotFoundException("外部工作簿中不存在工作表：" + sheet);
                }

                rng = ws.Range[range];
                // 单元格数守卫前移（M16 全量审查）：先读 Rows/Columns 计数（廉价、不编组数据），
                // 超限直接拒绝；原实现先 Value2 整块编组再判断，超大区域在守卫触发前就已 OOM。
                var totalRows = Convert.ToInt32(rng.Rows.Count);
                var totalCols = Convert.ToInt32(rng.Columns.Count);
                if ((long)totalRows * totalCols > MaxExternalCells)
                {
                    throw new DataTooLargeException("读取单元格数(" + totalRows + "×" + totalCols + ")超过上限 " + MaxExternalCells + "，请缩小 range 范围或分批读取。");
                }

                var value = rng.Value2;
                if (value is object[,] arr)
                {
                    var rows = arr.GetLength(0);
                    var cols = arr.GetLength(1);

                    var limit = maxRows > 0 && rows > maxRows ? maxRows : rows;
                    for (var r = 1; r <= limit; r++)
                    {
                        var row = new List<object>();
                        for (var c = 1; c <= cols; c++)
                        {
                            row.Add(Normalize(arr[r, c]));
                        }

                        result.Add(row);
                    }
                }
                else
                {
                    result.Add(new List<object> { Normalize(value) });
                }

                return result;
            }
            finally
            {
                // M16 全量审查：ws/rng RCW 及时释放（原实现从不释放、亦无 finally）。
                if (rng != null)
                {
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)rng); } catch { }
                }

                if (ws != null)
                {
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)ws); } catch { }
                }
            }
        }

        private static object Normalize(object v)
        {
            if (v == null)
            {
                return null;
            }

            if (v is double d && Math.Truncate(d) == d && Math.Abs(d) < 1e15)
            {
                return (long)d;
            }

            return v;
        }

        private static string SerializeTable(List<List<object>> rows)
        {
            var text = new List<string>();
            foreach (var row in rows)
            {
                text.Add(string.Join("\t", row.Select(cell => cell == null ? "" : Convert.ToString(cell))));
            }

            return JsonConvert.SerializeObject(new
            {
                rows = rows.Select(r => (object)r.ToArray()).ToArray(),
                row_count = rows.Count,
                preview = string.Join("\n", text.Take(20))
            });
        }

        private static string AsString(Dictionary<string, object> args, string key)
        {
            return args != null && args.TryGetValue(key, out var v) ? Convert.ToString(v) : null;
        }

        private static int ParseInt(Dictionary<string, object> args, string key, int fallback)
        {
            if (args != null && args.TryGetValue(key, out var v))
            {
                if (int.TryParse(Convert.ToString(v), out var n))
                {
                    return n;
                }
            }

            return fallback;
        }
    }

    /// <summary>表示单次读取的目标单元格超过上限的受控异常（映射为 TOO_MANY_CELLS，而非命中通用 READ_FAILED）。</summary>
    internal sealed class DataTooLargeException : Exception
    {
        public DataTooLargeException(string message) : base(message) { }
    }
}