using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// P0-2 L2 通用结构操作层（对齐 OfficeCLI 三层抽象的 L2 DOM 层）：
    /// 统一元素模型（cell/range/row/column/sheet/chart）上的 query/get/set/add/remove/move 六操作，
    /// 覆盖专用工具的长尾场景。全部走无状态 COM 约定 + 保护态守卫（写操作经宿主 WrapWriteTools）。
    /// </summary>
    public sealed class ExcelElementOps
    {
        private readonly dynamic application;

        public ExcelElementOps(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        private dynamic GetWorkbook()
        {
            return application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
        }

        private dynamic GetSheet(dynamic workbook, string name)
        {
            if (string.IsNullOrEmpty(name)) return workbook.ActiveSheet;
            try { return workbook.Worksheets[name]; }
            catch { throw new KeyNotFoundException("工作表不存在：" + name); }
        }

        // ===== query_elements：按谓词枚举元素 =====
        public ToolResult QueryElements(Dictionary<string, object> args)
        {
            var elementType = Required(args, "element_type").ToLowerInvariant(); // sheets | shapes | charts | usedrange_rows
            var sheet = GetOptionalString(args, "sheet");
            dynamic workbook = null;
            try
            {
                workbook = GetWorkbook();
                var items = new List<object>();
                if (elementType == "sheets")
                {
                    foreach (var s in workbook.Worksheets)
                    {
                        items.Add(new { type = "sheet", name = Convert.ToString(s.Name), visible = Convert.ToInt32(s.Visible) == -1 });
                    }
                }
                else if (elementType == "shapes")
                {
                    dynamic ws = GetSheet(workbook, sheet);
                    foreach (var sh in ws.Shapes)
                    {
                        items.Add(new { type = "shape", name = Convert.ToString(sh.Name), kind = Convert.ToString(sh.Type), top = Convert.ToDouble(sh.Top), left = Convert.ToDouble(sh.Left) });
                    }
                }
                else if (elementType == "charts")
                {
                    dynamic ws = GetSheet(workbook, sheet);
                    foreach (var ch in ws.ChartObjects())
                    {
                        items.Add(new { type = "chart", name = Convert.ToString(ch.Name), top = Convert.ToDouble(ch.Top), left = Convert.ToDouble(ch.Left) });
                    }
                }
                else if (elementType == "usedrange_rows")
                {
                    dynamic ws = GetSheet(workbook, sheet);
                    dynamic used = ws.UsedRange;
                    int rows = Convert.ToInt32(used.Rows.Count);
                    int cols = Convert.ToInt32(used.Columns.Count);
                    items.Add(new { type = "usedrange", rows = rows, columns = cols, address = Convert.ToString(used.Address) });
                }
                else
                {
                    return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "element_type 支持：sheets / shapes / charts / usedrange_rows" };
                }

                return new ToolResult { Ok = true, Data = new { count = items.Count, elements = items } };
            }
            finally
            {
                ReleaseIfCom(workbook);
            }
        }

        // ===== get_element：读取元素值/属性 =====
        public ToolResult GetElement(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheet = GetOptionalString(args, "sheet");
            dynamic workbook = null;
            dynamic ws = null;
            dynamic range = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                range = ws.Range[address];
                var values = range.Value2;
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                var data = new object[rows, cols];
                if (rows == 1 && cols == 1)
                {
                    // 单单元格：Value2 是标量（string/double 等），无 GetValue
                    data[0, 0] = values;
                }
                else if (values != null)
                {
                    for (int r = 1; r <= rows; r++)
                        for (int c = 1; c <= cols; c++)
                            data[r - 1, c - 1] = values.GetValue(r, c);
                }
                return new ToolResult { Ok = true, Data = new { address = Convert.ToString(range.Address), rows = rows, columns = cols, values = data, formula = rows * cols == 1 ? Convert.ToString(range.Formula) : null } };
            }
            finally
            {
                ReleaseIfCom(range); ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        // ===== set_element：设置元素值 =====
        public ToolResult SetElement(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheet = GetOptionalString(args, "sheet");
            var value = args.TryGetValue("value", out var v) ? v : null;
            dynamic workbook = null;
            dynamic ws = null;
            dynamic range = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                range = ws.Range[address];
                var before = Convert.ToString(range.Value2);
                range.Value2 = value;
                return new ToolResult
                {
                    Ok = true,
                    Data = new { address = Convert.ToString(range.Address), before = before, after = Convert.ToString(range.Value2) },
                    AffectedRanges = new List<string> { Convert.ToString(range.Address) }
                };
            }
            finally
            {
                ReleaseIfCom(range); ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        // ===== add_element：新增元素（sheet 已有专用工具，这里覆盖 shape textbox）=====
        public ToolResult AddElement(Dictionary<string, object> args)
        {
            var kind = Required(args, "kind").ToLowerInvariant(); // textbox
            var sheet = GetOptionalString(args, "sheet");
            dynamic workbook = null;
            dynamic ws = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                if (kind == "textbox")
                {
                    double top = GetOptionalDouble(args, "top", 20);
                    double left = GetOptionalDouble(args, "left", 20);
                    double width = GetOptionalDouble(args, "width", 160);
                    double height = GetOptionalDouble(args, "height", 40);
                    string text = GetOptionalString(args, "text") ?? "";
                    var sh = ws.Shapes.AddTextbox(1, left, top, width, height); // 1 = msoTextOrientationHorizontal
                    sh.TextFrame.Characters().Text = text;
                    return new ToolResult { Ok = true, Data = new { kind = "textbox", name = Convert.ToString(sh.Name) }, AffectedRanges = new List<string>() };
                }
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "kind 支持：textbox（sheet/chart 已有专用工具）" };
            }
            finally
            {
                ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        // ===== remove_element：删除元素 =====
        public ToolResult RemoveElement(Dictionary<string, object> args)
        {
            var kind = Required(args, "kind").ToLowerInvariant(); // shape | chart
            var name = Required(args, "name");
            var sheet = GetOptionalString(args, "sheet");
            dynamic workbook = null;
            dynamic ws = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                if (kind == "shape")
                {
                    dynamic sh = ws.Shapes[name] ?? throw new KeyNotFoundException("形状不存在：" + name);
                    sh.Delete();
                    return new ToolResult { Ok = true, Data = new { removed = name, kind = "shape" } };
                }
                if (kind == "chart")
                {
                    dynamic ch = ws.ChartObjects()[name] ?? throw new KeyNotFoundException("图表不存在：" + name);
                    ch.Delete();
                    return new ToolResult { Ok = true, Data = new { removed = name, kind = "chart" } };
                }
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "kind 支持：shape / chart（工作表删除用 delete_excel_worksheet）" };
            }
            finally
            {
                ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        // ===== move_element：移动元素 =====
        public ToolResult MoveElement(Dictionary<string, object> args)
        {
            var kind = Required(args, "kind").ToLowerInvariant(); // shape | chart
            var name = Required(args, "name");
            var sheet = GetOptionalString(args, "sheet");
            dynamic workbook = null;
            dynamic ws = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                dynamic target = kind == "shape" ? (object)ws.Shapes[name] : (object)ws.ChartObjects()[name];
                if (target == null) throw new KeyNotFoundException("元素不存在：" + name);
                if (args.TryGetValue("top", out var top)) target.Top = Convert.ToDouble(top);
                if (args.TryGetValue("left", out var left)) target.Left = Convert.ToDouble(left);
                return new ToolResult { Ok = true, Data = new { moved = name, top = Convert.ToDouble(target.Top), left = Convert.ToDouble(target.Left) } };
            }
            finally
            {
                ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }


        // ===== P1-3：外部数据源接入 =====
        public ToolResult ImportCsv(Dictionary<string, object> args)
        {
            var filePath = Required(args, "file_path");
            if (!System.IO.File.Exists(filePath))
                return new ToolResult { Ok = false, ErrorCode = "FILE_NOT_FOUND", Data = "文件不存在：" + filePath };
            var sheetName = GetOptionalString(args, "sheet_name") ?? ("CSV_" + System.IO.Path.GetFileNameWithoutExtension(filePath));
            var delimiter = GetOptionalString(args, "delimiter") ?? ",";
            dynamic workbook = null;
            dynamic ws = null;
            try
            {
                workbook = GetWorkbook();
                if (SheetExists(workbook, sheetName))
                    return new ToolResult { Ok = false, ErrorCode = "SHEET_EXISTS", Data = "工作表已存在：" + sheetName };
                ws = (dynamic)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                ws.Name = sheetName;

                // 编码探测（BOM→UTF8 严格→GBK 兜底，与 read_business_glossary 口径一致）
                var bytes = System.IO.File.ReadAllBytes(filePath);
                var enc = DetectEncoding(bytes);
                var text = enc.GetString(bytes);
                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                for (int r = 0; r < lines.Length && r < 100000; r++)
                {
                    var cells = lines[r].Split(new[] { delimiter }, StringSplitOptions.None);
                    for (int c = 0; c < cells.Length && c < 256; c++)
                    {
                        var cell = (dynamic)ws.Cells[r + 1, c + 1];
                        cell.Value2 = cells[c];
                        if (System.Runtime.InteropServices.Marshal.IsComObject(cell)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell);
                    }
                }
                return new ToolResult
                {
                    Ok = true,
                    Data = new { sheet = sheetName, rows = Math.Min(lines.Length, 100000), columns = lines.Length > 0 ? lines[0].Split(new[] { delimiter }, StringSplitOptions.None).Length : 0, encoding = enc.EncodingName },
                    AffectedRanges = new List<string> { "'" + sheetName + "'!A1" }
                };
            }
            finally
            {
                ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        private static System.Text.Encoding DetectEncoding(byte[] bytes)
        {
            // BOM 检测
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new System.Text.UTF8Encoding(false);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return System.Text.Encoding.Unicode;
            // UTF-8 严格解码
            try { var strict = new System.Text.UTF8Encoding(false, true); strict.GetString(bytes); return new System.Text.UTF8Encoding(false); }
            catch { }
            // GBK 兜底
            try { return System.Text.Encoding.GetEncoding(936); } catch { return System.Text.Encoding.Default; }
        }

        private static bool SheetExists(dynamic workbook, string name)
        {
            foreach (var s in workbook.Worksheets)
                if (string.Equals(Convert.ToString(s.Name), name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
        // ===== helpers =====
        private static string Required(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                var s = Convert.ToString(v);
                if (s.Length > 0) return s;
            }
            throw new ArgumentException("缺少必填参数：" + key);
        }

        private static string GetOptionalString(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null) return Convert.ToString(v);
            return null;
        }

        private static double GetOptionalDouble(Dictionary<string, object> args, string key, double def)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null && double.TryParse(Convert.ToString(v), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var d)) return d;
            return def;
        }

        private static void ReleaseIfCom(object o)
        {
            if (o != null && System.Runtime.InteropServices.Marshal.IsComObject(o))
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o); } catch { }
            }
        }
    }
}
