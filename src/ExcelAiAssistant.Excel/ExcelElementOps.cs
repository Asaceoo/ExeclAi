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
            // FIX（M28/R3）：仅标量；矩阵场景给明确指引而非 COM 裸错。
            // 校验必须在 Range 解析之前——区域地址（A1:B2）+ 矩阵值原先先抛 INVALID_ADDRESS，
            // 掩盖了「请改用 write_excel_range」的可执行指引。
            if (value is System.Collections.IEnumerable && !(value is string))
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "set_element 只接受标量值；写二维矩阵请用 write_excel_range 工具（values 参数）。" };
            dynamic workbook = null;
            dynamic ws = null;
            dynamic range = null;
            try
            {
                workbook = GetWorkbook();
                ws = GetSheet(workbook, sheet);
                range = ws.Range[address];
                int seRows = Convert.ToInt32(range.Rows.Count);
                int seCols = Convert.ToInt32(range.Columns.Count);
                var before = Convert.ToString(range.Value2);
                range.Value2 = value;
                return new ToolResult
                {
                    Ok = true,
                    Data = new { address = Convert.ToString(range.Address), rows = seRows, columns = seCols, before = before, after = Convert.ToString(range.Value2) },
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
                    // FIX（审查 B-3）：数值非法显式报错，不再静默回退默认值（原 top="abc" 会 ok:true 但位置不对，误导 AI）
                    var okTop = TryDoubleArg(args, "top", 20, out var top, out var topErr);
                    var okLeft = TryDoubleArg(args, "left", 20, out var left, out var leftErr);
                    var okWidth = TryDoubleArg(args, "width", 160, out var width, out var widthErr);
                    var okHeight = TryDoubleArg(args, "height", 40, out var height, out var heightErr);
                    if (!(okTop && okLeft && okWidth && okHeight))
                    {
                        return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = topErr ?? leftErr ?? widthErr ?? heightErr };
                    }
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
                    dynamic sh;
                    // FIX（审查 D-1）：COM 索引器先抛异常，原 ?? throw 是死代码 → 错误被外层记成 INVALID_ARGS
                    try { sh = ws.Shapes[name]; }
                    catch { return new ToolResult { Ok = false, ErrorCode = "ELEMENT_NOT_FOUND", Data = "形状不存在：" + name + "。请先调 query_elements(element_type=shapes) 查看现有形状名。" }; }
                    sh.Delete();
                    return new ToolResult { Ok = true, Data = new { removed = name, kind = "shape" } };
                }
                if (kind == "chart")
                {
                    dynamic ch;
                    try { ch = ws.ChartObjects()[name]; }
                    catch { return new ToolResult { Ok = false, ErrorCode = "ELEMENT_NOT_FOUND", Data = "图表不存在：" + name + "。请先调 query_elements(element_type=charts) 查看现有图表名。" }; }
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
                if (kind != "shape" && kind != "chart")
                    return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "kind 支持：shape / chart" };
                dynamic target;
                try { target = kind == "shape" ? (object)ws.Shapes[name] : (object)ws.ChartObjects()[name]; }
                catch { return new ToolResult { Ok = false, ErrorCode = "ELEMENT_NOT_FOUND", Data = "元素不存在：" + name + "。请先调 query_elements 查看现有名称。" }; }
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

            // FIX（审查 C2）：大小上限 50MB（对照 ExcelImportService 口径），防 OOM 与超大文件灌入
            var fileInfo = new System.IO.FileInfo(filePath);
            if (fileInfo.Length > 50L * 1024 * 1024)
                return new ToolResult { Ok = false, ErrorCode = "FILE_TOO_LARGE", Data = "CSV 文件超过 50MB 上限，请拆分后导入。" };

            // FIX（审查 C5）：表名清洗——Excel 上限 31 字符且禁 [ ] : * ? / \
            var rawName = GetOptionalString(args, "sheet_name") ?? ("CSV_" + System.IO.Path.GetFileNameWithoutExtension(filePath));
            var sheetName = rawName;
            foreach (var bad in new[] { '[', ']', ':', '*', '?', '/', '\\' }) sheetName = sheetName.Replace(bad, '_');
            if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
            if (sheetName.Length == 0) sheetName = "CSV_import";
            var delimiter = GetOptionalString(args, "delimiter") ?? ",";
            // FIX（M28/R3）：多字符 delimiter 原先静默取首字符切分（"::" 实际按 ":" 切且用户不知情），
            // 改为显式拒绝——静默截断语义比报错更危险。
            if (delimiter.Length != 1)
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "delimiter 仅支持单字符分隔符（当前：\"" + delimiter + "\"）。多字符分隔请先转换源文件。" };

            // FIX（M28/R3）：解析与统计前移到建表之前——矩阵超限早退时不残留空工作表（副作用后置）。
            // 编码探测（BOM→UTF8 严格→GBK 兜底，与 read_business_glossary 口径一致）
            var bytes = System.IO.File.ReadAllBytes(filePath);
            var enc = DetectEncoding(bytes);
            var text = enc.GetString(bytes);
            // FIX（审查 C4/BOM/性能）：去 BOM；NumberFormat="@" 前置防公式注入与日期强转（CLAUDE.md §7）；
            // 整块 object[,] 一次写（原逐格 COM 最坏 10万×256=2560万次调用会挂死）
            if (text.Length > 0 && text[0] == '\uFEFF') text = text.Substring(1);
            // FIX（M28/R3）：记录切分改状态机（SplitCsvRecords）——引号内的换行属于字段内容，
            // 原按 \n 文本 Split 会把 "a\nb" 拆成两条残缺记录（跨行引号字段数据损坏）。
            var records = SplitCsvRecords(text, delimiter, out var unbalancedQuotes);
            if (unbalancedQuotes)
            {
                // M28 审查：引号不闭合会把剩余全文吞进一条记录（静默数据损坏），必须 fail-closed
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "INVALID_ARGS",
                    Data = "CSV 引号未闭合（存在未配对的 \"），疑似文件损坏。请检查源文件或在源端转义后再导入。"
                };
            }

            int maxRows = Math.Min(records.Count, 100000);
            int totalMaxCols = 0;
            for (int r = 0; r < maxRows; r++)
            {
                var n = SplitCsvLine(records[r], delimiter).Length; // FIX（M27R2）：引号感知统计，裸 Split 会把 "a,b" 数成两列
                if (n > totalMaxCols) totalMaxCols = n;
            }
            var truncatedColumns = totalMaxCols > 256; // FIX（M28/R3）：列截断进结果标志，不再静默丢列
            int maxCols = Math.Min(totalMaxCols, 256);
            // FIX（M28/R3）：内存放大上限——rows×cols 乘积超限直接拒绝（10万×256=2560 万格 object[,] 数百 MB），
            // 且 50MB 文件体量上限挡不住「宽表×长表」组合。
            if ((long)maxRows * maxCols > 4000000L)
                return new ToolResult { Ok = false, ErrorCode = "MATRIX_TOO_LARGE", Data = "CSV 解析矩阵超过上限（rows×cols ≤ 4,000,000，当前 " + maxRows + "×" + maxCols + "）。请拆分文件或减少列数后重试。" };

            dynamic workbook = null;
            dynamic ws = null;
            try
            {
                workbook = GetWorkbook();
                if (SheetExists(workbook, sheetName))
                    return new ToolResult { Ok = false, ErrorCode = "SHEET_EXISTS", Data = "工作表已存在：" + sheetName };
                ws = (dynamic)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                ws.Name = sheetName;

                if (maxRows > 0 && maxCols > 0)
                {
                    var block = new object[maxRows, maxCols];
                    for (int r = 0; r < maxRows; r++)
                    {
                        var cells = SplitCsvLine(records[r], delimiter); // FIX（M27R2）：恢复引号感知切分（此前重放时退回裸 Split）
                        for (int c = 0; c < maxCols; c++) block[r, c] = c < cells.Length ? cells[c] : null;
                    }
                    dynamic target = (dynamic)ws.Range[ws.Cells[1, 1], ws.Cells[maxRows, maxCols]];
                    target.NumberFormat = "@";
                    target.Value2 = block;
                }
                return new ToolResult
                {
                    Ok = true,
                    Truncated = records.Count > maxRows || truncatedColumns, // M28 审查：列截断同样属 Truncated，防消费方漏看
                    Data = new { sheet = sheetName, rows = maxRows, columns = maxCols, truncated_columns = truncatedColumns, encoding = enc.EncodingName },
                    // FIX（审查 E5）：AffectedRanges 报完整覆盖区（原只报 A1 不足以支撑核对与回滚）
                    AffectedRanges = new List<string> { "'" + sheetName.Replace("'", "''") + "'!A1:" + ColLetter(maxCols) + maxRows }
                };
            }
            finally
            {
                ReleaseIfCom(ws); ReleaseIfCom(workbook);
            }
        }

        /// <summary>
        /// CSV 记录切分（RFC4180 子集，M28/R3 修复）：状态机扫全文本——引号内的 CR/LF 属于字段内容，
        /// 不结束记录；引号外遇 CR/LF 结束当前记录（CRLF 视为一个换行）。输出保留引号原样
        /// （含 "" 转义），后续仍交 SplitCsvLine 做字段级解析——两层状态机职责分离，不重复剥引号。
        /// 中间空行保留为空记录（对应一行空数据）；文本末尾的换行不产生多余空记录。
        /// </summary>
        private static List<string> SplitCsvRecords(string text)
        {
            return SplitCsvRecords(text, ",", out _);
        }

        private static List<string> SplitCsvRecords(string text, out bool unbalancedQuotes)
        {
            return SplitCsvRecords(text, ",", out unbalancedQuotes);
        }

        /// <summary>
        /// 记录级状态机（M28 审查修正）：只有「字段起始位」（记录开头或分隔符之后）的引号才开启引号段，
        /// 与 SplitCsvLine 的字段级规则一致——字段中间的裸引号（英寸号/人名等）按字面处理，
        /// 不会被误当成跨行字段起点而吞掉后续记录。引号内 CR/LF 仍保留；未闭合时 out 标记供 fail-closed。
        /// </summary>
        private static List<string> SplitCsvRecords(string text, string delimiter, out bool unbalancedQuotes)
        {
            var records = new List<string>();
            var cur = new System.Text.StringBuilder();
            char sep = string.IsNullOrEmpty(delimiter) ? ',' : delimiter[0];
            bool inQuotes = false;
            bool atFieldStart = true;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { cur.Append('"'); cur.Append('"'); i++; }
                        else { inQuotes = false; cur.Append('"'); atFieldStart = false; }
                    }
                    else
                    {
                        cur.Append(ch); // 含 CR/LF：跨行引号字段的核心场景
                    }
                }
                else if (ch == '"' && atFieldStart)
                {
                    inQuotes = true;
                    cur.Append('"');
                }
                else if (ch == '\r' || ch == '\n')
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') { i++; } // CRLF 合一
                    records.Add(cur.ToString());
                    cur.Clear();
                    atFieldStart = true;
                }
                else
                {
                    cur.Append(ch);
                    atFieldStart = (ch == sep);
                }
            }

            if (cur.Length > 0)
            {
                records.Add(cur.ToString()); // 尾部无换行的最后一条记录；以换行结尾时不补空记录
            }

            unbalancedQuotes = inQuotes;
            return records;
        }

        /// <summary>CSV 行切分（RFC4180 子集）：引号内分隔符不切、"" 转义、引号剥离。</summary>
        private static string[] SplitCsvLine(string line, string delimiter)
        {
            if (string.IsNullOrEmpty(delimiter)) { return new[] { line }; }
            char sep = delimiter[0];
            var fields = new List<string>();
            var cur = new System.Text.StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; }
                        else { inQuotes = false; }
                    }
                    else { cur.Append(ch); }
                }
                else if (ch == '"' && cur.Length == 0) { inQuotes = true; }
                else if (ch == sep) { fields.Add(cur.ToString()); cur.Clear(); }
                else { cur.Append(ch); }
            }
            fields.Add(cur.ToString());
            return fields.ToArray();
        }

        private static string ColLetter(int c)
        {
            var s = "";
            while (c > 0) { var m = (c - 1) % 26; s = (char)('A' + m) + s; c = (c - 1) / 26; }
            return s;
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
        private static bool TryDoubleArg(Dictionary<string, object> args, string key, double def, out double value, out string error)
        {
            error = null;
            value = def;
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                var s = Convert.ToString(v);
                if (s.Length > 0)
                {
                    if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value)) return true;
                    error = "参数 " + key + " 需为数值，收到：" + s;
                    return false;
                }
            }
            return true;
        }



    }
}
