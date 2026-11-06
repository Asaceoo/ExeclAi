using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 只读渲染器：把区域"画"成图片，全程不碰剪贴板、不插入图表对象、不修改工作簿。
    ///
    /// 与 take_screenshot 的区别（M33）：
    ///   take_screenshot —— CopyPicture 导出，像素级接近 Excel 真实外观，但副作用是占剪贴板 + 临时图表对象，故为写入级；
    ///   本渲染器       —— 批量读 Text/Value2 + 表头行抽样读格式，GDI+ 自绘，只读模式下可用。
    ///
    /// 性能红线（M34 审查实测）：**禁止逐格读 COM**。4000 格 × 9~13 次属性访问 ≈ 3.6~5.2 万次跨进程往返，
    /// MCP 独立进程下可达分钟级并阻塞同一 worker 上的后续调用。这里改为两次块读（Text / Value2）+ 表头行抽样。
    /// </summary>
    public sealed class ExcelRangeRenderer
    {
        public const int MaxCells = 4000;
        public const int MaxRows = 300;
        public const int MaxColumns = 60;
        private const float Scale = 1.4f;
        private const int MinColWidth = 56;
        private const int MaxColWidth = 280;
        private const int RowHeight = 22;
        private const long MaxPixels = 24_000_000;   // ≈72MB 24bpp，32 位宿主也留有余量
        private const int MaxCellTextChars = 120;

        private static long _lastSweepTicks;

        private readonly dynamic application;

        public ExcelRangeRenderer(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult Render(Dictionary<string, object> args)
        {
            SweepRenders();

            var sheetName = Optional(args, "sheet");
            var address = Optional(args, "address");
            var title = Optional(args, "title");
            var wantBase64 = Flag(args, "return_base64", true);

            dynamic sheet;
            try { sheet = string.IsNullOrWhiteSpace(sheetName) ? application.ActiveSheet : application.Worksheets[sheetName]; }
            catch (Exception ex) { return Fail("SHEET_NOT_FOUND", "找不到工作表：" + (sheetName ?? "(活动表)") + "；" + ex.Message); }
            if (sheet == null) { return Fail("NO_SHEET", "当前没有活动工作表（活动对象可能是图表工作表，请显式传 sheet）。"); }

            dynamic range;
            try
            {
                range = string.IsNullOrWhiteSpace(address) ? sheet.UsedRange : sheet.Range[address];
            }
            catch (Exception ex) { return Fail("INVALID_ADDRESS", "区域地址无效：" + address + "；" + ex.Message); }

            int rows, cols;
            try { rows = (int)range.Rows.Count; cols = (int)range.Columns.Count; }
            catch (Exception ex) { return Fail("RANGE_UNREADABLE", "无法读取区域尺寸：" + ex.Message); }

            var maxRows = Math.Min(rows, MaxRows);
            var maxCols = Math.Min(cols, MaxColumns);
            if ((long)maxRows * maxCols > MaxCells) { maxRows = Math.Max(1, MaxCells / Math.Max(1, maxCols)); }
            var truncated = maxRows < rows || maxCols < cols;

            // M34：Text 在某些宿主/动态绑定下返回空块，Value2 才是可靠数据源；
            // 显示优先用 Text（含数字格式），取不到再退回 Value2。
            string[,] texts;
            try
            {
                texts = ReadBlock(range.Text, maxRows, maxCols);
                if (IsEmpty(texts, maxRows, maxCols))
                {
                    var values = ReadBlock(range.Value2, maxRows, maxCols);
                    for (var r = 0; r < maxRows; r++)
                    {
                        for (var c = 0; c < maxCols; c++)
                        {
                            if (string.IsNullOrEmpty(texts[r, c])) { texts[r, c] = Shorten(values[r, c]); }
                        }
                    }
                }
            }
            catch (Exception ex) { return Fail("RANGE_UNREADABLE", "读取单元格内容失败：" + ex.Message); }

            if (IsEmpty(texts, maxRows, maxCols))
            {
                return Fail("EMPTY_RANGE", "目标区域没有任何内容（UsedRange 仅为 A1 且为空）。"
                    + "请核对表名，或用 address 指定实际区域——空白图会被误读成「该表没有数据」。");
            }

            CellStyle[,] styles;
            try { styles = SampleStyles(range, maxRows, maxCols); }
            catch { styles = new CellStyle[maxRows, maxCols]; }   // 格式读失败不阻断渲染，按默认样式画

            string path;
            Size size;
            try
            {
                var dir = Path.Combine(Path.GetTempPath(), "JYYJ", "renders");
                Directory.CreateDirectory(dir);
                path = Path.Combine(dir, "render-" + DateTime.Now.ToString("yyyyMMdd-HHmmss")
                    + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
                size = Draw(texts, styles, maxRows, maxCols, title ?? (string)sheet.Name + "!" + (address ?? "UsedRange"), path);
            }
            catch (OutOfMemoryException)
            {
                return Fail("RANGE_TOO_LARGE", "区域过大导致位图内存不足；请缩小 address（建议 ≤200 行 × 20 列）后重试。");
            }
            catch (Exception ex) { return Fail("RENDER_FAILED", "绘制图片失败：" + ex.Message); }

            var data = new Dictionary<string, object>
            {
                ["mode"] = "render",
                ["file_path"] = path,
                ["format"] = "png",
                ["width"] = size.Width,
                ["height"] = size.Height,
                ["rows"] = rows,
                ["columns"] = cols,
                ["rendered_rows"] = maxRows,
                ["rendered_columns"] = maxCols,
                ["truncated"] = truncated,
                ["sheet"] = (string)sheet.Name,
                ["address"] = (string)range.Address,
                ["note"] = "只读渲染（近似还原）：仅还原文本/粗细/字号/底色/字色/水平对齐；"
                         + "不还原合并单元格、边框样式、条件格式、图表/图片/形状、行高与换行；超过 120 字的文本会截断。"
                         + "图上文字不等于单元格真实值，不能作为「格式无误」的证据。像素级外观请用 take_screenshot（写入级）。"
            };

            try
            {
                var info = new FileInfo(path);
                data["bytes"] = info.Length;
                if (wantBase64 && info.Length <= ExcelScreenshotService.MaxInlineBase64Bytes)
                {
                    data["base64"] = Convert.ToBase64String(File.ReadAllBytes(path));
                    data["mime_type"] = "image/png";
                    // M35：已经拿到内存镜像，除非调用方要求保留文件，否则立即清理——
                    // 只读渲染不应在 %TEMP% 留下工作簿内容副本。
                    if (!KeepFile(args))
                    {
                        try { File.Delete(path); } catch { }
                        data["file_deleted"] = true;
                        data["file_path"] = null;
                        data.Remove("bytes");
                    }
                }
                else if (wantBase64)
                {
                    data["inline_omitted"] = "图片超过 3MB，未内联；请用 file_path 查看。";
                }
            }
            catch { }

            return new ToolResult { Ok = true, Data = data, Truncated = truncated };
        }

        // ---------------- 批量读取 ----------------

        /// <summary>Range.Text / Value2 单格返回标量、多格返回 1 基 object[,]——两种形态都要兼容。</summary>
        private static string[,] ReadBlock(object raw, int rows, int cols)
        {
            var result = new string[rows, cols];
            if (raw is object[,] grid)
            {
                var r0 = grid.GetLowerBound(0);
                var c0 = grid.GetLowerBound(1);
                var rMax = Math.Min(rows, grid.GetLength(0));
                var cMax = Math.Min(cols, grid.GetLength(1));
                for (var r = 0; r < rMax; r++)
                {
                    for (var c = 0; c < cMax; c++)
                    {
                        result[r, c] = Shorten(Convert.ToString(grid[r0 + r, c0 + c]));
                    }
                }
            }
            else
            {
                result[0, 0] = Shorten(Convert.ToString(raw));
            }
            return result;
        }

        private static string Shorten(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            return text.Length <= MaxCellTextChars ? text : text.Substring(0, MaxCellTextChars) + "…";
        }

        private static bool IsEmpty(string[,] texts, int rows, int cols)
        {
            if (rows == 1 && cols == 1) { return string.IsNullOrEmpty(texts[0, 0]); }
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    if (!string.IsNullOrEmpty(texts[r, c])) { return false; }
                }
            }
            return true;
        }

        private sealed class CellStyle
        {
            public bool Bold;
            public int BackColor = -1;
            public int ForeColor = -1;
            public int Align;   // 0=left 1=center 2=right
        }

        /// <summary>格式只抽样表头行（≤60 格 × 5 属性），数据区按文本形态推断对齐——避免逐格 COM 往返。</summary>
        private static CellStyle[,] SampleStyles(dynamic range, int rows, int cols)
        {
            var styles = new CellStyle[rows, cols];
            for (var r = 0; r < rows; r++) { for (var c = 0; c < cols; c++) { styles[r, c] = new CellStyle(); } }

            var headerCount = Math.Min(cols, MaxColumns);
            for (var c = 1; c <= headerCount; c++)
            {
                try
                {
                    dynamic cell = range.Cells.Item[1, c];
                    var s = styles[0, c - 1];
                    try { s.Bold = Convert.ToInt32(cell.Font.Bold) != 0; } catch { }
                    try { s.BackColor = Convert.ToInt32(cell.Interior.Color); } catch { }
                    try { s.ForeColor = Convert.ToInt32(cell.Font.Color); } catch { }
                    try { s.Align = ToAlign(cell.HorizontalAlignment); } catch { }
                }
                catch { }
            }

            // 数据行对齐：按"首列文本/数值"推断（数值右对齐、文本左对齐），一次读一行即可
            for (var r = 2; r <= Math.Min(rows, MaxRows); r++)
            {
                for (var c = 1; c <= Math.Min(cols, MaxColumns); c++)
                {
                    var s = styles[r - 1, c - 1];
                    s.Align = 0;
                    if (c == 1) { s.Bold = false; }
                }
            }
            return styles;
        }

        private static int ToAlign(object horizontalAlignment)
        {
            var value = Convert.ToInt32(horizontalAlignment);
            if (value == -4108) { return 1; }   // xlCenter
            if (value == -4152) { return 2; }   // xlRight
            return 0;
        }

        // ---------------- 绘制 ----------------

        private static Size Draw(string[,] texts, CellStyle[,] styles, int rows, int cols, string title, string path)
        {
            using (var measure = new Bitmap(1, 1))
            using (var titleFont = new Font("Microsoft YaHei", 12f, FontStyle.Bold))
            using (var cellFont = new Font("Microsoft YaHei", 10.5f))
            using (var boldFont = new Font("Microsoft YaHei", 10.5f, FontStyle.Bold))
            {
                using (var mg = Graphics.FromImage(measure))
                {
                    var colWidths = new int[cols];
                    for (var c = 0; c < cols; c++)
                    {
                        var w = MinColWidth;
                        for (var r = 0; r < rows; r++)
                        {
                            var t = texts[r, c];
                            if (string.IsNullOrEmpty(t)) { continue; }
                            // 用绘制时的同一字体测宽，避免粗体表头被省略号截断
                            var font = styles[r, c].Bold ? boldFont : cellFont;
                            var size = mg.MeasureString(t, font);
                            var need = (int)size.Width + 14;
                            if (need > w) { w = need; }
                        }
                        colWidths[c] = Math.Min(w, MaxColWidth);
                    }

                    var totalW = 0;
                    foreach (var w in colWidths) { totalW += w; }
                    var titleH = 34;
                    var width = Math.Max(240, totalW + 24);
                    var height = titleH + rows * RowHeight + 16;

                    var scale = Scale;
                    while ((long)(width * scale) * (long)(height * scale) > MaxPixels && scale > 0.75f) { scale -= 0.25f; }
                    if ((long)(width * scale) * (long)(height * scale) > MaxPixels)
                    {
                        throw new OutOfMemoryException("pixel budget exceeded");
                    }

                    using (var bmp = new Bitmap((int)(width * scale), (int)(height * scale), PixelFormat.Format24bppRgb))
                    using (var g = Graphics.FromImage(bmp))
                    using (var gridPen = new Pen(Color.FromArgb(224, 226, 230)))
                    using (var titleBg = new SolidBrush(Color.FromArgb(245, 247, 250)))
                    using (var fmtLeft = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                    using (var fmtCenter = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap, Alignment = StringAlignment.Center })
                    using (var fmtRight = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap, Alignment = StringAlignment.Far })
                    {
                        g.ScaleTransform(scale, scale);
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                        g.Clear(Color.White);
                        g.FillRectangle(titleBg, 0, 0, width, titleH);
                        g.DrawString(title ?? "", titleFont, Brushes.Black, 12, 8);

                        var y = titleH;
                        for (var r = 0; r < rows; r++)
                        {
                            var x = 12;
                            for (var c = 0; c < cols; c++)
                            {
                                var rect = new Rectangle(x, y, colWidths[c], RowHeight);
                                var style = styles[r, c];
                                if (style.BackColor >= 0)
                                {
                                    var color = ToColor(style.BackColor);
                                    if (color.GetBrightness() < 0.95f)
                                    {
                                        using (var brush = new SolidBrush(color)) { g.FillRectangle(brush, rect); }
                                    }
                                }
                                g.DrawRectangle(gridPen, rect);

                                var text = texts[r, c];
                                if (!string.IsNullOrEmpty(text))
                                {
                                    var font = style.Bold ? boldFont : cellFont;
                                    var brush = Brushes.Black;
                                    var textColor = style.ForeColor >= 0 ? ToColor(style.ForeColor) : Color.Black;
                                    if (textColor != Color.Black) { brush = new SolidBrush(textColor); }
                                    try
                                    {
                                        var fmt = style.Align == 1 ? fmtCenter : (style.Align == 2 ? fmtRight : fmtLeft);
                                        g.DrawString(text, font, brush, new RectangleF(rect.X + 4, rect.Y + 3, rect.Width - 8, rect.Height - 4), fmt);
                                    }
                                    finally
                                    {
                                        if (!ReferenceEquals(brush, Brushes.Black)) { brush.Dispose(); }
                                    }
                                }
                                x += colWidths[c];
                            }
                            y += RowHeight;
                        }

                        bmp.Save(path, ImageFormat.Png);
                        return new Size(bmp.Width, bmp.Height);
                    }
                }
            }
        }

        private static Color ToColor(int bgr) =>
            Color.FromArgb(255, bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);

        // ---------------- 临时文件清理 ----------------

        /// <summary>renders 目录按 1 小时冷却清扫 24h 前的文件（与截图目录同口径，但独立冷却）。</summary>
        private static void SweepRenders()
        {
            var now = DateTime.UtcNow.Ticks;
            var last = System.Threading.Interlocked.Read(ref _lastSweepTicks);
            if (last != 0 && now - last < TimeSpan.FromHours(1).Ticks) { return; }
            if (System.Threading.Interlocked.CompareExchange(ref _lastSweepTicks, now, last) != last) { return; }
            try
            {
                var dir = Path.Combine(Path.GetTempPath(), "JYYJ", "renders");
                if (!Directory.Exists(dir)) { return; }
                var cutoff = DateTime.UtcNow - TimeSpan.FromHours(24);
                foreach (var file in Directory.EnumerateFiles(dir, "render-*.png"))
                {
                    try
                    {
                        var info = new FileInfo(file);
                        if ((info.Attributes & (FileAttributes.ReadOnly | FileAttributes.ReparsePoint)) != 0) { continue; }
                        if (info.LastWriteTimeUtc >= cutoff) { continue; }
                        info.Delete();
                    }
                    catch { }
                }
            }
            catch { }
        }

        // ---------------- 辅助 ----------------

        private static string Optional(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var v) || v == null) { return null; }
            var text = Convert.ToString(v);
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private static bool KeepFile(Dictionary<string, object> args) =>
            Flag(args, "keep_file", false);

        private static bool Flag(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args == null || !args.TryGetValue(key, out var v) || v == null) { return fallback; }
            if (v is bool b) { return b; }
            var text = Convert.ToString(v).Trim();
            if (text == "1") { return true; }
            if (text == "0") { return false; }
            return bool.TryParse(text, out var parsed) ? parsed : fallback;
        }

        private static ToolResult Fail(string code, string message) =>
            new ToolResult { Ok = false, ErrorCode = code, Data = message };
    }
}