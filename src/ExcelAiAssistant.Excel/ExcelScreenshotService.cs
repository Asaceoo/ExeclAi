using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Security;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 截图能力（AI 视觉闭环的地基）：把工作表/区域/Excel 窗口渲染成图片文件，供模型"看结果"自检。
    ///
    /// 实现口径：
    /// - sheet/range：Range.CopyPicture → 临时 ChartObject 承接 → Chart.Export 落盘，全程不需要窗口可见，
    ///   适合 MCP/CLI 的 Visible=false 无头场景；临时图表在 finally 中删除，不污染工作簿。
    /// - active_window：Win32 PrintWindow 抓 Excel 主窗口（含功能区/编辑区）；窗口不可见时打印不出内容，
    ///   自动降级为 active sheet 截图并给出 warning（不静默返回黑图）。
    ///
    /// 安全：只读行为，不修改单元格；用户可传 file_path，一律 OGuardOutputPath 校验（拒绝 .. / UNC / 网络盘）。
    /// </summary>
    public sealed class ExcelScreenshotService
    {
        /// <summary>单次截图单元格上限：超过则拒绝，避免 CopyPicture 把巨型区域渲染成图片拖垮 Excel。</summary>
        public const int MaxCaptureCells = 200000;

        /// <summary>返回 base64 的体积上限（默认 3MB）：超过则不返回内联数据，只给文件路径，防 MCP 响应爆炸。</summary>
        public const int MaxInlineBase64Bytes = 3 * 1024 * 1024;

        private readonly dynamic application;

        public ExcelScreenshotService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult TakeScreenshot(Dictionary<string, object> args)
        {
            var target = NormalizeTarget(Optional(args, "target"));
            var format = NormalizeFormat(Optional(args, "format"));
            if (format == null)
            {
                return Invalid("INVALID_FORMAT", "format 仅支持 png/jpeg，当前为：" + Optional(args, "format"));
            }

            if (target == null)
            {
                // M29 审查（A6/U1）：未知 target 报错而非静默截整表——拼写错误必须显式暴露
                return Invalid("INVALID_TARGET", "target 仅支持 sheet/range/active_window，当前为："
                    + Optional(args, "target") + "。请修正拼写后重试。");
            }

            var extension = format == "png" ? ".png" : ".jpg";
            var filterName = format == "png" ? "PNG" : "JPG";
            var returnBase64 = PBool(args, "return_base64", false);
            var sheetName = Optional(args, "sheet");
            var address = Optional(args, "address");

            string filePath;
            try
            {
                filePath = ResolveOutputPath(Optional(args, "file_path"), extension);
            }
            catch (Exception pathException)
            {
                return Invalid("INVALID_PATH", "输出路径不合法：" + pathException.Message);
            }

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                try
                {
                    Directory.CreateDirectory(directory);
                }
                catch (Exception createException)
                {
                    return Invalid("INVALID_PATH", "输出目录无法创建：" + createException.Message);
                }
            }

            dynamic worksheet = null;
            dynamic range = null;
            var releaseWorksheet = false;
            var warning = (string)null;
            try
            {
                if (target == "active_window")
                {
                    var windowResult = CaptureActiveWindow(filePath, format, filterName);
                    if (windowResult != null)
                    {
                        return BuildSuccess(windowResult, target, null, null, returnBase64);
                    }

                    // 窗口不可见/抓取失败：降级为活动表截图，并如实告知
                    warning = "Excel 窗口当前不可见（无头模式），active_window 已降级为活动工作表截图。";
                    target = "sheet";
                }

                worksheet = ResolveWorksheet(sheetName, out releaseWorksheet);

                string rangeAddress;
                if (target == "range")
                {
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        return Invalid("INVALID_ARGS", "target=range 时必须提供 address（如 A1:F20）。");
                    }

                    range = worksheet.Range[address];
                    rangeAddress = Convert.ToString(range.Address).Replace("$", string.Empty);
                }
                else
                {
                    range = worksheet.UsedRange;
                    rangeAddress = Convert.ToString(range.Address).Replace("$", string.Empty);
                }

                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                if (rows <= 0 || columns <= 0)
                {
                    return Invalid("EMPTY_RANGE", "目标区域为空，没有可截图的内容。");
                }

                // M29 审查（U1）：Excel 空表 UsedRange 恒为 $A$1（1×1），rows/columns 判空永不命中——
                // 空表白图会被 AI 误读为"内容确实如此"（视觉自检静默假阳性）。
                // 追加内容检测：1×1 且单元格空（Value2 为 null/空串）即判空表，报错并给恢复路径。
                if (rows == 1 && columns == 1)
                {
                    var single = Convert.ToString(range.Value2);
                    if (string.IsNullOrEmpty(single))
                    {
                        return Invalid(
                            "EMPTY_RANGE",
                            "目标工作表没有任何内容（UsedRange 仅为 A1 且为空）。"
                            + "请先用 get_excel_sheet_list 核对表名拼写，或改用 target=range 并提供具体 address。");
                    }
                }

                var cells = (long)rows * columns;
                if (cells > MaxCaptureCells)
                {
                    return Invalid(
                        "RANGE_TOO_LARGE",
                        "目标区域 " + cells + " 格超过截图上限 " + MaxCaptureCells + " 格；请先用 address 缩小到具体区域。");
                }

                ExportRange(worksheet, range, filePath, filterName, rows, columns);
                var shot = new CaptureOutcome
                {
                    FilePath = filePath,
                    Format = format,
                    Sheet = Convert.ToString(worksheet.Name),
                    Address = rangeAddress,
                    Rows = rows,
                    Columns = columns
                };
                return BuildSuccess(shot, target, warning, null, returnBase64);
            }
            catch (COMException comException)
            {
                return Invalid(
                    "SCREENSHOT_FAILED",
                    "截图失败（COM 0x" + unchecked((uint)comException.HResult).ToString("X8") + "）：" + comException.Message
                    + "。若目标表被保护，请先解除保护；无头模式下可改用 target=sheet。");
            }
            catch (Exception exception)
            {
                return Invalid("SCREENSHOT_FAILED", "截图失败：" + exception.Message);
            }
            finally
            {
                // M29 审查（A4）：worksheet/range 均为共享 RCW（ActiveSheet 或同名命中缓存），
                // 不做 FinalRelease（断连会殃及其他持有者），交 GC 按身份回收。
            }
        }

        // ===================== 区域导出（无头，剪贴板 + 临时图表） =====================

        /// <summary>把区域渲染成图片：CopyPicture → 临时 ChartObject.Paste → Chart.Export，finally 删除临时图表。</summary>
        private static void ExportRange(dynamic worksheet, dynamic range, string filePath, string filterName, int rows, int columns)
        {
            dynamic chartObjects = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                // xlScreen=1（按屏幕外观）、xlBitmap=2。Chart.Paste 需要位图格式。
                range.CopyPicture(1, 2);

                chartObjects = worksheet.ChartObjects();
                // 尺寸按区域估算（列宽约 56pt、行高约 18pt），夹在可导出范围内。
                var width = Math.Min(1600, Math.Max(240, columns * 56));
                var height = Math.Min(1200, Math.Max(160, rows * 18));
                chartObject = chartObjects.Add(0, 0, width, height);
                chart = chartObject.Chart;
                chart.Paste();
                chart.Export(filePath, filterName);
            }
            finally
            {
                if (chart != null)
                {
                    try { Marshal.FinalReleaseComObject(chart); } catch { }
                }

                if (chartObject != null)
                {
                    try { chartObject.Delete(); } catch { }
                    try { Marshal.FinalReleaseComObject(chartObject); } catch { }
                }

                if (chartObjects != null)
                {
                    try { Marshal.FinalReleaseComObject(chartObjects); } catch { }
                }
            }
        }

        // ===================== 窗口捕获（PrintWindow） =====================

        private sealed class CaptureOutcome
        {
            public string FilePath;
            public string Format;
            public string Sheet;
            public string Address;
            public int Rows;
            public int Columns;
        }

        /// <summary>抓 Excel 主窗口；不可见或抓取失败返回 null（由调用方降级）。</summary>
        private CaptureOutcome CaptureActiveWindow(string filePath, string format, string filterName)
        {
            try
            {
                var hwnd = new IntPtr(Convert.ToInt64(application.Hwnd));
                if (hwnd == IntPtr.Zero || !NativeMethods.IsWindowVisible(hwnd))
                {
                    return null;
                }

                NativeMethods.RECT bounds;
                if (!NativeMethods.GetWindowRect(hwnd, out bounds))
                {
                    return null;
                }

                var width = bounds.Right - bounds.Left;
                var height = bounds.Bottom - bounds.Top;
                if (width <= 0 || height <= 0)
                {
                    return null;
                }

                using (var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        var hdc = graphics.GetHdc();
                        try
                        {
                            // PW_RENDERFULLCONTENT：让 DirectComposition 渲染的窗口内容也能抓到。
                            if (!NativeMethods.PrintWindow(hwnd, hdc, 2))
                            {
                                return null;
                            }
                        }
                        finally
                        {
                            graphics.ReleaseHdc(hdc);
                        }
                    }

                    SaveBitmap(bitmap, filePath, format);
                }

                return new CaptureOutcome
                {
                    FilePath = filePath,
                    Format = format,
                    Sheet = null,
                    Address = null,
                    Rows = 0,
                    Columns = 0
                };
            }
            catch
            {
                // 窗口抓取属尽力而为：任何失败都交给降级路径，不阻断整体截图。
                return null;
            }
            finally
            {
                // 占位：window 为 dynamic 局部，无本地 RCW 需要释放（ActiveWindow 是共享对象，禁止 FinalRelease）。
            }
        }

        private static void SaveBitmap(Bitmap bitmap, string filePath, string format)
        {
            if (format == "jpeg")
            {
                var encoder = GetEncoder(ImageFormat.Jpeg);
                if (encoder != null)
                {
                    using (var parameters = new EncoderParameters(1))
                    {
                        parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 88L);
                        bitmap.Save(filePath, encoder, parameters);
                        return;
                    }
                }

                bitmap.Save(filePath, ImageFormat.Jpeg);
                return;
            }

            bitmap.Save(filePath, ImageFormat.Png);
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            foreach (var codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }

            return null;
        }

        // ===================== 结果与路径 =====================

        private ToolResult BuildSuccess(CaptureOutcome outcome, string target, string warning, string extra, bool returnBase64)
        {
            if (!File.Exists(outcome.FilePath))
            {
                return Invalid("SCREENSHOT_FAILED", "截图文件未生成：" + outcome.FilePath);
            }

            var info = new FileInfo(outcome.FilePath);
            if (info.Length <= 0)
            {
                try { File.Delete(outcome.FilePath); } catch { }
                return Invalid("SCREENSHOT_FAILED", "截图文件为空（导出失败）：" + outcome.FilePath);
            }

            int imageWidth;
            int imageHeight;
            try
            {
                using (var image = Image.FromFile(outcome.FilePath))
                {
                    imageWidth = image.Width;
                    imageHeight = image.Height;
                }
            }
            catch
            {
                imageWidth = 0;
                imageHeight = 0;
            }

            var data = new Dictionary<string, object>
            {
                ["file_path"] = outcome.FilePath,
                ["format"] = outcome.Format,
                ["bytes"] = info.Length,
                ["width"] = imageWidth,
                ["height"] = imageHeight,
                ["target"] = target,
                ["sheet"] = outcome.Sheet,
                ["address"] = outcome.Address,
                ["rows"] = outcome.Rows,
                ["columns"] = outcome.Columns
            };

            var finalWarning = warning;
            if (returnBase64)
            {
                if (info.Length <= MaxInlineBase64Bytes)
                {
                    try
                    {
                        data["base64"] = Convert.ToBase64String(File.ReadAllBytes(outcome.FilePath));
                        data["mime_type"] = outcome.Format == "png" ? "image/png" : "image/jpeg";
                    }
                    catch (Exception readException)
                    {
                        finalWarning = Join(finalWarning, "读取截图内容失败：" + readException.Message);
                    }
                }
                else
                {
                    finalWarning = Join(
                        finalWarning,
                        "截图体积 " + info.Length + " 字节超过内联上限 " + MaxInlineBase64Bytes + "，未返回 base64（请用 file_path 读取）。");
                }
            }

            return new ToolResult
            {
                Ok = true,
                Data = data,
                Warning = finalWarning
            };
        }

        private static string ResolveOutputPath(string raw, string extension)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                var directory = Path.Combine(Path.GetTempPath(), "JYYJ", "screenshots");
                // M29 审查（S4）：文件名含随机段——纯时间戳可预测，junction/同毫秒竞争可被预置利用
                var fileName = "shot-" + DateTime.Now.ToString("yyyyMMdd-HHmmss")
                    + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + extension;
                return Path.Combine(directory, fileName);
            }

            return PathGuard.GuardAndComplete(raw, extension);
        }

        /// <summary>target 归一化：容忍大小写与常见别名；空缺省 sheet。
        /// M29 审查（A6/U1）：非空未知值返回 null（由调用方报 INVALID_TARGET）——静默回退 sheet 会把
        /// 拼写错误截成整表图，语义漂移且与 schema enum 严格性矛盾。</summary>
        private static string NormalizeTarget(string raw)
        {
            var value = (raw ?? string.Empty).Trim().ToLowerInvariant();
            switch (value)
            {
                case "active_window":
                case "window":
                case "activewindow":
                    return "active_window";
                case "range":
                case "selection":
                    return "range";
                case "sheet":
                case "active_sheet":
                case "worksheet":
                case "":
                    return "sheet";
                default:
                    return null; // 未知值：报错而非静默降级
            }
        }

        private static string NormalizeFormat(string raw)
        {
            var value = (raw ?? string.Empty).Trim().ToLowerInvariant();
            if (value == "" || value == "png")
            {
                return "png";
            }

            if (value == "jpeg" || value == "jpg")
            {
                return "jpeg";
            }

            return null;
        }

        private dynamic ResolveWorksheet(string sheetName, out bool release)
        {
            // M29 审查（A4）：Worksheets[name] 与 ActiveSheet 可能命中同一共享 RCW（显式传活动表名时），
            // FinalRelease 会把其他持有者一并断连（"separated from its underlying RCW"）——统一不强制释放，交 GC 兜底。
            release = false;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return application.ActiveSheet;
            }

            return application.Worksheets[sheetName];
        }

        private static string Join(string left, string right)
        {
            if (string.IsNullOrEmpty(left))
            {
                return right;
            }

            if (string.IsNullOrEmpty(right))
            {
                return left;
            }

            return left + " " + right;
        }

        // ===================== 参数辅助 =====================

        private static string Optional(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var value) && value != null)
            {
                return Convert.ToString(value);
            }

            return null;
        }

        private static bool PBool(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args != null && args.TryGetValue(key, out var value) && value != null
                && bool.TryParse(Convert.ToString(value), out var parsed))
            {
                return parsed;
            }

            return fallback;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(dynamic com)
        {
            if (com == null || !Marshal.IsComObject(com))
            {
                return;
            }

            try { Marshal.FinalReleaseComObject(com); } catch { }
        }

        /// <summary>Win32 窗口捕获（PrintWindow）。</summary>
        internal static class NativeMethods
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

            [DllImport("user32.dll")]
            public static extern bool IsWindowVisible(IntPtr hWnd);
        }
    }
}