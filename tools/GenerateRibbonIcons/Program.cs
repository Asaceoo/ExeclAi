using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace GenerateRibbonIcons
{
    /// <summary>
    /// 为 JYYJ助手 Ribbon 全功能生成统一扁平风 PNG 图标：
    /// 96×96、分类配色圆角底 + 白色剪影，输出到 ../src/ExcelAiAssistant.Addin/RibbonIcons/{controlId}.png。
    /// 分类色：打开面板=蓝、视图=绿、数据处理=蓝、文本处理=紫、常用=橙、更多=红。
    /// </summary>
    internal static class Program
    {
        private static readonly int N = 96;
        private static readonly string OutDir = @"D:\execl1.0\src\ExcelAiAssistant.Addin\RibbonIcons";

        private static Color Blue = Color.FromArgb(47, 128, 237);
        private static Color Green = Color.FromArgb(39, 174, 96);
        private static Color Purple = Color.FromArgb(155, 81, 224);
        private static Color Orange = Color.FromArgb(242, 153, 74);
        private static Color Red = Color.FromArgb(235, 87, 87);
        private static Color White = Color.White;

        private static int Main()
        {
            Directory.CreateDirectory(OutDir);
            var icons = new (string key, string file, Color color, Action<Graphics> draw)[]
            {
                ("ToggleTaskPane", "ToggleTaskPane", Blue, DrawPanel),
                ("ToggleSpotlight", "ToggleSpotlight", Green, DrawSpotlight),
                ("TogglePointer", "TogglePointer", Green, DrawPointer),
                ("JQClean", "JQClean", Blue, DrawSparkle),
                ("JQDedup", "JQDedup", Blue, DrawDedup),
                ("JQSort", "JQSort", Blue, DrawSort),
                ("JQTranspose", "JQTranspose", Blue, DrawTranspose),
                ("JQFilter", "JQFilter", Blue, DrawFilter),
                ("JQPivot", "JQPivot", Blue, DrawPivot),
                ("JQTextSplit", "JQTextSplit", Purple, DrawSplit),
                ("JQTextMerge", "JQTextMerge", Purple, DrawMerge),
                ("JQTextClean", "JQTextClean", Purple, DrawClean),
                ("JQFillFormula", "JQFillFormula", Purple, g => DrawText(g, "fx")),
                ("JQInsertText", "JQInsertText", Orange, DrawInsertText),
                ("JQDelText", "JQDelText", Orange, DrawDelText),
                ("JQRandNum", "JQRandNum", Orange, DrawDie),
                ("JQSequence", "JQSequence", Orange, DrawSequence),
                ("JQInsertBlank", "JQInsertBlank", Orange, DrawInsertBlank),
                ("JQDelBlank", "JQDelBlank", Orange, DrawDelBlank),
                ("JQFillDown", "JQFillDown", Orange, DrawFillDown),
                ("JQMergeDup", "JQMergeDup", Orange, DrawMergeDup),
                ("JQMacroBox", "JQMacroBox", Red, DrawMacroBox),
                ("JQFuncLib", "JQFuncLib", Red, DrawFuncLib),
                ("JQSqlQuery", "JQSqlQuery", Red, DrawSql)
            };

            var errors = new List<string>();
            foreach (var ic in icons)
            {
                try
                {
                    Render(ic.file, ic.color, ic.draw);
                    Console.WriteLine("OK  " + ic.file);
                }
                catch (Exception ex)
                {
                    errors.Add($"{ic.file}: {ex.Message}");
                    Console.WriteLine("ERR " + ic.file + " " + ex.Message);
                }
            }

            Console.WriteLine(errors.Count == 0
                ? $"完成：生成 {icons.Length} 个图标 -> {OutDir}"
                : "存在失败：" + string.Join(" | ", errors));
            return errors.Count == 0 ? 0 : 2;
        }

        private static void Render(string file, Color color, Action<Graphics> draw)
        {
            using var bmp = new Bitmap(N, N, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            DrawBackground(g, color);
            using (var clip = RoundedPath(6, 6, N - 12, N - 12, 22))
            {
                g.SetClip(clip);
                draw(g);
                g.ResetClip();
            }
            bmp.Save(Path.Combine(OutDir, file + ".png"), ImageFormat.Png);
        }

        private static void DrawBackground(Graphics g, Color color)
        {
            var top = Lighten(color, 0.28f);
            using var lg = new LinearGradientBrush(new Rectangle(0, 0, N, N), top, color, 90f);
            using var path = RoundedPath(4, 4, N - 8, N - 8, 24);
            g.FillPath(lg, path);
            // 底部细微纵深
            using var shade = new LinearGradientBrush(new Rectangle(0, N / 2, N, N / 2),
                Color.FromArgb(0, Color.Black), Color.FromArgb(40, Color.Black), 90f);
            g.FillPath(shade, path);
        }

        private static Color Lighten(Color c, float f)
        {
            return Color.FromArgb(
                255,
                (int)(c.R + (255 - c.R) * f),
                (int)(c.G + (255 - c.G) * f),
                (int)(c.B + (255 - c.B) * f));
        }

        private static GraphicsPath RoundedPath(float x, float y, float w, float h, float r)
        {
            var p = new GraphicsPath();
            p.AddArc(x, y, r * 2, r * 2, 180, 90);
            p.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
            p.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
            p.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
            p.CloseFigure();
            return p;
        }

        private static Pen P(float w = 6f) { return new Pen(White, w); }
        private static Brush B() { return new SolidBrush(White); }
        private static Font F(float em) { return new Font("Segoe UI Semibold", em, FontStyle.Bold, GraphicsUnit.Pixel); }
        private static Pen RoundCap(Pen pen) { pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round; return pen; }

        private static void DrawText(Graphics g, string s)
        {
            var f = F(28);
            var sz = g.MeasureString(s, f);
            g.DrawString(s, f, B(), (N - sz.Width) / 2, (N - sz.Height) / 2 - 2);
        }

        // ===== 打开面板：右半面板 + 左侧列表 =====
        private static void DrawPanel(Graphics g)
        {
            using var pen = P(6);
            var r = new Rectangle(24, 26, 48, 44);
            RoundedRect(g, pen, r.X, r.Y, r.Width, r.Height, 8);
            g.DrawLine(pen, 48, 28, 48, 68);
            for (var i = 0; i < 3; i++)
                g.DrawLine(pen, 28 + (float)(i * 0), 36 + i * 10, 44, 36 + i * 10);
            g.DrawLine(pen, 52, 40, 68, 40);
            g.DrawLine(pen, 52, 52, 68, 52);
        }

        // ===== 聚光灯：十字准星 =====
        private static void DrawSpotlight(Graphics g)
        {
            using var pen = P(6);
            g.DrawLine(pen, 26, 48, 70, 48);
            g.DrawLine(pen, 48, 26, 48, 70);
            using var rp = P(6);
            g.DrawEllipse(RoundCap(rp), 39, 39, 18, 18);
        }

        // ===== 指针：斜向光标箭头 =====
        private static void DrawPointer(Graphics g)
        {
            var pts = new[]
            {
                new Point(66, 26), new Point(52, 60), new Point(36, 48),
                new Point(28, 66), new Point(22, 60), new Point(32, 40)
            };
            g.FillPolygon(B(), pts);
        }

        // ===== 清洗：光点/闪光 =====
        private static void DrawSparkle(Graphics g)
        {
            var c = new PointF(48, 46);
            g.FillEllipse(B(), c.X - 9, c.Y - 9, 18, 18);
            using var pen = P(6);
            g.DrawLine(pen, 48, 66, 48, 76);
            g.DrawLine(pen, 38, 58, 30, 52);
            g.DrawLine(pen, 58, 58, 66, 52);
        }

        // ===== 去重：两个交叠菱形 =====
        private static void DrawDedup(Graphics g)
        {
            var d1 = new[] { new Point(36, 30), new Point(54, 40), new Point(36, 50), new Point(18, 40) };
            var d2 = new[] { new Point(46, 44), new Point(64, 54), new Point(46, 64), new Point(28, 54) };
            g.FillPolygon(B(), d1);
            g.FillPolygon(B(), d2);
        }

        // ===== 排序：升序柱状 =====
        private static void DrawSort(Graphics g)
        {
            using var b = B();
            g.FillRectangle(b, 24, 44, 12, 28);
            g.FillRectangle(b, 42, 34, 12, 38);
            g.FillRectangle(b, 60, 24, 12, 48);
        }

        // ===== 转置：双向箭头 =====
        private static void DrawTranspose(Graphics g)
        {
            using var pen = P(6);
            g.DrawLine(pen, 28, 40, 68, 40);
            g.DrawLine(pen, 28, 56, 68, 56);
            // 左箭头（指左）
            g.DrawLine(pen, 28, 48, 22, 42); g.DrawLine(pen, 28, 48, 22, 54);
            // 右箭头（指右）
            g.DrawLine(pen, 68, 48, 74, 42); g.DrawLine(pen, 68, 48, 74, 54);
        }

        // ===== 筛选：漏斗 =====
        private static void DrawFilter(Graphics g)
        {
            g.FillPolygon(B(), new[]
            {
                new Point(24, 26), new Point(72, 26), new Point(56, 46), new Point(56, 70), new Point(40, 70), new Point(40, 46)
            });
        }

        // ===== 透视：3×3 网格一笔高亮 =====
        private static void DrawPivot(Graphics g)
        {
            using var pen = P(5);
            g.DrawRectangle(pen, 24, 24, 48, 48);
            g.DrawLine(pen, 24, 40, 72, 40);
            g.DrawLine(pen, 24, 56, 72, 56);
            g.DrawLine(pen, 40, 24, 40, 72);
            g.DrawLine(pen, 56, 24, 56, 72);
            g.FillRectangle(B(), 42, 42, 12, 12);
        }

        // ===== 拆分：单元格中竖线 =====
        private static void DrawSplit(Graphics g)
        {
            using var pen = P(6);
            g.DrawRectangle(pen, 26, 26, 44, 44);
            g.DrawLine(pen, 48, 26, 48, 70);
        }

        // ===== 合并：两格 + 汇入箭头 =====
        private static void DrawMerge(Graphics g)
        {
            using var pen = P(6);
            g.DrawRectangle(pen, 24, 26, 48, 18);   // 顶格
            g.DrawLine(pen, 48, 26, 48, 44);
            g.DrawLine(pen, 48, 48, 30, 66);        // 汇入
            g.DrawLine(pen, 48, 48, 66, 66);
            g.DrawLine(pen, 30, 66, 66, 66);
        }

        // ===== 清理：橡皮 =====
        private static void DrawClean(Graphics g)
        {
            using var pen = P(7);
            var d = new PointF(30, 66);
            g.DrawLine(pen, 30, 62, 62, 28);
            g.DrawLine(pen, 34, 70, 66, 36);
            g.DrawLine(pen, 62, 64, 66, 60);
        }

        // ===== 插文本：T + 加号 =====
        private static void DrawInsertText(Graphics g)
        {
            using var pen = P(7);
            using var b = B();
            g.DrawLine(pen, 26, 30, 60, 30);
            g.DrawLine(pen, 43, 30, 43, 62);
            g.FillEllipse(b, 58, 50, 18, 18);
            g.DrawLine(pen, 61, 59, 73, 59);
            g.DrawLine(pen, 67, 53, 67, 65);
        }

        // ===== 删文本：T + 叉 =====
        private static void DrawDelText(Graphics g)
        {
            using var pen = P(7);
            g.DrawLine(pen, 26, 30, 60, 30);
            g.DrawLine(pen, 43, 30, 43, 62);
            g.DrawLine(pen, 56, 52, 72, 68);
            g.DrawLine(pen, 56, 68, 72, 52);
        }

        // ===== 随机数：骰子 =====
        private static void DrawDie(Graphics g)
        {
            using var pen = P(6);
            g.DrawRectangle(pen, 26, 26, 44, 44);
            using var b = B();
            g.FillEllipse(b, 38, 38, 8, 8);
            g.FillEllipse(b, 50, 38, 8, 8);
            g.FillEllipse(b, 44, 44, 8, 8);
            g.FillEllipse(b, 38, 50, 8, 8);
            g.FillEllipse(b, 50, 50, 8, 8);
        }

        // ===== 序列：1 2 3 =====
        private static void DrawSequence(Graphics g)
        {
            using var b = B();
            g.FillEllipse(b, 34, 26, 6, 6);
            g.FillEllipse(b, 44, 26, 6, 6);
            g.FillEllipse(b, 54, 26, 6, 6);
            g.FillEllipse(b, 30, 40, 6, 6);
            g.FillEllipse(b, 40, 40, 6, 6);
            g.FillEllipse(b, 50, 40, 6, 6);
            g.FillEllipse(b, 34, 54, 6, 6);
            g.FillEllipse(b, 44, 54, 6, 6);
            g.FillEllipse(b, 54, 54, 6, 6);
        }

        // ===== 插空行：横线 + 下加号 =====
        private static void DrawInsertBlank(Graphics g)
        {
            using var pen = P(7);
            g.DrawLine(pen, 26, 34, 70, 34);
            g.DrawLine(pen, 26, 62, 70, 62);
            g.DrawLine(pen, 48, 38, 48, 58);
            g.DrawLine(pen, 38, 48, 58, 48);
        }

        // ===== 删空行：两线 + 叉 =====
        private static void DrawDelBlank(Graphics g)
        {
            using var pen = P(7);
            g.DrawLine(pen, 26, 34, 70, 34);
            g.DrawLine(pen, 26, 62, 70, 62);
            g.DrawLine(pen, 40, 40, 56, 56);
            g.DrawLine(pen, 40, 56, 56, 40);
        }

        // ===== 填下：向下箭头到底线 =====
        private static void DrawFillDown(Graphics g)
        {
            using var pen = P(7);
            g.DrawLine(pen, 48, 26, 48, 58);
            g.DrawLine(pen, 36, 46, 48, 58);
            g.DrawLine(pen, 60, 46, 48, 58);
            g.DrawLine(pen, 26, 66, 70, 66);
        }

        // ===== 合并相同：等号 =====
        private static void DrawMergeDup(Graphics g)
        {
            using var pen = P(8);
            g.DrawLine(pen, 26, 40, 70, 40);
            g.DrawLine(pen, 26, 54, 70, 54);
        }

        // ===== 宏收纳箱：盒子 =====
        private static void DrawMacroBox(Graphics g)
        {
            using var pen = P(6);
            g.DrawRectangle(pen, 26, 30, 44, 16);
            g.DrawLine(pen, 26, 46, 70, 46);
            g.DrawLine(pen, 26, 46, 32, 66);
            g.DrawLine(pen, 26, 46, 20, 40);
            g.DrawLine(pen, 70, 46, 76, 40);
            g.DrawLine(pen, 32, 66, 64, 66);
            g.DrawLine(pen, 70, 46, 64, 66);
            g.DrawLine(pen, 20, 40, 26, 30);
            g.DrawLine(pen, 76, 40, 70, 30);
        }

        // ===== 函数库：翻开的书 =====
        private static void DrawFuncLib(Graphics g)
        {
            var spine = new PointF(48, 30);
            var left = new[] { spine, new PointF(26, 34), new PointF(26, 62), new PointF(48, 58) };
            var right = new[] { spine, new PointF(70, 34), new PointF(70, 62), new PointF(48, 58) };
            g.FillPolygon(B(), left);
            g.FillPolygon(B(), right);
        }

        private static void DrawSql(Graphics g)
        {
            using var pen = P(6);
            g.DrawEllipse(pen, 24, 26, 48, 18);
            g.DrawLine(pen, 24, 35, 24, 60);
            g.DrawLine(pen, 72, 35, 72, 60);
            g.DrawArc(pen, 24, 52, 48, 18, 0, 180);
        }

        private static void RoundedRect(Graphics g, Pen pen, float x, float y, float w, float h, float r)
        {
            using var path = RoundedPath(x, y, w, h, r);
            g.DrawPath(pen, path);
        }
    }
}