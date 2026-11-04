using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 向导矢量示意图控件（v3.7.1.55 放大版）：GDI+ 矢量绘制，随窗口/DPI 无级缩放不糊不裁。
    /// v3.7.1.55：设计空间 640×108→640×140，字号全面上调（标题 9.5F/块 9F/注 8.5F），
    /// Board 场景 KPI 卡 30→40 高消除"卡名与取值说明"文字拥挤（用户截图红框反馈）。
    /// 三型：Mapping=你的数据列→看板六角色；Pivot=行/值维度拼表效果；Board=看板布局与取值高亮。
    /// </summary>
    internal sealed class WizardDiagram : Control
    {
        internal enum DiagramKind { Mapping = 0, Pivot = 1, Board = 2 }

        internal DiagramKind Kind { get; set; } = DiagramKind.Mapping;

        private static readonly Color CRole1 = Color.FromArgb(68, 114, 196);    // 地区 蓝
        private static readonly Color CRole2 = Color.FromArgb(112, 173, 71);    // 品类 绿
        private static readonly Color CRole3 = Color.FromArgb(255, 192, 0);     // 月份 琥珀
        private static readonly Color CRole4 = Color.FromArgb(237, 125, 49);    // 金额 橙
        private static readonly Color CRole5 = Color.FromArgb(112, 48, 160);    // 数量 紫
        private static readonly Color CRole6 = Color.FromArgb(192, 0, 0);       // 利润 红
        private static readonly Color CGray = Color.FromArgb(110, 120, 136);

        public WizardDiagram()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Color.White;
            Height = 150;
            // 对抗性审查（审查者视角）：自定义控件未覆写 GetPreferredSize，TableLayoutPanel AutoSize 行
            // 可能按空首选尺寸塌缩——MinimumSize 硬兜底保证说明区示意图永远有高度
            MinimumSize = new Size(0, 150);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            // 设计坐标系 640×140，按控件实际尺寸等比缩放（k=每逻辑像素的物理像素数）
            float k = Math.Min(ClientSize.Width / 640f, ClientSize.Height / 140f);
            if (k <= 0.05f) { return; }
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.ScaleTransform(k, k);
            if (Kind == DiagramKind.Mapping) { PaintMapping(g); }
            else if (Kind == DiagramKind.Pivot) { PaintPivot(g); }
            else { PaintBoard(g); }
        }

        // ---- 绘制小工具（逻辑坐标；v3.7.1.55 字号上调）----

        private static void Box(Graphics g, float x, float y, float w, float h, string title, Color border)
        {
            using (var bg = new SolidBrush(Color.FromArgb(248, 250, 253)))
            using (var p = new Pen(border, 1.6f))
            {
                g.FillRectangle(bg, x, y, w, h);
                g.DrawRectangle(p, x, y, w, h);
            }
            if (!string.IsNullOrEmpty(title))
            {
                using (var f = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold))
                using (var b = new SolidBrush(Color.FromArgb(55, 65, 80)))
                {
                    g.DrawString(title, f, b, x + 7, y + 4);
                }
            }
        }

        private static void Chip(Graphics g, float x, float y, float w, float h, string text, Color color, bool filled)
        {
            using (var bg = new SolidBrush(filled ? color : Color.White))
            using (var p = new Pen(color, 1.3f))
            {
                g.FillRectangle(bg, x, y, w, h);
                g.DrawRectangle(p, x, y, w, h);
            }
            using (var f = new Font("Microsoft YaHei UI", 9F))
            using (var b = new SolidBrush(filled ? Color.White : color))
            {
                var sz = g.MeasureString(text, f);
                g.DrawString(text, f, b, x + (w - sz.Width) / 2f, y + (h - sz.Height) / 2f);
            }
        }

        private static void Arrow(Graphics g, float x1, float y1, float x2, float y2, string label)
        {
            using (var p = new Pen(Color.FromArgb(150, 160, 175), 2.2f))
            {
                p.EndCap = LineCap.ArrowAnchor;
                g.DrawLine(p, x1, y1, x2, y2);
            }
            if (!string.IsNullOrEmpty(label))
            {
                using (var f = new Font("Microsoft YaHei UI", 8.5F))
                using (var b = new SolidBrush(CGray))
                {
                    var sz = g.MeasureString(label, f);
                    g.DrawString(label, f, b, (x1 + x2) / 2f - sz.Width / 2f, y1 - sz.Height - 3);
                }
            }
        }

        private static void Note(Graphics g, float x, float y, string text)
        {
            using (var f = new Font("Microsoft YaHei UI", 8.5F))
            using (var b = new SolidBrush(CGray))
            {
                g.DrawString(text, f, b, x, y);
            }
        }

        // ---- 场景一（Step1）：你的数据列 → 六个角色 → 超级数据表 ----
        private void PaintMapping(Graphics g)
        {
            Box(g, 8, 6, 212, 128, "① 你的数据表（原始选区）", Color.FromArgb(200, 206, 216));
            float cx = 16, cy = 34, cw = 96, ch = 22;
            Chip(g, cx, cy, cw, ch, "地区/人员→分类1", CRole1, true);
            Chip(g, cx + cw + 8, cy, cw, ch, "品类→分类2", CRole2, true);
            Chip(g, cx, cy + ch + 6, cw, ch, "月份/日期→时间", CRole3, true);
            Chip(g, cx + cw + 8, cy + ch + 6, cw, ch, "金额→核心数", CRole4, true);
            Chip(g, cx, cy + (ch + 6) * 2, cw, ch, "数量→笔数", CRole5, true);
            Chip(g, cx + cw + 8, cy + (ch + 6) * 2, cw, ch, "利润(可选)", CRole6, true);
            Note(g, cx, cy + (ch + 6) * 2 + ch + 6, "六种颜色=向导里六个下拉，各选你表里的一列");
            Arrow(g, 228, 70, 260, 70, "自动");
            Box(g, 266, 6, 366, 128, "② 超级数据表（自动生成，原表不动）", CRole4);
            float bx = 274, by = 36, bw = 114, bh = 48;
            Box(g, bx, by, bw, bh, "月度汇总", CRole3);
            Box(g, bx + bw + 7, by, bw, bh, "品类汇总", CRole2);
            Box(g, bx + (bw + 7) * 2, by, bw, bh, "地区汇总", CRole1);
            Note(g, bx, by + bh + 10, "类型已规整：文字格式的数字→真数值 · 日期统一防错乱");
            Note(g, bx, by + bh + 28, "看板、透视表、KPI 卡都从这张表取数");
        }

        // ---- 场景二（Step2）：行维度 × 值字段 → 透视分析 ----
        private void PaintPivot(Graphics g)
        {
            Box(g, 8, 6, 182, 128, "行维度（按什么分类）", CRole1);
            Chip(g, 16, 38, 68, 22, "月份 ✓", CRole3, true);
            Chip(g, 92, 38, 68, 22, "地区 ✓", CRole1, true);
            Chip(g, 16, 66, 68, 22, "品类", CRole2, false);
            Note(g, 16, 100, "放在透视表左侧做行标签");
            Arrow(g, 196, 70, 222, 70, "组合");
            Box(g, 228, 6, 162, 128, "值字段（算什么数）", CRole4);
            Chip(g, 238, 38, 62, 22, "金额 ✓", CRole4, true);
            Chip(g, 308, 38, 62, 22, "数量", CRole5, false);
            Note(g, 238, 100, "自动求和，填进数值区");
            Arrow(g, 396, 70, 422, 70, "生成");
            Box(g, 428, 6, 204, 128, "透视分析（生成效果）", Color.FromArgb(200, 206, 216));
            // 迷你交叉表：行=月份 列=地区
            float tx = 438, ty = 38, tw = 62, th = 22;
            string[] heads = { "", "华北", "华东" };
            for (int c = 0; c < 3; c++)
            {
                Box(g, tx + c * tw, ty, tw - 3, th, "", CRole1);
                Note(g, tx + c * tw + 8, ty + 3, heads[c]);
            }
            string[][] rowsData = { new[] { "1月", "82万", "65万" }, new[] { "2月", "90万", "71万" } };
            for (int r = 0; r < 2; r++)
            {
                for (int c = 0; c < 3; c++)
                {
                    Box(g, tx + c * tw, ty + (r + 1) * th, tw - 3, th, "", Color.FromArgb(226, 230, 238));
                    Note(g, tx + c * tw + 8, ty + (r + 1) * th + 3, rowsData[r][c]);
                }
            }
            Note(g, 438, ty + 3 * th + 6, "月份每行 × 地区每列，格子=金额求和");
        }

        // ---- 场景三（Step3）：看板布局与「取的是哪些值」高亮（v3.7.1.55 放大：KPI 卡 40 高，卡名与取值说明分行不拥挤）----
        private void PaintBoard(Graphics g)
        {
            Box(g, 8, 6, 624, 128, "超级看板布局（生成效果）—— KPI 卡的数字就是对应列求和", Color.FromArgb(200, 206, 216));
            // 顶部 5 张 KPI 卡（第一张高亮=销售总额取金额列）
            string[] kpis = { "销售总额", "销售量", "成交笔数", "平均客单价", "最大单笔" };
            for (int i = 0; i < 5; i++)
            {
                float kx = 16 + i * 94;
                bool hot = i == 0;
                Box(g, kx, 32, 88, 42, "", hot ? CRole4 : Color.FromArgb(210, 215, 224));
                using (var f = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold))
                using (var b = new SolidBrush(hot ? CRole4 : Color.FromArgb(80, 90, 105)))
                {
                    g.DrawString(kpis[i], f, b, kx + 6, 36);
                }
                using (var f = new Font("Microsoft YaHei UI", 8.5F))
                using (var b = new SolidBrush(hot ? CRole4 : CGray))
                {
                    g.DrawString(hot ? "=金额列求和" : "自动计算", f, b, kx + 6, 56);
                }
            }
            // 中部四宫格图表
            string[] charts = { "柱状·月度", "饼图·品类", "折线·趋势", "条形·地区" };
            Color[] ccs = { CRole1, CRole2, CRole3, CRole5 };
            for (int i = 0; i < 4; i++)
            {
                float gx = 16 + i * 116;
                Box(g, gx, 84, 108, 44, "", ccs[i]);
                Note(g, gx + 7, 88, charts[i]);
            }
            // 右侧切片器
            Box(g, 492, 84, 132, 44, "", CRole6);
            Note(g, 498, 87, "切片器（筛选按钮）");
            Chip(g, 498, 106, 38, 18, "地区", CRole1, false);
            Chip(g, 540, 106, 38, 18, "品类", CRole2, false);
            Chip(g, 582, 106, 38, 18, "月份", CRole3, false);
        }
    }
}
