using System;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 看板洞察卡布局纯逻辑（add_insight_card，2026-09-14）：
    /// 按标题+正文与卡宽估算建议卡高（中文全角/ASCII 半角字宽近似），超长正文截断省略。
    /// 形状实际创建在服务层（COM），本类只做确定性布局计算，供单测。
    /// </summary>
    public static class InsightCardLayout
    {
        public const float TitleFontSize = 11f;
        public const float BodyFontSize = 10f;
        public const float PaddingPt = 10f;
        public const float LineHeightTitle = 18f;
        public const float LineHeightBody = 15f;
        /// <summary>正文中文字符宽（10pt 字号下全角字宽近似 pt）。</summary>
        public const float CjkCharWidth = 10f;
        /// <summary>ASCII 字符按半角近似。</summary>
        public const float AsciiCharWidth = 5.5f;
        /// <summary>默认卡宽（pt）。</summary>
        public const float DefaultWidthPt = 220f;
        /// <summary>正文最大行数（超出截断加省略号），防洞察文字撑爆看板。</summary>
        public const int MaxBodyLines = 8;

        public sealed class CardLayout
        {
            public float WidthPt { get; set; }
            public float HeightPt { get; set; }
            public int BodyLines { get; set; }
            public bool Truncated { get; set; }
        }

        /// <summary>字符显示宽度（pt）：CJK/全角按全宽，ASCII 按半宽。</summary>
        public static float CharWidth(char c)
        {
            return c > 0xFF ? CjkCharWidth : AsciiCharWidth;
        }

        /// <summary>按卡宽估算一行可容纳的显示宽度（pt）。</summary>
        public static float LineCapacity(float widthPt)
        {
            return Math.Max(widthPt - PaddingPt * 2, CjkCharWidth * 4);
        }

        /// <summary>正文按显示宽度换行所需行数。</summary>
        public static int CountBodyLines(string body, float widthPt)
        {
            if (string.IsNullOrEmpty(body)) return 0;
            var capacity = LineCapacity(widthPt);
            var lines = 1;
            float acc = 0;
            foreach (var ch in body)
            {
                if (ch == '\n')
                {
                    lines++;
                    acc = 0;
                    continue;
                }

                acc += CharWidth(ch);
                if (acc > capacity)
                {
                    lines++;
                    acc = CharWidth(ch);
                }
            }

            return lines;
        }

        /// <summary>超长正文截断：保留前 maxLines 行宽度内的字符，超出以「…」结尾。</summary>
        public static string TruncateBody(string body, float widthPt, int maxLines)
        {
            if (string.IsNullOrEmpty(body)) return body;
            var capacity = LineCapacity(widthPt);
            var sb = new System.Text.StringBuilder();
            var lines = 1;
            float acc = 0;
            var truncated = false;
            foreach (var ch in body)
            {
                if (lines > maxLines)
                {
                    truncated = true;
                    break;
                }

                if (ch == '\n')
                {
                    if (lines == maxLines)
                    {
                        truncated = true;
                        break;
                    }

                    lines++;
                    acc = 0;
                    sb.Append(ch);
                    continue;
                }

                var w = CharWidth(ch);
                if (acc + w > capacity && lines == maxLines)
                {
                    truncated = true;
                    break;
                }

                acc += w;
                sb.Append(ch);
            }

            if (truncated) sb.Append('…');
            return sb.ToString();
        }

        /// <summary>完整布局：自动高度 = 内边距 + 标题行 + 正文行数×行高；正文超限截断。</summary>
        public static CardLayout Estimate(string title, string body, float widthPt)
        {
            var w = Math.Max(widthPt, 120f);
            var truncated = false;
            var bodyText = body ?? string.Empty;
            var lines = CountBodyLines(bodyText, w);
            if (lines > MaxBodyLines)
            {
                bodyText = TruncateBody(bodyText, w, MaxBodyLines);
                lines = MaxBodyLines;
                truncated = true;
            }

            var height = PaddingPt * 2 + LineHeightTitle + lines * LineHeightBody;
            return new CardLayout
            {
                WidthPt = w,
                HeightPt = height,
                BodyLines = lines,
                Truncated = truncated
            };
        }

        /// <summary>主题色映射：返回 Excel ForeColor.RGB 的 Long 值（&amp;HBBGGRR 字节序）。name: blue/green/orange/red/gray，缺省 blue。</summary>
        public static int ColorBgr(string name)
        {
            switch ((name ?? "blue").Trim().ToLowerInvariant())
            {
                case "green": return 0x62A85A;   // RGB(90,168,98)
                case "orange": return 0x2E9ED9;  // RGB(217,158,46)
                case "red": return 0x4A4AC4;     // RGB(196,74,74)
                case "gray": return 0x8F8F8F;    // RGB(143,143,143)
                default: return 0xC4722E;        // 蓝 RGB(46,114,196)
            }
        }
    }
}
