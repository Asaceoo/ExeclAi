using System;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 看板洞察卡布局纯逻辑单测（add_insight_card，2026-09-14）：
    /// 中文全角/ASCII 半角宽度换行估算、卡高自动、超长截断省略、主题色映射。
    /// </summary>
    public class InsightCardLayoutTests
    {
        [Fact]
        public void CharWidth_CjkFullAsciiHalf()
        {
            Assert.Equal(InsightCardLayout.CjkCharWidth, InsightCardLayout.CharWidth('东'));
            Assert.Equal(InsightCardLayout.CjkCharWidth, InsightCardLayout.CharWidth('，')); // 全角标点
            Assert.Equal(InsightCardLayout.AsciiCharWidth, InsightCardLayout.CharWidth('A'));
            Assert.Equal(InsightCardLayout.AsciiCharWidth, InsightCardLayout.CharWidth('1'));
        }

        [Fact]
        public void Estimate_HeightGrowsWithText()
        {
            var shortCard = InsightCardLayout.Estimate("结论", "华东贡献 42%。", 220f);
            var longCard = InsightCardLayout.Estimate("结论", "华东区域贡献了接近一半的销售额，主要动力来自新客户放量与老客复购提升，建议追加投放。", 220f);

            Assert.True(shortCard.HeightPt < longCard.HeightPt, "文字更多应更高");
            Assert.True(shortCard.HeightPt >= InsightCardLayout.PaddingPt * 2 + InsightCardLayout.LineHeightTitle);
            Assert.False(shortCard.Truncated);
        }

        [Fact]
        public void Truncate_OverMaxLines_AddsEllipsis()
        {
            var body = string.Join("。", Enumerable.Repeat("这是一段比较长的洞察正文，用来测试换行与截断逻辑", 20)) + "。";
            var layout = InsightCardLayout.Estimate("标题", body, 220f);

            Assert.True(layout.Truncated);
            Assert.Equal(InsightCardLayout.MaxBodyLines, layout.BodyLines);
            // 截断后的高度应封顶（不再随原文长度增长）
            var longer = InsightCardLayout.Estimate("标题", body + body, 220f);
            Assert.Equal(layout.HeightPt, longer.HeightPt);
        }

        [Fact]
        public void TruncateBody_ManualLineBreaks_Counted()
        {
            var body = "第一行结论。\n第二行依据。\n第三行建议。";
            Assert.Equal(3, InsightCardLayout.CountBodyLines(body, 400f));
            var truncated = InsightCardLayout.TruncateBody(body, 400f, 2);
            Assert.StartsWith("第一行结论。", truncated);
            Assert.EndsWith("…", truncated);
        }

        [Fact]
        public void ColorBgr_KnownMappings()
        {
            Assert.Equal(0xC4722E, InsightCardLayout.ColorBgr("blue"));   // RGB(46,114,196)
            Assert.Equal(0x62A85A, InsightCardLayout.ColorBgr("green"));  // RGB(90,168,98)
            Assert.Equal(0x4A4AC4, InsightCardLayout.ColorBgr("red"));    // RGB(196,74,74)
            Assert.Equal(0xC4722E, InsightCardLayout.ColorBgr(null));     // 缺省 blue
            Assert.Equal(0xC4722E, InsightCardLayout.ColorBgr("unknown"));
        }

        [Fact]
        public void Estimate_NarrowFloorAndEmptyBody()
        {
            // 极窄卡按最小可用宽度兜底
            var tiny = InsightCardLayout.Estimate("标", "文", 40f);
            Assert.True(tiny.WidthPt >= 120f);
            var empty = InsightCardLayout.Estimate("只有标题", "", 220f);
            Assert.Equal(0, empty.BodyLines);
        }
    }
}
