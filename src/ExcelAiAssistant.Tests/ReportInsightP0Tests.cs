using ExcelAiAssistant.Excel;
using Xunit;
using System.Collections.Generic;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// P0-1/P0-2 验收锁定（需求-报告洞察增强.md Q4/Q6 拍板）：
    /// - 无基线数据注入 → 概览声明必须出现（变异测试）；
    /// - 自检负例（空段/占位符/覆盖缺口/数字对账）必须被抓到（U29 篡改必抓先例）。
    /// </summary>
    public class ReportInsightP0Tests
    {
        private static object[,] Matrix(object[,] m) => m;

        [Fact]
        public void No_baseline_data_declares_no_baseline()
        {
            // 变异测试：无同期/目标/预算列 → 声明必须出现在概览
            var m = new object[4, 2]
            {
                { "月份", "销售额" },
                { "1月", 100 },
                { "2月", 200 },
                { "3月", 300 }
            };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:B4", 2, 20);
            var overview = report.Sections.Find(s => s.Title == "概览");
            Assert.NotNull(overview);
            Assert.Contains(overview.Lines, l => l.Contains("未发现对比基线列") && l.Contains("仅基于绝对值"));
            Assert.DoesNotContain(report.Sections, s => s.Title == "同环比对比");
        }

        [Fact]
        public void Target_baseline_produces_achievement_rate()
        {
            // 目标列语义（销售额目标 配对 销售额）→ 达成率入「同环比对比」节
            var m = new object[4, 2]
            {
                { "销售额", "销售额目标" },
                { 100, 120 },
                { 200, 200 },
                { 300, 300 }
            };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:B4", 2, 20);
            var section = report.Sections.Find(s => s.Title == "同环比对比");
            Assert.NotNull(section);
            Assert.Contains(section.Lines, l => l.Contains("达成率") && l.Contains("配对依据"));
        }

        [Fact]
        public void Period_baseline_produces_yoy()
        {
            // 同期列语义（本期销售额 配对 同期销售额）→ 同环比
            var m = new object[4, 2]
            {
                { "本期销售额", "同期销售额" },
                { 110, 100 },
                { 220, 200 },
                { 330, 300 }
            };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:B4", 2, 20);
            var section = report.Sections.Find(s => s.Title == "同环比对比");
            Assert.NotNull(section);
            Assert.Contains(section.Lines, l => l.Contains("同环比") && l.Contains("配对依据"));
        }

        [Fact]
        public void Selfcheck_catches_placeholder_line()
        {
            // 自检负例：占位符行必须被抓到
            var m = new object[2, 1]
            {
                { "列" },
                { 5 }
            };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:A2", 2, 20);
            // 注入占位符（模拟组装缺陷）
            report.Sections.Add(new ReportSection { Title = "测试节", Lines = new List<string> { "TODO 待补" } });
            var check = ReportTextBuilder.SelfCheck(report, null, m, new[] { "列" });
            Assert.Equal("issues", check.Verdict);
            Assert.Contains(check.Issues, x => x.Contains("占位符") || x.Contains("TODO"));
        }

        [Fact]
        public void Selfcheck_catches_empty_section()
        {
            var m = new object[2, 1] { { "列" }, { 5 } };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:A2", 2, 20);
            report.Sections.Add(new ReportSection { Title = "空节", Lines = new List<string>() });
            var check = ReportTextBuilder.SelfCheck(report, null, m, new[] { "列" });
            Assert.Equal("issues", check.Verdict);
            Assert.Contains(check.Issues, x => x.Contains("空段"));
        }

        [Fact]
        public void Selfcheck_catches_coverage_gap()
        {
            var m = new object[2, 1] { { "列" }, { 5 } };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:A2", 2, 20);
            var allow = new HashSet<string> { "概览", "不存在的章节" };
            var check = ReportTextBuilder.SelfCheck(report, allow, m, new[] { "列" });
            Assert.Contains(check.Issues, x => x.Contains("覆盖缺口") && x.Contains("不存在的章节"));
        }

        [Fact]
        public void Selfcheck_ok_on_clean_report()
        {
            var m = new object[4, 2]
            {
                { "月份", "销售额" },
                { "1月", 100 },
                { "2月", 200 },
                { "3月", 300 }
            };
            var report = ReportTextBuilder.Build(m, "Sheet1", "A1:B4", 2, 20);
            Assert.Equal("ok", report.SelfCheck.Verdict);
            Assert.Empty(report.SelfCheck.Issues);
            Assert.True(report.SelfCheck.CheckedCount > 0);
        }
    }
}
