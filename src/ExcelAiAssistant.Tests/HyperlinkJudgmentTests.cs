using System.Collections.Generic;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// verify_hyperlink 判定器单测（M15）：内容断言 fail-closed、present/内容互斥、
    /// 地址包含匹配、宽松比对（url 大小写/末尾斜杠、subaddress 定位符、text 大小写空白）。
    /// </summary>
    public sealed class HyperlinkJudgmentTests
    {
        private static List<ExcelVerifyJudgments.HyperlinkInfo> Links(
            params (string cell, string url, string sub, string text)[] items)
        {
            var list = new List<ExcelVerifyJudgments.HyperlinkInfo>();
            foreach (var (cell, url, sub, text) in items)
            {
                list.Add(new ExcelVerifyJudgments.HyperlinkInfo
                {
                    Cell = cell,
                    Url = url ?? string.Empty,
                    SubAddress = sub ?? string.Empty,
                    Text = text ?? string.Empty
                });
            }

            return list;
        }

        // ===================== 地址定位 =====================

        [Fact]
        public void Address_without_expectations_requires_a_link_in_range()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, null, null,
                Links(("B2", "https://a.com", "", "官网")));
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);

            var (missing, missingIssues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, null, null,
                Links(("C9", "https://a.com", "", "官网")));
            Assert.Contains("no_target", missing);
            Assert.Single(missingIssues);
        }

        [Fact]
        public void Range_address_matches_links_inside_bounds()
        {
            var (ok, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "A1:B5", null, null, null, null,
                Links(("B3", "https://a.com", "", "x")));
            Assert.Equal("ok", ok);

            var (outside, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "A1:B5", null, null, null, null,
                Links(("C1", "https://a.com", "", "x")));
            Assert.Contains("no_target", outside);
        }

        [Fact]
        public void Invalid_address_is_reported_not_silent()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "不是地址", null, null, null, null, Links());
            Assert.Contains("invalid_address", verdict);
            Assert.Single(issues);
        }

        // ===================== fail-closed 组合 =====================

        [Fact]
        public void Content_assertion_without_address_is_fail_closed()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, null, null, "https://a.com", null, null,
                Links(("B2", "https://a.com", "", "x")));
            Assert.Contains("target_without_address", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Expected_present_without_address_is_fail_closed()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, null, false, null, null, null, Links());
            Assert.Contains("invalid_expectation", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Expected_absent_with_content_assertion_is_contradictory()
        {
            // 矛盾组合 fail-closed：既报组合矛盾，也如实报目标上仍存在链接（两条 issue 并存）。
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", false, "https://a.com", null, null,
                Links(("B2", "https://a.com", "", "x")));
            Assert.Contains("invalid_expectation", verdict);
            Assert.Equal(2, issues.Count);

            var (absentClean, absentCleanIssues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", false, "https://a.com", null, null, Links());
            Assert.Contains("invalid_expectation", absentClean);
            Assert.Single(absentCleanIssues);
        }

        // ===================== 内容断言 =====================

        [Fact]
        public void Url_mismatch_is_flagged()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, "https://expected.com", null, null,
                Links(("B2", "https://actual.com", "", "x")));
            Assert.Contains("url_mismatch", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Url_comparison_ignores_case_and_trailing_slash()
        {
            var (verdict, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, "HTTPS://A.com/", null, null,
                Links(("B2", "https://a.com", "", "x")));
            Assert.Equal("ok", verdict);
        }

        [Fact]
        public void Subaddress_comparison_ignores_quotes_and_dollars()
        {
            var (ok, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, "'Sheet1'!$A$1", null,
                Links(("B2", "", "Sheet1!A1", "x")));
            Assert.Equal("ok", ok);

            var (mismatch, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, "Sheet2!A1", null,
                Links(("B2", "", "Sheet1!A1", "x")));
            Assert.Contains("subaddress_mismatch", mismatch);
        }

        [Fact]
        public void Text_comparison_ignores_case_and_surrounding_blank()
        {
            var (ok, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, null, " 官网 ",
                Links(("B2", "https://a.com", "", "官网")));
            Assert.Equal("ok", ok);

            var (mismatch, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, null, null, "官网",
                Links(("B2", "https://a.com", "", "门户")));
            Assert.Contains("text_mismatch", mismatch);
        }

        // ===================== present / count =====================

        [Fact]
        public void Expected_absent_passes_only_when_range_is_clear()
        {
            var (ok, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", false, null, null, null, Links());
            Assert.Equal("ok", ok);

            var (leftover, leftoverIssues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", false, null, null, null,
                Links(("B2", "https://a.com", "", "x")));
            Assert.Contains("should_be_absent", leftover);
            Assert.Single(leftoverIssues);
        }

        [Fact]
        public void Expected_count_checks_sheet_total()
        {
            var links = Links(
                ("B2", "https://a.com", "", "x"),
                ("B3", "", "Sheet1!A1", "y"));
            var (ok, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                2, null, null, null, null, null, links);
            Assert.Equal("ok", ok);

            var (mismatch, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                1, null, null, null, null, null, links);
            Assert.Contains("count_mismatch", mismatch);
        }

        [Fact]
        public void Count_and_content_assertions_compose()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                1, "B2", null, "https://expected.com", null, null,
                Links(
                    ("B2", "https://actual.com", "", "x"),
                    ("B9", "https://a.com", "", "y")));
            Assert.Contains("url_mismatch", verdict);
            Assert.Contains("count_mismatch", verdict);
            Assert.Equal(2, issues.Count);
        }
    }
}
