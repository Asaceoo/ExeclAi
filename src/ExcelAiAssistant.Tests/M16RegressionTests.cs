using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M16 全量分类审查回归锁：
    /// 1. ExcelErrorValues（CVErr 归一，单一来源）；
    /// 2. 判定器"无效期望值"fail-closed 统一（不得静默放行）；
    /// 3. verify_hyperlink 多格链接按矩形相交匹配。
    /// </summary>
    public sealed class M16RegressionTests
    {
        // ===================== ExcelErrorValues =====================

        [Theory]
        [InlineData(-2146826246, "#N/A")]
        [InlineData(-2146826281, "#DIV/0!")]
        [InlineData(-2146826288, "#NULL!")]
        [InlineData(-2146826281 - 1, null)] // 非 CVErr 码 → 原样
        public void Normalize_maps_cveerr_codes_to_error_text(int code, string expected)
        {
            var result = ExcelErrorValues.Normalize(code);
            if (expected == null)
            {
                Assert.Equal(code, result);
            }
            else
            {
                Assert.Equal(expected, result);
            }
        }

        [Fact]
        public void Normalize_passes_through_non_error_values()
        {
            Assert.Equal(42.0, ExcelErrorValues.Normalize(42.0));
            Assert.Equal("abc", ExcelErrorValues.Normalize("abc"));
            Assert.Null(ExcelErrorValues.Normalize(null));
            Assert.True(ExcelErrorValues.IsCveErr(-2146826246));
            Assert.False(ExcelErrorValues.IsCveErr(-1));
        }

        // ===================== 无效期望值 fail-closed =====================

        [Fact]
        public void Conditional_format_unknown_expected_rule_is_invalid_expectation()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                "cellIs", "1", null, 3, null, null,
                "greter" /* 拼写错误 */, null, null, null, null, null);
            Assert.Contains("invalid_expectation", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Conditional_format_no_expected_rule_is_no_expected_fail_closed()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                "cellIs", "1", null, 3, null, null,
                null, null, null, null, null, null);
            // CLAUDE.md §6：断言型 verify 无期望值必须 fail-closed，不得静默放行。
            Assert.Equal("no_expected", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Sort_unknown_direction_is_invalid_expectation()
        {
            var matrix = new object[2, 1] { { 1.0 }, { 2.0 } };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeSortVerdict(
                matrix, 0, 1, "descc" /* 拼写错误 */, null, null);
            Assert.Contains("invalid_expectation", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Cell_format_unknown_alignment_is_invalid_expectation()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeCellFormatVerdict(
                "General", true, false, "宋体", 11, 0, 16777215, -4108, -4108, 1, false,
                null, null, null, null, null, null, null,
                "middel" /* 拼写错误 */, null, null, null);
            Assert.Contains("invalid_expectation", verdict);
            Assert.Single(issues);
        }

        // ===================== 多格超链接矩形匹配 =====================

        [Fact]
        public void Hyperlink_multicell_anchor_matches_address_inside_span()
        {
            // 链接锚点 A1:B2，address=B2（非左上角格）也应命中（M16：与写侧口径对齐）。
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "B2", null, "https://a.com", null, null,
                new System.Collections.Generic.List<ExcelVerifyJudgments.HyperlinkInfo>
                {
                    new ExcelVerifyJudgments.HyperlinkInfo
                    {
                        Cell = "A1:B2",
                        Url = "https://a.com",
                        SubAddress = string.Empty,
                        Text = string.Empty
                    }
                });
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);

            // 行跨度覆盖：锚点 A1:A3、address=A2 命中
            var (spanOk, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "A2", null, "https://a.com", null, null,
                new System.Collections.Generic.List<ExcelVerifyJudgments.HyperlinkInfo>
                {
                    new ExcelVerifyJudgments.HyperlinkInfo
                    {
                        Cell = "A1:A3",
                        Url = "https://a.com",
                        SubAddress = string.Empty,
                        Text = string.Empty
                    }
                });
            Assert.Equal("ok", spanOk);

            // 完全不相交则仍报 no_target
            var (missing, _, _) = ExcelVerifyJudgments.ComputeHyperlinkVerdict(
                null, "C5", null, null, null, null,
                new System.Collections.Generic.List<ExcelVerifyJudgments.HyperlinkInfo>
                {
                    new ExcelVerifyJudgments.HyperlinkInfo
                    {
                        Cell = "A1:A3",
                        Url = "https://a.com",
                        SubAddress = string.Empty,
                        Text = string.Empty
                    }
                });
            Assert.Contains("no_target", missing);
        }
    }
}
