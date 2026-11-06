using System.Collections.Generic;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>FR-K3 分析结论附计算依据：AnalysisBasis 纯逻辑（地址解析/列区间/公式示例/成对完整观测）。</summary>
    public class AnalysisBasisTests
    {
        [Fact]
        public void TryParseSpan_parses_absolute_com_address()
        {
            var ok = AnalysisBasis.TryParseSpan("$B$2:$H$21", out var c1, out var r1, out var c2, out var r2);
            Assert.True(ok);
            Assert.Equal(2, c1);
            Assert.Equal(2, r1);
            Assert.Equal(8, c2);
            Assert.Equal(21, r2);
        }

        [Fact]
        public void TryParseSpan_parses_single_cell()
        {
            var ok = AnalysisBasis.TryParseSpan("$B$2", out var c1, out var r1, out var c2, out var r2);
            Assert.True(ok);
            Assert.Equal(c1, c2);
            Assert.Equal(r1, r2);
        }

        [Theory]
        [InlineData(1, "A")]
        [InlineData(26, "Z")]
        [InlineData(27, "AA")]
        [InlineData(52, "AZ")]
        [InlineData(703, "AAA")]
        public void ColumnLetter_converts_base26(int column, string expected)
        {
            Assert.Equal(expected, AnalysisBasis.ColumnLetter(column));
        }

        [Fact]
        public void ColumnSpan_skips_header_row_and_quotes_sheet()
        {
            // 数据行从地址起始行+1 开始（首行是表头）。
            Assert.Equal("'Sheet1'!$C$3:$C$21", AnalysisBasis.ColumnSpan("Sheet1", "$B$2:$H$21", 3));
        }

        [Fact]
        public void ColumnSpan_out_of_range_returns_empty()
        {
            Assert.Equal(string.Empty, AnalysisBasis.ColumnSpan("Sheet1", "$B$2:$H$21", 9));
            Assert.Equal(string.Empty, AnalysisBasis.ColumnSpan("Sheet1", "$B$2:$B$2", 1));
        }

        [Fact]
        public void Source_quotes_sheet_with_spaces()
        {
            Assert.Equal("'My Sheet'!B2:H21", AnalysisBasis.Source("My Sheet", "$B$2:$H$21"));
        }

        [Fact]
        public void FormulaExamples_caps_and_skips_unresolvable_columns()
        {
            var examples = AnalysisBasis.FormulaExamples(
                "Sheet1", "$B$2:$H$21", "AVERAGE", new[] { 2, 9, 3 }, 2);
            Assert.Equal(2, examples.Length);
            Assert.Equal("AVERAGE('Sheet1'!$B$3:$B$21)", examples[0]);
            Assert.Equal("AVERAGE('Sheet1'!$C$3:$C$21)", examples[1]);
        }

        [Fact]
        public void DetailRows_uses_actual_sheet_rows()
        {
            // 地址 $B$2:$H$21（20 行含表头）→ 明细 = 第 3 行～第 21 行，共 19 行。
            Assert.Contains("第 3 行～第 21 行", AnalysisBasis.DetailRows("$B$2:$H$21", 20));
            Assert.Contains("共 19 行", AnalysisBasis.DetailRows("$B$2:$H$21", 20));
            Assert.Equal("无明细行", AnalysisBasis.DetailRows("$B$2:$H$21", 1));
        }

        [Fact]
        public void PairAligned_drops_rows_missing_in_either_column()
        {
            var a = new Dictionary<int, double> { [1] = 1, [2] = 2, [3] = 3 };
            var b = new Dictionary<int, double> { [2] = 10, [3] = 20, [4] = 40 };
            var (xs, ys) = AnalysisBasis.PairAligned(a, b);
            Assert.Equal(new[] { 2.0, 3.0 }, xs);
            Assert.Equal(new[] { 10.0, 20.0 }, ys);
        }

        [Fact]
        public void PairAligned_keeps_all_rows_when_fully_paired()
        {
            var a = new Dictionary<int, double> { [1] = 1, [2] = 2, [3] = 3 };
            var b = new Dictionary<int, double> { [1] = 2, [2] = 4, [3] = 6 };
            var (xs, ys) = AnalysisBasis.PairAligned(a, b);
            // 完全线性 y=2x → 配对后相关系数应为 1（数值口径 sanity）。
            Assert.Equal(new[] { 1.0, 2.0, 3.0 }, xs);
            Assert.Equal(new[] { 2.0, 4.0, 6.0 }, ys);
        }

        [Fact]
        public void StatisticsBasis_contains_spec_mandated_fields()
        {
            var basis = AnalysisBasis.StatisticsBasis("Sheet1", "$B$2:$H$21", 20, new List<int> { 2, 3 });
            Assert.Contains("'Sheet1'!B2:H21", (string)basis["source"]);
            Assert.Contains("第 3 行", (string)basis["detail_rows"]);
            Assert.Contains("80%", (string)basis["numeric_rule"]);
            Assert.NotEmpty((string[])basis["formula_examples"]);
        }
    }
}
