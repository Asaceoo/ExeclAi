using System;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ChartTypeMapperTests
    {
        [Theory]
        [InlineData(null, "ColumnClustered")]
        [InlineData("", "ColumnClustered")]
        [InlineData("column", "ColumnClustered")]
        [InlineData("ColumnClustered", "ColumnClustered")]
        [InlineData(" bar ", "BarClustered")]
        [InlineData("line", "Line")]
        [InlineData("pie", "Pie")]
        [InlineData("area", "Area")]
        [InlineData("scatter", "xlXYScatter")]
        [InlineData("xyscatter", "xlXYScatter")]
        [InlineData("histogram", "xlHistogram")]
        [InlineData("unknown", "ColumnClustered")]
        public void Resolves_chart_type_names(string input, string expected)
        {
            Assert.Equal(expected, ChartTypeMapper.Resolve(input));
        }

        [Fact]
        public void Detects_histogram_case_insensitively()
        {
            Assert.True(ChartTypeMapper.IsHistogram("Histogram"));
            Assert.False(ChartTypeMapper.IsHistogram("column"));
            Assert.False(ChartTypeMapper.IsHistogram(null));
        }
    }

    public sealed class HistogramBinnerTests
    {
        [Fact]
        public void Builds_bins_with_header_and_totals()
        {
            var result = HistogramBinner.BuildBins(new[] { 1D, 2D, 3D, 4D, 5D }, 5);

            Assert.Equal(6, result.GetLength(0));
            Assert.Equal("区间", result[0, 0]);
            Assert.Equal("频数", result[0, 1]);
            var total = 0;
            for (var i = 1; i < result.GetLength(0); i++)
            {
                total += Convert.ToInt32(result[i, 1]);
            }

            Assert.Equal(5, total);
        }

        [Fact]
        public void Collapses_single_value_into_one_bin()
        {
            var result = HistogramBinner.BuildBins(new[] { 7D, 7D, 7D }, 3);

            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(3, Convert.ToInt32(result[1, 1]));
        }

        [Fact]
        public void Empty_input_returns_only_header()
        {
            var result = HistogramBinner.BuildBins(Array.Empty<double>(), 3);

            Assert.Equal(1, result.GetLength(0));
            Assert.Equal("区间", result[0, 0]);
        }

        [Fact]
        public void Null_input_returns_only_header()
        {
            var result = HistogramBinner.BuildBins(null, 3);

            Assert.Equal(1, result.GetLength(0));
            Assert.Equal("频数", result[0, 1]);
        }
    }

    public sealed class ReportTextBuilderTests
    {
        private static object[,] TrendMatrix()
        {
            return new object[,]
            {
                { "类别", "数量", "金额" },
                { "A", 10D, 100D },
                { "B", 20D, 200D },
                { "C", 30D, 300D },
                { "D", 40D, 400D }
            };
        }

        [Fact]
        public void Build_produces_all_expected_sections()
        {
            var result = ReportTextBuilder.Build(TrendMatrix(), "Sheet1", "A1:C5", 2, 20);

            var titles = result.Sections.Select(s => s.Title).ToArray();
            foreach (var expected in new[] { "概览", "趋势", "异常", "重点列摘要", "相关性", "建议" })
            {
                Assert.Contains(titles, t => t == expected);
            }

            Assert.True(result.WordCount > 0);
        }

        [Fact]
        public void Trend_section_reports_rising_values()
        {
            var result = ReportTextBuilder.Build(TrendMatrix(), "Sheet1", "A1:C5", 2, 20);

            var trend = result.Sections.First(s => s.Title == "趋势").Lines;
            Assert.Contains(trend, line => line.Contains("上升") && line.Contains("数量"));
        }

        [Fact]
        public void Anomaly_section_flags_outlier_with_given_sigma()
        {
            var matrix = new object[,]
            {
                { "ID", "值" },
                { "r1", 1D },
                { "r2", 2D },
                { "r3", 3D },
                { "r4", 100D }
            };

            var result = ReportTextBuilder.Build(matrix, "Sheet1", "A1:B5", 1, 20);

            var anomaly = result.Sections.First(s => s.Title == "异常").Lines;
            Assert.Contains(anomaly, line => line.Contains("离群") && line.Contains("值") && line.Contains("r4"));
        }

        [Fact]
        public void Markdown_output_has_title_and_sections()
        {
            var result = ReportTextBuilder.Build(TrendMatrix(), "Sheet1", "A1:C5", 2, 20);

            Assert.StartsWith("# XAI 分析报告", result.Markdown);
            Assert.Contains("## 概览", result.Markdown);
            Assert.Contains("## 建议", result.Markdown);
        }

        [Fact]
        public void Html_output_is_escaped_and_structured()
        {
            var result = ReportTextBuilder.Build(TrendMatrix(), "Sheet1", "A1:C5", 2, 20);

            Assert.Contains("<h1>XAI 分析报告</h1>", result.Html);
            Assert.Contains("<h2>概览</h2><ul>", result.Html);
            Assert.Contains("</html>", result.Html);
            Assert.DoesNotContain("<script>", result.Html);
        }

        [Fact]
        public void Worksheet_grid_starts_with_report_header()
        {
            var result = ReportTextBuilder.Build(TrendMatrix(), "Sheet1", "A1:C5", 2, 20);

            Assert.Contains("XAI 分析报告", Convert.ToString(result.ReportGrid[0, 0]));
            Assert.True(result.ReportGrid.GetLength(0) > 5);
        }

        [Fact]
        public void Correlation_section_appears_for_linked_columns()
        {
            var matrix = new object[,]
            {
                { "类别", "x", "y" },
                { "A", 1D, 2D },
                { "B", 2D, 4D },
                { "C", 3D, 6D },
                { "D", 4D, 8D },
                { "E", 5D, 10D }
            };

            var result = ReportTextBuilder.Build(matrix, "Sheet1", "A1:C6", 2, 20);

            var correlation = result.Sections.First(s => s.Title == "相关性").Lines;
            Assert.Contains(correlation, line => line.Contains("r=") && line.Contains("正相关"));
        }
    }

    public sealed class PhaseCRegistryTests
    {
        [Fact]
        public void Registry_includes_phase_c_tools()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            Assert.Contains(registry.Tools, tool => tool.Name == "insert_recommended_chart");
            Assert.Contains(registry.Tools, tool => tool.Name == "generate_analysis_report");

            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("insert_recommended_chart"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("generate_analysis_report"));
        }
    }
}