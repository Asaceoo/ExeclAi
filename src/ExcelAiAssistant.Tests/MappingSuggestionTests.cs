using System;
using System.Linq;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// FR-6.2 suggest_chart_mapping 选型启发式单测（纯函数，无 COM）。
    /// 每条断言 = 数据形态 → 期望工具/图表类型；理由与数据要求字段随建议强制返回。
    /// </summary>
    public sealed class MappingSuggestionTests
    {
        private static object[,] Build(string[] headers, object[][] rows)
        {
            var block = new object[rows.Length + 1, headers.Length];
            for (int c = 0; c < headers.Length; c++) { block[0, c] = headers[c]; }
            for (int r = 0; r < rows.Length; r++)
            {
                for (int c = 0; c < headers.Length; c++) { block[r + 1, c] = rows[r][c]; }
            }
            return block;
        }

        [Fact]
        public void Negative_values_suggest_waterfall_with_reason()
        {
            var block = Build(new[] { "科目", "金额" }, new[]
            {
                new object[] { "收入", 120000.0 },
                new object[] { "成本", -60000.0 },
                new object[] { "净利润", 30000.0 },
            });
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out var cols);
            Assert.Contains(sugg, s => s.Tool == "create_chart_advanced" && s.ChartType == "waterfall");
            var wf = sugg.First(s => s.ChartType == "waterfall");
            Assert.False(string.IsNullOrWhiteSpace(wf.Reason));
            Assert.False(string.IsNullOrWhiteSpace(wf.Requires));
        }

        [Fact]
        public void Decreasing_values_suggest_funnel()
        {
            var block = Build(new[] { "阶段", "人数" }, new[]
            {
                new object[] { "访问", 1000.0 },
                new object[] { "下单", 600.0 },
                new object[] { "支付", 300.0 },
                new object[] { "复购", 100.0 },
            });
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "create_chart_advanced" && s.ChartType == "funnel");
        }

        [Fact]
        public void Two_numeric_columns_suggest_combo_chart()
        {
            var block = Build(new[] { "月份", "销售额", "销量" }, new[]
            {
                new object[] { "1月", 100.0, 10.0 },
                new object[] { "2月", 130.0, 12.0 },
                new object[] { "3月", 90.0, 9.0 },
            });
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "create_combo_chart");
        }

        [Fact]
        public void Category_time_value_suggest_rank_dynamic()
        {
            var rows = new System.Collections.Generic.List<object[]>();
            foreach (var m in new[] { "1月", "2月" })
            {
                for (int i = 0; i < 4; i++)
                {
                    rows.Add(new object[] { "产品" + (char)('A' + i), m, 100.0 - i * 10 });
                }
            }
            var block = Build(new[] { "类别", "月份", "数值" }, rows.ToArray());
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "render_rank_dynamic");
        }

        [Fact]
        public void Five_column_sales_shape_suggests_sales_overview()
        {
            var rows = new System.Collections.Generic.List<object[]>();
            for (int i = 0; i < 6; i++)
            {
                rows.Add(new object[] { "华北", "品类" + (i % 2), (i % 2 + 1) + "月", 100.0 + i, 10.0 + i });
            }
            var block = Build(new[] { "地区", "品类", "月份", "销售额", "销量" }, rows.ToArray());
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "render_dashboard" && s.ChartType == "sales_overview");
        }

        [Fact]
        public void Few_categories_suggest_doughnut_many_suggest_column()
        {
            var few = Build(new[] { "类别", "占比" }, new[]
            {
                new object[] { "A", 0.5 }, new object[] { "B", 0.3 }, new object[] { "C", 0.2 },
            });
            Assert.Contains(ExcelDashboardService.AnalyzeChartMapping(few, out _),
                s => s.Tool == "create_doughnut");

            var many = Build(new[] { "类别", "数值" }, Enumerable.Range(0, 12)
                .Select(i => new object[] { "类" + i, (double)i }).ToArray());
            Assert.Contains(ExcelDashboardService.AnalyzeChartMapping(many, out _),
                s => s.Tool == "create_excel_chart");
        }

        [Fact]
        public void Pure_numeric_column_suggests_histogram()
        {
            var block = Build(new[] { "分数" }, Enumerable.Range(0, 20)
                .Select(i => new object[] { 60.0 + i }).ToArray());
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "create_chart_advanced" && s.ChartType == "histogram");
        }

        [Fact]
        public void Header_only_block_returns_no_suggestions()
        {
            var block = new object[1, 2] { { "A", "B" } };
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out var cols);
            Assert.Empty(sugg);
            Assert.Equal(2, cols.Count);
        }

        [Fact]
        public void Fallback_column_suggestion_always_present()
        {
            var block = Build(new[] { "类别", "数值" }, new[]
            {
                new object[] { "A", 1.0 }, new object[] { "B", 2.0 },
            });
            var sugg = ExcelDashboardService.AnalyzeChartMapping(block, out _);
            Assert.Contains(sugg, s => s.Tool == "create_excel_chart" && s.ChartType == "column");
        }
    }
}
