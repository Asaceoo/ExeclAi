using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 趋势预测纯逻辑 + 报告章节白名单/相关性阈值参数化单测（2026-09-14，完整全面数据分析批次）。
    /// </summary>
    public class TrendForecastTests
    {
        private static List<double> RisingSeries(int n = 12, double start = 100, double step = 10)
        {
            return Enumerable.Range(0, n).Select(i => start + step * i).ToList();
        }

        [Fact]
        public void Linear_RisingSeries_PredictsContinuation()
        {
            var values = RisingSeries();
            var r = TrendForecaster.Compute(values, "linear", 3, 3);
            Assert.Equal("上升", r.Trend);
            Assert.Equal(10.0, r.Slope, 5);
            // y=90+10x（x=1..12 时 y=100..210），截距 90 而非首点值 100。
            Assert.Equal(90.0, r.Intercept, 5);
            Assert.Equal(1.0, r.R2, 5);
            // 第 13 期 = 100 + 10*12 = 220
            Assert.Equal(220.0, r.ForecastValues[0], 3);
            Assert.Equal(230.0, r.ForecastValues[1], 3);
            Assert.Equal(3, r.ForecastValues.Count);
        }

        [Fact]
        public void MovingAverage_UsesLastWindowMean()
        {
            var values = new List<double> { 10, 20, 30, 40, 50, 60 };
            var r = TrendForecaster.Compute(values, "moving_average", 3, 2);
            Assert.Equal(50.0, r.ForecastValues[0], 5);
            Assert.Equal(50.0, r.ForecastValues[1], 5);
            Assert.Contains("水平", r.Trend);
        }

        [Fact]
        public void InsufficientPoints_Throws()
        {
            var values = new List<double> { 1, 2, 3, 4, 5 };
            Assert.Throws<ArgumentException>(() => TrendForecaster.Compute(values, "linear", 3, 3));
        }

        [Fact]
        public void NextLabel_ExtendsTrailingNumber()
        {
            // 契约：标签末尾为数字才延伸，且保留前导零位宽；其余形态退化为 T+k。
            Assert.Equal("2026-04", TrendForecaster.NextLabel("2026-03", 1));
            Assert.Equal("Q13", TrendForecaster.NextLabel("Q12", 1));
            Assert.Equal("T+1", TrendForecaster.NextLabel("一季度", 1));
            Assert.Equal("T+2", TrendForecaster.NextLabel(null, 2));
        }

        [Fact]
        public void Report_SectionWhitelist_FiltersAllOutputSections()
        {
            var matrix = new object[5, 2]
            {
                { "月份", "销量" },
                { "1月", 100.0 },
                { "2月", 120.0 },
                { "3月", 90.0 },
                { "4月", 150.0 }
            };
            var report = ReportTextBuilder.Build(matrix, "Sheet1", "A1:B5", 2, 20, 0.7, new[] { "概览", "建议" });
            var titles = report.Sections.Select(s => s.Title).ToList();
            Assert.Contains("概览", titles);
            Assert.Contains("建议", titles);
            Assert.DoesNotContain("相关性", titles);
            Assert.DoesNotContain("重点列摘要", titles);
        }

        [Fact]
        public void Report_LowerCorrelationThreshold_CatchesWeakPair()
        {
            // 仅两个数值列：A 递增、B 前半跟随后半走平，手工核算 r≈0.79（介于 0.3 与 0.99 之间）。
            var matrix = new object[11, 3]
            {
                { "序号", "A", "B" },
                { "第1行", 1.0, 2.0 },
                { "第2行", 2.0, 4.0 },
                { "第3行", 3.0, 6.0 },
                { "第4行", 4.0, 8.0 },
                { "第5行", 5.0, 10.0 },
                { "第6行", 6.0, 10.0 },
                { "第7行", 7.0, 10.0 },
                { "第8行", 8.0, 10.0 },
                { "第9行", 9.0, 10.0 },
                { "第10行", 10.0, 10.0 }
            };
            var strict = ReportTextBuilder.Build(matrix, "Sheet1", "A1:C11", 2, 20, 0.99, null);
            var strictSection = strict.Sections.First(s => s.Title == "相关性");
            Assert.Contains("未发现", strictSection.Lines[0]);

            var loose = ReportTextBuilder.Build(matrix, "Sheet1", "A1:C11", 2, 20, 0.3, null);
            var looseSection = loose.Sections.First(s => s.Title == "相关性");
            Assert.DoesNotContain("未发现", looseSection.Lines[0]);
        }
    }
}
