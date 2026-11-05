using System.Collections.Generic;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>异常检测器 OutlierDetector 纯逻辑测试（可脱离 Excel 运行）。</summary>
    public sealed class OutlierDetectorTests
    {
        [Fact]
        public void IQR_detects_single_outlier_in_small_sample()
        {
            // 中位数固定，极端值不会污染分位数 → 9999（第 4 位）检出，其余不误伤
            var res = OutlierDetector.DetectOutliers(new double[] { 100, 150, 200, 9999 });
            Assert.Equal(new List<int> { 4 }, res);
        }

        [Fact]
        public void IQR_no_outlier_when_data_normal()
        {
            var res = OutlierDetector.DetectOutliers(new double[] { 10, 12, 11, 13, 12, 11, 10, 13, 12, 11 });
            Assert.Empty(res);
        }

        [Fact]
        public void IQR_ignores_nan_and_infinity()
        {
            var res = OutlierDetector.DetectOutliers(new double[] { 1, 2, 3, 4, 1, double.NaN, double.PositiveInfinity, 90 });
            Assert.Single(res);
        }

        [Theory]
        [InlineData(null)]
        [InlineData(new double[0])]
        public void IQR_empty_or_null_returns_empty(double[] values)
        {
            Assert.Empty(OutlierDetector.DetectOutliers(values));
        }

        [Fact]
        public void Zscore_detects_outlier()
        {
            var res = OutlierDetector.DetectOutliers(new double[] { 2, 3, 4, 5, 5, 5, 5, 6, 7, 8, 50 }, "zscore");
            Assert.NotEmpty(res);
        }

        [Fact]
        public void Text_as_number_rows_identified()
        {
            // 2 列文本列：含可解析数字的字符串
            object[,] block = new object[,] {
                { "abc", "12.5" },
                { "def", "xyz" },
                { "",     "3"   }
            };
            var res = OutlierDetector.DetectTextAsNumber(block, 2);
            // 行1 "12.5"、行3 "3" 为文本型数字；行2 "xyz" 不是
            Assert.Equal(new List<int> { 1, 3 }, res);
        }

        [Fact]
        public void Bad_dates_rows_identified_in_date_column()
        {
            object[,] block = new object[,]
            {
                { "2026-01-01" },
                { "2026-01-05" },
                { "not-a-date" },
                { "2026-02-01" }
            };
            var res = OutlierDetector.DetectBadDates(block, 1);
            Assert.Equal(new List<int> { 3 }, res);
        }

        [Fact]
        public void BuildPlan_aggregates_algorithms_by_flag_type()
        {
            object[,] block = new object[,]
            {
                { "编号", 100, "2026-01-01", "ok" },
                { "编号", 200, "2026-01-05", ""   },   // empty @(2,4)
                { "编号", 9999, "not-a-date", "x" },   // outlier @(3,2), bad-date @(3,3)
                { "编号", 150, "2026-02-01", "ok" }
            };
            var plan = OutlierDetector.BuildPlan(block, new[] { "outlier", "format", "empty" });

            // outlier：金额列异常 9999
            Assert.True(plan.ContainsKey("outlier"));
            var o = Assert.Single(plan["outlier"]);
            Assert.Equal(3, o.Row); Assert.Equal(2, o.Col); Assert.Equal("9999", o.Value);

            // format：坏日期 一处
            Assert.True(plan.ContainsKey("format"));
            var f = Assert.Single(plan["format"]);
            Assert.Equal(3, f.Row); Assert.Equal(3, f.Col);

            // empty：一处 @(2,4)
            Assert.True(plan.ContainsKey("empty"));
            var e = Assert.Single(plan["empty"]);
            Assert.Equal(2, e.Row); Assert.Equal(4, e.Col);
        }

        [Fact]
        public void BuildPlan_flag_filter_excludes_unwanted_types()
        {
            object[,] block = new object[,] { { 1, 100 }, { 2, 100 }, { 3, 9999 }, { 4, 100 } };
            var plan = OutlierDetector.BuildPlan(block, new[] { "outlier" });
            Assert.True(plan.ContainsKey("outlier"));
            Assert.False(plan.ContainsKey("empty"));
        }
    }
}