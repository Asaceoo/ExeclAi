using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 假设检验纯逻辑单测（analyze_hypothesis_test，2026-09-14）：
    /// 特殊函数对已知统计表值断言（容差 1e-3），四类检验对手工可验数据断言。
    /// </summary>
    public class HypothesisTesterTests
    {
        [Fact]
        public void TTest_TwoSidedP_KnownTableValue()
        {
            // t=2.0, df=10：双侧 p≈0.0734（统计表值）
            Assert.Equal(0.0734, HypothesisTester.TTestTwoSidedP(2.0, 10.0), 3);
            // t=0 → p=1
            Assert.Equal(1.0, HypothesisTester.TTestTwoSidedP(0.0, 10.0), 6);
        }

        [Fact]
        public void ChiSquare_KnownTableValue()
        {
            // chi2=3.84, df=1：p≈0.05004（临界值 0.05 对应）
            Assert.Equal(0.0500, HypothesisTester.ChiSquareUpperP(3.84, 1.0), 3);
            // chi2=0 → p=1
            Assert.Equal(1.0, HypothesisTester.ChiSquareUpperP(0.0, 1.0), 6);
        }

        [Fact]
        public void FDist_CentralF_OneIsHalf()
        {
            // F=1，d1=d2=10：p=0.5（中心 F 的对称性）
            Assert.Equal(0.5, HypothesisTester.FUpperP(1.0, 10.0, 10.0), 3);
        }

        [Fact]
        public void WelchT_KnownData()
        {
            // A=[10,12,14] B=[20,22,24]：m1=12,m2=22,s²=4 → t=-10/√(8/3)≈-6.1237；Welch df=(8/3)²/[(16/9)/2×2]=4
            var r = HypothesisTester.WelchTTest(new[] { 10.0, 12.0, 14.0 }, new[] { 20.0, 22.0, 24.0 }, "A", "B");
            Assert.Equal(-6.1237, r.Statistic, 3);
            Assert.Equal(4.0, r.Df, 3);
            Assert.True(r.PValue > 0 && r.PValue < 0.05);
            Assert.Equal("welch_t", r.Method);
        }

        [Fact]
        public void PairedT_KnownData()
        {
            // 差值恒定 [1,1,1,1]：sd=0 → t=0（退化但不应崩）；改用已知差值 [2,4,6,8]：mean=5, s=2.582, t=5/ (2.582/2)=3.873
            var r = HypothesisTester.PairedTTest(new[] { 11.0, 12.0, 13.0, 14.0 }, new[] { 9.0, 8.0, 7.0, 6.0 }, "后", "前");
            // 差值 [2,4,6,8]
            Assert.Equal(3.873, r.Statistic, 3);
            Assert.Equal(3.0, r.Df, 5);
            Assert.True(r.PValue < 0.05);
        }

        [Fact]
        public void PairedT_LengthMismatch_Throws()
        {
            Assert.Throws<ArgumentException>(() =>
                HypothesisTester.PairedTTest(new[] { 1.0, 2.0, 3.0 }, new[] { 1.0, 2.0 }, "A", "B"));
        }

        [Fact]
        public void Anova_KnownData_FEquals75()
        {
            // 三组各 3 点：SSB=3×[(12-22)²+0+(32-22)²]=600，MSB=300；SSW=2×4×3=24，MSW=4；F=75
            var groups = new List<(string, List<double>)>
            {
                ("G1", new List<double> { 10, 12, 14 }),
                ("G2", new List<double> { 20, 22, 24 }),
                ("G3", new List<double> { 30, 32, 34 })
            };
            var r = HypothesisTester.AnovaFTest(groups);
            Assert.Equal(75.0, r.Statistic, 3);
            Assert.Equal(2.0, r.Df, 5);
            Assert.Equal(6.0, r.Df2, 5);
            Assert.True(r.PValue < 1e-4);
            Assert.Equal(3, r.Groups.Count);
        }

        [Fact]
        public void ChiSquare_KnownTable()
        {
            // 2×2 表 [[30,10],[20,40]]：chi2=16.667，df=1，p<0.001
            var pairs = new List<(string, string)>();
            for (var i = 0; i < 30; i++) pairs.Add(("男", "买"));
            for (var i = 0; i < 10; i++) pairs.Add(("男", "不买"));
            for (var i = 0; i < 20; i++) pairs.Add(("女", "买"));
            for (var i = 0; i < 40; i++) pairs.Add(("女", "不买"));
            var r = HypothesisTester.ChiSquareIndependence(pairs);
            Assert.Equal(16.667, r.Statistic, 3);
            Assert.Equal(1.0, r.Df, 5);
            Assert.True(r.PValue < 0.001);
            Assert.Equal(100, r.TotalN);
            // Cramér's V = sqrt(16.667/100) = 0.408
            Assert.Equal(0.408, r.EffectSize ?? 0, 3);
        }

        [Fact]
        public void ChiSquare_TooFewPairs_Throws()
        {
            var pairs = new List<(string, string)> { ("A", "X"), ("A", "Y"), ("B", "X") };
            Assert.Throws<ArgumentException>(() => HypothesisTester.ChiSquareIndependence(pairs));
        }

        [Fact]
        public void MarkSignificance_RespectsAlpha()
        {
            var r = HypothesisTester.WelchTTest(new[] { 10.0, 12.0, 14.0 }, new[] { 20.0, 22.0, 24.0 }, "A", "B");
            HypothesisTester.MarkSignificance(r, 0.05);
            Assert.True(r.Significant);  // t≈-6.12, df=4 → p≈0.0036 < 0.05
            HypothesisTester.MarkSignificance(r, 0.001);
            Assert.False(r.Significant); // p≈0.0036 ≥ 0.001
        }
    }
}
