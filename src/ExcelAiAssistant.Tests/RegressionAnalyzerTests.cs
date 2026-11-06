using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 多元线性回归纯逻辑单测（analyze_regression，2026-09-14）：
    /// 与 Python 沙箱锁定 statsmodels 参考值全指标对照、精确数据系数还原、
    /// 完全共线/样本不足 fail-closed、诊断明细对齐。
    /// </summary>
    public class RegressionAnalyzerTests
    {
        private static readonly string[] Names = { "x1", "x2" };

        /// <summary>
        /// 参考数据集（Python 沙箱与 statsmodels/scipy 全参数对齐验证，误差 ≤1e-12）：
        /// y = 3 + 2x1 + 1.5x2 + 确定性噪声。
        /// </summary>
        private static (double[] Y, List<double[]> X) ReferenceData()
        {
            var x1 = new[] { 1.0, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            var x2 = new[] { 2.0, 1, 4, 3, 6, 5, 8, 7, 10, 9 };
            var noise = new[] { 0.3, -0.2, 0.5, -0.4, 0.1, -0.3, 0.4, -0.1, 0.2, -0.5 };
            var y = x1.Select((v, i) => 3.0 + 2.0 * v + 1.5 * x2[i] + noise[i]).ToArray();
            var x = x1.Select((v, i) => new[] { v, x2[i] }).ToList();
            return (y, x);
        }

        [Fact]
        public void ReferenceValues_MatchStatsmodels()
        {
            var (y, x) = ReferenceData();
            var r = RegressionAnalyzer.Compute(y, x, Names);

            // 系数/标准误/t（statsmodels 锁定值）
            Assert.Equal(3, r.Coefficients.Count);
            Assert.Equal(3.0825, r.Coefficients[0].Estimate, 10);
            Assert.Equal(1.6925, r.Coefficients[1].Estimate, 10);
            Assert.Equal(1.7925, r.Coefficients[2].Estimate, 10);
            Assert.Equal(0.1114955156, r.Coefficients[0].StdError, 8);
            Assert.Equal(0.0517807879, r.Coefficients[1].StdError, 8);
            Assert.Equal(0.0517807879, r.Coefficients[2].StdError, 8);
            Assert.Equal(27.6468518332, r.Coefficients[0].TStat, 6);
            Assert.Equal(32.6858680069, r.Coefficients[1].TStat, 6);
            Assert.Equal(34.6170862053, r.Coefficients[2].TStat, 6);

            // p 值为 1e-8 量级且极小（t 分布连分式与 scipy 对齐误差 ~1e-12）
            Assert.True(r.Coefficients[0].PValue < 1e-6 && r.Coefficients[0].PValue > 0);
            Assert.True(r.Coefficients[1].PValue < 1e-6 && r.Coefficients[1].PValue > 0);

            // 模型层面
            Assert.Equal(0.999812723484, r.R2, 10);
            Assert.Equal(0.999759215908, r.AdjR2, 10);
            Assert.Equal(18685.4423076925, r.FStat, 4);
            Assert.True(r.FPValue < 1e-9);
            Assert.Equal(7, r.DfResidual);
            Assert.Equal(2, r.DfModel);
            Assert.Equal(0.161245154966, r.ResidualStdError, 8);
            Assert.Equal(1.492307692308, r.DurbinWatson, 8);

            // VIF（x1 与 x2 明显正相关 → 8.5）
            Assert.Equal(8.5078125, r.Vifs[0].Vif, 8);
            Assert.Equal(8.5078125, r.Vifs[1].Vif, 8);

            // Cook's D 最大值（statsmodels 影响度对齐）
            Assert.Equal(0.301424501425, r.Observations.Max(o => o.CooksD), 8);
        }

        [Fact]
        public void ExactLinearData_RecoversCoefficients()
        {
            var x1 = new[] { 1.0, 2, 3, 4, 5, 6 };
            var x2 = new[] { 10.0, 8, 6, 5, 2, 0 };
            var y = x1.Select((v, i) => 5.0 + 2.0 * v + 3.0 * x2[i]).ToArray();
            var x = x1.Select((v, i) => new[] { v, x2[i] }).ToList();

            var r = RegressionAnalyzer.Compute(y, x, Names);
            Assert.Equal(5.0, r.Coefficients[0].Estimate, 6);
            Assert.Equal(2.0, r.Coefficients[1].Estimate, 6);
            Assert.Equal(3.0, r.Coefficients[2].Estimate, 6);
            Assert.True(r.R2 > 0.999999);
        }

        [Fact]
        public void ExactCollinearity_Throws()
        {
            // x3 = x1 + x2 完全共线
            var x1 = new[] { 1.0, 2, 3, 4, 5, 6 };
            var x2 = new[] { 2.0, 4, 3, 5, 7, 8 };
            var y = x1.Select(v => 1.0 + v).ToArray();
            var x = x1.Select((v, i) => new[] { v, x2[i], v + x2[i] }).ToList();

            var ex = Assert.Throws<ArgumentException>(() => RegressionAnalyzer.Compute(y, x, new[] { "x1", "x2", "x3" }));
            Assert.Contains("奇异", ex.Message);
        }

        [Fact]
        public void TooFewRows_Throws()
        {
            // 3 个特征 → 需 3+1+2=6 行；只给 5 行
            var y = new[] { 1.0, 2, 3, 4, 5 };
            var x = y.Select((v, i) => new[] { v, i + 1.0, v * 0.5 }).ToList();

            var ex = Assert.Throws<ArgumentException>(() => RegressionAnalyzer.Compute(y, x, Names.Concat(new[] { "x3" }).ToArray()));
            Assert.Contains("完整观测", ex.Message);
        }

        [Fact]
        public void Diagnostics_AlignedWithRows()
        {
            var (y, x) = ReferenceData();
            var r = RegressionAnalyzer.Compute(y, x, Names);

            Assert.Equal(y.Length, r.Observations.Count);
            for (var i = 0; i < y.Length; i++)
            {
                var o = r.Observations[i];
                Assert.Equal(i + 1, o.Row);
                Assert.Equal(y[i], o.Actual, 12);
                Assert.Equal(y[i] - o.Fitted, o.Residual, 9);
                Assert.True(o.CooksD >= 0);
            }
        }
    }
}
