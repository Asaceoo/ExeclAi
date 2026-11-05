using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 多元线性回归纯逻辑（analyze_regression，2026-09-14）：
    /// 普通最小二乘（正规方程 + 部分主元高斯消元求逆），输出系数表（估计/标准误/双侧 t/p）、
    /// R²/调整 R²/F 整体检验、残差标准误/Durbin-Watson/VIF/Cook's D 诊断。
    /// 全部公式已于 2026-09-14 在 Python 沙箱与 statsmodels/scipy 全参数对齐验证（误差 ≤1e-12）。
    /// 不接触 COM，供服务层与单测使用。非线性/逻辑回归/正则化走 python-advanced-analysis 技能。
    /// </summary>
    public static class RegressionAnalyzer
    {
        /// <summary>残差自由度至少 2，否则 F 检验与残差诊断严重失真，调用方须 fail-closed。</summary>
        public const int MinResidualDf = 2;

        /// <summary>求逆时主元绝对值低于该阈值视为奇异（完全共线/常数列/重复列）。</summary>
        private const double SingularityThreshold = 1e-12;

        public sealed class CoefficientRow
        {
            public string Name { get; set; }
            public double Estimate { get; set; }
            public double StdError { get; set; }
            public double TStat { get; set; }
            public double PValue { get; set; }
        }

        public sealed class VifRow
        {
            public string Name { get; set; }
            public double Vif { get; set; }
        }

        public sealed class ObservationRow
        {
            public int Row { get; set; }
            public double Actual { get; set; }
            public double Fitted { get; set; }
            public double Residual { get; set; }
            public double CooksD { get; set; }
        }

        public sealed class RegressionResult
        {
            public int TotalObs { get; set; }
            public int FeatureCount { get; set; }
            public int DfModel { get; set; }
            public int DfResidual { get; set; }
            public double R2 { get; set; }
            public double AdjR2 { get; set; }
            public double FStat { get; set; }
            public double FPValue { get; set; }
            public double ResidualStdError { get; set; }
            public double DurbinWatson { get; set; }
            public List<CoefficientRow> Coefficients { get; } = new List<CoefficientRow>();
            public List<VifRow> Vifs { get; } = new List<VifRow>();
            public List<ObservationRow> Observations { get; } = new List<ObservationRow>();
        }

        /// <summary>
        /// features 为逐行特征向量（不含截距列），y 与 features 行数一致；
        /// 求解路径：正规方程 (X'X)β = X'y → 部分主元高斯消元对 [X'X|I] 求逆。
        /// </summary>
        public static RegressionResult Compute(IReadOnlyList<double> y, IReadOnlyList<double[]> features, IReadOnlyList<string> featureNames)
        {
            if (y == null || features == null || y.Count != features.Count)
            {
                throw new ArgumentException("因变量与特征数据行数不一致。");
            }

            var m = featureNames == null ? 0 : featureNames.Count;
            if (m == 0)
            {
                throw new ArgumentException("至少需要 1 个特征列。");
            }

            var n = y.Count;
            var k = m + 1; // 参数个数（含截距）
            if (n - k < MinResidualDf)
            {
                throw new ArgumentException($"有效观测 {n} 行不足：{m} 个特征至少需要 {k + MinResidualDf} 行完整观测（残差自由度≥{MinResidualDf}）。");
            }

            // 设计矩阵：第 0 列为截距
            var X = new double[n, k];
            for (var i = 0; i < n; i++)
            {
                X[i, 0] = 1.0;
                for (var j = 0; j < m; j++)
                {
                    X[i, j + 1] = features[i][j];
                }
            }

            var inv = SolveWithInverse(X, k, n);

            // β = (X'X)^-1 X'y：先算 X'y，再逐分量内积
            var xty = new double[k];
            for (var b = 0; b < k; b++)
            {
                var s = 0.0;
                for (var i = 0; i < n; i++)
                {
                    s += X[i, b] * y[i];
                }
                xty[b] = s;
            }

            var beta = new double[k];
            for (var a = 0; a < k; a++)
            {
                var s = 0.0;
                for (var b = 0; b < k; b++)
                {
                    s += inv[a, b] * xty[b];
                }
                beta[a] = s;
            }

            // 拟合值/残差/平方和
            var fitted = new double[n];
            var resid = new double[n];
            var yMean = y.Average();
            double sse = 0, sst = 0;
            for (var i = 0; i < n; i++)
            {
                var f = 0.0;
                for (var a = 0; a < k; a++)
                {
                    f += X[i, a] * beta[a];
                }
                fitted[i] = f;
                resid[i] = y[i] - f;
                sse += resid[i] * resid[i];
                sst += (y[i] - yMean) * (y[i] - yMean);
            }

            var dfModel = k - 1;
            var dfResid = n - k;
            var sigma2 = sse / dfResid;
            var r2 = sst == 0 ? 1.0 : 1.0 - sse / sst;
            var adjR2 = 1.0 - (1.0 - r2) * (n - 1) / dfResid;
            var fStat = sse > 0 ? (sst - sse) / dfModel / sigma2 : 0.0;

            var result = new RegressionResult
            {
                TotalObs = n,
                FeatureCount = m,
                DfModel = dfModel,
                DfResidual = dfResid,
                R2 = r2,
                AdjR2 = adjR2,
                FStat = fStat,
                FPValue = HypothesisTester.FUpperP(fStat, dfModel, dfResid),
                ResidualStdError = Math.Sqrt(sigma2),
                DurbinWatson = sse > 0 ? DurbinWatsonOf(resid) : double.NaN
            };

            // 系数表：SE_j = sqrt(σ² · ((X'X)^-1)_jj)，t 为双侧，p 用 Core 既有 t 分布（与 scipy 对齐）
            result.Coefficients.Add(new CoefficientRow
            {
                Name = "截距",
                Estimate = beta[0],
                StdError = Math.Sqrt(sigma2 * inv[0, 0]),
                TStat = beta[0] / Math.Sqrt(sigma2 * inv[0, 0])
            });
            for (var j = 0; j < m; j++)
            {
                var se = Math.Sqrt(sigma2 * inv[j + 1, j + 1]);
                result.Coefficients.Add(new CoefficientRow
                {
                    Name = featureNames[j],
                    Estimate = beta[j + 1],
                    StdError = se,
                    TStat = beta[j + 1] / se
                });
            }
            foreach (var c in result.Coefficients)
            {
                c.PValue = HypothesisTester.TTestTwoSidedP(Math.Abs(c.TStat), dfResid);
            }

            // VIF_j = ((X'X)^-1)_jj · S_j，其中 S_j = Σ(x_j - x̄_j)²（含截距设计下的闭式解）
            for (var j = 0; j < m; j++)
            {
                var mean = 0.0;
                for (var i = 0; i < n; i++)
                {
                    mean += features[i][j];
                }
                mean /= n;
                var sj = 0.0;
                for (var i = 0; i < n; i++)
                {
                    sj += (features[i][j] - mean) * (features[i][j] - mean);
                }

                result.Vifs.Add(new VifRow
                {
                    Name = featureNames[j],
                    Vif = sj > 0 ? inv[j + 1, j + 1] * sj : 1.0
                });
            }

            // Cook's D_i = e_i²/(k·σ²) · h_ii/(1-h_ii)²，h_ii = x_i'(X'X)^-1 x_i
            for (var i = 0; i < n; i++)
            {
                var h = 0.0;
                for (var a = 0; a < k; a++)
                {
                    for (var b = 0; b < k; b++)
                    {
                        h += X[i, a] * inv[a, b] * X[i, b];
                    }
                }

                result.Observations.Add(new ObservationRow
                {
                    Row = i + 1,
                    Actual = y[i],
                    Fitted = fitted[i],
                    Residual = resid[i],
                    CooksD = (resid[i] * resid[i]) / (k * sigma2) * h / Math.Pow(1.0 - h, 2)
                });
            }

            return result;
        }

        /// <summary>Durbin-Watson = Σ(e_i - e_{i-1})² / Σe_i²（≈2 无自相关，<2 正自相关，>2 负自相关）。</summary>
        private static double DurbinWatsonOf(double[] resid)
        {
            double num = 0;
            for (var i = 1; i < resid.Length; i++)
            {
                var d = resid[i] - resid[i - 1];
                num += d * d;
            }

            var den = resid.Sum(e => e * e);
            return den > 0 ? num / den : double.NaN;
        }

        /// <summary>对增广矩阵 [X'X | I] 做部分主元高斯-约当消元，返回 (X'X)^-1；奇异时 fail-closed。</summary>
        private static double[,] SolveWithInverse(double[,] X, int k, int n)
        {
            var aug = new double[k, 2 * k];
            for (var a = 0; a < k; a++)
            {
                for (var b = 0; b < k; b++)
                {
                    var s = 0.0;
                    for (var i = 0; i < n; i++)
                    {
                        s += X[i, a] * X[i, b];
                    }
                    aug[a, b] = s;
                }
                aug[a, k + a] = 1.0;
            }

            for (var col = 0; col < k; col++)
            {
                // 部分主元
                var pivotRow = col;
                var pivotMax = Math.Abs(aug[col, col]);
                for (var r = col + 1; r < k; r++)
                {
                    if (Math.Abs(aug[r, col]) > pivotMax)
                    {
                        pivotMax = Math.Abs(aug[r, col]);
                        pivotRow = r;
                    }
                }

                if (pivotMax < SingularityThreshold)
                {
                    throw new ArgumentException("设计矩阵奇异：特征列存在常数列（无变异）或完全共线（某列可由其他列线性表出），请剔除冗余特征后重试。");
                }

                if (pivotRow != col)
                {
                    for (var c = 0; c < 2 * k; c++)
                    {
                        (aug[col, c], aug[pivotRow, c]) = (aug[pivotRow, c], aug[col, c]);
                    }
                }

                var pivot = aug[col, col];
                for (var c = 0; c < 2 * k; c++)
                {
                    aug[col, c] /= pivot;
                }

                for (var r = 0; r < k; r++)
                {
                    if (r == col) continue;
                    var factor = aug[r, col];
                    if (factor == 0) continue;
                    for (var c = 0; c < 2 * k; c++)
                    {
                        aug[r, c] -= factor * aug[col, c];
                    }
                }
            }

            var inv = new double[k, k];
            for (var a = 0; a < k; a++)
            {
                for (var b = 0; b < k; b++)
                {
                    inv[a, b] = aug[a, k + b];
                }
            }

            return inv;
        }
    }
}
