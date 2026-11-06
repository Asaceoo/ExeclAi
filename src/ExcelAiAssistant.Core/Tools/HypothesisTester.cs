using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 假设检验纯逻辑（analyze_hypothesis_test，2026-09-14）：
    /// Welch t / 配对 t / 单因素 ANOVA(F) / 卡方独立性，含特殊函数（lgamma/不完全 gamma/不完全 beta，Numerical Recipes 口径）。
    /// 不接触 COM，供服务层与单测使用。非参检验/复杂设计仍路由 python-advanced-analysis 技能。
    /// </summary>
    public static class HypothesisTester
    {
        // ---- 特殊函数 ----

        private const double Epsilon = 3.0e-9;
        private const int MaxIterations = 200;

        private static double LogGamma(double x)
        {            // Lanczos 近似（g=7, n=9 系数），double 精度 ~1e-13。
            double[] cof =
            {
                0.99999999999980993, 676.5203681218851, -1259.1392167224028,
                771.32342877765313, -176.61502916214059, 12.507343278686905,
                -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7
            };
            if (x < 0.5)
            {
                return Math.Log(Math.PI / Math.Sin(Math.PI * x)) - LogGamma(1.0 - x);
            }

            x -= 1.0;
            var a = cof[0];
            var t = x + 7.5;
            for (var i = 1; i < 9; i++)
            {
                a += cof[i] / (x + i);
            }

            return 0.5 * Math.Log(2 * Math.PI) + (x + 0.5) * Math.Log(t) - t + Math.Log(a);
        }

        /// <summary>正则化下不完全 gamma P(a,x)（级数 x&lt;a+1；连分式否则）。</summary>
        public static double GammaP(double a, double x)
        {
            if (a <= 0 || x < 0) return double.NaN;
            if (x == 0) return 0;
            if (x < a + 1.0)
            {
                // 级数
                var ap = a;
                var sum = 1.0 / a;
                var del = sum;
                for (var n = 1; n <= MaxIterations; n++)
                {
                    ap += 1.0;
                    del *= x / ap;
                    sum += del;
                    if (Math.Abs(del) < Math.Abs(sum) * Epsilon) break;
                }

                return sum * Math.Exp(-x + a * Math.Log(x) - LogGamma(a));
            }

            // 连分式（Lentz）
            var b = x + 1.0 - a;
            var c = 1.0 / Epsilon;
            var d = 1.0 / b;
            var h = d;
            for (var i = 1; i <= MaxIterations; i++)
            {
                var an = -i * (i - a);
                b += 2.0;
                d = an * d + b;
                if (Math.Abs(d) < Epsilon) d = Epsilon;
                c = b + an / c;
                if (Math.Abs(c) < Epsilon) c = Epsilon;
                d = 1.0 / d;
                var del2 = d * c;
                h *= del2;
                if (Math.Abs(del2 - 1.0) < Epsilon) break;
            }

            var q = Math.Exp(-x + a * Math.Log(x) - LogGamma(a)) * h;
            return 1.0 - q;
        }

        /// <summary>
        /// 正则化不完全 beta I_x(a,b)：连分式反向求值（d 系数与 scipy.special.betainc 全参数对齐验证）。
        /// I_x = bt/a · 1/(1 + d1/(1 + d2/(1+...)))，d_{2m+1}=-(a+m)(a+b+m)x/((a+2m)(a+2m+1))，d_{2m}=m(b-m)x/((a+2m-1)(a+2m))。
        /// </summary>
        public static double BetaI(double a, double b, double x)
        {
            if (a <= 0 || b <= 0) return double.NaN;
            if (x <= 0) return 0;
            if (x >= 1) return 1;
            var bt = Math.Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b)
                               + a * Math.Log(x) + b * Math.Log(1 - x));
            const int N = 400;
            var ds = new double[N];
            for (int m = 0, i = 0; m < N / 2; m++, i += 2)
            {
                ds[i] = -(a + m) * (a + b + m) * x / ((a + 2.0 * m) * (a + 2.0 * m + 1.0));
                ds[i + 1] = (m + 1) * (b - m - 1.0) * x / ((a + 2.0 * (m + 1) - 1.0) * (a + 2.0 * (m + 1)));
            }

            var t = 1.0;
            for (var i = N - 1; i >= 0; i--)
            {
                t = 1.0 + ds[i] / t;
            }

            return bt / a * (1.0 / t);
        }

        /// <summary>Student t 双侧 p 值（df 自由度）。</summary>
        public static double TTestTwoSidedP(double t, double df)
        {
            if (double.IsNaN(t) || df <= 0) return double.NaN;
            var x = df / (df + t * t);
            var p = BetaI(df / 2.0, 0.5, x);
            return p;
        }

        /// <summary>卡方右上侧 p 值（k 自由度）。</summary>
        public static double ChiSquareUpperP(double chi2, double k)
        {
            if (chi2 < 0 || k <= 0) return double.NaN;
            return 1.0 - GammaP(k / 2.0, chi2 / 2.0);
        }

        /// <summary>F 右上侧 p 值（d1/d2 自由度）。</summary>
        public static double FUpperP(double f, double d1, double d2)
        {
            if (f < 0 || d1 <= 0 || d2 <= 0) return double.NaN;
            var x = d1 * f / (d1 * f + d2);
            return 1.0 - BetaI(d1 / 2.0, d2 / 2.0, x);
        }

        // ---- 检验 ----

        public sealed class GroupStats
        {
            public string Label { get; set; }
            public int N { get; set; }
            public double Mean { get; set; }
            public double Variance { get; set; }
        }

        public sealed class TestResult
        {
            public string Method { get; set; }
            public double Statistic { get; set; }
            public double Df { get; set; }
            public double Df2 { get; set; }
            public double PValue { get; set; }
            public double Alpha { get; set; }
            public bool Significant { get; set; }
            public double? EffectSize { get; set; }
            public string EffectName { get; set; }
            public List<GroupStats> Groups { get; set; } = new List<GroupStats>();
            public List<string> PremiseNotes { get; set; } = new List<string>();
            public int TotalN { get; set; }
        }

        private static double SampleVariance(IReadOnlyList<double> v)
        {
            var n = v.Count;
            if (n < 2) return 0;
            var mean = v.Average();
            return v.Sum(x => (x - mean) * (x - mean)) / (n - 1);
        }

        private static GroupStats Describe(string label, IReadOnlyList<double> v)
        {
            return new GroupStats
            {
                Label = label,
                N = v.Count,
                Mean = v.Count > 0 ? v.Average() : double.NaN,
                Variance = SampleVariance(v)
            };
        }

        /// <summary>两独立样本 t 检验（Welch，不假设方差齐性）。</summary>
        public static TestResult WelchTTest(IReadOnlyList<double> a, IReadOnlyList<double> b, string labelA, string labelB)
        {
            if (a.Count < 2 || b.Count < 2)
            {
                throw new ArgumentException("每组至少需要 2 个有效数值。");
            }

            var result = new TestResult { Method = "welch_t" };
            var m1 = a.Average();
            var m2 = b.Average();
            var v1 = SampleVariance(a);
            var v2 = SampleVariance(b);
            var se = Math.Sqrt(v1 / a.Count + v2 / b.Count);
            var t = (m1 - m2) / se;
            var df = Math.Pow(v1 / a.Count + v2 / b.Count, 2) /
                     (Math.Pow(v1 / a.Count, 2) / (a.Count - 1) + Math.Pow(v2 / b.Count, 2) / (b.Count - 1));
            var sp = Math.Sqrt(((a.Count - 1) * v1 + (b.Count - 1) * v2) / (a.Count + b.Count - 2));
            result.Statistic = t;
            result.Df = df;
            result.PValue = TTestTwoSidedP(Math.Abs(t), df);
            result.EffectName = "cohen_d";
            result.EffectSize = sp > 0 ? (m1 - m2) / sp : 0;
            result.Groups.Add(Describe(labelA, a));
            result.Groups.Add(Describe(labelB, b));
            result.TotalN = a.Count + b.Count;
            result.PremiseNotes.Add("Welch t 不要求方差齐性；两组样本建议各≥30 或近似正态（正态性可用 Python 技能复核）");
            return result;
        }

        /// <summary>配对 t 检验（逐行配对差值）。</summary>
        public static TestResult PairedTTest(IReadOnlyList<double> a, IReadOnlyList<double> b, string labelA, string labelB)
        {
            if (a.Count != b.Count)
            {
                throw new ArgumentException("配对检验要求两列有效数值行数一致（成对完整观测）。");
            }

            if (a.Count < 2)
            {
                throw new ArgumentException("配对样本至少需要 2 对有效数值。");
            }

            var diff = a.Zip(b, (x, y) => x - y).ToList();
            var result = new TestResult { Method = "paired_t" };
            var mean = diff.Average();
            var sd = Math.Sqrt(SampleVariance(diff));
            var t = sd > 0 ? mean / (sd / Math.Sqrt(diff.Count)) : 0;
            var df = diff.Count - 1;
            result.Statistic = t;
            result.Df = df;
            result.PValue = TTestTwoSidedP(Math.Abs(t), df);
            result.EffectName = "cohen_dz";
            result.EffectSize = sd > 0 ? mean / sd : 0;
            result.Groups.Add(Describe(labelA, a));
            result.Groups.Add(Describe(labelB, b));
            result.TotalN = diff.Count;
            result.PremiseNotes.Add("配对 t 假设差值近似正态；差值方向为「" + labelA + " − " + labelB + "」");
            return result;
        }

        /// <summary>单因素方差分析（F 检验，≥2 组）。</summary>
        public static TestResult AnovaFTest(IReadOnlyList<(string Label, List<double> Values)> groups)
        {
            var valid = groups.Where(g => g.Values.Count >= 2).ToList();
            if (valid.Count < 2)
            {
                throw new ArgumentException("ANOVA 至少需要 2 个各含 2 个有效数值的组。");
            }

            var result = new TestResult { Method = "anova_f" };
            var all = valid.SelectMany(g => g.Values).ToList();
            var grand = all.Average();
            var k = valid.Count;
            var n = all.Count;
            double ssb = 0, ssw = 0;
            foreach (var g in valid)
            {
                var m = g.Values.Average();
                ssb += g.Values.Count * (m - grand) * (m - grand);
                ssw += g.Values.Sum(x => (x - m) * (x - m));
                result.Groups.Add(Describe(g.Label, g.Values));
            }

            var df1 = k - 1;
            var df2 = n - k;
            var msb = ssb / df1;
            var msw = ssw / df2;
            var f = msw > 0 ? msb / msw : 0;
            result.Statistic = f;
            result.Df = df1;
            result.Df2 = df2;
            result.PValue = FUpperP(f, df1, df2);
            result.EffectName = "eta_squared";
            result.EffectSize = (ssb + ssw) > 0 ? ssb / (ssb + ssw) : 0;
            result.TotalN = n;
            result.PremiseNotes.Add("ANOVA 假设各组近似正态且方差齐；组间方差差异大时结论须谨慎（可用 Python 技能做 Welch ANOVA 复核）");
            return result;
        }

        /// <summary>卡方独立性检验（两个类别列的列联表）。</summary>
        public static TestResult ChiSquareIndependence(IReadOnlyList<(string RowLabel, string ColLabel)> pairs)
        {
            if (pairs.Count < 5)
            {
                throw new ArgumentException("卡方检验至少需要 5 条有效（类别A, 类别B）配对。");
            }

            var rows = pairs.Select(p => p.RowLabel).Distinct().ToList();
            var cols = pairs.Select(p => p.ColLabel).Distinct().ToList();
            var rowIdx = rows.Select((v, i) => (v, i)).ToDictionary(x => x.v, x => x.i);
            var colIdx = cols.Select((v, i) => (v, i)).ToDictionary(x => x.v, x => x.i);
            var table = new double[rows.Count, cols.Count];
            foreach (var p in pairs)
            {
                table[rowIdx[p.RowLabel], colIdx[p.ColLabel]] += 1;
            }

            var rowTotals = new double[rows.Count];
            var colTotals = new double[cols.Count];
            double total = 0;
            for (var i = 0; i < rows.Count; i++)
            {
                for (var j = 0; j < cols.Count; j++)
                {
                    rowTotals[i] += table[i, j];
                    colTotals[j] += table[i, j];
                    total += table[i, j];
                }
            }

            double chi2 = 0;
            var lowExpected = false;
            for (var i = 0; i < rows.Count; i++)
            {
                for (var j = 0; j < cols.Count; j++)
                {
                    var expected = rowTotals[i] * colTotals[j] / total;
                    if (expected < 5) lowExpected = true;
                    var diff = table[i, j] - expected;
                    chi2 += expected > 0 ? diff * diff / expected : 0;
                }
            }

            var df = (rows.Count - 1) * (cols.Count - 1);
            if (df <= 0)
            {
                throw new ArgumentException("卡方独立性检验要求两个类别列各有至少 2 个取值（单取值列无变异）。");
            }

            var result = new TestResult { Method = "chi_square" };
            result.Statistic = chi2;
            result.Df = df;
            result.PValue = ChiSquareUpperP(chi2, df);
            result.EffectName = "cramers_v";
            result.EffectSize = total > 0 ? Math.Sqrt(chi2 / (total * Math.Min(rows.Count - 1, cols.Count - 1))) : 0;
            result.Groups.Add(Describe("行类别", rows.Select(r => (double)rowIdx[r] + 1).ToList()));
            result.TotalN = (int)total;
            result.PremiseNotes.Add("类别交叉表：行=" + string.Join("/", rows.Take(6)) + (rows.Count > 6 ? "…" : "") + "；列=" + string.Join("/", cols.Take(6)) + (cols.Count > 6 ? "…" : ""));
            if (lowExpected)
            {
                result.PremiseNotes.Add("存在期望频数<5 的单元格，卡方近似偏保守——合并稀薄类别或改用 Fisher 精确检验（Python 技能）更稳健");
            }

            return result;
        }

        /// <summary>按 alpha 标注显著性并输出结论行。</summary>
        public static void MarkSignificance(TestResult r, double alpha)
        {
            r.Alpha = alpha;
            r.Significant = !double.IsNaN(r.PValue) && r.PValue < alpha;
        }
    }
}
