using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 聚类分群纯逻辑（analyze_cluster，2026-09-14）：
    /// Z-score 标准化 + K-Means++ 初始化 + Lloyd 迭代（多种子重启取最优惯性）+ 轮廓系数自动定 k。
    /// 不接触 COM，供服务层与单测使用。DBSCAN/层次聚类/高阶场景仍路由 python-advanced-analysis 技能。
    /// </summary>
    public static class ClusterAnalyzer
    {
        /// <summary>自动定 k 的下限/上限约束：k 范围 [2, kMax]，且每簇至少 2 个样本。</summary>
        public const int MinK = 2;
        public const int MaxK = 6;
        public const int MinPointsPerCluster = 2;

        public sealed class ClusterResult
        {
            public int K { get; set; }
            public double BestSilhouette { get; set; }
            public List<KSilhouette> KTrace { get; set; } = new List<KSilhouette>();
            /// <summary>每个样本（输入顺序）的簇编号，0 起。</summary>
            public int[] Labels { get; set; }
            /// <summary>各簇样本数（原顺序统计）。</summary>
            public List<ClusterProfile> Profiles { get; set; } = new List<ClusterProfile>();
            public bool Standardized { get; set; }
            public int TotalPoints { get; set; }
        }

        public sealed class KSilhouette
        {
            public int K { get; set; }
            public double Silhouette { get; set; }
        }

        public sealed class ClusterProfile
        {
            public int Cluster { get; set; }
            public int Size { get; set; }
            /// <summary>各特征均值（原始量纲）。</summary>
            public double[] MeanOriginal { get; set; }
        }

        /// <summary>
        /// 聚类主入口：features 为每行一个特征向量；explicitK 传入则固定 k（仍回填该 k 的轮廓系数）。
        /// 自动定 k 在 [2, kMax] 内逐 k 跑 KMeans+轮廓系数，取轮廓最大者；平局取更小 k。
        /// </summary>
        public static ClusterResult Compute(IReadOnlyList<double[]> features, string[] featureNames,
            int? explicitK, int kMax = MaxK, bool standardize = true, int seed = 42)
        {
            var n = features.Count;
            if (n < MinK * MinPointsPerCluster)
            {
                throw new ArgumentException($"有效样本仅 {n} 行，聚类至少需要 {MinK * MinPointsPerCluster} 行完整观测（每簇至少 {MinPointsPerCluster} 个）。");
            }

            var dim = features[0].Length;
            var data = standardize ? Standardize(features) : features.Select(f => f.ToArray()).ToList();
            var kUpper = Math.Min(Math.Max(MinK, kMax), Math.Max(MinK, n / MinPointsPerCluster));

            var trace = new List<KSilhouette>();
            int bestK = explicitK ?? MinK;
            double bestSil = -2;
            int[] bestLabels = null;
            var rand = new Random(seed);

            var kList = explicitK.HasValue
                ? new[] { Math.Max(MinK, explicitK.Value) }
                : Enumerable.Range(MinK, kUpper - MinK + 1).ToArray();
            foreach (var k in kList)
            {
                var labels = BestKMeans(data, k, rand);
                var sil = MeanSilhouette(data, labels, k);
                trace.Add(new KSilhouette { K = k, Silhouette = sil });
                if (sil > bestSil + 1e-12)
                {
                    bestSil = sil;
                    bestK = k;
                    bestLabels = labels;
                }
            }

            if (bestLabels == null)
            {
                bestLabels = BestKMeans(data, bestK, rand);
                bestSil = MeanSilhouette(data, bestLabels, bestK);
            }

            var result = new ClusterResult
            {
                K = bestK,
                BestSilhouette = bestSil,
                KTrace = trace,
                Labels = bestLabels,
                Standardized = standardize,
                TotalPoints = n
            };

            // 各簇画像：样本数 + 原始量纲均值
            for (var c = 0; c < bestK; c++)
            {
                var members = Enumerable.Range(0, n).Where(i => bestLabels[i] == c).ToList();
                var meanOrig = new double[dim];
                foreach (var d in Enumerable.Range(0, dim))
                {
                    meanOrig[d] = members.Count > 0 ? members.Average(i => features[i][d]) : 0;
                }

                result.Profiles.Add(new ClusterProfile
                {
                    Cluster = c,
                    Size = members.Count,
                    MeanOriginal = meanOrig
                });
            }

            return result;
        }

        /// <summary>Z-score 标准化（样本标准差，常数维度置 0）。</summary>
        public static List<double[]> Standardize(IReadOnlyList<double[]> features)
        {
            var n = features.Count;
            var dim = features[0].Length;
            var output = Enumerable.Range(0, n).Select(_ => new double[dim]).ToList();
            for (var d = 0; d < dim; d++)
            {
                var col = features.Select(f => f[d]).ToList();
                var mean = col.Average();
                var std = Math.Sqrt(col.Sum(x => (x - mean) * (x - mean)) / Math.Max(n - 1, 1));
                for (var i = 0; i < n; i++)
                {
                    output[i][d] = std > 0 ? (col[i] - mean) / std : 0;
                }
            }

            return output;
        }

        /// <summary>多种子重启跑 KMeans，取惯性（簇内平方和）最小的一轮；k-means++ 初始化保证确定性（种子递增）。</summary>
        private static int[] BestKMeans(IReadOnlyList<double[]> data, int k, Random rand)
        {
            int[] bestLabels = null;
            double bestInertia = double.MaxValue;
            for (var restart = 0; restart < 5; restart++)
            {
                var centers = KMeansPlusPlusInit(data, k, rand);
                var labels = Lloyd(data, centers, k);
                var inertia = Inertia(data, centers, labels, k);
                if (inertia < bestInertia)
                {
                    bestInertia = inertia;
                    bestLabels = labels;
                }
            }

            return bestLabels;
        }

        /// <summary>K-Means++ 概率初始化（D² 加权抽样）。</summary>
        private static double[][] KMeansPlusPlusInit(IReadOnlyList<double[]> data, int k, Random rand)
        {
            var n = data.Count;
            var dim = data[0].Length;
            var centers = new double[k][];
            centers[0] = (double[])data[rand.Next(n)].Clone();
            var dist = new double[n];
            for (var c = 1; c < k; c++)
            {
                double sum = 0;
                for (var i = 0; i < n; i++)
                {
                    var dMin = Distance(data[i], centers[0]);
                    for (var j = 1; j < c; j++)
                    {
                        var d = Distance(data[i], centers[j]);
                        if (d < dMin) dMin = d;
                    }

                    dist[i] = dMin * dMin;
                    sum += dist[i];
                }

                var target = rand.NextDouble() * sum;
                var acc = 0.0;
                var pick = n - 1;
                for (var i = 0; i < n; i++)
                {
                    acc += dist[i];
                    if (acc >= target)
                    {
                        pick = i;
                        break;
                    }
                }

                centers[c] = (double[])data[pick].Clone();
            }

            return centers;
        }

        /// <summary>Lloyd 迭代至指派稳定或 100 轮。</summary>
        private static int[] Lloyd(IReadOnlyList<double[]> data, double[][] centers, int k)
        {
            var n = data.Count;
            var labels = new int[n];
            for (var iter = 0; iter < 100; iter++)
            {
                var moved = false;
                for (var i = 0; i < n; i++)
                {
                    var best = 0;
                    var bestD = Distance(data[i], centers[0]);
                    for (var c = 1; c < k; c++)
                    {
                        var d = Distance(data[i], centers[c]);
                        if (d < bestD)
                        {
                            bestD = d;
                            best = c;
                        }
                    }

                    if (labels[i] != best) moved = true;
                    labels[i] = best;
                }

                for (var c = 0; c < k; c++)
                {
                    var members = Enumerable.Range(0, n).Where(i => labels[i] == c).ToList();
                    if (members.Count == 0) continue;
                    var dim = data[0].Length;
                    for (var d = 0; d < dim; d++)
                    {
                        centers[c][d] = members.Average(i => data[i][d]);
                    }
                }

                if (!moved) break;
            }

            return labels;
        }

        private static double Inertia(IReadOnlyList<double[]> data, double[][] centers, int[] labels, int k)
        {
            double sum = 0;
            for (var i = 0; i < data.Count; i++)
            {
                sum += Math.Pow(Distance(data[i], centers[labels[i]]), 2);
            }

            return sum;
        }

        /// <summary>平均轮廓系数（O(n²)，样本量上限由服务层控制）。</summary>
        public static double MeanSilhouette(IReadOnlyList<double[]> data, int[] labels, int k)
        {
            if (k <= 1) return 0;
            var n = data.Count;
            if (n <= k) return 0;
            double total = 0;
            var counted = 0;
            for (var i = 0; i < n; i++)
            {
                var same = new List<double>();
                var otherMin = double.MaxValue;
                for (var j = 0; j < n; j++)
                {
                    if (j == i) continue;
                    var d = Distance(data[i], data[j]);
                    if (labels[j] == labels[i]) same.Add(d);
                    else if (d < otherMin) otherMin = d;
                }

                if (same.Count == 0 || double.IsInfinity(otherMin)) continue;
                var a = same.Average();
                var b = otherMin;
                var denom = Math.Max(a, b);
                if (denom <= 0) continue;
                total += (b - a) / denom;
                counted++;
            }

            return counted > 0 ? total / counted : 0;
        }

        private static double Distance(double[] a, double[] b)
        {
            double sum = 0;
            for (var i = 0; i < a.Length; i++)
            {
                var diff = a[i] - b[i];
                sum += diff * diff;
            }

            return Math.Sqrt(sum);
        }
    }
}
