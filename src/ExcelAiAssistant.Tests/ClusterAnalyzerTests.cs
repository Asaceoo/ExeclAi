using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 聚类分群纯逻辑单测（analyze_cluster，2026-09-14）：分离簇自动定 k、显式 k、标准化效果、样本不足 fail-closed。
    /// </summary>
    public class ClusterAnalyzerTests
    {
        /// <summary>三个远离的 2D 簇，每簇 3 点（确定性抖动）。</summary>
        private static List<double[]> ThreeClusters()
        {
            var pts = new List<double[]>();
            foreach (var (cx, cy) in new[] { (0.0, 0.0), (20.0, 20.0), (40.0, 5.0) })
            {
                pts.Add(new[] { cx, cy });
                pts.Add(new[] { cx + 0.5, cy + 0.3 });
                pts.Add(new[] { cx - 0.4, cy - 0.2 });
            }

            return pts;
        }

        [Fact]
        public void SeparatedClusters_AutoK_Equals3()
        {
            var pts = ThreeClusters();
            var r = ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, null, 6, true);
            Assert.Equal(3, r.K);
            Assert.True(r.BestSilhouette > 0.5);
            // 同簇点标签一致：前 3 点同簇、中 3 点同簇且与前 3 不同
            Assert.Equal(r.Labels[0], r.Labels[1]);
            Assert.Equal(r.Labels[1], r.Labels[2]);
            Assert.NotEqual(r.Labels[0], r.Labels[3]);
            Assert.NotEqual(r.Labels[0], r.Labels[6]);
            // k_trace 受「每簇至少 2 点」约束：n=9 → k 上限 4，覆盖 2..4
            Assert.Equal(3, r.KTrace.Count);
            Assert.Equal(2, r.KTrace[0].K);
            // 画像尺寸合计=总点数
            Assert.Equal(9, r.Profiles.Sum(p => p.Size));
        }

        [Fact]
        public void ExplicitK_Respected()
        {
            var pts = ThreeClusters();
            var r = ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, 2, 6, true);
            Assert.Equal(2, r.K);
            Assert.Single(r.KTrace);
            Assert.True(r.BestSilhouette > 0);
        }

        [Fact]
        public void Standardization_RecoversClusterOnSmallScaleDim()
        {
            // 簇结构只在小量纲维度上（y 方向差 4），x 方向差 1000——不标准化会被 x 主导丢簇。
            var pts = new List<double[]>
            {
                new[] { 0.0,    0.0 }, new[] { 1.0,   0.5 }, new[] { -0.5,  -0.3 },
                new[] { 1000.0, 4.0 }, new[] { 1001.0, 4.4 }, new[] { 999.0,  3.8 }
            };
            var std = ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, null, 6, true);
            Assert.True(std.BestSilhouette > 0.3);
            // 标准化后两簇正确分离：前 3 点同簇且与后 3 点不同
            Assert.Equal(std.Labels[0], std.Labels[1]);
            Assert.Equal(std.Labels[1], std.Labels[2]);
            Assert.NotEqual(std.Labels[0], std.Labels[3]);
        }

        [Fact]
        public void InsufficientPoints_Throws()
        {
            var pts = new List<double[]> { new[] { 1.0, 2.0 }, new[] { 3.0, 4.0 }, new[] { 5.0, 6.0 } };
            Assert.Throws<ArgumentException>(() => ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, null, 6, true));
        }

        [Fact]
        public void Silhouette_InRange_AndKMeansStable()
        {
            var pts = ThreeClusters();
            var r = ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, 3, 6, false);
            Assert.True(r.BestSilhouette >= -1 && r.BestSilhouette <= 1);
            // 确定性：同参重跑标签一致
            var r2 = ClusterAnalyzer.Compute(pts, new[] { "x", "y" }, 3, 6, false);
            Assert.Equal(r.Labels, r2.Labels);
        }

        [Fact]
        public void SingleDimension_Works()
        {
            // 1 维两簇：[0,1,2] 与 [100,101,102]
            var pts = new List<double[]>
            {
                new[] { 0.0 }, new[] { 1.0 }, new[] { 2.0 },
                new[] { 100.0 }, new[] { 101.0 }, new[] { 102.0 }
            };
            var r = ClusterAnalyzer.Compute(pts, new[] { "v" }, null, 6, true);
            Assert.True(r.K >= 2);
            // 0 附近与 100 附近的点不同簇
            Assert.NotEqual(r.Labels[0], r.Labels[5]);
        }
    }
}
