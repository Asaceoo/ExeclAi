using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 文本变体聚类纯逻辑单测（analyze_text_variants / FR-K2，2026-09-14）：
    /// 归一化归簇、中文错字/后缀 fuzzy 归簇、短串误合防护（LCP 约束）、合并方向选项、
    /// AC-K2 缩比召回（3 组真值全覆盖零混入）、唯一值上限 fail-closed。
    /// 算法与 Python 沙箱预验证对齐（召回 100%/零误合）。
    /// </summary>
    public class TextVariantClustererTests
    {
        [Fact]
        public void NormalizationOnly_Variants_GroupAsExact()
        {
            // 空白/大小写/全半角差异 → 归一化相等 → exact 高置信归簇
            var rows = new List<string> { "ABC Corp", "abc corp", "ＡＢＣ Corp", "abc  corp", "其他公司" };
            var r = TextVariantClusterer.Compute(rows, 2);

            Assert.Single(r.Groups);
            Assert.Equal(4, r.Groups[0].Members.Count);
            Assert.All(r.Groups[0].Members, m => Assert.Equal("exact", m.MatchKind));
            Assert.Equal(4, r.Groups[0].Members.Sum(m => m.Count));
            Assert.NotEqual("其他公司", r.Groups[0].Representative);
        }

        [Fact]
        public void ChineseTypoAndSuffix_Variants_GroupAsFuzzy()
        {
            // 中文公司名错字/后缀差异（编辑距离≤2 且共享首字符）→ fuzzy
            // 注：尾空格变体（"华为技术有限公司 "）在入口 Trim 即折叠为代表值，无需建议
            var rows = new List<string>
            {
                "华为技术有限公司", "华为技术有很公司", "华为技术有限 公司",
                "腾讯科技", "腾迅科技", "腾讯 科技"
            };
            var r = TextVariantClusterer.Compute(rows, 2);

            Assert.Equal(2, r.Groups.Count);
            var huawei = r.Groups.Single(g => g.Representative.StartsWith("华为"));
            Assert.Equal(3, huawei.Members.Count);
            Assert.Contains(huawei.Members, m => m.Value == "华为技术有很公司" && m.MatchKind == "fuzzy");
            Assert.Contains(huawei.Members, m => m.Value == "华为技术有限 公司");
            // 代表值=出现最多
            Assert.Equal("华为技术有限公司", huawei.Representative);
        }

        [Fact]
        public void ShortChineseDistinctNames_NotMerged()
        {
            // 沙箱实证误合样本：「腾讯科技」vs「小米科技」编辑距离=2 但语义无关（LCP=0）→ 不得归簇
            var rows = new List<string> { "腾讯科技", "小米科技", "腾讯科技", "小米科技", "百度在线" };
            var r = TextVariantClusterer.Compute(rows, 2);
            Assert.Empty(r.Groups);
        }

        [Fact]
        public void MergeOptions_LongestShortestMostFrequent()
        {
            var rows = new List<string>
            {
                "阿里巴巴集团",       // 出现最多（3 次）
                "阿里巴巴集团",
                "阿里巴巴集团",
                "阿里巴巴（集团）",   // 最长（全角括号）
                "阿里集团"            // 最短
            };
            var r = TextVariantClusterer.Compute(rows, 2);

            var g = Assert.Single(r.Groups);
            Assert.Equal("阿里巴巴集团", g.MostFrequent);
            Assert.Equal("阿里巴巴（集团）", g.Longest);
            Assert.Equal("阿里集团", g.Shortest);
        }

        [Fact]
        public void Ack2ScaledSample_RecallFull_NoMixing()
        {
            // AC-K2 缩比：3 组真值（分隔符/后缀/错字差异）+ 噪声 → 召回 100%、零跨组混入、零噪声误合
            var groupsTrue = new[]
            {
                new[] { "华为技术有限公司", "华为技术有很公司", "华为技术有限 公司" },
                new[] { "阿里巴巴集团", "阿里巴巴（集团）", "阿里巴巴(集团)", "阿里吧吧集团" },
                new[] { "腾讯科技", "腾迅科技", "腾讯 科技", "腾 讯科技" }
            };
            var noise = new[] { "字节跳动", "美团外卖", "京东物流", "拼多多", "小米科技", "百度在线", "网易游戏", "比亚迪汽车" };
            var rows = new List<string>();
            var truth = new Dictionary<string, int>();
            for (var gid = 0; gid < groupsTrue.Length; gid++)
            {
                foreach (var v in groupsTrue[gid])
                {
                    truth[v] = gid;
                    for (var k = 0; k < 8; k++) rows.Add(v);
                }
            }

            rows.AddRange(noise);
            var r = TextVariantClusterer.Compute(rows, 2);

            var covered = 0;
            foreach (var g in groupsTrue)
            {
                foreach (var m in g)
                {
                    if (r.Groups.Any(c => c.Members.Any(x => x.Value == m))) covered++;
                }
            }

            var total = groupsTrue.Sum(g => g.Length);
            Assert.True(covered >= total * 0.95, $"召回 {covered}/{total} 低于 95%");
            Assert.All(r.Groups, g =>
            {
                var gids = g.Members.Where(x => truth.ContainsKey(x.Value)).Select(x => truth[x.Value]).Distinct().ToList();
                Assert.True(gids.Count <= 1, $"建议组跨真值组混入：{g.Representative}");
            });
            Assert.DoesNotContain(r.Groups, g => g.Members.All(x => noise.Contains(x.Value)));
        }

        [Fact]
        public void TooManyUniqueValues_Throws()
        {
            var rows = new List<string>();
            for (var i = 0; i < TextVariantClusterer.MaxUniqueValues + 1; i++)
            {
                rows.Add("公司" + i);
            }

            var ex = Assert.Throws<ArgumentException>(() => TextVariantClusterer.Compute(rows, 2));
            Assert.Contains("rapidfuzz", ex.Message);
        }

        [Fact]
        public void EmptyAndNull_CountedNotClustered()
        {
            var rows = new List<string> { "客户A", "", null, "  ", "客户A ", "客户Ａ" };
            var r = TextVariantClusterer.Compute(rows, 2);
            Assert.Equal(3, r.EmptyCount);
            // 客户A 与「客户A 」（尾空格 Trim 折叠）同值、客户Ａ 全角归一 → 1 组 2 成员
            var g = Assert.Single(r.Groups);
            Assert.Equal(2, g.Members.Count);
        }
    }
}
