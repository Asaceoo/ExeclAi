using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 文本列变体聚类纯逻辑（analyze_text_variants，FR-K2，2026-09-14）：
    /// 探索式清洗建议——对文本列输出疑似变体组（空白/大小写/全半角归一化差异 + 编辑距离≤阈值且共享首字符）。
    /// 算法与 Python 沙箱预验证对齐（2026-09-14：3 组真值召回 100%、零跨组混入、零噪声误合；
    /// LCP≥1 约束用于挡短中文串编辑距离误合——「腾讯科技/小米科技」距离 2 但语义无关）。
    /// 指纹算法仅对拉丁字符有效，中文不做承诺（需求 v0.2 审查口径）。
    /// 合并动作不在此实现：AI 经用户确认后走 write_excel_cells_batch 改值语义（改值≠删行），本类只建议。
    /// 不接触 COM，供服务层与单测使用。
    /// </summary>
    public static class TextVariantClusterer
    {
        /// <summary>唯一值上限：贪心聚类 O(u²)，超出建议走 python-advanced-analysis（rapidfuzz）。</summary>
        public const int MaxUniqueValues = 1500;

        /// <summary>编辑距离 DP 的串长上限，超长串只做归一化精确归簇。</summary>
        private const int MaxNormalizeLength = 64;

        public sealed class VariantMember
        {
            public string Value { get; set; }
            public int Count { get; set; }
            /// <summary>exact=仅归一化差异（空白/大小写/全半角）；fuzzy=编辑距离差异（错字/后缀）。</summary>
            public string MatchKind { get; set; }
        }

        public sealed class VariantGroup
        {
            /// <summary>代表值：簇内出现次数最多的原值（平频取先出现）。</summary>
            public string Representative { get; set; }
            public List<VariantMember> Members { get; } = new List<VariantMember>();
            public int TotalCount { get; set; }
            /// <summary>合并方向选项：三个目标值（可能相同，相同则只有一个自然选择）。</summary>
            public string Longest { get; set; }
            public string Shortest { get; set; }
            public string MostFrequent { get; set; }
        }

        public sealed class ClusterResult
        {
            public int TotalRows { get; set; }
            public int UniqueValues { get; set; }
            public int EmptyCount { get; set; }
            public List<VariantGroup> Groups { get; } = new List<VariantGroup>();
        }

        /// <summary>
        /// values 为文本列逐行原值（允许 null/空白，计入 EmptyCount 不参与聚类）；
        /// maxDistance 为归一化后编辑距离阈值（1~4）。
        /// 唯一值超 MaxUniqueValues 抛 ArgumentException（调用方 fail-closed 引导 Python）。
        /// </summary>
        public static ClusterResult Compute(IReadOnlyList<string> values, int maxDistance)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            var maxDist = Math.Max(1, Math.Min(4, maxDistance));

            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            var firstSeen = new Dictionary<string, int>(StringComparer.Ordinal);
            var result = new ClusterResult();
            var order = 0;
            foreach (var raw in values)
            {
                result.TotalRows++;
                if (string.IsNullOrWhiteSpace(raw))
                {
                    result.EmptyCount++;
                    continue;
                }

                var v = raw.Trim();
                if (counts.TryGetValue(v, out var c))
                {
                    counts[v] = c + 1;
                }
                else
                {
                    counts[v] = 1;
                    firstSeen[v] = order++;
                }
            }

            result.UniqueValues = counts.Count;
            if (counts.Count > MaxUniqueValues)
            {
                throw new ArgumentException(
                    $"文本列唯一值 {counts.Count} 个，超过本地聚类上限 {MaxUniqueValues}（贪心两两比较防超时）；请缩小区域，或走 python-advanced-analysis 技能（rapidfuzz 高效批量比对）。");
            }

            // 归一化缓存
            var norms = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var v in counts.Keys)
            {
                norms[v] = Normalize(v);
            }

            // 贪心单遍聚类：按频次降序、平频按出现序
            var sorted = counts.Keys.OrderByDescending(v => counts[v])
                .ThenBy(v => firstSeen[v])
                .ToList();
            var clusters = new List<(string Rep, string Norm, List<string> Members)>();
            foreach (var v in sorted)
            {
                var nv = norms[v];
                var placed = false;
                foreach (var cl in clusters)
                {
                    var kind = MatchKindOf(nv, cl.Norm, maxDist);
                    if (kind == null) continue;
                    cl.Members.Add(v);
                    placed = true;
                    break;
                }

                if (!placed)
                {
                    clusters.Add((v, nv, new List<string> { v }));
                }
            }

            foreach (var cl in clusters.Where(c => c.Members.Count >= 2))
            {
                var g = new VariantGroup { Representative = cl.Rep };
                foreach (var m in cl.Members)
                {
                    var kind = MatchKindOf(norms[m], cl.Norm, maxDist) ?? "fuzzy";
                    g.Members.Add(new VariantMember { Value = m, Count = counts[m], MatchKind = kind });
                    g.TotalCount += counts[m];
                }

                var byLen = cl.Members.OrderByDescending(m => m.Length).ThenBy(m => firstSeen[m]).ToList();
                g.Longest = byLen.First();
                g.Shortest = byLen.Last();
                g.MostFrequent = cl.Rep;
                result.Groups.Add(g);
            }

            return result;
        }

        /// <summary>
        /// 归簇判定：返回 "exact"（归一化相等）/ "fuzzy"（编辑距离≤阈值且共享首字符）/ null（不同簇）。
        /// LCP≥1 约束：中文变体几乎总共享首字，可挡「腾讯科技/小米科技」类短串距离误合（沙箱实证）。
        /// </summary>
        private static string MatchKindOf(string normA, string normB, int maxDist)
        {
            if (string.Equals(normA, normB, StringComparison.Ordinal)) return "exact";
            var dist = EditDistance(normA, normB);
            if (dist <= maxDist && LongestCommonPrefix(normA, normB) >= 1) return "fuzzy";
            return null;
        }

        /// <summary>归一化：去空白（含全角空格）+ 全角转半角 + 小写折叠（保留中文原样）。</summary>
        public static string Normalize(string s)
        {
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (ch == ' ' || ch == '\t' || ch == '\u3000') continue;
                var c = ch;
                var code = (int)c;
                if (code >= 0xFF01 && code <= 0xFF5E)
                {
                    c = (char)(code - 0xFEE0);
                }

                sb.Append(char.ToLowerInvariant(c));
            }

            return sb.ToString();
        }

        /// <summary>经典 DP 编辑距离；长度差 >4 剪枝；任一串超 MaxNormalizeLength 时只判相等与否。</summary>
        public static int EditDistance(string a, string b)
        {
            if (string.Equals(a, b, StringComparison.Ordinal)) return 0;
            if (Math.Abs(a.Length - b.Length) > 4) return int.MaxValue;
            if (a.Length > MaxNormalizeLength || b.Length > MaxNormalizeLength)
            {
                return int.MaxValue; // 超长串不做模糊归簇（exact 已在 MatchKindOf 前置判定）
            }

            var prev = new int[b.Length + 1];
            for (var j = 0; j <= b.Length; j++)
            {
                prev[j] = j;
            }

            for (var i = 1; i <= a.Length; i++)
            {
                var cur = new int[b.Length + 1];
                cur[0] = i;
                for (var j = 1; j <= b.Length; j++)
                {
                    var sub = prev[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1);
                    cur[j] = Math.Min(Math.Min(prev[j] + 1, cur[j - 1] + 1), sub);
                }

                prev = cur;
            }

            return prev[b.Length];
        }

        private static int LongestCommonPrefix(string a, string b)
        {
            var n = Math.Min(a.Length, b.Length);
            for (var i = 0; i < n; i++)
            {
                if (a[i] != b[i]) return i;
            }

            return n;
        }
    }
}
