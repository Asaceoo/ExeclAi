using System;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M27R2 审查：覆盖审计两条清单的条数必须与 CLAUDE.md §6 的陈述一致。
    /// 新增工具时若忘了同步文档，这里会亮红（防口径再漂）。
    /// </summary>
    public class AutoVerifyPolicyCountLockTests
    {
        // 与 CLAUDE.md §6 当前陈述对齐（改清单时必须同步改这里 + CLAUDE.md）
        // M29：+1 apply_correction（组合修正工具，自带 verify 判定，豁免外层自动自检）→ 89→90
        private const int DocumentedMappedCount = 48;
        private const int DocumentedExemptCount = 91;

        [Fact]
        public void Mapped_primary_tools_count_matches_documentation()
        {
            Assert.Equal(DocumentedMappedCount, AutoVerifyPolicy.MappedPrimaryTools.Count);
        }

        private static System.Collections.IDictionary Exemptions()
        {
            var f = typeof(AutoVerifyPolicy).GetField("DocumentedWriteToolExemptions",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
            return (System.Collections.IDictionary)f.GetValue(null);
        }

        [Fact]
        public void Documented_exemptions_count_matches_documentation()
        {
            Assert.Equal(DocumentedExemptCount, Exemptions().Count);
        }

        [Fact]
        public void Two_lists_do_not_overlap()
        {
            var exemptNames = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
            foreach (var k in Exemptions().Keys) { exemptNames.Add(Convert.ToString(k)); }
            var overlap = System.Linq.Enumerable.ToList(System.Linq.Enumerable.Where(AutoVerifyPolicy.MappedPrimaryTools, n => exemptNames.Contains(n)));
            Assert.True(overlap.Count == 0, "两清单不应重叠：" + string.Join(",", overlap));
        }
    }
}
