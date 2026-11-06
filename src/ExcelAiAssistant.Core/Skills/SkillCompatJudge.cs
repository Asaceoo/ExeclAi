using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Skills
{
    /// <summary>
    /// 技能兼容性三档判定（AI Skill P1，Q2 拍板）：安装/展示时把技能 front matter 声明的
    /// tools 与本机注册表工具名集合做 diff，输出「可用 / 部分可用 / 知识参考」。
    /// 纯知识技能（未声明 tools）恒为「可用」——技能是 prompt 层指导，写操作仍走产品确认门。
    /// </summary>
    public static class SkillCompatJudge
    {
        public const string TierUsable = "usable";

        public const string TierPartial = "partial";

        public const string TierKnowledge = "knowledge";

        public static SkillCompatResult Judge(IEnumerable<string> declaredTools, IEnumerable<string> availableTools)
        {
            var declared = (declaredTools ?? Enumerable.Empty<string>())
                .Select(t => (t ?? string.Empty).Trim())
                .Where(t => t.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (declared.Count == 0)
            {
                return new SkillCompatResult(TierUsable, "可用（纯知识技能）", new List<string>());
            }

            var available = new HashSet<string>(
                (availableTools ?? Enumerable.Empty<string>()).Select(t => (t ?? string.Empty).Trim()),
                StringComparer.OrdinalIgnoreCase);
            var missing = declared.Where(t => !available.Contains(t)).ToList();
            if (missing.Count == 0)
            {
                return new SkillCompatResult(TierUsable, "可用", new List<string>());
            }

            if (missing.Count < declared.Count)
            {
                return new SkillCompatResult(TierPartial, "部分可用", missing);
            }

            return new SkillCompatResult(TierKnowledge, "知识参考", missing);
        }
    }

    /// <summary>兼容性判定结果：tier=usable/partial/knowledge，label 为中文展示文案，missing 为本机缺失的工具名。</summary>
    public sealed class SkillCompatResult
    {
        public SkillCompatResult(string tier, string label, IReadOnlyList<string> missing)
        {
            Tier = tier;
            Label = label;
            Missing = missing;
        }

        public string Tier { get; }

        public string Label { get; }

        public IReadOnlyList<string> Missing { get; }
    }
}
