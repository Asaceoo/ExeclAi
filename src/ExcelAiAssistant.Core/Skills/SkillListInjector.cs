using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAiAssistant.Core.Skills
{
    /// <summary>
    /// 技能清单注入文本构建（AI Skill P1，FR-S4/Q5 拍板）：把已装技能清单注入会话系统提示，
    /// 让模型知道有哪些技能、何时建议用户显式加载（/技能名）。渐进披露口径：清单只含
    /// 名称+一句话描述+来源（每技能一行，上限 30 行），技能全文只在用户确认加载后进入上下文。
    /// </summary>
    public static class SkillListInjector
    {
        public const int MaxLines = 30;

        private const int MaxDescriptionLength = 60;

        public static string BuildInjectionText(IReadOnlyList<Skill> skills)
        {
            if (skills == null || skills.Count == 0)
            {
                return string.Empty;
            }

            var valid = skills
                .Where(s => s != null && s.Valid && !string.IsNullOrWhiteSpace(s.Name))
                .OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (valid.Count == 0)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            var shown = 0;
            foreach (var skill in valid)
            {
                if (shown >= MaxLines)
                {
                    builder.Append("（另有 ").Append(valid.Count - shown).Append(" 个技能未列出）");
                    break;
                }

                var source = skill.IsBuiltIn ? "内置" : (skill.Source == "market" ? "市场" : "本地");
                var description = (skill.Description ?? string.Empty).Trim();
                if (description.Length > MaxDescriptionLength)
                {
                    description = description.Substring(0, MaxDescriptionLength) + "…";
                }

                builder.Append("- /").Append(skill.Name)
                    .Append("（").Append(source).Append("）：").Append(description)
                    .AppendLine();
                shown++;
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }
    }
}
