using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Skills
{
    /// <summary>
    /// 市场来源技能文本标记（AI Skill P1，Q4/Q7 拍板）：外部（市场）技能安装时在 front matter
    /// 注入 <c>source: market</c> 与 <c>risk_policy: high</c>（外部技能强制 external 信任级——
    /// 第三方自声明的低风险策略不可信，一律覆盖为 high）。产品契约优先声明由安装器追加到正文尾部。
    /// </summary>
    public static class SkillMarketText
    {
        public const string MarketSource = "market";

        public const string ForcedRiskPolicy = "high";

        /// <summary>来源与风险策略标记：无 front matter 时包裹；有则按行重建——移除既有 risk_policy 行、缺 source 补 source: market、末尾注入 risk_policy: high（行级重建免疫删除后的索引漂移）。</summary>
        public static string EnsureMarketSource(string markdown)
        {
            var text = markdown ?? string.Empty;
            var trimmed = text.TrimStart();
            if (!trimmed.StartsWith("---", StringComparison.Ordinal))
            {
                var name = ParseSkillName(text) ?? "market-skill";
                return "---\nsource: " + MarketSource + "\nrisk_policy: " + ForcedRiskPolicy + "\nname: " + name + "\n---\n\n" + text;
            }

            var end = text.IndexOf("\n---", 3, StringComparison.Ordinal);
            if (end < 0)
            {
                return text; // front matter 未闭合：SkillCatalog 同样无法解析，保持原样
            }

            var block = text.Substring(3, end - 3);
            var lines = new List<string>(block.Split('\n'));
            bool IsKey(string line, string key)
            {
                var t = line.Trim();
                return t.StartsWith(key, StringComparison.OrdinalIgnoreCase) && t.IndexOf(':') > 0;
            }

            var hasSource = lines.Any(l => IsKey(l, "source"));
            lines.RemoveAll(l => IsKey(l, "risk_policy"));
            var insertIdx = lines.Count > 0 ? 1 : 0; // 跳过 block 首个空段（--- 后的换行）
            if (!hasSource)
            {
                lines.Insert(Math.Min(insertIdx, lines.Count), "source: " + MarketSource);
            }

            lines.Insert(Math.Min(insertIdx, lines.Count), "risk_policy: " + ForcedRiskPolicy);
            return "---" + string.Join("\n", lines) + text.Substring(end);
        }

        /// <summary>快速读取 SKILL.md 文本的首个 /^name\s*:/ 字段（不依赖完整解析）。</summary>
        public static string ParseSkillName(string text)
        {
            var trimmed = (text ?? string.Empty).TrimStart();
            if (!trimmed.StartsWith("---", StringComparison.Ordinal))
            {
                return null;
            }

            var match = Regex.Match(trimmed, @"(?m)^name\s*:\s*([^\r\n]+)");
            return match.Success ? match.Groups[1].Value.Trim().Trim('"', '\'') : null;
        }

        /// <summary>解析 front matter 的 tools 列表（供兼容性三档判定），格式兼容 [a, b] 与 a, b。</summary>
        public static List<string> ParseDeclaredTools(string markdown)
        {
            var result = new List<string>();
            var trimmed = (markdown ?? string.Empty).TrimStart();
            if (!trimmed.StartsWith("---", StringComparison.Ordinal))
            {
                return result;
            }

            var end = trimmed.IndexOf("\n---", 3, StringComparison.Ordinal);
            if (end < 0)
            {
                return result;
            }

            var block = trimmed.Substring(3, end - 3);
            foreach (var line in block.Split('\n'))
            {
                var t = line.Trim();
                if (!t.StartsWith("tools", StringComparison.OrdinalIgnoreCase) || t.IndexOf(':') < 0)
                {
                    continue;
                }

                var value = t.Substring(t.IndexOf(':') + 1).Trim().Trim('[', ']');
                result.AddRange(value
                    .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(item => item.Trim().Trim('"', '\''))
                    .Where(item => item.Length > 0));
                break;
            }

            return result;
        }

        /// <summary>产品契约优先声明（追加到技能正文尾部；契约清单与冻结稿 FR-S5 一致）。</summary>
        public static string ContractAppendix =>
            "\n\n> 本技能在本产品内执行时，产品既有契约优先：中文字体走 msyh.ttc 探测链（fail-closed）、" +
            "写操作走确认门与回滚（自动自检 auto-verify 照常）、统计口径以产品分析工具为准、禁止编造数据。";

        private static bool HasFrontMatterKey(string text, string key)
        {
            var end = text.IndexOf("\n---", 3, StringComparison.Ordinal);
            if (end < 0)
            {
                return false;
            }

            var block = text.Substring(3, end - 3);
            foreach (var line in block.Split('\n'))
            {
                var t = line.Trim();
                if (t.StartsWith(key, StringComparison.OrdinalIgnoreCase) && t.IndexOf(':') > 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
