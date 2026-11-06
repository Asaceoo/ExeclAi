using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Security
{
    /// <summary>策略判定结果：允许 / 拒绝 / 不在白名单。</summary>
    public enum PolicyDecision
    {
        Allow,
        Deny,
        NotAllowlisted
    }

    /// <summary>
    /// 受控执行的白/黑名单策略（纯逻辑，可单元测试）：
    /// 黑名单优先，命中即拒绝；白名单为空视为"一律不允许"；否则必须命中白名单。
    /// 匹配忽略大小写（路径匹配同时对全路径做规范化比较）。
    /// </summary>
    public static class ControlledExecutionPolicy
    {
        public static PolicyDecision Evaluate(
            string identity,
            IEnumerable<string> allowlist,
            IEnumerable<string> denylist,
            bool fullPath = false)
        {
            if (string.IsNullOrWhiteSpace(identity))
            {
                return PolicyDecision.NotAllowlisted;
            }

            var normalized = fullPath ? NormalizePath(identity) : identity.Trim().ToLowerInvariant();
            var denyList = (denylist ?? Array.Empty<string>()).ToList();
            var allowList = (allowlist ?? Array.Empty<string>()).ToList();

            if (Matches(normalized, denyList, fullPath))
            {
                return PolicyDecision.Deny;
            }

            if (allowList.Count == 0)
            {
                return PolicyDecision.NotAllowlisted;
            }

            return Matches(normalized, allowList, fullPath)
                ? PolicyDecision.Allow
                : PolicyDecision.NotAllowlisted;
        }

        private static bool Matches(string identity, List<string> rules, bool fullPath)
        {
            foreach (var rule in rules)
            {
                if (string.IsNullOrWhiteSpace(rule))
                {
                    continue;
                }

                var candidate = fullPath ? NormalizePath(rule) : rule.Trim().ToLowerInvariant();
                if (string.Equals(identity, candidate, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizePath(string value)
        {
            return PathHelpers.Normalize(value);
        }
    }

    internal static class PathHelpers
    {
        public static string Normalize(string value)
        {
            var trimmed = (value ?? string.Empty).Replace('/', '\\').Trim().TrimEnd('\\');
            try
            {
                return System.IO.Path.GetFullPath(trimmed).TrimEnd('\\').ToLowerInvariant();
            }
            catch (Exception)
            {
                return trimmed.ToLowerInvariant();
            }
        }
    }
}