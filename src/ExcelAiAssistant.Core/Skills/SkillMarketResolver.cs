using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Skills
{
    /// <summary>
    /// 技能市场输入解析（AI Skill P1，FR-S1）：把设置页输入的
    /// 「owner/repo」「owner/repo/skill」「GitHub 链接」「skills.sh 技能页链接」统一解析为
    /// GitHub 仓库坐标与可选单技能名，供安装器下载 codeload zip（分支先 main 后 master 回退）。
    /// 单仓多技能仓库（如 K-Dense-AI/scientific-agent-skills 143 技能）依赖 skill 指定形态避免全量安装。
    /// 其它站点的 .zip / SKILL.md 直链不走本解析器（保持既有 InstallSkillFromUrl 通道）。
    /// </summary>
    public static class SkillMarketResolver
    {
        public sealed class RepoTarget
        {
            public RepoTarget(string owner, string repo, string skill)
            {
                Owner = owner;
                Repo = repo;
                Skill = skill;
            }

            public string Owner { get; }

            public string Repo { get; }

            /// <summary>可选：仓库内的单个技能目录名（目录名须与 SKILL.md 所在文件夹同名）。</summary>
            public string Skill { get; }
        }

        /// <summary>解析失败返回 false；不支持 http(s) 直链以外的主机（交给 URL 安装通道）。</summary>
        public static bool TryParse(string input, out RepoTarget target)
        {
            target = null;
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }

            var text = input.Trim().TrimEnd('/');
            if (text.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                Uri uri;
                try
                {
                    uri = new Uri(text);
                }
                catch (UriFormatException)
                {
                    return false;
                }

                var host = uri.Host.ToLowerInvariant();
                var isSkillsSh = host.EndsWith("skills.sh", StringComparison.OrdinalIgnoreCase);
                var isGithub = host.Equals("github.com", StringComparison.OrdinalIgnoreCase) ||
                               host.Equals("www.github.com", StringComparison.OrdinalIgnoreCase);
                if (!isSkillsSh && !isGithub)
                {
                    return false;
                }

                var segments = uri.AbsolutePath
                    .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                    .ToList();
                if (isSkillsSh)
                {
                    // skills.sh/owner/repo 或 skills.sh/owner/repo/skill
                    if (segments.Count < 2)
                    {
                        return false;
                    }

                    target = new RepoTarget(segments[0], segments[1], NormalizeSkill(segments.Count >= 3 ? segments[2] : null));
                    return true;
                }

                // github.com/owner/repo(/tree/branch/.../skillDir)
                if (segments.Count < 2)
                {
                    return false;
                }

                var repo = segments[1];
                if (repo.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
                {
                    repo = repo.Substring(0, repo.Length - 4);
                }

                var skillName = segments.Count >= 4 && segments[2].Equals("tree", StringComparison.OrdinalIgnoreCase)
                    ? segments[segments.Count - 1]
                    : null;
                target = new RepoTarget(segments[0], repo, NormalizeSkill(skillName));
                return true;
            }

            // 裸形态：owner/repo 或 owner/repo/skill（亦支持 owner/repo@skill）
            var parts = text.Split(new[] { '/', '@' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2 || parts.Length > 3)
            {
                return false;
            }

            foreach (var part in parts)
            {
                if (part.Contains("..") || part.IndexOfAny(new[] { '\\', ' ', '?', '#', '%' }) >= 0)
                {
                    return false;
                }
            }

            var repoPart = parts[1];
            if (repoPart.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
            {
                repoPart = repoPart.Substring(0, repoPart.Length - 4);
            }

            target = new RepoTarget(parts[0], repoPart, NormalizeSkill(parts.Length == 3 ? parts[2] : null));
            return true;
        }

        /// <summary>codeload zip 候选地址：分支先 main 后 master，安装器逐个回退。</summary>
        public static List<string> BuildCodeloadUrls(string owner, string repo)
        {
            return new List<string>
            {
                "https://codeload.github.com/" + owner + "/" + repo + "/zip/refs/heads/main",
                "https://codeload.github.com/" + owner + "/" + repo + "/zip/refs/heads/master"
            };
        }

        private static string NormalizeSkill(string skill)
        {
            if (string.IsNullOrWhiteSpace(skill))
            {
                return null;
            }

            var trimmed = skill.Trim().Trim('/');
            return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed;
        }
    }
}
