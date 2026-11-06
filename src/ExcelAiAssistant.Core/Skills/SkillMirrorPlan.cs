using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Skills
{
    /// <summary>
    /// 技能仓库镜像下载计划（AI Skill P1 补强，用户真机实证：国内网络对 codeload.github.com
    /// 直连常被 TLS 握手掐断「未能创建 SSL/TLS 安全通道」）。jsDelivr CDN 在国内可达且本会话
    /// 已实证：清单 API（data.jsdelivr.com flat）+ 按文件 CDN（cdn.jsdelivr.net/gh）两段式
    /// 可以替代整仓 zip。纯逻辑：从 flat listing JSON 计算需下载的文件计划。
    /// </summary>
    public static class SkillMirrorPlan
    {
        /// <summary>未指定单技能时允许安装的技能目录数上限（与主安装通道护栏一致）。</summary>
        public const int MaxSkillDirs = 20;

        public static string FlatListingUrl(string owner, string repo, string branch)
        {
            return "https://data.jsdelivr.com/v1/package/gh/" + owner + "/" + repo + "@" + branch + "/flat";
        }

        public static string CdnFileUrl(string owner, string repo, string branch, string remotePath)
        {
            var clean = remotePath.StartsWith("/") ? remotePath.Substring(1) : remotePath;
            return "https://cdn.jsdelivr.net/gh/" + owner + "/" + repo + "@" + branch + "/" + clean;
        }

        /// <summary>
        /// 从 jsDelivr flat listing JSON 计算下载计划。skillDir 非空时只取路径包含
        /// “/{skillDir}/SKILL.md”的那一个技能目录（兼容仓库深层目录）；为空时取全部 SKILL.md
        /// 所在目录（超过 MaxSkillDirs 抛 InvalidOperationException，引导 owner/repo/技能名 形态）。
        /// 返回 (远端路径, 相对本地路径) 列表，远端路径不含前导斜杠。
        /// </summary>
        public static List<SkillMirrorFile> BuildDownloadPlan(string flatJson, string skillDir)
        {
            var result = new List<SkillMirrorFile>();
            var root = JToken.Parse(flatJson);
            var files = root["files"] as JArray;
            if (files == null)
            {
                throw new InvalidOperationException("jsDelivr 清单格式异常：缺少 files 数组。");
            }

            var skillDirs = new List<string>();
            foreach (var entry in files)
            {
                var name = (string)entry["name"];
                var type = (string)entry["type"];
                // jsDelivr flat 清单两种形态：带 type（file/directory）或不带 type（条目全为文件）——
                // 仅显式 directory 才跳过，缺失 type 按文件处理（本仓清单实证 1251 条均无 type 字段）
                if (string.IsNullOrEmpty(name) || type == "directory" ||
                    !name.EndsWith("/SKILL.md", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var dir = name.Substring(0, name.Length - "/SKILL.md".Length).TrimStart('/');
                if (!string.IsNullOrEmpty(skillDir))
                {
                    // 深层目录匹配：路径段中出现 /{skillDir}/ 即命中（cli-tool/components/skills/x/excel-analysis/SKILL.md）；
                    // net48 string.Contains 无比较重载，用 IndexOf
                    if (name.IndexOf("/" + skillDir + "/SKILL.md", StringComparison.OrdinalIgnoreCase) < 0 &&
                        !dir.Equals(skillDir, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    return BuildForDir(files, dir);
                }

                skillDirs.Add(dir);
            }

            if (skillDirs.Count > MaxSkillDirs)
            {
                throw new InvalidOperationException(
                    "该仓库经镜像探测包含 " + skillDirs.Count + " 个技能，全量安装体量过大。请用 owner/repo/技能名 形态指定安装（例如 owner/repo/" +
                    (skillDirs.Count > 0 ? LastSegment(skillDirs[0]) : "技能名") + "）。");
            }

            if (!string.IsNullOrEmpty(skillDir))
            {
                throw new InvalidOperationException(
                    "镜像清单中未找到 SKILL.md：仓库清单共 " + files.Count + " 项，其中没有技能目录「" + skillDir + "」，请核对技能目录名（owner/repo/技能名）。");
            }

            foreach (var dir in skillDirs)
            {
                result.AddRange(BuildForDir(files, dir));
            }

            return result;
        }

        /// <summary>收集单个技能目录下的全部文件计划（本地相对路径以技能目录为根）。</summary>
        public static List<SkillMirrorFile> BuildForDir(JArray files, string dir)
        {
            var result = new List<SkillMirrorFile>();
            var prefix = dir.StartsWith("/") ? dir : "/" + dir;
            foreach (var entry in files)
            {
                var name = (string)entry["name"];
                var type = (string)entry["type"];
                if (string.IsNullOrEmpty(name) || type == "directory")
                {
                    continue;
                }

                var normalized = name.StartsWith("/") ? name : "/" + name;
                if (!normalized.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var localRelative = normalized.Substring(prefix.Length + 1);
                result.Add(new SkillMirrorFile(normalized.TrimStart('/'), localRelative));
            }

            if (!result.Any(f => f.LocalRelative.Equals("SKILL.md", StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException("技能目录 " + dir + " 下未找到 SKILL.md。");
            }

            return result;
        }

        private static string LastSegment(string dir)
        {
            var idx = dir.LastIndexOf('/');
            return idx < 0 ? dir : dir.Substring(idx + 1);
        }
    }

    /// <summary>镜像下载计划中的单个文件：远端仓库路径（不含前导斜杠）与技能目录内的相对本地路径。</summary>
    public sealed class SkillMirrorFile
    {
        public SkillMirrorFile(string remotePath, string localRelative)
        {
            RemotePath = remotePath;
            LocalRelative = localRelative;
        }

        public string RemotePath { get; }

        public string LocalRelative { get; }
    }
}
