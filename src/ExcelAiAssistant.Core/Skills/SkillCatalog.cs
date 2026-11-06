using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAiAssistant.Core.Skills
{
    public sealed class SkillCatalog
    {
        private readonly List<DirectoryEntry> searchDirectories = new List<DirectoryEntry>();
        private readonly Dictionary<string, Skill> skillsByName =
            new Dictionary<string, Skill>(StringComparer.OrdinalIgnoreCase);

        public SkillCatalog()
        {
        }

        public SkillCatalog AddSearchDirectory(string directory, bool builtIn = false)
        {
            if (!string.IsNullOrWhiteSpace(directory))
            {
                searchDirectories.Add(new DirectoryEntry(directory, builtIn));
            }

            return this;
        }

        public void Reload()
        {
            skillsByName.Clear();
            var entries = new List<Skill>();
            foreach (var entry in searchDirectories)
            {
                if (!Directory.Exists(entry.Path))
                {
                    continue;
                }

                foreach (var skillDirectory in Directory.EnumerateDirectories(entry.Path))
                {
                    var markdownPath = Path.Combine(skillDirectory, "SKILL.md");
                    if (!File.Exists(markdownPath))
                    {
                        continue;
                    }

                    var skill = Parse(markdownPath);
                    if (skill != null)
                    {
                        skill.IsBuiltIn = entry.BuiltIn;
                        skill.Source = entry.BuiltIn ? "builtin" : (string.IsNullOrWhiteSpace(skill.Source) ? "local" : skill.Source);
                        entries.Add(skill);
                    }
                }
            }

            foreach (var skill in entries)
            {
                if (skill == null || string.IsNullOrWhiteSpace(skill.Name))
                {
                    continue; // 无效条目（缺 name/description）不占用字典，避免 null 键解析异常
                }

                skillsByName[skill.Name] = skill;
            }
        }

        public IReadOnlyList<Skill> All => skillsByName.Values.ToList();

        /// <summary>按分类分组的技能目录（用于侧边栏分组展示），组内按名称排序，未分类归入「通用」。</summary>
        public IReadOnlyList<SkillGroup> GroupByCategory()
        {
            var order = new[] { "数据", "分析", "格式", "自动化", "人事" };
            return All
                .Where(skill => skill.Valid)
                .GroupBy(skill => skill.Category ?? "通用")
                .Select(group => new SkillGroup(
                    group.Key,
                    group.OrderBy(skill => skill.Name, StringComparer.OrdinalIgnoreCase).ToList()))
                .OrderBy(group =>
                {
                    var idx = Array.IndexOf(order, group.Category);
                    return idx < 0 ? order.Length : idx;
                })
                .ToList();
        }

        public bool TryGet(string name, out Skill skill)
        {
            return skillsByName.TryGetValue(name ?? string.Empty, out skill);
        }

        public IReadOnlyList<Skill> Search(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return All;
            }

            var keywords = query.Split(new[] { ' ', '，', ',', '、' }, StringSplitOptions.RemoveEmptyEntries);
            return All
                .Where(skill => keywords.Any(keyword =>
                    skill.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    skill.Description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    skill.Triggers.Any(trigger => trigger.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)))
                .ToList();
        }

        private static Skill Parse(string markdownPath)
        {
            var text = File.ReadAllText(markdownPath, Encoding.UTF8);
            var frontMatter = ExtractFrontMatter(text);
            if (frontMatter == null)
            {
                return new Skill
                {
                    Name = Path.GetFileName(Path.GetDirectoryName(markdownPath)),
                    Directory = Path.GetDirectoryName(markdownPath),
                    Valid = false,
                    InvalidReason = "缺少 front matter"
                };
            }

            var name = GetValue(frontMatter, "name");
            var description = GetValue(frontMatter, "description");
            var version = GetValue(frontMatter, "version");
            var riskPolicy = GetValue(frontMatter, "risk_policy") ?? "normal";
            var tools = GetList(frontMatter, "tools");
            var category = GetValue(frontMatter, "category");
            var type = GetValue(frontMatter, "type");
            var source = GetValue(frontMatter, "source");
            var triggers = GetList(frontMatter, "triggers");
            var examples = GetList(frontMatter, "examples");
            var body = StripFrontMatter(text);

            var skill = new Skill
            {
                Name = name,
                Description = description,
                Version = version,
                RiskPolicy = riskPolicy,
                Tools = tools,
                Category = string.IsNullOrWhiteSpace(category) ? "通用" : category.Trim(),
                Type = string.IsNullOrWhiteSpace(type) ? "quick" : type.Trim(),
                Triggers = triggers,
                Examples = examples,
                Directory = Path.GetDirectoryName(markdownPath),
                Body = body,
                Source = source,
                Valid = true
            };

            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(description))
            {
                skill.Valid = false;
                skill.InvalidReason = "缺少 name 或 description";
            }

            return skill;
        }

        private static Dictionary<string, string> ExtractFrontMatter(string text)
        {
            if (!text.StartsWith("---", StringComparison.Ordinal))
            {
                return null;
            }

            var end = text.IndexOf("\n---", 3, StringComparison.Ordinal);
            if (end < 0)
            {
                return null;
            }

            var block = text.Substring(3, end - 3);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in block.Split('\n'))
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    continue;
                }

                var colon = trimmed.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                var key = trimmed.Substring(0, colon).Trim();
                var value = trimmed.Substring(colon + 1).Trim();
                result[key] = value.Trim('"', '\'');
            }

            return result;
        }

        private static string StripFrontMatter(string text)
        {
            if (!text.StartsWith("---", StringComparison.Ordinal))
            {
                return text;
            }

            var end = text.IndexOf("\n---", 3, StringComparison.Ordinal);
            if (end < 0)
            {
                return text;
            }

            return text.Substring(end + 4).Trim();
        }

        private static string GetValue(Dictionary<string, string> frontMatter, string key)
        {
            return frontMatter.TryGetValue(key, out var value) ? value : null;
        }

        private static List<string> GetList(Dictionary<string, string> frontMatter, string key)
        {
            var value = GetValue(frontMatter, key);
            if (string.IsNullOrWhiteSpace(value))
            {
                return new List<string>();
            }

            return value
                .Trim('[', ']')
                .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim().Trim('"', '\''))
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToList();
        }

        private sealed class DirectoryEntry
        {
            public DirectoryEntry(string path, bool builtIn)
            {
                Path = path;
                BuiltIn = builtIn;
            }

            public string Path { get; }

            public bool BuiltIn { get; }
        }
    }

    /// <summary>技能目录中的一个分类分组。</summary>
    public sealed class SkillGroup
    {
        public SkillGroup(string category, IReadOnlyList<Skill> skills)
        {
            Category = category;
            Skills = skills;
        }

        public string Category { get; }

        public IReadOnlyList<Skill> Skills { get; }
    }
}