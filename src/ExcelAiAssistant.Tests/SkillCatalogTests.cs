using System;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Skills;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class SkillCatalogTests : IDisposable
    {
        private readonly string root;

        public SkillCatalogTests()
        {
            root = Path.Combine(Path.GetTempPath(), $"excel-ai-skills-{Guid.NewGuid():N}");
        }

        [Fact]
        public void Parses_front_matter_name_description_and_tools()
        {
            WriteSkill("data-clean", @"---
name: data-clean
description: 数据清洗助手。
version: 1.0.0
risk_policy: normal
tools: [clean_excel_data, undo_excel_operation]
---

# 数据清洗
按步骤执行。
");
            var catalog = CreateCatalog();

            var skill = Assert.Single(catalog.All);
            Assert.Equal("data-clean", skill.Name);
            Assert.Equal("数据清洗助手。", skill.Description);
            Assert.Equal("1.0.0", skill.Version);
            Assert.Equal("normal", skill.RiskPolicy);
            Assert.True(skill.Valid);
            Assert.Equal(2, skill.Tools.Count);
            Assert.Contains("clean_excel_data", skill.Tools);
            Assert.Contains("# 数据清洗", skill.Body);
        }

        [Fact]
        public void Marks_missing_front_matter_as_invalid()
        {
            WriteSkill("broken", "只有正文，没有 front matter。");
            var catalog = CreateCatalog();

            var skill = Assert.Single(catalog.All);
            Assert.False(skill.Valid);
            Assert.Contains("front matter", skill.InvalidReason);
        }

        [Fact]
        public void Marks_missing_description_as_invalid()
        {
            WriteSkill("no-desc", "---\nname: no-desc\n---\n正文");
            var catalog = CreateCatalog();

            var skill = Assert.Single(catalog.All);
            Assert.False(skill.Valid);
            Assert.Contains("name", skill.InvalidReason);
        }

        [Fact]
        public void Deduplicates_by_name_between_directories()
        {
            WriteSkill("data-clean", "---\nname: data-clean\ndescription: 用户版\ntools: []\n---\nA");
            var userDir = Path.Combine(root, "user");
            Directory.CreateDirectory(userDir);
            Directory.CreateDirectory(Path.Combine(userDir, "data-clean"));
            File.WriteAllText(
                Path.Combine(userDir, "data-clean", "SKILL.md"),
                "---\nname: data-clean\ndescription: 内置版解析替换\ntools: []\n---\nB");

            var catalog = new SkillCatalog()
                .AddSearchDirectory(Path.Combine(root, "builtin"))
                .AddSearchDirectory(userDir);
            catalog.Reload();

            var skill = Assert.Single(catalog.All);
            Assert.Equal("B", skill.Body.Trim());
        }

        [Fact]
        public void Search_matches_by_keyword()
        {
            WriteSkill("data-clean", "---\nname: data-clean\ndescription: 清洗数据、去空行\ntools: []\n---\n");
            var catalog = CreateCatalog();

            Assert.Single(catalog.Search("清洗"));
            Assert.Empty(catalog.Search("图表"));
        }

        [Fact]
        public void Parses_category_triggers_and_examples_metadata()
        {
            WriteSkill("report", @"---
name: report
description: 报表分析。
category: 分析
triggers: [分析, 报告, 洞察]
examples: [分析这份数据并出报告]
tools: []
---

正文");
            var catalog = CreateCatalog();

            var skill = Assert.Single(catalog.All);
            Assert.Equal("分析", skill.Category);
            Assert.Equal(new[] { "分析", "报告", "洞察" }, skill.Triggers);
            Assert.Equal(new[] { "分析这份数据并出报告" }, skill.Examples);
        }

        [Fact]
        public void Missing_category_defaults_to_generic_and_is_grouped()
        {
            WriteSkill("no-cat", "---\nname: no-cat\ndescription: 无分类\ncategory:  \ntools: []\n---\n");
            var catalog = CreateCatalog();

            var groups = catalog.GroupByCategory();
            var group = Assert.Single(groups);
            Assert.Equal("通用", group.Category);
            Assert.Single(group.Skills);
        }

        [Fact]
        public void Group_by_category_groups_and_sorts_skills()
        {
            WriteSkill("data-clean", "---\nname: data-clean\ndescription: 清洗\ncategory: 数据\ntools: []\n---\n");
            WriteSkill("report", "---\nname: report\ndescription: 报表\ncategory: 分析\ntools: []\n---\n");
            WriteSkill("vba", "---\nname: vba\ndescription: 自动化\ncategory: 自动化\ntools: []\n---\n");
            var catalog = CreateCatalog();

            var groups = catalog.GroupByCategory();
            Assert.Equal(3, groups.Count);
            Assert.Equal("数据", groups[0].Category);
            Assert.Equal("分析", groups[1].Category);
            Assert.Equal("自动化", groups[2].Category);
        }

        [Fact]
        public void Search_matches_by_trigger_not_in_name_or_description()
        {
            WriteSkill("data-clean", "---\nname: data-clean\ndescription: 数据清洗\ntriggers: [修日期, 脏数据]\ntools: []\n---\n");
            var catalog = CreateCatalog();

            Assert.Single(catalog.Search("修日期"));
            Assert.Empty(catalog.Search("出图"));
        }

        private SkillCatalog CreateCatalog()
        {
            var catalog = new SkillCatalog().AddSearchDirectory(root);
            catalog.Reload();
            return catalog;
        }

        private void WriteSkill(string name, string content)
        {
            var directory = Path.Combine(root, name);
            Directory.CreateDirectory(directory);
            File.WriteAllText(Path.Combine(directory, "SKILL.md"), content);
        }

        public void Dispose()
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}