using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Skills;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// AI Skill P1（v3.8.1.77）回归锁：市场输入解析、兼容性三档、清单注入文本、市场来源标记。
    /// </summary>
    public sealed class SkillMarketResolverTests
    {
        [Theory]
        [InlineData("K-Dense-AI/scientific-agent-skills/seaborn", "K-Dense-AI", "scientific-agent-skills", "seaborn")]
        [InlineData("davila7/claude-code-templates", "davila7", "claude-code-templates", null)]
        [InlineData("owner/repo@seaborn", "owner", "repo", "seaborn")]
        [InlineData("https://www.skills.sh/vercel-labs/skills/find-skills", "vercel-labs", "skills", "find-skills")]
        [InlineData("https://github.com/anthropics/skills", "anthropics", "skills", null)]
        [InlineData("https://github.com/anthropics/skills/tree/main/skills/seaborn", "anthropics", "skills", "seaborn")]
        public void TryParse_parses_repo_inputs(string input, string owner, string repo, string skill)
        {
            Assert.True(SkillMarketResolver.TryParse(input, out var target));
            Assert.Equal(owner, target.Owner);
            Assert.Equal(repo, target.Repo);
            Assert.Equal(skill, target.Skill);
        }

        [Theory]
        [InlineData("")]
        [InlineData("onlyone")]
        [InlineData("https://example.com/x.zip")]
        [InlineData("owner/repo/extra/deep")]
        [InlineData("a..b/c")]
        public void TryParse_rejects_invalid_inputs(string input)
        {
            Assert.False(SkillMarketResolver.TryParse(input, out _));
        }

        [Fact]
        public void TryParse_trims_git_suffix_and_trailing_slash()
        {
            Assert.True(SkillMarketResolver.TryParse("https://github.com/anthropics/skills.git/", out var target));
            Assert.Equal("skills", target.Repo);
        }

        [Fact]
        public void BuildCodeloadUrls_tries_main_then_master()
        {
            var urls = SkillMarketResolver.BuildCodeloadUrls("anthropics", "skills");
            Assert.Equal(2, urls.Count);
            Assert.Contains("refs/heads/main", urls[0]);
            Assert.Contains("refs/heads/master", urls[1]);
        }
    }

    public sealed class SkillCompatJudgeTests
    {
        [Fact]
        public void Pure_knowledge_skill_is_usable()
        {
            var result = SkillCompatJudge.Judge(new string[0], new[] { "read_excel_range" });
            Assert.Equal(SkillCompatJudge.TierUsable, result.Tier);
            Assert.Empty(result.Missing);
        }

        [Fact]
        public void Declared_subset_is_usable()
        {
            var result = SkillCompatJudge.Judge(new[] { "read_excel_range", "write_excel_range" }, new[] { "read_excel_range", "write_excel_range", "undo" });
            Assert.Equal(SkillCompatJudge.TierUsable, result.Tier);
        }

        [Fact]
        public void Partial_missing_is_partial_with_missing_list()
        {
            var result = SkillCompatJudge.Judge(new[] { "read_excel_range", "bash" }, new[] { "read_excel_range" });
            Assert.Equal(SkillCompatJudge.TierPartial, result.Tier);
            Assert.Equal(new[] { "bash" }, result.Missing);
        }

        [Fact]
        public void All_missing_is_knowledge()
        {
            var result = SkillCompatJudge.Judge(new[] { "bash", "browser" }, new[] { "read_excel_range" });
            Assert.Equal(SkillCompatJudge.TierKnowledge, result.Tier);
            Assert.Equal(2, result.Missing.Count);
        }

        [Fact]
        public void Null_inputs_are_safe()
        {
            Assert.Equal(SkillCompatJudge.TierUsable, SkillCompatJudge.Judge(null, null).Tier);
        }
    }

    public sealed class SkillListInjectorTests
    {
        private static Skill Make(string name, string desc, bool builtIn = false, string source = null)
        {
            return new Skill { Name = name, Description = desc, Valid = true, IsBuiltIn = builtIn, Source = source };
        }

        [Fact]
        public void Empty_input_returns_empty_text()
        {
            Assert.Equal(string.Empty, SkillListInjector.BuildInjectionText(null));
            Assert.Equal(string.Empty, SkillListInjector.BuildInjectionText(new List<Skill>()));
        }

        [Fact]
        public void Lines_carry_source_and_description()
        {
            var text = SkillListInjector.BuildInjectionText(new List<Skill>
            {
                Make("seaborn", "统计图表配方", source: "market"),
                Make("excel-live", "活工作簿方法论", builtIn: true)
            });
            Assert.Contains("- /seaborn（市场）：统计图表配方", text);
            Assert.Contains("- /excel-live（内置）：活工作簿方法论", text);
        }

        [Fact]
        public void Long_description_is_truncated()
        {
            var text = SkillListInjector.BuildInjectionText(new List<Skill> { Make("a", new string('x', 100)) });
            Assert.Contains("…", text);
            Assert.DoesNotContain(new string('x', 61), text);
        }

        [Fact]
        public void More_than_30_skills_are_summarized()
        {
            var skills = new List<Skill>();
            for (var i = 0; i < 35; i++)
            {
                skills.Add(Make("skill-" + i.ToString("00"), "d"));
            }

            var text = SkillListInjector.BuildInjectionText(skills);
            Assert.Equal(31, text.Split('\n').Length); // 30 行 + 末尾换行后的空段
            Assert.Contains("未列出", text);
            Assert.DoesNotContain("/skill-34", text);
        }

        [Fact]
        public void Invalid_skills_are_excluded()
        {
            var invalid = new Skill { Name = "bad", Valid = false };
            var text = SkillListInjector.BuildInjectionText(new List<Skill> { invalid, Make("good", "d") });
            Assert.Contains("/good", text);
            Assert.DoesNotContain("/bad", text);
        }
    }

    public sealed class SkillMarketTextTests
    {
        [Fact]
        public void Wraps_plain_body_with_market_marks()
        {
            var result = SkillMarketText.EnsureMarketSource("# Title\nbody");
            Assert.StartsWith("---\nsource: market\nrisk_policy: high\nname: market-skill\n---", result);
            Assert.Contains("# Title", result);
        }

        [Fact]
        public void Injects_source_and_forced_risk_into_existing_front_matter()
        {
            var md = "---\nname: seaborn\ndescription: d\n---\n\nbody";
            var result = SkillMarketText.EnsureMarketSource(md);
            Assert.Contains("source: market", result);
            Assert.Contains("risk_policy: high", result);
            Assert.Contains("name: seaborn", result);
            Assert.DoesNotContain("产品既有契约优先", result); // 契约声明在安装落盘时追加，EnsureMarketSource 不附带
        }

        [Fact]
        public void Existing_low_risk_policy_is_overridden_to_high()
        {
            var md = "---\nname: x\nsource: market\nrisk_policy: low\n---\n\nbody";
            var result = SkillMarketText.EnsureMarketSource(md);
            Assert.DoesNotContain("risk_policy: low", result);
            Assert.Contains("risk_policy: high", result);
            var count = result.Split(new[] { "risk_policy:" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(1, count);
        }

        [Fact]
        public void Existing_source_is_not_duplicated()
        {
            var md = "---\nname: x\nsource: market\n---\n\nbody";
            var result = SkillMarketText.EnsureMarketSource(md);
            Assert.Equal(1, result.Split(new[] { "source:" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void Source_absent_with_existing_risk_policy_marks_correctly()
        {
            // 审查 F3：删除 risk_policy 行后原 insertAt 索引过期——组合场景回归锁
            var md = "---\nname: x\nrisk_policy: low\ndescription: d\n---\n\nbody";
            var result = SkillMarketText.EnsureMarketSource(md);
            Assert.DoesNotContain("risk_policy: low", result);
            Assert.Contains("source: market", result);
            Assert.Contains("risk_policy: high", result);
            Assert.Contains("name: x", result);
            Assert.Contains("description: d", result);
            var count = result.Split(new[] { "risk_policy:" }, StringSplitOptions.None).Length - 1;
            Assert.Equal(1, count);
        }

        [Fact]
        public void ParseDeclaredTools_supports_both_forms()
        {
            Assert.Equal(new[] { "a", "b" }, SkillMarketText.ParseDeclaredTools("---\nname: x\ntools: [a, b]\n---\n"));
            Assert.Equal(new[] { "a", "b" }, SkillMarketText.ParseDeclaredTools("---\nname: x\ntools: a, b\n---\n"));
            Assert.Empty(SkillMarketText.ParseDeclaredTools("---\nname: x\n---\n"));
        }

        [Fact]
        public void ParseSkillName_reads_front_matter_name()
        {
            Assert.Equal("seaborn", SkillMarketText.ParseSkillName("---\nname: seaborn\ndescription: d\n---\nbody"));
            Assert.Null(SkillMarketText.ParseSkillName("no front matter"));
        }
    }
}
