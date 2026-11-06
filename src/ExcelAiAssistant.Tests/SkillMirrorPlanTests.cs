using System;
using System.Linq;
using ExcelAiAssistant.Core.Skills;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// AI Skill 镜像下载计划回归锁（v3.8.1.79，用户真机实证：codeload 直连 TLS 握手被掐，
    /// jsDelivr 清单+按文件 CDN 兜底）。关键口径：jsDelivr flat 清单条目无 type 字段（条目全为文件），
    /// 计划器不得按 type 过滤；深层技能目录按路径段匹配。
    /// </summary>
    public sealed class SkillMirrorPlanTests
    {
        private const string RealShapedListing = @"{""default"":null,""files"":[
            {""name"":""/.claude/agents/a.md"",""hash"":""x1"",""size"":1},
            {""name"":""/cli-tool/components/skills/ent/excel-analysis/SKILL.md"",""hash"":""y"",""size"":2},
            {""name"":""/cli-tool/components/skills/ent/excel-analysis/scripts/gen.py"",""hash"":""z"",""size"":3},
            {""name"":""/skills/other/SKILL.md"",""hash"":""w"",""size"":4},
            {""name"":""/skills/other/data.json"",""hash"":""v"",""size"":5},
            {""name"":""/skills/other"",""type"":""directory""}
        ]}";

        [Fact]
        public void Skill_filter_matches_deep_directory_and_keeps_aux_files()
        {
            var plan = SkillMirrorPlan.BuildDownloadPlan(RealShapedListing, "excel-analysis");
            Assert.Equal(2, plan.Count);
            // 实现口径：RemotePath 不含前导斜杠（CdnFileUrl 两种形态均兼容）
            Assert.Contains(plan, f => f.RemotePath == "cli-tool/components/skills/ent/excel-analysis/SKILL.md"
                                        && f.LocalRelative == "SKILL.md");
            Assert.Contains(plan, f => f.RemotePath.EndsWith("/scripts/gen.py", StringComparison.Ordinal) && f.LocalRelative == "scripts/gen.py");
        }

        [Fact]
        public void No_skill_plans_all_skill_dirs()
        {
            var plan = SkillMirrorPlan.BuildDownloadPlan(RealShapedListing, null);
            Assert.Equal(4, plan.Count);
            Assert.NotNull(plan.FirstOrDefault(f => f.LocalRelative == "data.json"));
        }

        [Fact]
        public void Missing_skill_dir_reports_not_found()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                SkillMirrorPlan.BuildDownloadPlan(RealShapedListing, "no-such-skill"));
            Assert.Contains("未找到 SKILL.md", ex.Message);
        }

        [Fact]
        public void Over_cap_without_skill_is_rejected_with_guidance()
        {
            var sb = new System.Text.StringBuilder("{\"files\":[");
            for (var i = 0; i < 25; i++)
            {
                sb.Append("{\"name\":\"/skills/pack" + i.ToString("00") + "/SKILL.md\"},");
            }

            sb.Append("]}");
            var ex = Assert.Throws<InvalidOperationException>(() =>
                SkillMirrorPlan.BuildDownloadPlan(sb.ToString(), null));
            Assert.Contains("25 个技能", ex.Message);
            Assert.Contains("owner/repo/技能名", ex.Message);
        }

        [Fact]
        public void Urls_use_cdn_and_flat_listing()
        {
            Assert.Contains("data.jsdelivr.com/v1/package/gh/o/r@main/flat", SkillMirrorPlan.FlatListingUrl("o", "r", "main"));
            Assert.Equal("https://cdn.jsdelivr.net/gh/o/r@main/p/a.md", SkillMirrorPlan.CdnFileUrl("o", "r", "main", "/p/a.md"));
            Assert.Equal("https://cdn.jsdelivr.net/gh/o/r@main/p/a.md", SkillMirrorPlan.CdnFileUrl("o", "r", "main", "p/a.md"));
        }

        [Fact]
        public void BuildForDir_without_skill_md_throws()
        {
            var arr = JArray.Parse("[{\"name\":\"/d/readme.txt\"}]");
            Assert.Throws<InvalidOperationException>(() => SkillMirrorPlan.BuildForDir(arr, "/d"));
        }
    }
}
