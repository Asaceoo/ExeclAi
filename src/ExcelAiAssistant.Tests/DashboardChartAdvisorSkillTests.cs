using System;
using System.Linq;
using System.IO;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class DashboardChartAdvisorSkillTests
    {
        private static string LoadSkill()
        {
            var root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", ".."));
            var path = Path.Combine(root, "ExcelAiAssistant.Addin", "skills", "dashboard-chart-advisor", "SKILL.md");
            Assert.True(File.Exists(path), "图表顾问 Skill 源文件不存在：" + path);
            return File.ReadAllText(path);
        }

        [Fact]
        public void Skill_contains_four_dimension_selection_contract()
        {
            var text = LoadSkill();
            Assert.Contains("数据要求", text);
            Assert.Contains("适合场景", text);
            Assert.Contains("主要作用", text);
            Assert.Contains("分析目的", text);
            Assert.Contains("按分析目的选图", text);
        }

        [Theory]
        [InlineData("排序条形图")]
        [InlineData("折线图")]
        [InlineData("100%堆积柱形图")]
        [InlineData("树状图")]
        [InlineData("直方图")]
        [InlineData("柏拉图")]
        [InlineData("箱线图")]
        [InlineData("散点图")]
        [InlineData("瀑布图")]
        [InlineData("漏斗图")]
        [InlineData("填充地图")]
        [InlineData("数据透视图")]
        public void Skill_covers_key_chart_families(string chartName)
        {
            Assert.Contains(chartName, LoadSkill());
        }

        [Fact]
        public void Skill_preserves_tool_routes_and_rejection_rules()
        {
            var text = LoadSkill();
            Assert.Contains("detect_dashboard_data_readiness", text);
            Assert.Contains("create_chart_advanced", text);
            Assert.Contains("verify_dashboard", text);
            Assert.Contains("verify_chart", text);
            Assert.Contains("拒绝漏斗图", text);
            Assert.Contains("不得称为Excel独立原生图表类型", text);
        }

        [Fact]
        public void Skill_requires_mapping_suggestion_with_reasons_before_render()
        {
            var text = LoadSkill();
            Assert.Contains("suggest_chart_mapping", text);
            Assert.Contains("哪个做数据源", text);
            Assert.Contains("理由", text);
        }

        [Fact]
        public void Skill_requires_selfcheck_reply_contract_with_evidence()
        {
            var text = LoadSkill();
            Assert.Contains("verify_dashboard_deep", text);
            Assert.Contains("自检四问", text);
            Assert.Contains("数据源", text);
            Assert.Contains("依据", text);
        }

        [Fact]
        public void Skill_rejects_map_charts_for_china()
        {
            var text = LoadSkill();
            Assert.Contains("中国区", text);
            Assert.Contains("停用", text);
            Assert.Contains("色阶", text);
        }

        [Fact]
        public void ChartAdvanced_schema_type_enum_matches_implementation_valid_types()
        {
            // FR-4.2 一致性兜底：Schema 枚举必须与 ExcelPivotChartService.CreateChartAdvanced 的
            // validTypes 完全一致（2026-09-11 审查发现 histogram 漂移：代码 11 类、Schema 10 类，
            // AI 按 Schema 传 histogram 被拒）。本测试锁定全集，防单侧增删再漂移。
            var schema = Newtonsoft.Json.Linq.JObject.FromObject(ExcelAiAssistant.Excel.M3ToolSchemas.ChartAdvanced());
            var enumProp = (Newtonsoft.Json.Linq.JArray)schema["properties"]["type"]["enum"];
            var schemaTypes = enumProp.Select(t => (string)t).OrderBy(t => t).ToList();
            var expected = new[] { "bullet", "waffle", "slider", "funnel", "butterfly", "gantt",
                "waterfall", "treemap", "sunburst", "pareto", "histogram" }.OrderBy(t => t).ToList();
            Assert.Equal(expected, schemaTypes);
        }

        [Fact]
        public void Skill_requires_screenshot_contract_after_render()
        {
            var text = LoadSkill();
            Assert.Contains("screenshot_dashboard", text);
            // 2026-09-13 契约升级：screenshot_dashboard 返回 image_data_uri 缩略图，AI 须直接读图自审（旧"无法直接看图"口径已废止）
            Assert.Contains("image_data_uri", text);
            Assert.Contains("读图", text);
        }
    }
}