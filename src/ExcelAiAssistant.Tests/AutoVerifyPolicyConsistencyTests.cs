using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 自检覆盖审计的一致性锁：MappedPrimaryTools 列表必须与 TryPlan/TryPlanAll 的实际映射行为同步——
    /// 给列表里每个工具喂一份最小可行实参，必须产出至少一条自检计划。
    /// 谁往 switch 加了映射却忘了同步列表（或反过来），这个测试直接打红。
    /// </summary>
    public sealed class AutoVerifyPolicyConsistencyTests
    {
        private static Dictionary<string, object> TemplateArgsFor(string tool)
        {
            switch (tool)
            {
                case "write_excel_cell":
                    return Args("address", "A1", "value", 1);
                case "write_excel_range":
                    return Args("address", "A1:B2", "values", new[] { new[] { 1, 2 } });
                case "write_excel_cells_batch":
                    return Args("cells", new object[]
                    {
                        new Dictionary<string, object> { ["address"] = "A1", ["value"] = 1 }
                    });
                case "append_excel_rows":
                    return Args("start_cell", "A1", "values", new[] { new[] { 1 } });
                case "clean_excel_data":
                    return Args("address", "A1:A9", "operations", new object[]
                    {
                        new Dictionary<string, object> { ["tool"] = "remove_blank_rows" }
                    });
                case "dedup_range":
                    return Args("address", "A1:A9");
                case "sort_range":
                    return Args("address", "A1:A9", "sort_column", 1, "direction", "asc");
                case "format_excel_range":
                    return Args("address", "A1", "bold", true);
                case "format_excel_conditional":
                    return Args("address", "A1", "rule", "greater", "value", 1);
                case "create_excel_pivot_summary":
                    return Args("source_address", "A1:B9", "label_column", 1, "value_column", 2, "target_address", "D1");
                case "flag_cells":
                    return Args("address", "A1:A9");
                case "create_excel_chart":
                    return Args("source_address", "A1:B9", "chart_type", "ColumnClustered");
                case "create_combo_chart":
                    return Args("source_address", "A1:C9", "secondary_series", "2");
                case "create_doughnut":
                    return Args("source_address", "A1:B9", "hole_size", "50");
                case "create_dashboard_kpi_brief":
                    return Args("source_address", "A1:F21");
                case "render_dashboard":
                    return Args("source_address", "A1:F21");
                case "build_pivot_sheet":
                    return Args("source_address", "A1:F21", "pivot_rows", "月份,品类", "pivot_measures", "金额,利润");
                case "render_super_dashboard":
                    return Args("source_address", "A1:F21", "dashboard_sheet", "超级看板");
                case "create_truncated_column":
                    return Args("source_address", "A1:B9", "truncate_factor", "3.0");
                case "set_linkage":
                    return Args("source_address", "A1:F21", "mode", "native", "dashboard_sheet", "销售看板");
                case "render_dim_cross":
                    return Args("source_address", "A1:F21", "row_dim_column", 1, "col_dim_column", 2);
                case "repair_chart":
                    return Args("sheet", "Sheet1", "chart_name", "Chart 1", "source_address", "A1:B7");
                case "create_super_table":
                    return Args("source_address", "A1:F21", "sheet", "Sheet1", "super_sheet", "超级数据表");
                case "check_chart_version":
                    return Args("chart_type", "treemap");
                case "delete_dashboard":
                    return Args("dashboard_sheet", "销售看板");
                case "create_chart_advanced":
                    return Args("source_address", "A1:C9", "type", "bullet");
                case "render_pnl_brief":
                    return Args("source_address", "A1:C9");
                case "render_target_track":
                    return Args("source_address", "A1:C9");
                case "render_rank_dynamic":
                    return Args("source_address", "A1:C9");
                case "edit_dashboard_theme":
                    return Args("dashboard_sheet", "销售看板", "theme", "blue");
                case "delete_dashboard_chart":
                    return Args("dashboard_sheet", "销售看板", "chart_name", "Chart 1");
                case "insert_recommended_chart":
                    return Args("source_address", "A1:B9");
                case "add_excel_worksheet":
                    return Args("name", "一致性测试表");
                case "rename_excel_worksheet":
                    return Args("sheet", "旧名", "name", "新名");
                case "delete_excel_worksheet":
                    return Args("sheet", "旧名");
                case "clear_excel_range":
                    return Args("address", "A1:B2");
                case "delete_text_matches":
                    return Args("address", "A1:B2", "pattern", "临时");
                case "data_validation":
                    return Args("address", "A1:A5", "type", "list", "source", "A1:A3");
                case "merge_or_split_cells":
                    return Args("action", "merge", "address", "A1:B2");
                case "merge_duplicate_cells":
                    return Args("address", "A1:A5");
                case "create_excel_sparkline":
                    return Args("data_range", "A1:E1", "dest_range", "G1", "type", "column");
                case "create_excel_table":
                    return Args("sheet", "HF4", "address", "A1:C4", "name", "tbl_一致性");
                case "insert_image":
                    return Args("sheet", "HF4", "image_path", "C:\\tmp\\一致性.png", "anchor", "F2", "width", 40, "height", 30);
                case "freeze_panes":
                    return Args("sheet", "HF4", "cell", "B2");
                case "manage_hyperlinks":
                    return Args("sheet", "HF4", "action", "add", "address", "A1", "link_url", "https://example.com", "display_text", "官网");
                default:
                    return null;
            }
        }

        [Fact]
        public void Every_mapped_tool_actually_produces_a_plan()
        {
            foreach (var tool in AutoVerifyPolicy.MappedPrimaryTools)
            {
                var args = TemplateArgsFor(tool);
                Assert.True(args != null,
                    "MappedPrimaryTools 里的 " + tool + " 没有一致性测试模板实参，请补 TemplateArgsFor 分支。");

                var plans = new List<AutoVerifyPlan>();
                Assert.True(AutoVerifyPolicy.TryPlanAll(tool, args, plans),
                    "MappedPrimaryTools 列出了 " + tool + "，但 TryPlanAll 用模板实参没有产出任何自检计划——映射列表与 switch 已失同步。");
                Assert.True(plans.Count > 0);
                Assert.StartsWith("verify_", plans[0].VerifyTool);
            }
        }

        [Fact]
        public void Mapped_list_has_no_duplicates_and_no_verify_tools()
        {
            var seen = new HashSet<string>();
            foreach (var tool in AutoVerifyPolicy.MappedPrimaryTools)
            {
                Assert.True(seen.Add(tool), "MappedPrimaryTools 存在重复项：" + tool);
                Assert.False(AutoVerifyPolicy.IsVerifyTool(tool));
            }
        }

        [Fact]
        public void Exemptions_all_carry_reasons_and_never_overlap_mapped_tools()
        {
            var mapped = new HashSet<string>(AutoVerifyPolicy.MappedPrimaryTools);
            foreach (var entry in AutoVerifyPolicy.DocumentedWriteToolExemptions)
            {
                Assert.False(string.IsNullOrWhiteSpace(entry.Value),
                    "豁免 " + entry.Key + " 没有写明理由——覆盖审计要求每条豁免可评审。");
                Assert.False(mapped.Contains(entry.Key),
                    "工具 " + entry.Key + " 同时出现在映射列表与豁免清单，自相矛盾。");
            }
        }

        [Fact]
        public void Insert_image_mapping_skips_non_positive_size_sentinels()
        {
            // 审查 F2 回归锁：width/height ≤ 0（含 -1 原生尺寸哨兵）不得透传为 expected_*，否则尺寸必然误报。
            var plans = new List<AutoVerifyPlan>();
            var sentinelArgs = Args("sheet", "HF4", "image_path", "C:\\tmp\\一致性.png", "anchor", "F2", "width", -1, "height", -1);
            Assert.True(AutoVerifyPolicy.TryPlanAll("insert_image", sentinelArgs, plans));
            Assert.Single(plans);
            Assert.False(plans[0].Args.ContainsKey("expected_width"), "width=-1 哨兵不应透传为 expected_width");
            Assert.False(plans[0].Args.ContainsKey("expected_height"), "height=-1 哨兵不应透传为 expected_height");
            Assert.Equal("F2", plans[0].Args["anchor"]);

            // 零尺寸同样剔除（InsertImage 工具侧会拒绝，映射层双保险）。
            plans.Clear();
            var zeroArgs = Args("sheet", "HF4", "image_path", "C:\\tmp\\一致性.png", "anchor", "F2", "width", 0);
            Assert.True(AutoVerifyPolicy.TryPlanAll("insert_image", zeroArgs, plans));
            Assert.False(plans[0].Args.ContainsKey("expected_width"), "width=0 不应透传为 expected_width");

            // 正常显式尺寸仍然透传，保持自动核对尺寸的能力。
            plans.Clear();
            var normalArgs = Args("sheet", "HF4", "image_path", "C:\\tmp\\一致性.png", "anchor", "F2", "width", 40, "height", 30);
            Assert.True(AutoVerifyPolicy.TryPlanAll("insert_image", normalArgs, plans));
            Assert.Equal(40d, Convert.ToDouble(plans[0].Args["expected_width"]));
            Assert.Equal(30d, Convert.ToDouble(plans[0].Args["expected_height"]));
        }

        private static Dictionary<string, object> Args(params object[] keyValues)
        {
            var map = new Dictionary<string, object>();
            for (var i = 0; i < keyValues.Length; i += 2)
            {
                map[(string)keyValues[i]] = keyValues[i + 1];
            }

            return map;
        }
    }
}
