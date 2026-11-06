using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>执行层自动自检策略：主工具实参 → verify 预期参数的映射推导。</summary>
    public sealed class AutoVerifyPolicyTests
    {
        [Fact]
        public void Value_cell_write_plans_write_result_verify()
        {
            var ok = AutoVerifyPolicy.TryPlan("write_excel_cell", Args(
                "sheet", "数据",
                "address", "B3",
                "value", "hello"), out var plan);

            Assert.True(ok);
            Assert.Equal("verify_write_result", plan.VerifyTool);
            Assert.Equal("B3", plan.Args["address"]);
            Assert.Equal("数据", plan.Args["sheet"]);
            var expected = Assert.IsType<List<object>>(plan.Args["expected_values"]);
            var row = Assert.IsType<List<object>>(expected.Single());
            Assert.Equal("hello", row.Single());
        }

        [Fact]
        public void Formula_cell_write_plans_formula_verify()
        {
            Assert.True(AutoVerifyPolicy.TryPlan("write_excel_cell", Args(
                "address", "C1",
                "value", "=B2*C2"), out var plan));
            Assert.Equal("verify_formula", plan.VerifyTool);
            Assert.Equal("=B2*C2", plan.Args["expected_formula"]);

            Assert.True(AutoVerifyPolicy.TryPlan("write_excel_cell", Args(
                "address", "C1",
                "value", "=B2*C2",
                "value_mode", "formula"), out plan));
            Assert.Equal("verify_formula", plan.VerifyTool);
        }

        [Fact]
        public void Plain_range_write_plans_write_result_verify()
        {
            var values = new JArray { new JArray { 1, 2 }, new JArray { 3, 4 } };
            var ok = AutoVerifyPolicy.TryPlan("write_excel_range", Args(
                "address", "$A$1:$B$2",
                "values", values), out var plan);

            Assert.True(ok);
            Assert.Equal("verify_write_result", plan.VerifyTool);
            Assert.Equal(values, plan.Args["expected_values"]);
        }

        [Fact]
        public void Formula_range_write_is_skipped()
        {
            var values = new JArray { new JArray { "=A1*2" } };
            Assert.False(AutoVerifyPolicy.TryPlan("write_excel_range", Args(
                "address", "A1", "values", values, "value_mode", "formula"), out _));
            Assert.False(AutoVerifyPolicy.TryPlan("write_excel_range", Args(
                "address", "A1", "values", values, "value_mode", "auto"), out _));
        }

        [Fact]
        public void Clean_operations_map_to_verify_checks()
        {
            var operations = new JArray
            {
                new JObject { ["tool"] = "remove_blank_rows" },
                new JObject { ["tool"] = "fill_blanks", ["column"] = 2 },
                new JObject { ["tool"] = "fix_dates" }
            };
            var ok = AutoVerifyPolicy.TryPlan("clean_excel_data", Args(
                "address", "A1:D20",
                "operations", operations), out var plan);

            Assert.True(ok);
            Assert.Equal("verify_data_clean", plan.VerifyTool);
            var checks = Assert.IsType<List<object>>(plan.Args["checks"]);
            Assert.Equal(3, checks.Count);
            Assert.Equal("no_blank_rows", ((Dictionary<string, object>)checks[0])["type"]);
            var fillCheck = (Dictionary<string, object>)checks[1];
            Assert.Equal("no_blanks", fillCheck["type"]);
            var fillColumns = (List<object>)fillCheck["columns"];
            Assert.Equal(2, Convert.ToInt64(fillColumns.Single()));
            Assert.Equal("dates_valid", ((Dictionary<string, object>)checks[2])["type"]);
        }

        [Fact]
        public void Dedup_plans_no_duplicates_check()
        {
            Assert.True(AutoVerifyPolicy.TryPlan("dedup_range", Args(
                "address", "A1:C9"), out var plan));
            Assert.Equal("verify_data_clean", plan.VerifyTool);
            var checks = Assert.IsType<List<object>>(plan.Args["checks"]);
            Assert.Equal("no_duplicates", ((Dictionary<string, object>)checks.Single())["type"]);
        }

        [Fact]
        public void Sort_plans_sort_filter_verify_and_filter_is_left_to_model()
        {
            AutoVerifyPolicy.TryPlan("sort_range", Args(
                "address", "A1:C10", "sort_column", 3, "direction", "desc"), out var plan);
            Assert.True(plan != null, "排序应产生自检计划");
            Assert.Equal("verify_sort_filter", plan.VerifyTool);
            Assert.Equal("A1:C10", plan.Args["address"]);
            Assert.Equal(3, plan.Args["expected_sort_column"]);
            Assert.Equal("desc", plan.Args["expected_sort_direction"]);

            // 筛选的源区域保留不匹配行，区域级谓词断言必然误报，交给模型按输出区域显式核对。
            Assert.False(AutoVerifyPolicy.TryPlan("filter_excel_data", Args(
                "address", "A1:C10", "column", 1, "value", "北京", "operator", "contains"), out _));
        }

        [Fact]
        public void Format_only_plans_when_expectation_present()
        {
            AutoVerifyPolicy.TryPlan("format_excel_range", Args(
                "address", "A1:A5", "bold", true, "number_format", "0.00"), out var plan);
            Assert.True(plan != null, "有可断言预期时应产生自检计划");
            Assert.Equal("verify_cell_format", plan.VerifyTool);
            Assert.True((bool)plan.Args["expected_bold"]);
            Assert.Equal("0.00", plan.Args["expected_number_format"]);

            // 只有地址没有任何可断言预期时不产生空自检。
            Assert.False(AutoVerifyPolicy.TryPlan("format_excel_range", Args(
                "address", "A1:A5", "auto_fit_columns", true), out _));
        }

        [Fact]
        public void Conditional_format_maps_rule_and_colors()
        {
            AutoVerifyPolicy.TryPlan("format_excel_conditional", Args(
                "address", "B1:B20", "rule", "greater", "value", 100,
                "fill_color", "#ffd9d9", "text_color", "#9c0006"), out var plan);
            Assert.True(plan != null, "greater 规则应产生自检计划");
            Assert.Equal("verify_conditional_format", plan.VerifyTool);
            Assert.Equal("greater", plan.Args["expected_rule"]);
            Assert.Equal("100", plan.Args["expected_formula1"]);
            Assert.Equal("#ffd9d9", plan.Args["expected_fill_color"]);
            Assert.Equal("#9c0006", plan.Args["expected_font_color"]);

            Assert.False(AutoVerifyPolicy.TryPlan("format_excel_conditional", Args(
                "address", "B1:B20", "rule", "clear"), out _));

            // data_bar 颜色在 BarColor 上，verify 比对 Interior 会误报，只断言规则类型。
            AutoVerifyPolicy.TryPlan("format_excel_conditional", Args(
                "address", "C1:C20", "rule", "data_bar", "bar_color", "#63be7b"), out plan);
            Assert.Equal("data_bar", plan.Args["expected_rule"]);
            Assert.False(plan.Args.ContainsKey("expected_fill_color"));
            Assert.False(plan.Args.ContainsKey("expected_font_color"));
            Assert.False(plan.Args.ContainsKey("expected_formula1"));
        }

        [Fact]
        public void Pivot_plans_and_enriches_from_result()
        {
            // 未传 target_address（自动落点）也出计划，执行后由结果富化补齐地址与行数断言。
            Assert.True(AutoVerifyPolicy.TryPlan("create_excel_pivot_summary", Args(
                "source_address", "A1:C9"), out var plan));
            Assert.Equal("verify_pivot", plan.VerifyTool);

            var resultData = new { source_rows = 8, row_count = 3, target_address = "$E$5" };
            AutoVerifyPolicy.EnrichFromResult(plan, "create_excel_pivot_summary", resultData);
            Assert.Equal("$E$5", plan.Args["target_address"]);
            Assert.Equal(3, plan.Args["expected_count"]);

            // 其他工具不做结果富化。
            AutoVerifyPolicy.TryPlan("sort_range", Args("address", "A1:A3"), out var sortPlan);
            AutoVerifyPolicy.EnrichFromResult(sortPlan, "sort_range", resultData);
            Assert.False(sortPlan.Args.ContainsKey("expected_count"));
        }

        [Fact]
        public void Pivot_result_enrichment_tolerates_bad_shape()
        {
            Assert.True(AutoVerifyPolicy.TryPlan("create_excel_pivot_summary", Args(
                "source_address", "A1:C9", "target_address", "E1"), out var plan));
            AutoVerifyPolicy.EnrichFromResult(plan, "create_excel_pivot_summary", "not-json-object");
            Assert.Equal("E1", plan.Args["target_address"]);
            Assert.False(plan.Args.ContainsKey("expected_count"));
        }

        [Fact]
        public void Flag_and_chart_plan_respective_verifies()
        {
            var colors = new JObject { ["outlier"] = "#ffd9d9" };
            AutoVerifyPolicy.TryPlan("flag_cells", Args(
                "address", "A1:D200", "fill_colors", colors), out var plan);
            Assert.True(plan != null, "标记应产生自检计划");
            Assert.Equal("verify_marker", plan.VerifyTool);
            Assert.Equal(colors, plan.Args["expected_fill_colors"]);

            AutoVerifyPolicy.TryPlan("create_excel_chart", Args(
                "source_address", "A1:B8", "chart_type", "ColumnClustered", "title", "销量"), out plan);
            Assert.Equal("verify_chart", plan.VerifyTool);
            Assert.Equal("ColumnClustered", plan.Args["expected_chart_type"]);
            Assert.Equal("销量", plan.Args["expected_title"]);
            Assert.Equal("A1:B8", plan.Args["expected_data_source"]);

            AutoVerifyPolicy.TryPlan("insert_recommended_chart", Args(
                "source_address", "A1:B8"), out plan);
            Assert.Equal("verify_chart", plan.VerifyTool);
        }

        [Fact]
        public void Verify_tools_and_unknown_tools_never_plan()
        {
            Assert.False(AutoVerifyPolicy.TryPlan("verify_write_result", Args("address", "A1"), out _));
            Assert.False(AutoVerifyPolicy.TryPlan("preview_range", Args("address", "A1"), out _));
            Assert.False(AutoVerifyPolicy.TryPlan(null, Args(), out _));
        }

        [Fact]
        public void Plain_clr_arrays_plan_without_jtoken()
        {
            // JSON 反序列化外，模型实参也可能是纯 CLR 数组（脚本化调用），映射不得依赖 JToken 形态。
            var values = new[] { new[] { 1, 2 } };
            Assert.True(AutoVerifyPolicy.TryPlan("write_excel_range", Args(
                "address", "A1:B2", "values", values), out var plan));
            Assert.Equal(values, plan.Args["expected_values"]);
        }

        [Fact]
        public void Batch_write_plans_per_cell_with_formula_split()
        {
            var cells = new JArray
            {
                new JObject { ["address"] = "A1", ["value"] = "名称" },
                new JObject { ["address"] = "B1", ["value"] = "=SUM(C1:C9)", ["value_mode"] = "formula" }
            };
            var plans = new List<AutoVerifyPlan>();
            Assert.True(AutoVerifyPolicy.TryPlanAll("write_excel_cells_batch", Args("cells", cells), plans));
            Assert.Equal(2, plans.Count);
            Assert.Equal("verify_write_result", plans[0].VerifyTool);
            Assert.Equal("A1", plans[0].Args["address"]);
            Assert.True((bool)plans[0].Args["type_aware"]);
            Assert.Equal("verify_formula", plans[1].VerifyTool);
            Assert.Equal("=SUM(C1:C9)", plans[1].Args["expected_formula"]);
        }

        [Fact]
        public void Batch_write_over_cap_is_skipped()
        {
            var cells = new JArray();
            for (var i = 0; i < AutoVerifyPolicy.MaxAutoVerifyCellsPerCall + 1; i++)
            {
                cells.Add(new JObject { ["address"] = "A" + (i + 1), ["value"] = i });
            }

            var plans = new List<AutoVerifyPlan>();
            Assert.False(AutoVerifyPolicy.TryPlanAll("write_excel_cells_batch", Args("cells", cells), plans));
            Assert.Empty(plans);
        }

        [Fact]
        public void Append_plans_and_enriches_address_from_result()
        {
            var values = new JArray { new JArray { "苹果", 8 }, new JArray { "梨", 2 } };
            var plans = new List<AutoVerifyPlan>();
            Assert.True(AutoVerifyPolicy.TryPlanAll("append_excel_rows", Args(
                "start_cell", "A1", "values", values), plans));
            var plan = Assert.Single(plans);
            Assert.Equal("verify_write_result", plan.VerifyTool);
            Assert.False(plan.Args.ContainsKey("address"), "追加行目标地址必须由结果富化得出");

            var result = new ExcelAiAssistant.Core.Models.ToolResult { Ok = true };
            result.AffectedRanges.Add("数据!$A$5:$B$6");
            Assert.True(AutoVerifyPolicy.EnrichFromResult(plan, "append_excel_rows", result.Data, result.AffectedRanges));
            Assert.Equal("$A$5:$B$6", plan.Args["address"]);
            Assert.Equal("数据", plan.Args["sheet"]);
        }

        [Fact]
        public void Append_without_affected_range_is_dropped()
        {
            var values = new JArray { new JArray { "x", 1 } };
            var plans = new List<AutoVerifyPlan>();
            AutoVerifyPolicy.TryPlanAll("append_excel_rows", Args("start_cell", "A1", "values", values), plans);
            var plan = Assert.Single(plans);
            Assert.False(AutoVerifyPolicy.EnrichFromResult(plan, "append_excel_rows", null, null));
        }

        [Fact]
        public void Structural_writes_map_to_structure_verify()
        {
            var plans = new List<AutoVerifyPlan>();
            AutoVerifyPolicy.TryPlanAll("add_excel_worksheet", Args("name", "汇总"), plans);
            var plan = Assert.Single(plans);
            Assert.Equal("verify_structure", plan.VerifyTool);
            Assert.Equal("汇总", plan.Args["expected_sheet_exists"]);

            plans.Clear();
            AutoVerifyPolicy.TryPlanAll("rename_excel_worksheet", Args("sheet", "旧名", "name", "新名"), plans);
            plan = Assert.Single(plans);
            Assert.Equal("新名", plan.Args["expected_sheet_exists"]);
            Assert.Equal("旧名", plan.Args["expected_sheet_absent"]);

            plans.Clear();
            AutoVerifyPolicy.TryPlanAll("delete_excel_worksheet", Args("sheet", "草稿"), plans);
            plan = Assert.Single(plans);
            Assert.Equal("草稿", plan.Args["expected_sheet_absent"]);
        }

        [Fact]
        public void Pivot_enrichment_carries_conservation_source()
        {
            AutoVerifyPolicy.TryPlan("create_excel_pivot_summary", Args(
                "source_address", "A1:B9", "value_column", 2, "target_address", "E1"), out var plan);
            var resultData = new { row_count = 3, target_address = "$E$1" };
            AutoVerifyPolicy.EnrichFromResult(plan, "create_excel_pivot_summary", resultData);
            Assert.Equal("A1:B9", plan.Args["source_address"]);
            Assert.Equal(2, plan.Args["source_value_column"]);
            Assert.Equal(3, plan.Args["expected_count"]);
        }

        [Fact]
        public void Write_plans_enable_type_aware_readback()
        {
            AutoVerifyPolicy.TryPlan("write_excel_cell", Args("address", "B2", "value", 3), out var plan);
            Assert.True((bool)plan.Args["type_aware"]);
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
