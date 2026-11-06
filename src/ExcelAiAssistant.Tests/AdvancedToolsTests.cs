using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;
using static ExcelAiAssistant.Excel.ExcelReadService;

namespace ExcelAiAssistant.Tests
{
    public sealed class AdvancedToolsTests
    {
        private static readonly string[] RequiredAdvancedTools =
        {
            "get_excel_sheet_list", "preview_excel_sheet", "get_excel_range_values_and_formulas",
            "get_current_selection_values_and_formulas", "select_excel_range", "find_excel_cells",
            "manage_excel_filter", "manage_excel_named_range", "write_excel_cells_batch",
            "append_calculation_scratchpad", "clear_calculation_scratchpad",
            "insert_excel_dimension", "delete_excel_dimension", "copy_paste_excel_range", "create_toc_sheet"
        };

        [Fact]
        public void Registry_registers_all_advanced_tools_with_expected_risk()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("get_excel_sheet_list"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("preview_excel_sheet"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("get_excel_range_values_and_formulas"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("get_current_selection_values_and_formulas"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("select_excel_range"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("find_excel_cells"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("manage_excel_filter"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("manage_excel_named_range"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("write_excel_cells_batch"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("append_calculation_scratchpad"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("clear_calculation_scratchpad"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("insert_excel_dimension"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("delete_excel_dimension"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("copy_paste_excel_range"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("create_toc_sheet"));
        }

        [Fact]
        public void Every_advanced_tool_carries_a_parameter_schema()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();
            foreach (var name in RequiredAdvancedTools)
            {
                Assert.True(registry.TryGet(name, out var tool), $"缺少工具：{name}");
                Assert.NotNull(tool.Parameters);
                Assert.Equal("object", tool.Parameters["type"]);
                Assert.Contains("properties", tool.Parameters.Keys);
            }
        }

        [Fact]
        public void Schemas_declare_required_arguments()
        {
            Assert.Equal(new[] { "address" }, RequiredKeys(AdvancedToolSchemas.ReadRange()));
            Assert.Equal(new[] { "what" }, RequiredKeys(AdvancedToolSchemas.FindCells()));
            Assert.Equal(new[] { "action" }, RequiredKeys(AdvancedToolSchemas.Filter()));
            Assert.Equal(new[] { "action", "name" }, RequiredKeys(AdvancedToolSchemas.NamedRange()));
            Assert.Equal(new[] { "cells" }, RequiredKeys(AdvancedToolSchemas.BatchWrite()));
            Assert.Equal(new[] { "address" }, RequiredKeys(AdvancedToolSchemas.Dimension()));
            Assert.Equal(new[] { "source_address", "target_address" }, RequiredKeys(AdvancedToolSchemas.CopyPaste()));
        }

        [Fact]
        public void ReadSelection_schema_exposes_only_max_cells_no_dead_sheet()
        {
            // 实现 read.ReadSelection 只使用 application.Selection，从不读取 sheet；schema 暴露 sheet=死参数。
            var props = Properties(AdvancedToolSchemas.ReadSelection());
            Assert.False(props.ContainsKey("sheet"));
            Assert.True(props.ContainsKey("max_cells"));
        }

        [Fact]
        public void FindCells_mode_enum_restricts_to_value_and_formula()
        {
            var props = Properties(AdvancedToolSchemas.FindCells());
            var mode = Assert.IsType<Dictionary<string, object>>(props["mode"]);
            var modeEnum = Assert.IsType<string[]>(mode["enum"]);
            Assert.Equal(new[] { "value", "formula" }, modeEnum);
        }

        [Fact]
        public void ToAbsolute_builds_absolute_references()
        {
            Assert.Equal("$A$1:$B$3", ExcelRangeHelper.ToAbsolute("A1:B3"));
            Assert.Equal("$C$2", ExcelRangeHelper.ToAbsolute("C2"));
            Assert.Equal("$A$1", ExcelRangeHelper.ToAbsolute("'Data'!A1"));
        }

        [Fact]
        public void CellAddress_and_Resize_produce_expected_addresses()
        {
            Assert.Equal("B3", ExcelRangeHelper.CellAddress(3, 2));
            Assert.Equal("A1:B2", ExcelRangeHelper.Resize("A1", 2, 2));
            Assert.Equal("$A$1:$H$1", ExcelRangeHelper.ToAbsolute("A1:H1"));
        }

        [Fact]
        public void InferColumnTypes_classifies_number_text_and_mixed()
        {
            var sample = new object[,]
            {
                { "收入", "城市", "日期", "备注", null },
                { 100, "北京", "2024-01-01", "电话", null },
                { 200, "上海", "2024-01-02", "电话", null },
                { 300, "深圳", "2024-01-03", "邮件", null }
            };

            var types = InferColumnTypes(sample, 5);
            Assert.Equal("number", types[0]);
            Assert.Equal("text", types[1]);
            Assert.Equal("blank", types[4]); // 全空列
        }

        [Fact]
        public void FilterMatcher_evaluates_all_operators()
        {
            Assert.True(DataTransformService.Matches("Apple", "equals", "apple")); // 忽略大小写
            Assert.True(DataTransformService.Matches("Apple", "contains", "ppl"));
            Assert.True(DataTransformService.Matches("10", "gt", "5"));
            Assert.False(DataTransformService.Matches("10", "lt", "5"));
            Assert.True(DataTransformService.Matches(null, "empty", null));
            Assert.True(DataTransformService.Matches("x", "not_empty", null));
        }

        private static string[] RequiredKeys(Dictionary<string, object> schema)
        {
            return schema.ContainsKey("required") ? (string[])schema["required"] : new string[0];
        }

        private static Dictionary<string, object> Properties(Dictionary<string, object> schema)
        {
            return (Dictionary<string, object>)schema["properties"];
        }
    }
}