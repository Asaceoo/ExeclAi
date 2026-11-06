using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class SecondPhaseToolsTests
    {
        [Fact]
        public void Registry_registers_all_second_phase_tools_with_expected_risk()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("export_selection_to_markdown"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("import_markdown_to_excel"));

            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("merge_or_split_cells"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("reset_cell_value_type"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("transform_excel_formulas"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("convert_formula_references"));

            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("trace_formula_references"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("drill_excel_range"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("copy_visible_rows_to_sheet"));

            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("apply_page_setup"));
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("export_sheets_to_pdf"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("merge_workbooks"));
            Assert.Equal(RiskLevel.StructuralChange, registry.GetRiskLevel("split_workbook"));
            Assert.Equal(RiskLevel.ReadOnly, registry.GetRiskLevel("list_external_links"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("replace_external_links"));

            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("validate_vba_snippet"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("run_vba_snippet"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("run_python_code"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("run_cli_command"));
            Assert.Equal(RiskLevel.Dangerous, registry.GetRiskLevel("run_exe_command"));
        }

        [Fact]
        public void Every_second_phase_tool_carries_a_parameter_schema()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();
            string[] names =
            {
                "export_selection_to_markdown", "import_markdown_to_excel", "merge_or_split_cells",
                "reset_cell_value_type", "transform_excel_formulas", "convert_formula_references",
                "trace_formula_references", "drill_excel_range", "copy_visible_rows_to_sheet",
                "apply_page_setup", "export_sheets_to_pdf", "merge_workbooks", "split_workbook",
                "list_external_links", "replace_external_links", "validate_vba_snippet", "run_vba_snippet",
                "run_python_code", "run_cli_command", "run_exe_command"
            };
            foreach (var name in names)
            {
                Assert.True(registry.TryGet(name, out var tool), "缺少工具：" + name);
                Assert.NotNull(tool.Parameters);
                Assert.Equal("object", tool.Parameters["type"]);
                Assert.Contains("properties", tool.Parameters.Keys);
            }
        }

        [Fact]
        public void Schemas_declare_required_arguments()
        {
            Assert.Equal(new[] { "markdown", "address" }, RequiredKeys(SecondPhaseSchemas.ImportMarkdown()));
            Assert.Equal(new[] { "action", "address" }, RequiredKeys(SecondPhaseSchemas.MergeSplit()));
            Assert.Equal(new[] { "address", "target_type" }, RequiredKeys(SecondPhaseSchemas.ResetValueType()));
            Assert.Equal(new[] { "address" }, RequiredKeys(SecondPhaseSchemas.ConvertRefs()));
            Assert.Equal(new[] { "source_address" }, RequiredKeys(SecondPhaseSchemas.CopyVisibleRows()));
            Assert.Equal(new[] { "file_path" }, RequiredKeys(SecondPhaseSchemas.ExportPdf()));
            Assert.Equal(new[] { "old_link", "new_link" }, RequiredKeys(SecondPhaseSchemas.ReplaceExternalLinks()));
            Assert.Equal(new[] { "command" }, RequiredKeys(SecondPhaseSchemas.RunCli()));
            Assert.Equal(new[] { "path" }, RequiredKeys(SecondPhaseSchemas.RunExe()));
        }

        [Fact]
        public void ParseMarkdownToMatrix_builds_rectangular_matrix()
        {
            var service = new ExcelMarkdownService(new object());
            var markdown = "| 城市 | 销量 |\n| --- | --- |\n| 北京 | 100 |\n| 上海 | 200 |";
            var result = service.ParseMarkdownToMatrix(new Dictionary<string, object> { ["markdown"] = markdown });

            Assert.True(result.Ok);
            Assert.Equal(3, (int)Prop(result.Data, "rows"));
            Assert.Equal(2, (int)Prop(result.Data, "columns"));
            var values = (object[,])Prop(result.Data, "values");
            Assert.Equal("城市", values[0, 0]);
            Assert.Equal("100", values[1, 1]);
            Assert.Equal("200", values[2, 1]);
        }

        [Fact]
        public void Script_tools_refuse_when_disabled()
        {
            var service = new ScriptRunnerService(new object(), new ScriptOptions());
            var result = service.ValidateVbaSnippet(new Dictionary<string, object> { ["snippet"] = "Sub A()\nEnd Sub" });

            Assert.False(result.Ok);
            Assert.Equal("DISABLED", result.ErrorCode);
        }

        [Fact]
        public void Script_tools_static_validate_when_vbe_unavailable()
        {
            var service = new ScriptRunnerService(new object(), new ScriptOptions { Enabled = true });
            var result = service.ValidateVbaSnippet(new Dictionary<string, object> { ["snippet"] = "Public Sub Foo()\nEnd Sub" });

            // 无 VBE 时降级为静态校验，Ok=true。
            Assert.True(result.Ok);
            Assert.Equal("static", (string)Prop(result.Data, "validation_mode"));
            Assert.True((bool)Prop(result.Data, "valid"));
        }

        [Fact]
        public void PreRunVbaGate_blocks_rect_block_autofit()
        {
            // 用户实测根因：对跨行跨列矩形 Range 直接调 .AutoFit → 运行时报 1004 中断（0x800A9C68）。
            // 应在执行前拦截为 VBA_INVALID，而非注入 Excel 白跑一次再撞运行时错误。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    ws.Range(\"G12:K18\").AutoFit\nEnd Sub");
            Assert.NotNull(msg);
            Assert.Contains("AutoFit", msg);
            Assert.Contains("1004", msg);
        }

        [Fact]
        public void PreRunVbaGate_blocks_bare_range_rect_autofit()
        {
            // 无接收者前缀的裸 Range(...) 同样拦截（旧正则漏网）。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    Range(\"A1:C3\").AutoFit\nEnd Sub");
            Assert.NotNull(msg);
            Assert.Contains("1004", msg);
        }

        [Fact]
        public void PreRunVbaGate_blocks_range_alignment()
        {
            // 用户真机实证：Range 没有 Alignment 属性（运行时 438），须用 HorizontalAlignment/VerticalAlignment。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    ws.Range(\"B2:C9\").Alignment = xlHAlignCenter\nEnd Sub");
            Assert.NotNull(msg);
            Assert.Contains("Alignment", msg);
            Assert.Contains("HorizontalAlignment", msg);
        }

        [Fact]
        public void PreRunVbaGate_allows_safe_rows_columns_autofit()
        {
            // 安全改写：按行/列调用 .Rows(...).AutoFit / .Columns(...).AutoFit，不得被误判为 error 阻断。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    ws.Rows(\"12:18\").AutoFit\n    ws.Columns(\"G:K\").AutoFit\nEnd Sub");
            Assert.Null(msg);
        }

        [Fact]
        public void PreRunVbaGate_allows_safe_range_chain_rows_autofit()
        {
            // 矩形 Range 上显式中介调用 Rows/Columns → 安全，不阻断。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    ws.Range(\"G12:K18\").Rows(\"12:18\").AutoFit\nEnd Sub");
            Assert.Null(msg);
        }

        [Fact]
        public void PreRunVbaGate_allows_single_cell_and_single_strip_autofit()
        {
            // 单格、整行、整列（单列/单行）调用 AutoFit 不视为必错（仅 suggestion），gate 不阻断。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    ws.Range(\"A1\").AutoFit\n    ws.Range(\"A1:A5\").AutoFit\n    ws.Range(\"2:2\").AutoFit\nEnd Sub");
            Assert.Null(msg);
        }

        [Fact]
        public void ValidateVbaSnippet_static_flags_rect_autofit_as_error()
        {
            var service = new ScriptRunnerService(new object(), new ScriptOptions { Enabled = true });
            var result = service.ValidateVbaSnippet(new Dictionary<string, object>
            {
                ["snippet"] = "Sub T()\n    ws.Range(\"G12:K18\").AutoFit\nEnd Sub"
            });

            Assert.True(result.Ok);
            Assert.False((bool)Prop(result.Data, "valid"));
            dynamic issues = Prop(result.Data, "issues");
            // 至少存在一条 AutoFit error。
            var texts = new List<string>();
            foreach (var it in (System.Collections.IEnumerable)issues)
            {
                texts.Add((string)it.GetType().GetProperty("level").GetValue(it, null) + ":" +
                          (string)it.GetType().GetProperty("message").GetValue(it, null));
            }

            Assert.Contains(texts, t => t.StartsWith("error:") && t.Contains("AutoFit") && t.Contains("1004"));
        }

        [Fact]
        public void PreRunVbaGate_blocks_module_compile_broken_snippets()
        {
            // 结构级必失败（过程完全缺闭合）：闸门在执行前拦截为 VBA_INVALID，不注入 Excel 白跑。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    Dim x As Integer\n    x = 1");
            Assert.NotNull(msg);
        }

        [Fact]
        public void PreRunVbaGate_lets_compile_only_broken_snips_to_runtime()
        {
            // 结构上 End 配平、但变量/类型/引用层面编译失败（如 Sub 用 End Function 闭合）的片段，
            // 闸门不阻断（它只判 string 配对与结构配对）→ 放行到真正执行；运行时若整模块（含探针）
            // 无法编译，则归为第二道防线 VBA_SYNTAX_ERROR，而非把晦涩的 HRESULT 0x800A9C68 当运行时错误回喂。
            var msg = ScriptRunnerService.PreRunVbaGate(
                "Sub T()\n    Dim x As Integer\n    x = UndefinedVar\nEnd Function");
            Assert.Null(msg);
        }

        private static object Prop(object data, string name)
        {
            return data.GetType().GetProperty(name).GetValue(data, null);
        }

        private static string[] RequiredKeys(Dictionary<string, object> schema)
        {
            return schema.ContainsKey("required") ? (string[])schema["required"] : new string[0];
        }
    }
}