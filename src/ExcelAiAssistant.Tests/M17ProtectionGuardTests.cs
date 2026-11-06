using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M17 回归锁：写路径工具「保护态」集中守卫。
    /// 1. Evaluate 判定矩阵（保护表/活动表缺省/结构保护/参数覆盖/探测失败放行）；
    /// 2. 豁免清单纪律（理由非空、覆盖审计联动、ct_ 前缀规则）；
    /// 3. Wrap 元数据保真与守卫先行（拦截时不触达内层执行体）。
    /// 判定为纯逻辑（探测经 ISheetProtectionProbe 接缝注入），COM 探测实现由双宿主真机覆盖。
    /// </summary>
    public sealed class M17ProtectionGuardTests
    {
        /// <summary>假探测实现：直接给定表→保护态映射，替代 COM。</summary>
        private sealed class FakeProbe : ISheetProtectionProbe
        {
            private readonly Dictionary<string, bool> sheets = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            public string ActiveSheetName { get; set; }
            public bool StructureProtected { get; set; }
            public bool ContentsProbeFails { get; set; }
            public bool StructureProbeFails { get; set; }
            public int ContentsProbeCalls { get; private set; }

            public FakeProbe WithSheet(string name, bool protectedContents)
            {
                sheets[name] = protectedContents;
                return this;
            }

            public bool TryGetContentsProtection(string sheetName, out string resolvedName, out bool contentsProtected)
            {
                ContentsProbeCalls++;
                resolvedName = null;
                contentsProtected = false;
                if (ContentsProbeFails)
                {
                    return false;
                }

                var key = string.IsNullOrWhiteSpace(sheetName) ? ActiveSheetName : sheetName;
                if (key == null || !sheets.TryGetValue(key, out var prot))
                {
                    return false; // 表不存在 → 守卫放行
                }

                resolvedName = key;
                contentsProtected = prot;
                return true;
            }

            public bool TryGetStructureProtection(out bool structureProtected)
            {
                structureProtected = StructureProtected;
                return !StructureProbeFails;
            }
        }

        private static FakeProbe DefaultProbe(bool activeProtected = false, bool structureProtected = false)
        {
            return new FakeProbe
            {
                ActiveSheetName = "数据",
                StructureProtected = structureProtected
            }
            .WithSheet("数据", activeProtected)
            .WithSheet("保护表", true);
        }

        // ===================== Evaluate 判定矩阵 =====================

        [Fact]
        public void Write_on_protected_sheet_is_blocked()
        {
            var result = SheetProtectionGuard.Evaluate(
                "write_excel_range", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "保护表", ["address"] = "A1" });
            Assert.NotNull(result);
            Assert.False(result.Ok);
            Assert.Equal("SHEET_PROTECTED", result.ErrorCode);
            Assert.Contains("保护表", Convert.ToString(result.Data));
            Assert.Contains("verify_protection", Convert.ToString(result.Data));
        }

        [Fact]
        public void Write_on_unprotected_sheet_passes()
        {
            var result = SheetProtectionGuard.Evaluate(
                "write_excel_range", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "数据", ["address"] = "A1" });
            Assert.Null(result);
        }

        [Fact]
        public void Missing_sheet_arg_resolves_active_sheet()
        {
            // 活动表受保护 → 拦截
            var blocked = SheetProtectionGuard.Evaluate(
                "write_excel_cell", DefaultProbe(activeProtected: true),
                new Dictionary<string, object> { ["address"] = "A1" });
            Assert.NotNull(blocked);

            // 活动表未保护 → 放行
            var pass = SheetProtectionGuard.Evaluate(
                "write_excel_cell", DefaultProbe(),
                new Dictionary<string, object> { ["address"] = "A1" });
            Assert.Null(pass);
        }

        [Fact]
        public void Unknown_named_sheet_falls_through_to_tool_error()
        {
            var result = SheetProtectionGuard.Evaluate(
                "write_excel_range", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "NoSuchSheet", ["address"] = "A1" });
            Assert.Null(result);
        }

        [Fact]
        public void Copy_visible_rows_reads_target_sheet_not_source()
        {
            // target_sheet 受保护 → 拦截（即使源表未保护）
            var blocked = SheetProtectionGuard.Evaluate(
                "copy_visible_rows_to_sheet", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "数据", ["target_sheet"] = "保护表" });
            Assert.NotNull(blocked);

            // target_sheet 未保护 → 放行（源表保护与否不影响：读操作原生允许）
            var pass = SheetProtectionGuard.Evaluate(
                "copy_visible_rows_to_sheet", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "保护表", ["target_sheet"] = "数据" });
            Assert.Null(pass);
        }

        [Fact]
        public void Structure_tools_respect_protect_structure()
        {
            foreach (var tool in new[] { "add_excel_worksheet", "rename_excel_worksheet", "delete_excel_worksheet" })
            {
                var blocked = SheetProtectionGuard.Evaluate(
                    tool, DefaultProbe(structureProtected: true),
                    new Dictionary<string, object> { ["sheet"] = "数据", ["name"] = "新表" });
                Assert.NotNull(blocked);
                Assert.Equal("SHEET_PROTECTED", blocked.ErrorCode);
            }

            // 结构未保护 → 放行
            var pass = SheetProtectionGuard.Evaluate(
                "add_excel_worksheet", DefaultProbe(),
                new Dictionary<string, object> { ["name"] = "新表" });
            Assert.Null(pass);
        }

        [Fact]
        public void Delete_target_protected_sheet_blocked_without_structure_protection()
        {
            var blocked = SheetProtectionGuard.Evaluate(
                "delete_excel_worksheet", DefaultProbe(),
                new Dictionary<string, object> { ["sheet"] = "保护表" });
            Assert.NotNull(blocked);
            Assert.Equal("SHEET_PROTECTED", blocked.ErrorCode);
        }

        [Fact]
        public void Probe_failure_falls_through_to_tool_error_path()
        {
            // 守卫探测失败必须放行（加法前置，不制造新失败模式）
            var failing = new FakeProbe { ActiveSheetName = "数据", ContentsProbeFails = true }.WithSheet("数据", true);
            var result = SheetProtectionGuard.Evaluate(
                "write_excel_range", failing,
                new Dictionary<string, object> { ["sheet"] = "数据", ["address"] = "A1" });
            Assert.Null(result);

            var structureFailing = new FakeProbe { StructureProbeFails = true };
            var structureResult = SheetProtectionGuard.Evaluate(
                "add_excel_worksheet", structureFailing,
                new Dictionary<string, object> { ["name"] = "新表" });
            Assert.Null(structureResult);

            // 探测接缝为 null（防御）：放行
            Assert.Null(SheetProtectionGuard.Evaluate("write_excel_cell", null, new Dictionary<string, object>()));
        }

        // ===================== 豁免清单纪律 =====================

        [Fact]
        public void All_exemptions_have_documented_reasons()
        {
            Assert.NotEmpty(SheetProtectionGuard.Exemptions);
            foreach (var kv in SheetProtectionGuard.Exemptions)
            {
                Assert.False(string.IsNullOrWhiteSpace(kv.Value), $"豁免缺少理由：{kv.Key}");
            }
        }

        [Fact]
        public void Override_and_structure_entries_must_not_be_exempted()
        {
            foreach (var tool in SheetProtectionGuard.SheetArgOverrides.Keys)
            {
                Assert.False(SheetProtectionGuard.IsExempted(tool), $"参数覆盖工具不应豁免：{tool}");
            }

            foreach (var tool in SheetProtectionGuard.StructureTools)
            {
                Assert.False(SheetProtectionGuard.IsExempted(tool), $"结构守卫工具不应豁免：{tool}");
            }
        }

        [Fact]
        public void Custom_tool_prefix_is_auto_exempted()
        {
            Assert.True(SheetProtectionGuard.IsExempted("ct_vba_tool"));
            Assert.True(SheetProtectionGuard.IsExempted("ct_format_answer_sheet"));
            Assert.False(SheetProtectionGuard.IsExempted("write_excel_cell"));
            Assert.False(SheetProtectionGuard.IsExempted(null));
        }

        [Fact]
        public void Auto_verify_mapped_tools_must_not_be_guard_exempted()
        {
            // 覆盖审计联动：AutoVerifyPolicy 已映射自动自检的 26 个主写工具，
            // 必须全部受保护态守卫（误入豁免清单即盲区）。
            // 允许"自检映射 + 守卫豁免"并存的只有维度不同的工具：
            // - freeze_panes：视图级操作，保护态下原生可用，另有 verify_freeze_panes 独立自检；
            // - export_sheets_to_pdf / export_excel_chart_image：M28 起映射 verify_file（文件核验），
            //   但导出只产出文件、不写工作表单元格，保护态与其无关。
            var fileLevelMappedButGuardExempt = new HashSet<string>(StringComparer.Ordinal)
            {
                "freeze_panes", "export_sheets_to_pdf", "export_excel_chart_image"
            };
            foreach (var tool in AutoVerifyPolicy.MappedPrimaryTools)
            {
                if (fileLevelMappedButGuardExempt.Contains(tool))
                {
                    continue;
                }

                Assert.False(SheetProtectionGuard.IsExempted(tool), $"自动自检映射工具不应豁免守卫：{tool}");
            }
        }

        // ===================== Wrap 元数据保真与守卫先行 =====================

        [Fact]
        public void Wrap_preserves_metadata_and_guard_runs_before_inner()
        {
            var innerCalled = 0;
            var descriptor = new ToolDescriptor(
                "write_excel_range", "desc", RiskLevel.NormalWrite,
                _ => { innerCalled++; return new ToolResult { Ok = true }; });

            var wrapped = SheetProtectionGuard.WrapWithProbe(descriptor, DefaultProbe());
            Assert.Equal(descriptor.Name, wrapped.Name);
            Assert.Equal(descriptor.Description, wrapped.Description);
            Assert.Equal(descriptor.RiskLevel, wrapped.RiskLevel);
            Assert.Equal(descriptor.Parameters, wrapped.Parameters);

            // 命中保护表：拦截且内层不触达
            var blocked = wrapped.Execute(new Dictionary<string, object> { ["sheet"] = "保护表" });
            Assert.False(blocked.Ok);
            Assert.Equal("SHEET_PROTECTED", blocked.ErrorCode);
            Assert.Equal(0, innerCalled);

            // 未保护：放行到内层
            var pass = wrapped.Execute(new Dictionary<string, object> { ["sheet"] = "数据" });
            Assert.True(pass.Ok);
            Assert.Equal(1, innerCalled);
        }

        [Fact]
        public void WrapWriteTools_covers_all_write_tools_and_skips_readonly_and_exempt()
        {
            var probe = DefaultProbe(activeProtected: true);
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "write_excel_cell", "w", RiskLevel.NormalWrite,
                _ => new ToolResult { Ok = true }));
            registry.Register(new ToolDescriptor(
                "run_cli_command", "e", RiskLevel.Dangerous,
                _ => new ToolResult { Ok = true, ErrorCode = "INNER_REACHED" }));
            registry.Register(new ToolDescriptor(
                "preview_excel_range", "r", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, ErrorCode = "READONLY_REACHED" }));
            registry.Register(new ToolDescriptor(
                "ct_vba_tool", "c", RiskLevel.NormalWrite,
                _ => new ToolResult { Ok = true, ErrorCode = "CT_REACHED" }));

            SheetProtectionGuard.WrapWriteToolsWithProbe(registry, probe);

            // 受守卫：保护态活动表拦截（活动表=数据 且 activeProtected=true）
            var blocked = registry.TryGet("write_excel_cell", out var guarded)
                ? guarded.Execute(new Dictionary<string, object>())
                : null;
            Assert.NotNull(blocked);
            Assert.Equal("SHEET_PROTECTED", blocked.ErrorCode);

            // 豁免工具与 ReadOnly 工具不包装，内层可达
            Assert.Equal("INNER_REACHED", registry.TryGet("run_cli_command", out var exempt)
                ? exempt.Execute(new Dictionary<string, object>()).ErrorCode : null);
            Assert.Equal("READONLY_REACHED", registry.TryGet("preview_excel_range", out var ro)
                ? ro.Execute(new Dictionary<string, object>()).ErrorCode : null);
            Assert.Equal("CT_REACHED", registry.TryGet("ct_vba_tool", out var ct)
                ? ct.Execute(new Dictionary<string, object>()).ErrorCode : null);

            // 包装不改变工具总数
            Assert.Equal(4, registry.Tools.Count);
        }
    }
}
