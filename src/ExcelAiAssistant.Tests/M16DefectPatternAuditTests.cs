using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M16 缺陷原型自动审计（M17.3，grill 清单⑤）：
    /// 把 M13–M16 全量审查沉淀的缺陷原型从"文档清单 + 行为回归锁"升格为
    /// 门禁源码级自动检测——新增/重构代码若丢弃防护基础设施，单测门禁首轮即红。
    ///
    /// 审计对象：src/ExcelAiAssistant.Excel 生产源码（注释剥离后扫描）。
    /// 原型 → 规则映射（详见 docs/WPS工具分类审查报告.md §十二/§十四）：
    ///   R1  CVErr 魔数只允许存在于 ExcelErrorValues.cs（单一来源，禁止散落）
    ///   R2  ExcelErrorValues 必须接入全部曾接入的 7 个统计/加工文件
    ///   R3  Workbooks.Open 借用语义（userOwned）不得丢失
    ///   R4  run_cli_command 白名单必须按完整路径评估（basename 绕过修复不得回退）
    ///   R5  fetch_url SSRF 双防护（mapped-IPv6 复核 + DNS 全地址校验）不得丢失
    ///   R6  markdown 管道符导出转义/导入反转义往返配套不得丢失
    ///   R7  工作簿级文件操作必须过 PathGuard
    ///   R8  Application/ActiveWorkbook/ActiveSheet 级共享 RCW 禁止 FinalRelease（M16/M17 红线）
    ///   R9  SheetProtectionGuard 探测零释放（M17 WPS 真机实证红线）
    ///   R10 宿主契约：snake_case 投影 + write 形状拟合上报在位
    ///   R11 sheet 清单 VeryHidden 归因（hidden_state）在位
    /// </summary>
    public sealed class M16DefectPatternAuditTests
    {
        private static readonly string[] ExcelErrorIntegrations =
        {
            "ColumnProfileEngine",
            "ExcelDataToolsService",
            "ExcelUtilityToolsService",
            "ExcelAnalysisService",
            "ExcelJyyjtoolService",
            "ExcelAdvancedEditService",
            "ExcelPivotChartService"
        };

        private static readonly string[] PathGuardFiles =
        {
            "ExcelWorkbookService",
            "ExcelChartService",
            "ExcelReportService"
        };

        // ===================== 基础设施 =====================

        /// <summary>从测试基目录向上定位仓库根（含 src/ExcelAiAssistant.Excel）。</summary>
        private static string FindProductionSourceDir()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            for (int i = 0; i < 12 && dir != null; i++, dir = dir.Parent)
            {
                var candidate = Path.Combine(dir.FullName, "src", "ExcelAiAssistant.Excel");
                if (Directory.Exists(candidate))
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException("未定位到 src/ExcelAiAssistant.Excel（从 " + AppContext.BaseDirectory + " 向上查找失败）");
        }

        /// <summary>读取全部生产源码为 (文件名, 剥注释源码)。</summary>
        private static List<KeyValuePair<string, string>> LoadProductionSources()
        {
            var dir = FindProductionSourceDir();
            return Directory.GetFiles(dir, "*.cs")
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .Select(p => new KeyValuePair<string, string>(
                    Path.GetFileName(p),
                    StripLineComments(File.ReadAllText(p))))
                .ToList();
        }

        /// <summary>
        /// 逐行剥离 // 行注释（含 /// 文档注释）。不处理 /* */ 块注释——
        /// 当前生产代码无块注释包裹 CVErr 数字的场景；若未来引入，审计宁可漏报也不误报，
        /// 由代码评审兜底。字符串字面量中的 "//"（如 URL）会被截断，对数字扫描无影响。
        /// </summary>
        private static string StripLineComments(string source)
        {
            var lines = source.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                int idx = lines[i].IndexOf("//", StringComparison.Ordinal);
                if (idx >= 0)
                {
                    lines[i] = lines[i].Substring(0, idx);
                }
            }

            return string.Join("\n", lines);
        }

        private static string LoadOne(List<KeyValuePair<string, string>> sources, string fileName)
        {
            var hit = sources.FirstOrDefault(kv =>
                string.Equals(kv.Key, fileName, StringComparison.OrdinalIgnoreCase));
            Assert.True(hit.Key != null, $"生产源码中缺少 {fileName}（审计基础设施自身失效）");
            return hit.Value;
        }

        // ===================== R1 CVErr 魔数单一来源 =====================

        public static IEnumerable<object[]> CveErrMagicNumbers()
        {
            // 7 码的十进制形态（0x800A07xx 高位 Int32）
            yield return new object[] { "-2146826246" };
            yield return new object[] { "-2146826253" };
            yield return new object[] { "-2146826259" };
            yield return new object[] { "-2146826265" };
            yield return new object[] { "-2146826275" };
            yield return new object[] { "-2146826281" };
            yield return new object[] { "-2146826288" };
        }

        [Theory]
        [MemberData(nameof(CveErrMagicNumbers))]
        public void R1_cveerr_magic_numbers_exist_only_in_excel_error_values(string magic)
        {
            foreach (var kv in LoadProductionSources())
            {
                if (string.Equals(kv.Key, "ExcelErrorValues.cs", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // 单一来源本体
                }

                Assert.True(!kv.Value.Contains(magic),
                    $"R1 违规：{kv.Key} 的代码中出现 CVErr 魔数 {magic}。" +
                    "错误值识别必须引用 ExcelErrorValues（单一来源），禁止硬编码散落。");
            }
        }

        // ===================== R2 ExcelErrorValues 接入在位 =====================

        [Fact]
        public void R2_excel_error_values_integrated_into_all_consumer_files()
        {
            foreach (var fileName in ExcelErrorIntegrations)
            {
                var src = LoadOne(LoadProductionSources(), fileName + ".cs");
                Assert.True(src.Contains("ExcelErrorValues"),
                    $"R2 违规：{fileName} 不再引用 ExcelErrorValues。" +
                    "CVErr 归一接入被移除——错误值将静默变数字污染统计/排序/写回（M16 P1 原型回潮）。");
            }
        }

        // ===================== R3 Workbooks.Open 借用语义 =====================

        [Fact]
        public void R3_import_borrow_semantics_preserved()
        {
            var import = LoadOne(LoadProductionSources(), "ExcelImportService.cs");
            Assert.True(import.Contains("Workbooks.Open"),
                "R3 违规：ExcelImportService 不再使用 Workbooks.Open（审计假设失效，需复核规则）。");
            Assert.True(import.Contains("ownedMates"),
                "R3 违规：ExcelImportService 丢失 ownedMates 借用隔离——" +
                "对用户已打开文件将回到\"finally 无提示关闭用户工作簿\"的 P0 原型。");

            var registry = LoadOne(LoadProductionSources(), "ExternalBookRegistry.cs");
            Assert.True(registry.Contains("userOwned"),
                "R3 违规：ExternalBookRegistry 丢失 userOwned 标记，借用语义无法判定归属。");
        }

        // ===================== R4 CLI 白名单全路径评估 =====================

        [Fact]
        public void R4_cli_whitelist_evaluates_full_path()
        {
            var src = LoadOne(LoadProductionSources(), "ScriptRunnerService.cs");
            Assert.True(src.Contains("ControlledExecutionPolicy.Evaluate"),
                "R4 违规：CLI 执行不再过 ControlledExecutionPolicy（受控执行策略被绕过）。");
            Assert.True(src.Contains("CLI_DENIED"),
                "R4 违规：CLI 白名单拒绝错误码 CLI_DENIED 丢失。");
            // 源文件磁盘文本含字符字面量 '\\'（反斜杠转义），此处按磁盘文本匹配
            Assert.True(src.Contains("IndexOf('\\\\')"),
                "R4 违规：白名单评估丢失反斜杠路径分支——" +
                "回到\"basename 命中白名单、按原样路径执行\"的 M16 P1 绕过原型。");
            Assert.True(src.Contains("IndexOf('/')"),
                "R4 违规：路径成分分支只查反斜杠——正斜杠路径（C:/temp/python.exe）将绕过完整路径评估。");
            Assert.True(src.Contains("GetFullPath"),
                "R4 违规：完整路径评估丢失 GetFullPath 归一。");
        }

        // ===================== R5 SSRF 双防护 =====================

        [Fact]
        public void R5_fetch_url_ssrf_guards_preserved()
        {
            var src = LoadOne(LoadProductionSources(), "WebFetchService.cs");
            Assert.True(src.Contains("BLOCKED_PRIVATE_ADDR"),
                "R5 违规：BLOCKED_PRIVATE_ADDR 错误码丢失，SSRF 拒绝语义不可观测。");
            Assert.True(src.Contains("GetHostAddresses"),
                "R5 违规：DNS 解析校验丢失——主机名不解析即放行会回到 DNS 重绑定/xip.io 绕过原型。");
            Assert.True(src.Contains("IsPrivateOrReserved"),
                "R5 违规：IsPrivateOrReserved 私有地址判定丢失。");
        }

        // ===================== R6 markdown 管道符往返 =====================

        [Fact]
        public void R6_markdown_pipe_roundtrip_pair_preserved()
        {
            var src = LoadOne(LoadProductionSources(), "ExcelMarkdownService.cs");
            Assert.True(src.Contains("SplitUnescapedPipe"),
                "R6 违规：导入侧 SplitUnescapedPipe（按未转义 | 拆分并反转义）丢失——" +
                "含管道符单元格在导出→再导入往返中被错误拆列。");
            Assert.True(src.Contains("Replace(\"|\", \"\\\\|\")"),
                "R6 违规：导出侧 | → \\| 转义丢失，与导入侧反转义不再配套。");
        }

        // ===================== R7 工作簿级操作 PathGuard =====================

        [Fact]
        public void R7_workbook_level_operations_use_path_guard()
        {
            foreach (var fileName in PathGuardFiles)
            {
                var src = LoadOne(LoadProductionSources(), fileName + ".cs");
                Assert.True(src.Contains("PathGuard"),
                    $"R7 违规：{fileName} 丢失 PathGuard——工作簿级文件操作未限定路径范围。");
            }
        }

        // ===================== R8 共享 RCW FinalRelease 红线 =====================

        [Fact]
        public void R8_no_final_release_on_application_level_objects()
        {
            // M16/M17 真机实证：Workbook/Worksheet 级共享 RCW 禁止 FinalRelease
            //（WPS 下 RCW 按 IUnknown 身份缓存，强释放连带分离外部持有人）。
            // Application/Active*（activeWorkbook/activeSheet/activeCell 等任意 active
            // 前缀共享对象）级变量全项目一律禁止强释放；其余变量名走白名单制（宁漏报不误报）。
            var pattern = new Regex(
                @"FinalReleaseComObject\s*\(\s*(app|application|sharedapp|active[a-z]*)\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
            foreach (var kv in LoadProductionSources())
            {
                Assert.True(!pattern.IsMatch(kv.Value),
                    $"R8 违规：{kv.Key} 对 Application/Active 级共享 RCW 调用 FinalReleaseComObject。" +
                    "共享对象只读使用、交 GC 回收（M16 红线）。");
            }
        }

        // ===================== R9 SheetProtectionGuard 零释放 =====================

        [Fact]
        public void R9_protection_guard_probe_does_zero_com_release()
        {
            var src = LoadOne(LoadProductionSources(), "SheetProtectionGuard.cs");
            Assert.True(!src.Contains("FinalReleaseComObject"),
                "R9 违规：SheetProtectionGuard 探测路径出现 FinalReleaseComObject——" +
                "探测类 COM 代码必须零释放（枚举命中项强释放会分离外部持有人的 RCW，M17 WPS 真机实证）。");
        }

        // ===================== R10 宿主契约在位 =====================

        [Fact]
        public void R10_host_snake_case_projection_and_refit_reporting_preserved()
        {
            var host = LoadOne(LoadProductionSources(), "ExcelToolHost.cs");
            Assert.True(host.Contains("ToSnakeCase"),
                "R10 违规：宿主 snake_case 投影丢失，context 输出回到 PascalCase（M16 契约回潮）。");
            Assert.True(host.Contains("refit_source_cells"),
                "R10 违规：write 形状拟合上报 refit_source_cells 丢失，AI 使用方失去形状重拟合观测。");
        }

        // ===================== R11 VeryHidden 归因在位 =====================

        [Fact]
        public void R11_sheet_list_reports_hidden_state()
        {
            var src = LoadOne(LoadProductionSources(), "ExcelReadService.cs");
            Assert.True(src.Contains("hidden_state"),
                "R11 违规：sheet 清单 hidden_state 字段丢失——VeryHidden 表将误报为可见（M16 原型回潮）。");
        }
    }
}
