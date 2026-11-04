using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 保护态探测接缝：把「读保护态」与「守卫判定」解耦。
    /// 生产实现 ComSheetProtectionProbe 走 dynamic COM；单测用假实现直接给态。
    /// 返回 false = 无法确定（表不存在/探测异常），守卫一律放行——加法前置，
    /// 不改变工具自身的既有失败路径。
    /// </summary>
    internal interface ISheetProtectionProbe
    {
        /// <summary>
        /// 读目标工作表 Contents 保护态。sheetName 为空解析活动表；
        /// resolvedName 回填实际命中的表名（用于错误消息）。表不存在或探测异常返回 false。
        /// </summary>
        bool TryGetContentsProtection(string sheetName, out string resolvedName, out bool contentsProtected);

        /// <summary>读活动工作簿结构保护态（ProtectStructure）。探测异常返回 false。</summary>
        bool TryGetStructureProtection(out bool structureProtected);
    }

    /// <summary>生产探测实现：dynamic COM 读取，遵循 M16/M17 RCW 纪律（全程零强制释放，探测对象交 GC 回收）。</summary>
    internal sealed class ComSheetProtectionProbe : ISheetProtectionProbe
    {
        private readonly dynamic application;

        public ComSheetProtectionProbe(object application)
        {
            this.application = application;
        }

        public bool TryGetContentsProtection(string sheetName, out string resolvedName, out bool contentsProtected)
        {
            resolvedName = null;
            contentsProtected = false;
            try
            {
                dynamic app = application;
                dynamic ws = string.IsNullOrWhiteSpace(sheetName)
                    ? app.ActiveSheet
                    : FindSheet(app.ActiveWorkbook, sheetName);
                if (ws == null)
                {
                    // 表不存在等归工具自身报错（SHEET_NOT_FOUND 等），守卫不越权。
                    return false;
                }

                resolvedName = Convert.ToString(ws.Name);
                contentsProtected = Convert.ToBoolean(ws.ProtectContents);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            // 刻意不做任何 FinalRelease：M16 红线 + M17 WPS 真机实证——COM RCW 按 IUnknown
            // 身份缓存，枚举命中项可能与调用方/宿主持有的引用共享同一 RCW，守卫一旦强制
            // 释放会把外部引用一并「RCW 分离」（实证：harness 后续 Unprotect 直接抛
            // 「COM 对象与其基础 RCW 分开后就不能再使用」）。探测用完即弃的短命对象交 GC 回收。
        }

        public bool TryGetStructureProtection(out bool structureProtected)
        {
            structureProtected = false;
            try
            {
                structureProtected = Convert.ToBoolean(application.ActiveWorkbook.ProtectStructure);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>按名查找工作表；返回对象由调用方只读使用、不做强制释放（见 TryGetContentsProtection 注释）。</summary>
        private static dynamic FindSheet(dynamic workbook, string name)
        {
            try
            {
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    if (string.Equals(Convert.ToString(worksheet.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return null;
        }
    }

    /// <summary>
    /// M17：写路径工具「保护态」集中守卫。
    /// 对所有非 ReadOnly 工具的 Execute 包装一道前置探测：目标工作表处于保护状态
    /// （ProtectContents）或结构类工具对应工作簿结构受保护（ProtectStructure）时，
    /// 直接返回 SHEET_PROTECTED 结构化错误，不再让原始 COM 异常（0x800A03EC 等）
    /// 散落到 AI 使用方——此前全项目仅 ExcelStructureService 三处有显式保护检查。
    /// 设计原则：
    /// ① 默认全量覆盖（写能力工具不表态即受守卫），豁免必须带理由——与
    ///    AutoVerifyPolicy 覆盖审计同构，防止新工具带着「保护态裸奔」上线；
    /// ② 守卫是「尽力而为的加法前置」：探测阶段任何异常一律放行回工具自身
    ///    错误路径（行为只增不退，不制造新的失败模式）；
    /// ③ RCW 纪律：工作表/工作簿级共享 RCW 只读使用不强制释放（M16 红线）；
    ///    枚举未命中的短命工作表对象照常释放，命中项读完后释放。
    /// </summary>
    public static class SheetProtectionGuard
    {
        /// <summary>守卫统一错误码（与 ExcelStructureService 既有用法一致）。</summary>
        public const string ErrorCode = "SHEET_PROTECTED";

        /// <summary>
        /// 有写能力但豁免于保护态守卫的工具（工具名 → 理由）。豁免必须有明确理由，
        /// 未列出的写能力工具默认受守卫。ct_ 前缀用户沉淀工具按前缀规则自动豁免。
        /// </summary>
        public static readonly IReadOnlyDictionary<string, string> Exemptions =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                // —— 脚本/进程类：执行体语义自洽，args 无可解析的目标表写入 ——
                ["run_vba_snippet"] = "VBA 脚本执行结果形态不可预知，保护态冲突由脚本自身报错。",
                ["run_python_code"] = "同 run_vba_snippet。",
                ["run_python_excel"] = "同 run_vba_snippet。",
                ["run_cli_command"] = "进程级命令，无工作簿写入语义。",
                ["run_exe_command"] = "同 run_cli_command。",
                ["save_vba_tool"] = "只写本地工具库文件，不触碰工作簿。",
                ["run_workflow"] = "流程内部步骤各自命中具体工具的守卫，外层不重复拦截。",
                ["apply_correction"] = "M29 组合修正工具：attempts 内每个子工具已各自按 sheet 参数完成守卫判定，外层再包会误判活动表。",
                ["validate_vba_snippet"] = "静态校验工具，不修改工作簿状态。",
                ["undo_excel_operation"] = "快照恢复语义：恢复目标由历史快照决定，保护态冲突由恢复自身报错。",
                // —— 工作簿/文件级：不写目标工作表单元格 ——
                ["split_workbook"] = "工作簿文件级操作，不写目标工作表单元格。",
                ["merge_workbooks"] = "同 split_workbook。",
                ["import_delimited_file"] = "导入产出为新表，源表只读。",
                ["import_excel_file_sheet"] = "同 import_delimited_file。",
                ["merge_delimited_files"] = "同 import_delimited_file。",
                ["import_markdown_to_excel"] = "同 import_delimited_file。",
                ["export_excel_chart_image"] = "只导出图片文件，不写工作表。",
                ["export_sheets_to_pdf"] = "只导出 PDF 文件，不写工作表。",
                ["copy_range_as_image"] = "只产出图片文件，不写工作表。",
                ["preview_external_file"] = "外部文件预览，不修改工作簿。",
                ["open_excel_external_book"] = "外部簿句柄管理，不写当前簿工作表。",
                ["replace_external_links"] = "工作簿级链接维护，不写工作表单元格。",
                // M28/R3：落盘动作不触碰单元格内容，保护态与它无关（保存的是既有状态）。
                ["save_workbook"] = "落盘当前工作簿文件，不写工作表单元格。",
                // —— 视图级：保护态下原生可用 ——
                ["freeze_panes"] = "窗口视图设置，保护态下原生可用，不应拦截。",
                // —— 产出为新表：源只读、写入目标非 args 解析出的源表 ——
                ["split_sheet_by_field"] = "写入目标为新建工作表，源表只读。",
                ["create_sheet_by_value"] = "同 split_sheet_by_field。",
                ["create_toc_sheet"] = "写入目标为新建目录表。",
                ["generate_analysis_report"] = "报告产出为新表或文件。",
            };

        /// <summary>
        /// 工作簿结构守卫工具：写入对象是「工作表集合」本身，走 ProtectStructure 判定
        /// 而非单表 ProtectContents（delete 额外保留目标表 ProtectContents 检查）。
        /// </summary>
        public static readonly IReadOnlyCollection<string> StructureTools = new[]
        {
            "add_excel_worksheet",
            "rename_excel_worksheet",
            "delete_excel_worksheet",
            // M-D3：delete_dashboard 删除看板+辅助两张 sheet，属工作簿结构操作，
            // 走 ProtectStructure 判定（写能力默认受守卫，Dangerous 级另过确认门）
            "delete_dashboard",
            "start_draft", "discard_draft",
            // M27R4(#3)：import_csv 执行 Worksheets.Add——工作簿 ProtectStructure 时必须给 SHEET_PROTECTED
            "import_csv"
        };

        /// <summary>目标工作表参数名覆盖（缺省按 "sheet" 解析）。</summary>
        public static readonly IReadOnlyDictionary<string, string> SheetArgOverrides =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                // 写入目标是 target_sheet；args["sheet"] 语义为只读源表
                ["copy_visible_rows_to_sheet"] = "target_sheet",
                ["merge_sheets"] = "target_sheet",
            // M26 审查修复：草稿族与 import_csv 的目标表显式映射（B2/B4）
                // M27R4(#2)：apply_draft 实际回写目标是**源表**（草稿→源 Range 回写后删草稿）——
                // 原只探 draft_sheet：源表保护时放行吃裸 0x800A03EC、草稿保护时误拦。现双探。
                ["apply_draft"] = "draft_sheet",
                ["start_draft"] = "sheet",
                ["discard_draft"] = "draft_sheet",
                ["diff_draft"] = "draft_sheet",
                ["import_csv"] = "sheet_name",
            };

        /// <summary>工具是否豁免于保护态守卫：豁免清单命中，或 ct_ 前缀用户沉淀工具。</summary>
        public static bool IsExempted(string tool)
        {
            if (string.IsNullOrWhiteSpace(tool))
            {
                return false;
            }

            return Exemptions.ContainsKey(tool)
                || tool.StartsWith("ct_", StringComparison.Ordinal);
        }

        /// <summary>
        /// 对注册表内全部写能力工具统一包装守卫（在 CreateRegistry 收口处调用一次）。
        /// ReadOnly 工具与豁免工具原样保留；其余替换为带前置探测的描述符。
        /// </summary>
        public static void WrapWriteTools(ToolRegistry registry, object application)
        {
            WrapWriteToolsWithProbe(registry, new ComSheetProtectionProbe(application));
        }

        /// <summary>接缝注入版（单测用）：探测实现可替换，判定逻辑与生产行为完全一致。</summary>
        internal static void WrapWriteToolsWithProbe(ToolRegistry registry, ISheetProtectionProbe probe)
        {
            if (registry == null || probe == null)
            {
                return;
            }

            // Tools 返回快照，替换原位安全。
            foreach (var tool in registry.Tools)
            {
                if (tool.RiskLevel == RiskLevel.ReadOnly)
                {
                    continue;
                }

                if (IsExempted(tool.Name))
                {
                    continue;
                }

                registry.Replace(WrapWithProbe(tool, probe));
            }
        }

        /// <summary>包装单个工具：保留名称/描述/风险级别/参数 schema，仅替换执行体。</summary>
        public static ToolDescriptor Wrap(ToolDescriptor tool, object application)
        {
            return WrapWithProbe(tool, new ComSheetProtectionProbe(application));
        }

        /// <summary>接缝注入版（单测用）。</summary>
        internal static ToolDescriptor WrapWithProbe(ToolDescriptor tool, ISheetProtectionProbe probe)
        {
            if (tool == null)
            {
                throw new ArgumentNullException(nameof(tool));
            }

            var inner = tool.Execute;
            return new ToolDescriptor(
                tool.Name,
                tool.Description,
                tool.RiskLevel,
                args => ExecuteGuarded(tool.Name, probe, args, inner),
                tool.Parameters);
        }

        private static ToolResult ExecuteGuarded(
            string toolName,
            ISheetProtectionProbe probe,
            Dictionary<string, object> args,
            Func<Dictionary<string, object>, ToolResult> inner)
        {
            var blocked = Evaluate(toolName, probe, args);
            if (blocked != null)
            {
                return blocked;
            }

            return inner(args);
        }

        /// <summary>
        /// 守卫判定（内部可见供单测，纯逻辑无 COM）：返回 null = 放行。
        /// 判定顺序：结构类工具走 ProtectStructure（delete 追加目标表 ProtectContents），
        /// 其余工具解析目标工作表后读 ProtectContents。
        /// </summary>
        internal static ToolResult Evaluate(string toolName, ISheetProtectionProbe probe, IReadOnlyDictionary<string, object> args)
        {
            if (probe == null)
            {
                return null;
            }

            var map = args ?? new Dictionary<string, object>();

            if (IsStructureTool(toolName))
            {
                if (!probe.TryGetStructureProtection(out var structureProtected))
                {
                    return null;
                }

                if (structureProtected)
                {
                    return Blocked(toolName, null, protectStructure: true);
                }

                if (string.Equals(toolName, "delete_excel_worksheet", StringComparison.Ordinal))
                {
                    var target = GetString(map, "sheet");
                    if (!string.IsNullOrWhiteSpace(target)
                        && probe.TryGetContentsProtection(target, out _, out var targetProtected)
                        && targetProtected)
                    {
                        return Blocked(toolName, target, protectStructure: false);
                    }
                }

                // M-D3：delete_dashboard 目标表参数名是 dashboard_sheet（非 sheet），
                // 删除对象为看板表，同样追加目标表 ProtectContents 检查
                if (string.Equals(toolName, "delete_dashboard", StringComparison.Ordinal))
                {
                    var targetDashboard = GetString(map, "dashboard_sheet");
                    if (!string.IsNullOrWhiteSpace(targetDashboard)
                        && probe.TryGetContentsProtection(targetDashboard, out _, out var dashProtected)
                        && dashProtected)
                    {
                        return Blocked(toolName, targetDashboard, protectStructure: false);
                    }
                }

                return null;
            }

            var argKey = SheetArgOverrides.TryGetValue(toolName, out var overrideKey)
                ? overrideKey
                : "sheet";
            var sheetName = GetString(map, argKey);
            if (!probe.TryGetContentsProtection(sheetName, out var resolvedName, out var contentsProtected))
            {
                return null;
            }

            if (contentsProtected)
            {
                return Blocked(toolName, resolvedName, protectStructure: false);
            }

            // M27R4(#2)：apply_draft 真实写入目标是 draft_<源表>_<时间戳> 反解出的源表——追加探源表
            if (string.Equals(toolName, "apply_draft", StringComparison.Ordinal))
            {
                var sourceSheet = TryResolveDraftSource(sheetName);
                if (!string.IsNullOrWhiteSpace(sourceSheet)
                    && probe.TryGetContentsProtection(sourceSheet, out var srcResolved, out var srcProtected)
                    && srcProtected)
                {
                    return Blocked(toolName, srcResolved, protectStructure: false);
                }
            }

            return null;
        }

        /// <summary>M27R4：从草稿表名 draft_<源表>_<时间戳> 反解源表名；不符合约定返回 null。</summary>
        internal static string TryResolveDraftSource(string draftName)
        {
            if (string.IsNullOrWhiteSpace(draftName) || !draftName.StartsWith("draft_", StringComparison.Ordinal))
            {
                return null;
            }

            var body = draftName.Substring(6);
            var lastUnderscore = body.LastIndexOf('_');
            if (lastUnderscore <= 0)
            {
                return null;
            }

            var tail = body.Substring(lastUnderscore + 1);
            var allDigits = tail.Length > 0;
            foreach (var c in tail)
            {
                if (!char.IsDigit(c)) { allDigits = false; break; }
            }

            return allDigits ? body.Substring(0, lastUnderscore) : null;
        }

        private static bool IsStructureTool(string toolName)
        {
            foreach (var name in StructureTools)
            {
                if (string.Equals(name, toolName, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static ToolResult Blocked(string toolName, string sheetName, bool protectStructure)
        {
            var target = protectStructure
                ? "工作簿结构处于保护状态，工具 " + toolName + " 无法增删改名工作表"
                : "工作表 '" + sheetName + "' 处于保护状态，工具 " + toolName + " 无法写入";
            return new ToolResult
            {
                Ok = false,
                ErrorCode = ErrorCode,
                Data = target + "；请先解除保护（可用 verify_protection 探测保护状态）后重试。"
            };
        }

        private static string GetString(IReadOnlyDictionary<string, object> map, string key)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            if (value is Newtonsoft.Json.Linq.JValue jv)
            {
                return jv.Type == Newtonsoft.Json.Linq.JTokenType.Null ? null : Convert.ToString(jv.Value);
            }

            return Convert.ToString(value);
        }
    }
}
