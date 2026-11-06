using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Errors
{
    /// <summary>单个错误码的人话解释与下一步动作。</summary>
    public sealed class ErrorEntry
    {
        public ErrorEntry(string code, string message, string nextAction)
        {
            Code = code; Message = message; NextAction = nextAction;
}

        public string Code { get; }
        public string Message { get; }
        public string NextAction { get; }
    }

    /// <summary>
    /// 错误码目录（M36-3）：把散落在各工具里的字符串码收敛为「码 → 人话 → next_action」，
    /// MCP / CLI / 侧边栏共用同一份解释，避免同一个码在不同入口给出不同说法。
    /// 未登记的码不影响原有行为（TryGet 返回 false，调用方跳过富化）。
    /// </summary>
    public static class ErrorCatalog
    {
        private static readonly Dictionary<string, ErrorEntry> Map =
            new Dictionary<string, ErrorEntry>(StringComparer.OrdinalIgnoreCase)
        {
            ["EXCEL_WARMING"] = new ErrorEntry("EXCEL_WARMING", "Excel 实例还在启动（预热未完成），本次未执行。", "等 10-30 秒后重试同一调用；连续失败请检查是否有弹窗阻塞 Excel（受保护视图/宏提示）。"),
            ["PROFILE_FILTERED"] = new ErrorEntry("PROFILE_FILTERED", "该工具存在于服务端，但未在当前工具面中暴露。", "把环境变量 JYYJ_TOOL_PROFILE 设为 all 后重启 MCP 服务；或改用 tools/list 中可见的等价工具。"),
            ["DUPLICATE_WRITE_SUPPRESSED"] = new ErrorEntry("DUPLICATE_WRITE_SUPPRESSED", "相同工具与参数在 10 分钟内已执行过，本次未重复写入。", "若确需再执行同一操作，在 args 里加 allow_duplicate=true；否则忽略本次结果。"),
            ["TOOL_EXCEPTION"] = new ErrorEntry("TOOL_EXCEPTION", "工具内部抛出异常，本次执行中断。", "读 message 里的异常信息；若为参数类型/缺失字段，对照 inputSchema 修正后重试；持续出现请附日志反馈。"),
            ["UNKNOWN_TOOL"] = new ErrorEntry("UNKNOWN_TOOL", "工具名不存在，或当前模式/工具面下不可用（名字拼错也会走到这里）。", "调用 tools/list 核对工具名；写操作需以 --allow-write 启动。"),
            ["INVALID_ARGS"] = new ErrorEntry("INVALID_ARGS", "工具参数不合法（缺字段、类型不对，或主工具实参被平铺到顶层）。", "对照 tools/list 的 inputSchema 修正；run_self_check/apply_correction 需嵌套 {tool, args}。"),
            ["EMPTY_RANGE"] = new ErrorEntry("EMPTY_RANGE", "目标区域没有任何值，无法完成统计/渲染/清洗。", "先确认区域地址与工作表名；可用 get_excel_range_values_and_formulas 预览实际内容。"),
            ["NO_VERIFY_PLAN"] = new ErrorEntry("NO_VERIFY_PLAN", "该写工具没有可用的自检计划，无法验证写入结果。", "改用带 verify_* 的写工具，或后续显式调用对应 verify 工具核对。"),
            ["SELF_CHECK_FAILED"] = new ErrorEntry("SELF_CHECK_FAILED", "自检未通过，写入结果与期望不一致。", "读取 checks 内每条 verdict/data 定位差异单元格，用 apply_correction 修正后再验证。"),
            ["SELF_CHECK_SKIPPED"] = new ErrorEntry("SELF_CHECK_SKIPPED", "verify 工具未注册或在当前模式下不可用，自检被跳过。", "确认 verify 工具在 tools/list 中可见；只读模式或 core 工具面可能不含它。"),
            ["CORRECTION_FAILED"] = new ErrorEntry("CORRECTION_FAILED", "多轮自动修正仍未通过自检。", "查看各 attempt 的 main_error_code 与 verdict，调整参数或换工具，不要盲目增加轮数。"),
            ["CONFIRMATION_REQUIRED"] = new ErrorEntry("CONFIRMATION_REQUIRED", "该操作属危险级，需要用户显式确认。", "向用户展示操作内容并获得确认后，带确认码重试；不要在未确认时循环重试。"),
            ["CONFIRMATION_EXHAUSTED"] = new ErrorEntry("CONFIRMATION_EXHAUSTED", "确认码尝试次数已用尽。", "重新发起该操作获取新确认码，或改用非危险级替代方案。"),
            ["CONFIRMATION_CODE_INVALID"] = new ErrorEntry("CONFIRMATION_CODE_INVALID", "确认码无效或已过期。", "重新调用原工具获取新的确认码，注意确认码有 TTL。"),
            ["CONFIRMATION_ARGS_MISMATCH"] = new ErrorEntry("CONFIRMATION_ARGS_MISMATCH", "确认码与本次参数不匹配（确认后又改了参数）。", "确认后不要修改参数；如参数变化需重新走确认流程。"),
            ["RESOURCE_UNAVAILABLE"] = new ErrorEntry("RESOURCE_UNAVAILABLE", "所需资源（Excel 实例/工作表/图表等）当前不可用。", "检查目标工作簿是否已打开、表名是否正确，必要时重新初始化会话。"),
            ["WRITE_FAILED"] = new ErrorEntry("WRITE_FAILED", "写入失败（COM 异常或目标不可写）。", "检查工作表是否受保护、区域是否被合并单元格覆盖，重试前先读取现状。"),
            ["WRITE_OOM"] = new ErrorEntry("WRITE_OOM", "写入数据量超出内存/单元格上限。", "拆分为多次小批量写入，或先降采样再写。"),
            ["PARAM_MOVED"] = new ErrorEntry("PARAM_MOVED", "参数位置/层级不对（常见于把实参平铺到顶层）。", "按 inputSchema 嵌套参数；闭环类工具用 {tool, args}。"),
            ["SHEET_PROTECTED"] = new ErrorEntry("SHEET_PROTECTED", "目标工作表处于保护状态，写入被守卫拦截。", "请用户取消保护（审阅→撤销工作表保护）后重试；或在未保护的表里操作。"),
            ["DELETE_COUNT_MISMATCH"] = new ErrorEntry("DELETE_COUNT_MISMATCH", "实际删除行数与预期不一致（可能存在隐藏行/筛选）。", "取消筛选、展开隐藏行后重试，并复核删除范围。"),
            ["CONSERVATION_VIOLATED"] = new ErrorEntry("CONSERVATION_VIOLATED", "数据守恒校验失败（汇总值/行数对不上）。", "先备份，再逐段核对受影响区域，确认是否有多余的重复删除。"),
            ["BATCH_WRITE_PARTIAL"] = new ErrorEntry("BATCH_WRITE_PARTIAL", "批量写入部分成功、部分失败。", "读取 affected_ranges 确认成功范围，仅对失败部分重试。"),
            ["BLANK_RENDER"] = new ErrorEntry("BLANK_RENDER", "渲染结果为空图（区域无可见内容）。", "确认区域有值且未被隐藏；必要时先写入数据再渲染。"),
            ["CHART_NOT_FOUND"] = new ErrorEntry("CHART_NOT_FOUND", "指定图表不存在。", "用 list/diagnose 类工具确认图表名与所在工作表。"),
            ["ELEMENT_NOT_FOUND"] = new ErrorEntry("ELEMENT_NOT_FOUND", "指定元素（形状/草稿/看板等）不存在。", "刷新元素清单后重试，注意元素可能已被删除或改名。"),
            ["FILE_NOT_FOUND"] = new ErrorEntry("FILE_NOT_FOUND", "文件路径不存在。", "核对绝对路径与扩展名；相对路径请改为绝对路径。"),
            ["FILE_TOO_LARGE"] = new ErrorEntry("FILE_TOO_LARGE", "文件超过处理上限。", "拆分文件或先裁剪数据区域再导入。"),
            ["CLI_DENIED"] = new ErrorEntry("CLI_DENIED", "命令执行被策略拒绝。", "该命令不在允许清单内；请改用受支持的工具或让用户手动执行。"),
            ["EXE_DENIED"] = new ErrorEntry("EXE_DENIED", "可执行文件启动被策略拒绝。", "该程序不在允许清单内；请让用户手动启动或改用内置工具。"),
            ["BLOCKED_PRIVATE_ADDR"] = new ErrorEntry("BLOCKED_PRIVATE_ADDR", "目标地址属于内网/私有网段，被网络策略拦截。", "改用公网地址，或让用户提供已下载的本地文件。"),
        };

        public static bool TryGet(string code, out ErrorEntry entry)
        {
            entry = null;
            return !string.IsNullOrWhiteSpace(code) && Map.TryGetValue(code.Trim(), out entry);
}

        public static string MessageOf(string code) => TryGet(code, out var e) ? e.Message : null;

        public static string NextActionOf(string code) => TryGet(code, out var e) ? e.NextAction : null;

        public static IReadOnlyCollection<ErrorEntry> All => Map.Values;
    }
}
