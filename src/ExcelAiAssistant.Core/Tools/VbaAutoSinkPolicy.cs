using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Agent;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 聊天自动沉淀判定：本轮成功执行过 run_vba_snippet、且持有最近一次成功代码、
    /// 且本轮从未调用过 save_vba_tool 时，系统应自动把代码沉淀为可复用工具。
    /// 抽成纯函数以便单测覆盖，防止「条件写反」导致功能静默失效。
    /// </summary>
    public static class VbaAutoSinkPolicy
    {
        public static bool ShouldAutoSink(IReadOnlyList<ToolEventRecord> toolEvents, string lastVbaSnippet)
        {
            if (toolEvents == null || string.IsNullOrWhiteSpace(lastVbaSnippet))
            {
                return false;
            }

            var hadSuccessfulRunVba = toolEvents.Any(ev =>
                ev.Result?.Ok == true && string.Equals(ev.Tool, "run_vba_snippet", StringComparison.OrdinalIgnoreCase));
            if (!hadSuccessfulRunVba)
            {
                return false;
            }

            // 只要 save_vba_tool 被调用过（无论成败/是否被用户拒绝）就不再自动沉淀，避免重复打扰。
            var saveVbaToolCalled = toolEvents.Any(ev =>
                string.Equals(ev.Tool, "save_vba_tool", StringComparison.OrdinalIgnoreCase));
            return !saveVbaToolCalled;
        }
    }
}
