using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 外部 AI 智能体桥接工具族（M26）：workbuddy 记忆读写 + 项目文档查询。
    /// 纯文件操作（无 COM），供 MCP server / CLI / 加载项统一注册，
    /// 实现「DSH / workbuddy 等 AI 智能体 ⇄ JYYJ 工具链」双向连接。
    /// </summary>
    public static class AgentBridgeToolRegistration
    {
        public static void Register(ToolRegistry registry, string memoryRoot = null)
        {
            if (registry == null) throw new ArgumentNullException(nameof(registry));

            registry.Register(new ToolDescriptor(
                "buddy_read_memory",
                "读取 workbuddy 长期记忆：默认 MEMORY.md；可传 name 指定文件名或 date 指定日期日志（YYYY-MM-DD）。",
                RiskLevel.ReadOnly,
                args => Bubble(AgentMemoryBridge.Read(Get(args, "memory_root") ?? memoryRoot, Get(args, "name"), Get(args, "date"), GetInt(args, "max_chars", 20000)))));

            registry.Register(new ToolDescriptor(
                "buddy_search_memory",
                "跨 workbuddy 全部记忆文件搜索关键词，返回文件名+行号+命中行（外部智能体查项目历史用）。",
                RiskLevel.ReadOnly,
                args =>
                {
                    var q = Get(args, "query");
                    if (string.IsNullOrWhiteSpace(q)) return Fail("INVALID_ARGS", "缺少必填参数：query");
                    return Bubble(AgentMemoryBridge.Search(Get(args, "memory_root") ?? memoryRoot, q, GetInt(args, "limit", 30)));
                }));

            registry.Register(new ToolDescriptor(
                "buddy_append_memory",
                "向 workbuddy 记忆追加一条条目（只追加不改写；写操作有审计）。",
                RiskLevel.NormalWrite,
                args =>
                {
                    var text = Get(args, "text");
                    if (string.IsNullOrWhiteSpace(text)) return Fail("INVALID_ARGS", "缺少必填参数：text");
                    return Bubble(AgentMemoryBridge.Append(Get(args, "memory_root") ?? memoryRoot, text, Get(args, "name"), Get(args, "date")));
                }));

            registry.Register(new ToolDescriptor(
                "list_agent_bridges",
                "列出本工具链暴露给外部 AI 智能体（DSH/workbuddy 等）的桥接能力与接入方式（MCP/CLI）。",
                RiskLevel.ReadOnly,
                _ => Ok(new
                {
                    mcp_server = "tools/McpServer/ExcelAiMcpServer.exe（stdio，--allow-write 开放写）",
                    cli = "tools/CliServer/ExcelAiCli.exe（list/schema/exec，--allow-write/--confirmed）",
                    bridges = new[]
                    {
                        "buddy_read_memory：读 .workbuddy/memory（MEMORY.md / 日期日志 / 任意文件）",
                        "buddy_search_memory：跨记忆文件关键词搜索",
                        "buddy_append_memory：追加记忆条目（NormalWrite，走审计）",
                        "list_agent_bridges：本清单"
                    },
                    excel_tools = "同进程另有 167 个 Excel 工具（MCP 客户端可一并调用）",
                    memory_root = AgentMemoryBridge.ResolveMemoryRoot(memoryRoot)
                })));
        }

        private static string Get(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                var s = Convert.ToString(v);
                return string.IsNullOrEmpty(s) ? null : s;
            }
            return null;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int def)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null && int.TryParse(Convert.ToString(v), out var i)) return i;
            return def;
        }

        private static ToolResult Ok(object data) => new ToolResult { Ok = true, Data = data };

        /// <summary>FIX（审查 E-2）：内层 ok:false 必须冒泡成 ToolResult.Ok=false，否则 MCP transport
        /// 只看到 isError=false，AI 会把「记忆文件不存在」当成功并编造内容。</summary>
        private static ToolResult Bubble(object payload)
        {
            var okProp = payload == null ? null : payload.GetType().GetProperty("ok");
            var isOk = okProp == null || !(okProp.GetValue(payload) is bool b) || b;
            if (isOk) return new ToolResult { Ok = true, Data = payload };
            var errProp = payload.GetType().GetProperty("error");
            return new ToolResult { Ok = false, ErrorCode = "MEMORY_FAILED", Data = errProp == null ? payload : errProp.GetValue(payload), Warning = Convert.ToString(payload) };
        }
        private static ToolResult Fail(string code, string message) => new ToolResult { Ok = false, ErrorCode = code, Data = message };
    }
}
