using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Connectors;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 本机 AI 客户端接入工具（MCP host 探测 / 一键登记 / 连通性自测）。
    /// 设计意图：让"把 JYYJ 接到 DSH、WorkBuddy、Codex、Claude"从手改 JSON 变成一句话。
    /// 安全边界：
    ///   · detect / test 只读，不改任何客户端配置；
    ///   · connect 默认只预览（apply=false），落盘前把原文件备份成 *.jyyj-bak；
    ///   · connect 只写/改一个叫 jyyj-excel 的条目，不删不改别人的条目；
    ///   · 不启动未知命令——test 只会执行该客户端配置里 jyyj-excel 自己那条。
    /// </summary>
    public static class ExcelAiClientToolsRegistration
    {
        public static void RegisterAiClientTools(ToolRegistry registry)
        {
            if (registry == null) { throw new ArgumentNullException(nameof(registry)); }

            var service = new AiClientService();

            registry.Register(new ToolDescriptor(
                "detect_ai_clients",
                "扫描本机已安装的 AI 客户端（Claude Desktop / Claude Code / Codex CLI / WorkBuddy / DSH / Cursor / Windsurf / Trae CN），"
                + "逐个报告：是否安装、MCP 配置文件路径、是否已登记 jyyj-excel、当前指向的启动命令。只读，不修改任何文件。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.Detect(args)),
                DetectSchema()));

            registry.Register(new ToolDescriptor(
                "connect_ai_client",
                "把 JYYJ 的 MCP server 登记到某个 AI 客户端（默认只返回预览，apply=true 才写盘）。"
                + "写前自动备份原配置为 *.jyyj-bak；只新增/更新名为 jyyj-excel 的条目，不动其他条目。"
                + "client 传 all 时批量处理所有已安装客户端。参数 allow_write=true 会给 server 传 --allow-write（让 AI 具备写表格能力）。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.Connect(args)),
                ConnectSchema()));

            registry.Register(new ToolDescriptor(
                "test_ai_client_connection",
                "用某个 AI 客户端配置里 jyyj-excel 的真实 command+args 启动一次 MCP 握手，验证该客户端确实能连上 JYYJ。"
                + "返回 reachable / protocolVersion / 握手耗时；client 省略时测试所有已登记客户端。只读。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.Test(args)),
                TestSchema()));
        }

        // ---------------- 服务实现 ----------------

        private sealed class AiClientService
        {
            public ToolResult Detect(Dictionary<string, object> args)
            {
                var clients = AiClientDiscovery.Detect();
                var installed = clients.Where(c => c.Installed || c.Configured).ToList();
                return new ToolResult
                {
                    Ok = true,
                    Data = new Dictionary<string, object>
                    {
                        ["server_command"] = AiClientDiscovery.ResolveServerCommand(),
                        ["installed_count"] = installed.Count,
                        ["configured_count"] = clients.Count(c => c.Configured),
                        ["clients"] = installed.Select(c => new Dictionary<string, object>
                        {
                            ["id"] = c.Id,
                            ["name"] = c.DisplayName,
                            ["installed"] = c.Installed,
                            ["configured"] = c.Configured,
                            ["config_path"] = c.ConfigPath,
                            ["command"] = c.Command,
                            ["args"] = c.Arguments,
                            ["detail"] = c.Detail
                        }).ToList(),
                        // 未安装 / 目录不在约定路径的客户端也列出来，避免模型把"没探测到"读成"不支持"
                        ["supported_clients"] = AiClientDiscovery.ClientIds,
                        ["note"] = "clients 只列已安装或已登记项；不在其中不代表不支持，可用 id 见 supported_clients。"
                                 + "接入顺序：connect_ai_client(apply=true) → 重启该客户端 → test_ai_client_connection。"
                    }
                };
            }

            public ToolResult Connect(Dictionary<string, object> args)
            {
                var client = GetString(args, "client");
                if (string.IsNullOrWhiteSpace(client))
                {
                    return Fail("INVALID_ARGS", "缺少 client（如 workbuddy / dsh / codex / claude-desktop / all）");
                }

                var apply = GetBool(args, "apply");
                var allowWrite = GetBool(args, "allow_write");
                var command = GetString(args, "command");
                var extraArgs = GetStringArray(args, "args");
                var serverArgs = new List<string>();
                if (allowWrite) { serverArgs.Add("--allow-write"); }
                if (extraArgs != null) { serverArgs.AddRange(extraArgs); }

                var targets = string.Equals(client, "all", StringComparison.OrdinalIgnoreCase)
                    ? AiClientDiscovery.Detect().Where(c => c.Installed).Select(c => c.Id).ToList()
                    : new List<string> { client };

                if (targets.Count == 0)
                {
                    return Fail("NO_CLIENT", "没有检测到已安装的 AI 客户端");
                }

                var results = new List<Dictionary<string, object>>();
                var anyError = false;
                foreach (var id in targets)
                {
                    var r = AiClientDiscovery.Connect(id, command, serverArgs.ToArray(), apply);
                    if (!string.IsNullOrEmpty(r.Error)) { anyError = true; }
                    results.Add(new Dictionary<string, object>
                    {
                        ["client"] = id,
                        // status 是给模型判定"到底落盘没有"的硬字段：error 为空不等于已写入
                        ["status"] = !string.IsNullOrEmpty(r.Error) ? "failed" : (!apply ? "preview" : (r.Changed ? "written" : "unchanged")),
                        ["applied"] = r.Applied,
                        ["changed"] = r.Changed,
                        ["action"] = r.Action,
                        ["config_path"] = r.ConfigPath,
                        ["backup"] = r.BackupPath,
                        ["backup_created"] = r.BackupCreated,
                        ["preview"] = r.Preview,
                        ["error"] = r.Error
                    });
                }

                var succeeded = results.Count(x => Convert.ToString(x["status"]) != "failed");
                var failed = results.Count - succeeded;
                return new ToolResult
                {
                    Ok = !anyError,
                    ErrorCode = anyError ? (succeeded > 0 ? "PARTIAL_FAILURE" : "CONNECT_FAILED") : null,
                    Data = new Dictionary<string, object>
                    {
                        ["apply"] = apply,
                        ["allow_write"] = allowWrite,
                        ["resolved_command"] = command,
                        ["server_args"] = serverArgs,
                        ["succeeded_count"] = succeeded,
                        ["failed_count"] = failed,
                        ["partial"] = anyError && succeeded > 0,
                        ["results"] = results,
                        ["note"] = apply
                            ? "status=written 的项已落盘（备份只在首次创建，不会覆盖原始基线）。必须重启对应 AI 客户端才会加载新配置；若目标客户端就是当前会话宿主，重启会断开本会话——先告知用户。"
                            : "status=preview 表示只在内存里预览、未写盘。确认无误后传 apply=true 再调用一次。"
                    }
                };
            }

            public ToolResult Test(Dictionary<string, object> args)
            {
                var client = GetString(args, "client");
                var timeout = GetInt(args, "timeout_ms", 8000);
                if (timeout < 1000) { timeout = 1000; }
                if (timeout > 60000) { timeout = 60000; }

                var targets = string.IsNullOrWhiteSpace(client)
                    ? AiClientDiscovery.Detect().Where(c => c.Configured).Select(c => c.Id).ToList()
                    : (string.Equals(client, "all", StringComparison.OrdinalIgnoreCase)
                        ? AiClientDiscovery.Detect().Where(c => c.Configured).Select(c => c.Id).ToList()
                        : new List<string> { client });

                if (targets.Count == 0)
                {
                    return Fail("NOT_CONFIGURED",
                        "没有已登记 jyyj-excel 的客户端。先调 connect_ai_client({client, apply:true})；"
                        + "若 tools/list 里没有 connect_ai_client，说明本 server 是只读模式，需要用户加 --allow-write 后重启。");
                }

                var results = new List<Dictionary<string, object>>();
                foreach (var id in targets)
                {
                    var r = AiClientDiscovery.Test(id, timeout);
                    results.Add(new Dictionary<string, object>
                    {
                        ["client"] = id,
                        ["configured"] = r.Configured,
                        ["reachable"] = r.Reachable,
                        ["protocol_version"] = r.ProtocolVersion,
                        ["elapsed_ms"] = r.ElapsedMs,
                        ["reason"] = r.Reason,
                        ["error"] = r.Error
                    });
                }

                var reachable = results.Count(x => Convert.ToBoolean(x["reachable"]));
                var unreachable = results.Where(x => !Convert.ToBoolean(x["reachable"]))
                    .Select(x => Convert.ToString(x["client"])).ToList();
                return new ToolResult
                {
                    // ok 只表示"至少一个通"；要判断"全部接入"请看 all_reachable
                    Ok = reachable > 0,
                    ErrorCode = reachable > 0 ? null : "NO_REACHABLE_CLIENT",
                    Data = new Dictionary<string, object>
                    {
                        ["reachable_count"] = reachable,
                        ["tested_count"] = results.Count,
                        ["all_reachable"] = reachable == results.Count,
                        ["unreachable_clients"] = unreachable,
                        ["handshake_only"] = true,
                        ["results"] = results,
                        ["note"] = "handshake_only=true：只验证进程可启动 + initialize 握手，不代表 Excel 已预热、tools/list 可用或客户端 UI 已加载新配置。"
                                 + "改动配置后必须重启对应客户端再测；若 unreachable，先看 results[].reason（not_configured/spawn_failed/timeout/no_response/server_error/protocol_mismatch）。"
                    }
                };
            }

            private static ToolResult Fail(string code, string message) =>
                new ToolResult
            {
                Ok = false,
                ErrorCode = code,
                Data = new Dictionary<string, object> { ["message"] = message }
            };

            private static string GetString(Dictionary<string, object> args, string key)
            {
                if (args == null || !args.TryGetValue(key, out var value) || value == null) { return null; }
                return Convert.ToString(value);
            }

            private static bool GetBool(Dictionary<string, object> args, string key)
            {
                if (args == null || !args.TryGetValue(key, out var value) || value == null) { return false; }
                if (value is bool b) { return b; }
                return bool.TryParse(Convert.ToString(value), out var parsed) && parsed;
            }

            private static int GetInt(Dictionary<string, object> args, string key, int fallback)
            {
                if (args == null || !args.TryGetValue(key, out var value) || value == null) { return fallback; }
                return int.TryParse(Convert.ToString(value), out var parsed) ? parsed : fallback;
            }

            private static string[] GetStringArray(Dictionary<string, object> args, string key)
            {
                if (args == null || !args.TryGetValue(key, out var value) || value == null) { return null; }
                if (value is System.Collections.IEnumerable enumerable && !(value is string))
                {
                    return enumerable.Cast<object>().Select(Convert.ToString).ToArray();
                }
                return null;
            }
        }

        // ---------------- schema ----------------

        private static Dictionary<string, object> DetectSchema()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>()
            };
        }

        private static Dictionary<string, object> ConnectSchema()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["client"] = new Dictionary<string, object>
                    {
                        ["type"] = "string",
                        ["enum"] = AiClientDiscovery.ClientIds.Concat(new[] { "all" }).ToArray(),
                        ["description"] = "客户端 id（必须用 detect_ai_clients 返回的 id；大小写不敏感，不接受显示名/别名）。传 all 表示所有已安装客户端。"
                    },
                    ["apply"] = Prop("boolean", "false=只返回预览（status=preview）；true=写盘（先备份为 *.jyyj-bak，只在首次创建）。默认 false。"),
                    ["allow_write"] = Prop("boolean", "默认 false。仅当用户明确要求该客户端能修改 Excel 时才传 true——这会让该客户端之后所有会话长期具备写权限；撤销需 connect(allow_write=false) 并重启客户端。"),
                    ["command"] = Prop("string", "可选：显式指定 ExcelAiMcpServer.exe 路径（默认自动探测）。"),
                    ["args"] = new Dictionary<string, object>
                    {
                        ["type"] = "array",
                        ["items"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["description"] = "可选：附加启动参数（如 --require-confirmation）。"
                    }
                },
                ["required"] = new[] { "client" }
            };
        }

        private static Dictionary<string, object> TestSchema()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["client"] = Prop("string", "客户端 id；省略则测试所有已登记 jyyj-excel 的客户端。"),
                    ["timeout_ms"] = Prop("integer", "单次握手超时毫秒，默认 8000，范围 1000~60000。")
                }
            };
        }

        private static Dictionary<string, object> Prop(string type, string description) =>
            new Dictionary<string, object> { ["type"] = type, ["description"] = description };
    }
}