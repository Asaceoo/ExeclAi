using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.McpServer
{
    /// <summary>
    /// JYYJ助手 MCP Server（M-A 形态，docs/需求-MCP接入.md P1 落地）：
    /// 独立 out-of-proc Excel COM 宿主 + ToolRegistry 全工具面，stdio JSON-RPC 2.0 暴露 MCP。
    /// Claude Desktop / 任何 MCP 客户端直连，无需安装 Excel 加载项。
    ///
    /// 红线（需求-MCP接入.md §5，全部保留）：
    /// - 不绕过任何安全设施：保护态守卫/确认门随 ToolDescriptor.Execute 原样生效；
    ///   【已知缺口】MCP 模式未传 OperationRepository → 覆盖审计编排（AgentRunner 层）与
    ///   撤销链路（RecordStructuralUndo no-op）不生效，写入操作无审计留痕与回滚（N3，待议）；
    /// - 默认只读模式（--allow-write 才开放写工具）；用户文件只读红线：自建内存工作簿，全程不碰用户文件；
    /// - 会话结束强杀本进程新建的 EXCEL.EXE（SpawnedPids 红线：绝不动用户实例）；
    /// - Dangerous 工具二次确认：MCP 侧返回 requires_confirmation 载荷，客户端回传 confirmed=true 才执行。
    /// </summary>
    internal static class Program
    {
        private static readonly List<int> SpawnedPids = new List<int>();
        internal static ToolRegistry _registry;
        private static ExcelToolHost _host;
        private static OperationRepository _repo;
        // P0-3：确认码强制门——工具名→一次性确认码（人工从确认文件读码后回传）
        private static readonly Dictionary<string, string> PendingConfirmations = new Dictionary<string, string>(); // P0-1：审计/回滚落库（M21 前为 null 的 N3 缺口已收口）
        private static dynamic _app;
        internal static bool _readOnlyMode = true;
        private static bool _allowWrite;
        private static string _memoryRoot; // M26：workbuddy 记忆根（可选覆盖）
        private static readonly DateTime _serverStartUtc = DateTime.UtcNow;
        private static StreamWriter _log;
        private static int _spawnedPid;
        // M26 懒初始化同步原语：后台预热完成后置位；初始化异常记录在 _initError
        private static readonly System.Threading.ManualResetEventSlim _readyEvent = new System.Threading.ManualResetEventSlim(false);
        private static string _initError;

        private static int Main(string[] args)
        {
            _allowWrite = args.Contains("--allow-write");
            var logPath = (string)null;
            for (var i = 0; i < args.Length - 1; i++)
            {
                if (args[i] == "--log") { logPath = args[i + 1]; }
                if (args[i] == "--memory-root") { _memoryRoot = args[i + 1]; }
            }

            // 日志绝不上 stdout——stdout 是 MCP 协议通道，污染即协议崩
            if (logPath != null)
            {
                try { _log = new StreamWriter(logPath, false, new UTF8Encoding(false)) { AutoFlush = true }; } catch { }
            }

            // M26 懒初始化：后台线程预热 Excel（冷启动 15-25s），stdio 循环立即启动响应 MCP 握手——
            // 修复「宿主等待握手超时 → host-tools-pending 连不上」的根因。
            var initThread = new System.Threading.Thread(() =>
            {
                try { InitializeExcel(); }
                catch (Exception ex)
                {
                    _initError = ex.Message;
                    Log("启动失败: " + ex);
                    _readyEvent.Set(); // 释放等待者，由 EnsureReady 抛出错误
                }
            });
            initThread.IsBackground = true;
            initThread.Start();
            Log("MCP server：stdio 循环启动，Excel 后台预热中…");

            // R1 修复：stdio 主循环任何异常（客户端断连 IOException/畸形输入）都不得跳过 Cleanup——
            // 泄漏 Visible=false 的 EXCEL.EXE 是红线级故障。
            try
            {
                RunStdioLoop();
            }
            finally
            {
                // FIX（impl #4 竞态）：先等后台预热结束再清理——否则客户端连上即断时 _app 尚 null、
                // SpawnedPids 为空，Cleanup 不 Quit 不 Kill，预热线程随后建出的 EXCEL 成无主孤儿（红线）。
                try { _readyEvent.Wait(TimeSpan.FromSeconds(60)); } catch { }
                Log("stdio 结束，清理 Excel 进程");
                try { Cleanup(); } catch (Exception cleanupException) { Log("Cleanup 异常: " + cleanupException.Message); }
            }
            return 0;
        }

        /// <summary>M26 懒初始化：Excel COM 冷启动耗时 15-25 秒，宿主 MCP 握手窗口远小于此——
        /// 改为后台线程预热；stdio 循环立即响应握手，tools 调用经 EnsureReady 等待就绪。</summary>
        private static void InitializeExcel()
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                _initError = "未找到 Excel.Application COM 组件（本机未安装 Excel？）";
                _readyEvent.Set(); // FIX（审查 F3）：早退必须释放等待者，否则每次 tools/* 白等 90 秒
                return;
            }

            _app = Activator.CreateInstance(excelType);
            // R2 修复：从自身 COM 对象取确定 PID 记账（进程名差集法有快照-激活窗口竞态，
            // 且 WPS 兼容机上 ProgID 解析到 et.exe 时 "EXCEL" 进程名永不匹配——双失效）。
            try
            {
                var hwnd = (int)_app.Hwnd;
                var pid = new uint[1];
                Win32Native.GetWindowThreadProcessId(hwnd, pid);
                if (pid[0] != 0)
                {
                    SpawnedPids.Add((int)pid[0]);
                    _spawnedPid = (int)pid[0];
                    Log("spawn PID=" + pid[0] + "（来自自身实例 Hwnd 归因）");
                }
            }
            catch (Exception pidException)
            {
                Log("PID 归因失败（会话结束仅靠 Quit 兜底）: " + pidException.Message);
            }

            _app.Visible = false;
            try { _app.DisplayAlerts = false; } catch { }
            try { _app.AutomationSecurity = 1; } catch { }

            // 新建未存盘内存工作簿（真机 harness 同款）：全程不 SaveAs、不打开用户文件
            _app.Workbooks.Add();

            // P0-1：接入审计/回滚——独立 operations-mcp.db（与加载项 operations.db 分离，互不污染）
            var repoDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant");
            _repo = new OperationRepository(Path.Combine(repoDir, "operations-mcp.db"));
            _host = new ExcelToolHost(_app, _repo, null, "mcp-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            _registry = _host.CreateRegistry();
            // M26：外部 AI 智能体桥接（workbuddy 记忆读写）——DSH/workbuddy 等经 MCP 直连
            AgentBridgeToolRegistration.Register(_registry, _memoryRoot);
            _readOnlyMode = !_allowWrite;

            if (_readOnlyMode)
            {
                foreach (var tool in _registry.Tools.Where(t => t.RiskLevel != RiskLevel.ReadOnly).ToList())
                {
                    _registry.Remove(tool.Name);
                }
            }

            Log("Excel 预热完成：工具数=" + _registry.Tools.Count + "，模式=" + (_readOnlyMode ? "只读" : "读写") + "，Excel PID=" + _spawnedPid);
            _readyEvent.Set();
        }

        /// <summary>强杀本进程新建的 EXCEL.EXE；用户既有实例绝不动。</summary>
        private static void Cleanup()
        {
            // 先优雅退出（DisplayAlerts=false 下无对话框风险），再对拒绝退出的实例兜底强杀。
            if (_app != null && Marshal.IsComObject(_app))
            {
                try { _app.Quit(); } catch { }
            }

            // R2-1 修复：按 PID 直取强杀（原 GetProcessesByName("EXCEL") 在 WPS 宿主下永不命中 = 死代码）。
            // 启动时间校验：PID 记账来自自身 COM 对象 Hwnd 归因（不可能落到用户实例），但加一道
            // "进程启动时间晚于 server 启动"防线，把未来宿主语义漂移的假说性风险一并关掉。
            foreach (var pid in SpawnedPids)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(pid);
                    if (p.StartTime >= _serverStartUtc)
                    {
                        p.Kill();
                        Log("已强杀 spawn 进程 PID=" + pid + " name=" + p.ProcessName);
                    }
                    else
                    {
                        Log("跳过 PID=" + pid + "（启动时间早于 server，疑似宿主语义漂移，防误杀）");
                    }
                    p.Dispose();
                }
                catch (Exception killException)
                {
                    Log("强杀 PID=" + pid + " 未执行: " + killException.GetType().Name);
                }
            }

            if (_app != null && Marshal.IsComObject(_app))
            {
                try { Marshal.FinalReleaseComObject(_app); } catch { }
                _app = null;
            }

            _log?.Dispose();
        }

        private static void Log(string message)
        {
            var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message;
            System.Diagnostics.Debug.WriteLine(line);
            if (_log != null) { try { _log.WriteLine(line); } catch { } }
        }

        // ================= MCP stdio 协议 =================

        private static void RunStdioLoop()
        {
            var stdout = Console.OpenStandardOutput();
            using (var reader = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false)))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) { continue; }

                    // 兼容 LSP 帧：Content-Length: N + 空行 + N 字符 JSON
                    if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                    {
                        // R1/R5 修复：帧头畸形与超大帧不再崩进程——回协议错误应答并继续服务
                        if (!int.TryParse(line.Substring(15).Trim(), out var len) || len < 0 || len > 10 * 1024 * 1024)
                        {
                            Emit(stdout, new JObject
                            {
                                ["jsonrpc"] = "2.0",
                                ["id"] = null,
                                ["error"] = new JObject { ["code"] = -32600, ["message"] = "Invalid LSP frame length" }
                            });
                            continue;
                        }

                        string skip;
                        do { skip = reader.ReadLine(); } while (skip != null && skip.Length > 0);
                        var buf = new char[len];
                        var read = 0;
                        while (read < len)
                        {
                            var n = reader.Read(buf, read, len - read);
                            if (n <= 0) { break; }
                            read += n;
                        }

                        HandleMessage(new string(buf, 0, read), stdout);
                        continue;
                    }

                    // 标准 MCP stdio：换行分隔 JSON（Claude Desktop / MCP Inspector 默认）
                    HandleMessage(line, stdout);
                }
            }
        }

        private static void HandleMessage(string json, Stream stdout)
        {
            JToken parsed;
            // Y2 修复：禁用 ISO 日期串自动转 DateTime——"2024-01-15" 必须保持字符串原样进工具层
            //（写 ISO 日期文本先 NumberFormat="@" 的编码红线依赖拿到的是文本）。
            // Newtonsoft 13.0.3 的 JsonLoadSettings 无 DateParseHandling，走 JsonTextReader+JsonSerializer 层禁用。
            try
            {
                using var stringReader = new System.IO.StringReader(json);
                // R2-2 修复：MaxDepth 防 10MB 合法深嵌套 JSON 触发递归建树 StackOverflowException——
                // .NET Framework 该异常不可捕获、finally 不保证执行，Cleanup 会被打穿。超限走 -32700 路径。
                using var jsonReader = new JsonTextReader(stringReader) { DateParseHandling = DateParseHandling.None, MaxDepth = 64 };
                parsed = JToken.Load(jsonReader);
            }
            catch (Exception ex)
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = null,
                    ["error"] = new JObject { ["code"] = -32700, ["message"] = "Parse error: " + ex.Message }
                });
                return;
            }

            // JSON-RPC 2.0 batch：数组请求逐条处理（MCP 客户端少用，但规范允许）
            if (parsed is JArray batch)
            {
                // Y2-7 修复：非对象元素回 per-element -32600（原静默丢弃会让客户端悬挂等超时）
                foreach (var item in batch)
                {
                    if (item is JObject obj)
                    {
                        HandleMessage(obj.ToString(Formatting.None), stdout);
                    }
                    else
                    {
                        Emit(stdout, new JObject
                        {
                            ["jsonrpc"] = "2.0",
                            ["id"] = null,
                            ["error"] = new JObject { ["code"] = -32600, ["message"] = "Invalid Request: batch 元素必须是 JSON 对象" }
                        });
                    }
                }
                return;
            }

            if (!(parsed is JObject msg))
            {
                // 标量/字符串等非法请求体：按协议错误应答而非崩进程
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = null,
                    ["error"] = new JObject { ["code"] = -32600, ["message"] = "Invalid Request: 请求体必须是 JSON 对象" }
                });
                return;
            }

            var id = msg["id"];
            var method = msg["method"] != null ? (string)msg["method"] : null;

            // 通知（无 id 或显式 notifications/*）：initialized 等不回复
            if (id == null || id.Type == JTokenType.Null || (method != null && method.StartsWith("notifications/", StringComparison.Ordinal)))
            {
                Log("通知: " + method);
                return;
            }

            try
            {
                switch (method)
                {
                    case "initialize":
                        HandleInitialize(id, stdout);
                        break;
                    case "ping":
                        Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject() });
                        break;
                    case "tools/list":
                        HandleToolsList(id, stdout);
                        break;
                    case "tools/call":
                        HandleToolsCall(id, msg, stdout);
                        break;
                    case "resources/list":
                        Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["resources"] = new JArray() } });
                        break;
                    case "prompts/list":
                        Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["prompts"] = new JArray() } });
                        break;
                    default:
                        Emit(stdout, new JObject
                        {
                            ["jsonrpc"] = "2.0",
                            ["id"] = id,
                            ["error"] = new JObject { ["code"] = -32601, ["message"] = "Method not found: " + method }
                        });
                        break;
                }
            }
            catch (Exception ex)
            {
                Log("处理 " + method + " 异常: " + ex);
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["error"] = new JObject { ["code"] = -32603, ["message"] = "Internal error: " + ex.Message }
                });
            }
        }

        private static void HandleInitialize(JToken id, Stream stdout)
        {
            Emit(stdout, new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["result"] = new JObject
                {
                    ["protocolVersion"] = "2024-11-05", // 规范允许 server 声明支持版本；Claude Desktop 用此版本
                    ["capabilities"] = new JObject
                    {
                        ["tools"] = new JObject { ["listChanged"] = false }
                    },
                    ["serverInfo"] = new JObject
                    {
                        ["name"] = "jyyj-excel",
                        ["version"] = typeof(Program).Assembly.GetName().Version.ToString()
                    }
                }
            });
        }

        /// <summary>M3ToolSchemas 字典 → 标准 JSON Schema：剥掉内部路由键 name，其余透传。</summary>
        internal static JObject ToJsonSchema(Dictionary<string, object> parameters)
        {
            if (parameters == null)
            {
                return new JObject { ["type"] = "object", ["properties"] = new JObject() };
            }

            return JObject.FromObject(CleanSchemaNode(parameters));
        }

        internal static object CleanSchemaNode(object node)
        {
            return CleanSchemaNode(node, isPropertyDefinition: false);
        }

        /// <summary>
        /// R3 修复：name 剥除按层级区分——M3ToolSchemas 的 Property/Array 帮助方法在「属性定义对象」里
        /// 放 name 作内部路由键，需剥除；但 RenameSheet/AddSheet 等工具有真业务参数也叫 name，
        /// 顶层/任意业务参数的 name 必须保留。判据：字典含 type 键（schema 节点特征）且在 properties
        /// 值位置出现（isPropertyDefinition=true）才剥。
        /// </summary>
        internal static object CleanSchemaNode(object node, bool isPropertyDefinition)
        {
            if (node is Dictionary<string, object> dict)
            {
                var result = new Dictionary<string, object>();
                var isSchemaContainer = dict.ContainsKey("properties");
                foreach (var kv in dict)
                {
                    if (isPropertyDefinition && kv.Key == "name") { continue; } // 仅属性定义层的路由键剥除
                    var childIsProperty = isSchemaContainer && kv.Key == "properties";
                    result[kv.Key] = childIsProperty
                        ? CleanPropertiesValue(kv.Value)
                        : CleanSchemaNode(kv.Value, isPropertyDefinition: false);
                }

                return result;
            }

            if (node is IEnumerable<object> list && !(node is string))
            {
                return list.Select(item => CleanSchemaNode(item, isPropertyDefinition: false)).ToArray();
            }

            return node;
        }

        /// <summary>properties 字典：值位置是属性定义，其内部 name 是路由键（剥）；键名本身是业务参数名（保留）。</summary>
        private static object CleanPropertiesValue(object propertiesValue)
        {
            if (propertiesValue is Dictionary<string, object> props)
            {
                var result = new Dictionary<string, object>();
                foreach (var kv in props)
                {
                    result[kv.Key] = CleanSchemaNode(kv.Value, isPropertyDefinition: true);
                }

                return result;
            }

            return CleanSchemaNode(propertiesValue, isPropertyDefinition: false);
        }

        /// <summary>M26：等待 Excel 后台预热完成；失败抛人话异常。</summary>
        private static void EnsureReady(int timeoutMs = 90000)
        {
            if (!_readyEvent.Wait(timeoutMs))
            {
                throw new InvalidOperationException("Excel 初始化超时（" + (timeoutMs / 1000) + " 秒）。本机 Excel COM 可能被占用或响应缓慢，请稍后重试。");
            }
            if (!string.IsNullOrEmpty(_initError))
            {
                throw new InvalidOperationException("Excel 初始化失败：" + _initError);
            }
        }

        private static void HandleToolsList(JToken id, Stream stdout)
        {
            try { EnsureReady(); }
            catch (Exception readyEx)
            {
                Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JObject { ["code"] = -32000, ["message"] = readyEx.Message } });
                return;
            }

            var tools = new JArray();
            foreach (var tool in _registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
            {
                tools.Add(new JObject
                {
                    ["name"] = tool.Name,
                    ["description"] = RiskHint(tool) + tool.Description,
                    ["inputSchema"] = ToJsonSchema(tool.Parameters)
                });
            }

            Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["tools"] = tools } });
            Log("tools/list 返回 " + tools.Count + " 个工具");
        }

        internal static string RiskHint(ToolDescriptor tool)
        {
            switch (tool.RiskLevel)
            {
                case RiskLevel.Dangerous: return "[危险级·需确认] ";
                case RiskLevel.StructuralChange: return "[结构变更] ";
                case RiskLevel.NormalWrite: return "[写入] ";
                default: return "";
            }
        }

        private static void HandleToolsCall(JToken id, JObject msg, Stream stdout)
        {
            try { EnsureReady(); }
            catch (Exception readyEx)
            {
                Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JObject { ["code"] = -32000, ["message"] = readyEx.Message } });
                return;
            }

            var name = (string)msg["params"]?["name"];
            var rawArgs = msg["params"]?["arguments"] as JObject;

            if (string.IsNullOrWhiteSpace(name) || !_registry.TryGet(name, out var tool))
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = BuildUnknownToolText(name) } },
                        ["isError"] = true
                    }
                });
                return;
            }

            var args = new Dictionary<string, object>();
            if (rawArgs != null)
            {
                foreach (var kv in rawArgs)
                {
                    args[kv.Key] = kv.Value.Type == JTokenType.Null ? null : kv.Value.ToObject<object>();
                }
            }

            // Dangerous 级二次确认（需求-MCP接入.md §3）：首次调用返回确认载荷，confirmed=true 才放行
            // P0-3 确认门强制力：Dangerous 首调生成一次性确认码文件（人工读码），回传 confirmed=<code> 才放行
            if (tool.RiskLevel == RiskLevel.Dangerous && !IsConfirmedWithCode(name, args))
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = new JArray
                        {
                            new JObject
                            {
                                ["type"] = "text",
                                ["text"] = BuildConfirmationText(name, args)
                            }
                        },
                        ["isError"] = false,
                        ["_meta"] = new JObject { ["requires_confirmation"] = true }
                    }
                });
                Log("Dangerous 工具待确认: " + name);
                return;
            }

            // 协议层参数不透传工具层：confirmed 已完成确认门使命，留在 args 里会污染工具参数校验
            args.Remove("confirmed");

            try
            {
                var result = tool.Execute(args);
                var text = BuildResultText(result, tool);
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = text } },
                        ["isError"] = !result.Ok
                    }
                });
                Log("tools/call " + name + " ok=" + result.Ok);
            }
            catch (Exception ex)
            {
                Log("tools/call " + name + " 异常: " + ex);
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = "工具执行异常: " + ex.Message } },
                        ["isError"] = true
                    }
                });
            }
        }

        /// <summary>未知工具人话指引：只读模式下明确说明模式限制并列出可用只读工具族，防止 AI 反复重试写工具。</summary>
        internal static string BuildUnknownToolText(string name)
        {
            var msg = "未知工具：" + name + "。";
            if (_readOnlyMode)
            {
                // Y2-5 修复：纯名字清单有语义地雷（verify_*20个/select 会激活/screenshot 落盘PNG/eval_formula 执行求值），
                // 每项附描述（截断 40 字符），并前置 family 标注。
                var readOnlyTools = _registry.Tools
                    .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var listing = string.Join("; ", readOnlyTools.Select(t =>
                {
                    var desc = (t.Description ?? "").Replace("[只读] ", "").Replace("[危险级·需确认] ", "");
                    if (desc.Length > 40) { desc = desc.Substring(0, 40) + "…"; }
                    return t.Name + "（" + desc + "）";
                }));
                msg += "当前 server 以只读模式启动（未带 --allow-write），写入/结构变更/危险级工具已全部下线，" +
                       "请勿重试写操作；可用只读工具 " + readOnlyTools.Count + " 个（名字后括号内为功能说明）：" +
                       listing +
                       "。注意：verify_* 是写入后自检、screenshot_dashboard 会生成 PNG 文件、eval_formula 会执行公式。如需写操作，请重启 server 并加 --allow-write 参数。";
            }
            return msg;
        }

        /// <summary>P0-3：确认码校验——比对 PendingConfirmations 中的一次性码，命中即消费（防重放）。
        /// 兼容旧 confirmed=true 语义仅限 --legacy-confirm 启动参数（默认关闭）。</summary>
        private static bool IsConfirmedWithCode(string toolName, Dictionary<string, object> args)
        {
            if (!args.TryGetValue("confirmed", out var v) || v == null) { return false; }
            var supplied = Convert.ToString(v);
            if (PendingConfirmations.TryGetValue(toolName, out var expected))
            {
                if (string.Equals(supplied, expected, StringComparison.OrdinalIgnoreCase))
                {
                    PendingConfirmations.Remove(toolName); // 一次性消费
                    return true;
                }
                return false; // 有待确认码但不匹配
            }
            // 无待确认记录（如进程重启后直接带码）——fail-closed
            return false;
        }

        /// <summary>ToolResult → MCP text 载荷：snake_case 数据 + 警告 + 受影响区域 + 错误码全量透传。</summary>
        internal static string BuildResultText(ToolResult result) => BuildResultText(result, null);

        internal static string BuildResultText(ToolResult result, ToolDescriptor tool)
        {
            var payload = new JObject();
            payload["ok"] = result.Ok;
            if (result.Data != null)
            {
                try { payload["data"] = result.Data is string s ? JToken.Parse(s) : JToken.FromObject(result.Data); }
                catch { payload["data"] = Convert.ToString(result.Data); }
            }

            if (!string.IsNullOrEmpty(result.Warning)) { payload["warning"] = result.Warning; }
            if (result.AffectedRanges != null && result.AffectedRanges.Count > 0)
            {
                payload["affected_ranges"] = JArray.FromObject(result.AffectedRanges);
            }

            if (!string.IsNullOrEmpty(result.ErrorCode)) { payload["error_code"] = result.ErrorCode; }
            if (result.Truncated) { payload["truncated"] = true; }

            var text = payload.ToString(Formatting.Indented);
            // R4 修复：传输层限幅不得产出非法 JSON（原 Substring 拦腰砍断 JSON 结构 + 可能劈开 UTF-16 代理对）。
            // 结构化降级：重建精简载荷，data 换前缀预览，永远合法可解析。
            const int Limit = 512 * 1024;
            if (text.Length > Limit)
            {
                var preview = BuildResultText(payload, tool);
                return preview;
            }
            return text;
        }

        /// <summary>R4：超限时重建合法 JSON 载荷——data 降为前缀预览 + truncated 标记。hint 按工具 schema 动态生成（Y2-2：
        /// 全局 hint 对 verify_chart 等无 max_cells/address 的工具是无 效指引；图片类工具给专属建议）。</summary>
        private static string BuildResultText(JObject payload, ToolDescriptor tool)
        {
            const int PreviewLimit = 200 * 1024;
            var slim = new JObject
            {
                ["ok"] = payload["ok"],
                ["transport_truncated"] = true,
                ["hint"] = BuildTruncationHint(tool)
            };
            if (payload["error_code"] != null) { slim["error_code"] = payload["error_code"]; }
            // Y2-1 修复：warning 与 data 同规格限幅（实测 600KB warning 原样透传仍超 512KB）
            if (payload["warning"] != null)
            {
                var warningText = payload["warning"].ToString(Formatting.None);
                if (warningText.Length > PreviewLimit)
                {
                    var wcut = warningText.Substring(0, PreviewLimit);
                    if (char.IsHighSurrogate(wcut[wcut.Length - 1])) { wcut = wcut.Substring(0, wcut.Length - 1); }
                    slim["warning"] = wcut + "...[已截断]";
                }
                else
                {
                    slim["warning"] = payload["warning"];
                }
            }

            var data = payload["data"];
            if (data != null)
            {
                var dataText = data.ToString(Formatting.None);
                if (dataText.Length > PreviewLimit)
                {
                    // 代理对回退：截断点落在高代理上时回退一个字符
                    var cut = dataText.Substring(0, PreviewLimit);
                    if (char.IsHighSurrogate(cut[cut.Length - 1])) { cut = cut.Substring(0, cut.Length - 1); }
                    slim["data_preview"] = cut;
                    slim["data_preview_note"] = "完整结果已省略，data 为前 " + PreviewLimit + " 字符预览。";
                }
                else
                {
                    slim["data"] = data;
                }
            }

            return slim.ToString(Formatting.Indented);
        }

        private static void Emit(Stream stdout, JObject message)
        {
            var bytes = Encoding.UTF8.GetBytes(message.ToString(Formatting.None) + "\n");
            stdout.Write(bytes, 0, bytes.Length);
            stdout.Flush();
        }
        /// <summary>Dangerous 拦截文案（Y2-6）：参数回显截断 500 字符防上下文白烧/敏感值复述；声明 MCP 模式不可撤销。</summary>
        /// <summary>P0-3 确认门强制力：生成一次性确认码文件（人工可读），返回拦截文案。
        /// AI 无法自答——码只存在于文件中，人工读码后由客户端回传 confirmed=<code>。</summary>
        private static string BuildConfirmationText(string name, Dictionary<string, object> args)
        {
            var code = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
            PendingConfirmations[name] = code;
            var argsText = JsonConvert.SerializeObject(args);
            if (argsText.Length > 500) { argsText = argsText.Substring(0, 500) + "...[截断]"; }

            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "pending-confirmations");
            Directory.CreateDirectory(dir);
            var file = Path.Combine(dir, name + "-" + code + ".txt");
            File.WriteAllText(file,
                "危险操作确认\n" +
                "工具: " + name + "\n" +
                "确认码: " + code + "\n" +
                "时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\n" +
                "参数: " + argsText + "\n" +
                "\n确认方式: 人工核对影响范围后，在 arguments 中携带 \"confirmed\": \"" + code + "\" 重新调用。\n" +
                "该确认码一次性，使用后作废。");

            return "该工具为 Dangerous 级（" + name + "），已拦截等待人工确认。" +
                   "确认文件已生成: " + file + "\n" +
                   "请人工打开该文件核对操作影响范围与确认码，确认后再携带该码重新调用（码只在文件中，不随本响应回传）。" +
                   "注意：确认码一次性；写操作已接入审计库，执行后可用 undo_excel_operation 回滚。" +
                   "当前参数：" + argsText;
        }
        /// <summary>按工具 schema 生成截断后的自愈指引（Y2-2）：全局 hint 对无 max_cells/address 的工具无效。</summary>
        private static string BuildTruncationHint(ToolDescriptor tool)
        {
            var schemaText = tool?.Parameters != null ? JsonConvert.SerializeObject(tool.Parameters) : string.Empty;
            var parts = new List<string> { "结果超512KB已被传输层截断。" };
            if (schemaText.Contains("\"max_cells\""))
            {
                parts.Add("本工具支持 max_cells 限制返回行数（1-200000），请调小后重试。");
            }
            else if (schemaText.Contains("\"address\""))
            {
                parts.Add("请缩小 address 范围（分块读取）后重试。");
            }
            else if (tool != null && tool.Name != null && tool.Name.Contains("screenshot"))
            {
                parts.Add("截图类工具请缩小截图范围（address 参数）后重试。");
            }
            else
            {
                parts.Add("请缩小请求范围后重试；如为图片/渲染类输出，可改用返回文件路径的模式避免内联大载荷。");
            }
            return string.Join("", parts);
        }

    }
    /// <summary>Win32 互操作（MCP server 自用：PID 归因）。</summary>
    internal static class Win32Native
    {
        [DllImport("user32")]
        public static extern uint GetWindowThreadProcessId(int hWnd, uint[] processId);
    }

}