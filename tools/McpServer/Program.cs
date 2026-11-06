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
using ExcelAiAssistant.Core.Workflow;
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
        internal static readonly System.Collections.Concurrent.ConcurrentDictionary<string, PendingConfirmation> PendingConfirmations = new System.Collections.Concurrent.ConcurrentDictionary<string, PendingConfirmation>(); // M27R3：并发消费/覆盖原子化 // P0-1 审计接入；M27R2：带 TTL/失败计数

        /// <summary>M27R2 审查：确认码携带过期时间与失败计数——防长驻进程内无限穷举。</summary>
        internal sealed class PendingConfirmation
        {
            public string Code;
            public DateTime ExpiryUtc;
            public int Failures;
            public string ArgsHash; // M28 审查（第三轮）：确认码绑定参数，防"确认 A 执行 B"
        }

        private static dynamic _app;
        internal static bool _readOnlyMode = true;
        // M36-2：工具面分层（core 省上下文 / all 全量，默认 all 保持旧行为）
        internal static readonly string ToolProfile = ResolveToolProfile();
        internal static readonly HashSet<string> ProfileHiddenTools = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> CoreProfileTools = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "get_current_excel_context",
            "get_excel_sheet_list",
            "get_excel_range_values_and_formulas",
            "preview_excel_range",
            "preview_excel_sheet",
            "get_current_selection_values_and_formulas",
            "find_excel_cells",
            "export_selection_to_markdown",
            "diagnose_formula",
            "write_excel_range",
            "write_excel_cell",
            "write_excel_cells_batch",
            "append_excel_rows",
            "clear_excel_range",
            "fill_formula",
            "sort_range",
            "manage_excel_filter",
            "format_excel_range",
            "format_excel_conditional",
            "freeze_panes",
            "merge_or_split_cells",
            "add_excel_worksheet",
            "rename_excel_worksheet",
            "delete_excel_worksheet",
            "create_excel_table",
            "insert_blank_rows_every_n",
            "delete_blank_rows_range",
            "clean_excel_data",
            "dedup_range",
            "analyze_range_statistics",
            "analyze_trend_forecast",
            "analyze_column_profile",
            "filter_excel_data",
            "compare_two_sheets",
            "generate_analysis_report",
            "create_excel_chart",
            "recommend_excel_chart",
            "insert_recommended_chart",
            "create_pivot_table",
            "verify_write_result",
            "verify_cell_format",
            "verify_chart",
            "verify_formula",
            "verify_file",
            "run_self_check",
            "apply_correction",
            "render_range_image",
            "take_screenshot",
            "save_workbook",
            "undo_excel_operation",
            "export_sheets_to_pdf",
            "detect_ai_clients",
            "connect_ai_client",
            "test_ai_client_connection"
        };

        private static string ResolveToolProfile()
        {
            var raw = (Environment.GetEnvironmentVariable("JYYJ_TOOL_PROFILE") ?? "all").Trim().ToLowerInvariant();
            if (raw == "all" || raw == "full") { return "all"; }
            if (raw == "core") { return "core"; }
            throw new InvalidOperationException("JYYJ_TOOL_PROFILE 只能是 core 或 all，当前值：" + raw);
        }

        private static bool _allowWrite;
        private static bool _ownsInstance; // M28 审查：仅「自建且启动时间晚于 server」的实例允许 Visible/Add/Quit
        // M28：确认门硬门选项——LegacyConfirm=兼容旧 confirmed=true（仅显式 --legacy-confirm 开启）；
        // ConfirmMode=file（默认，人工读确认文件）| stderr（带外：码只进 server stderr，不落盘、不进响应）
        internal static bool LegacyConfirm;
        internal static string ConfirmMode = "stderr"; // M28 审查（第三轮）：默认带外硬门（file 需显式开启）
        /// <summary>M29 探测模式：只允许 initialize/ping，避免 tools/list 在未预热时白等 120s（审查 B2）。</summary>
        internal static bool _probeOnly;

        // 审查 R20（P0-2）：串行工作队列 + 出站锁。
        // 原先 tools/call 直接在 stdin 读循环里同步跑，一个 60s 的工具会让 ping/tools/list/取消
        // 全部排队，客户端 60s 超时后判死并重连。现在读循环只做解析/分派，执行统一交给单条
        // worker 线程（保持串行——Excel COM 不适合并发写），ping/initialize 仍在读循环即时应答。
        private static readonly System.Collections.Concurrent.BlockingCollection<Action> _workQueue
            = new System.Collections.Concurrent.BlockingCollection<Action>();
        private static readonly object _stdoutLock = new object();
        private static System.Threading.Thread _workerThread;

        private static void StartWorker()
        {
            if (_workerThread != null) { return; }
            _workerThread = new System.Threading.Thread(() =>
            {
                foreach (var job in _workQueue.GetConsumingEnumerable())
                {
                    try { job(); }
                    catch (Exception ex) { Log("worker 任务异常: " + ex); }
                }
            });
            _workerThread.IsBackground = true;
            _workerThread.Name = "jyyj-mcp-worker";
            _workerThread.Start();
        }
        // M29：AI 自主模式——默认开启（MCP 是 AI 驱动的接口面，自有内存工作簿，不碰用户文件），
        // 破坏性工具不再要求人工确认码；补偿控制：全量审计日志 + 无障碍撤销 + 明确启动警告。
        // 可用 --require-confirmation / --no-autonomous / 环境变量 JYYJ_AI_AUTONOMOUS=0 恢复确认门。
        internal static bool AutonomousMode = true;

        /// <summary>
        /// M29 审查（S2）：自主模式白名单——只放行「可 undo 的结构/数据操作」类 Dangerous 工具；
        /// 代码执行类（VBA/Python/CLI/EXE）、路径敏类（外部簿打开/外部链接/包安装）保留确认门——
        /// 这些工具的 Dangerous 定级理由（任意文件读写/绕过操作记录）在自主模式下依然成立，
        /// 放行它们等于默认打开任意本地文件读写通道。
        /// </summary>
        internal static readonly System.Collections.Generic.HashSet<string> AutonomousDeniedTools =
            new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal)
            {
                "run_vba_snippet",
                "run_python_code",
                "run_python_excel",
                "run_cli_command",
                "run_exe_command",
                "install_python_packages",
                "open_excel_external_book",
                "replace_external_links"
            };

        /// <summary>自主模式下该 Dangerous 工具是否仍需人工确认（代码执行/路径敏类 = 需要）。</summary>
        internal static bool RequiresConfirmationInAutonomous(string tool)
        {
            return AutonomousDeniedTools.Contains(tool);
        }
        internal static string JobError; // M28 审查（第三轮）：Job 失败阶段 + Win32Error（供初始化报错可执行）
        private static string _memoryRoot; // M26：workbuddy 记忆根（可选覆盖）
        private static readonly DateTime _serverStartUtc = DateTime.UtcNow;
        private static StreamWriter _log;
        private static int _spawnedPid;
        private static IntPtr _jobHandle; // M28：Excel 所属 Job（KILL_ON_JOB_CLOSE），保持句柄存活至进程退出
        private static bool _jobAttached;  // M28 审查：Job 建立成功才允许初始化（否则硬杀兜底缺失，拒绝启动）
        // M26 懒初始化同步原语：后台预热完成后置位；初始化异常记录在 _initError
        private static readonly System.Threading.ManualResetEventSlim _readyEvent = new System.Threading.ManualResetEventSlim(false);
        private static string _initError;
        // M28 审查：关停单飞 + 完成信号——stdio finally 与看门狗可能同时进入 Shutdown，
        // 双线程并发 Cleanup 会对同一 RCW 并发 Quit/Kill；先到者执行、后到者等待完成。
        private static int _shutdownStarted;
        private static readonly System.Threading.ManualResetEventSlim _shutdownDone = new System.Threading.ManualResetEventSlim(false);

        private static int Main(string[] args)
        {
            _allowWrite = args.Contains("--allow-write");
            LegacyConfirm = args.Contains("--legacy-confirm");
            // M29：AI 自主模式默认开启；--require-confirmation / --no-autonomous / JYYJ_AI_AUTONOMOUS=0 恢复确认门。
            AutonomousMode = !args.Contains("--require-confirmation") && !args.Contains("--no-autonomous");
            var autonomousEnv = Environment.GetEnvironmentVariable("JYYJ_AI_AUTONOMOUS");
            if (!string.IsNullOrWhiteSpace(autonomousEnv))
            {
                // M29 审查（S1）：环境变量值必须显式表态——"0/false"关闭，"1/true"开启（不覆盖命令行关闭），
                // 其他值一律启动失败（与 --confirm-mode 非法值口径一致，防宽松解析方向反了被静默忽略）。
                var envValue = autonomousEnv.Trim();
                if (envValue == "0" || string.Equals(envValue, "false", StringComparison.OrdinalIgnoreCase))
                {
                    AutonomousMode = false;
                }
                else if (envValue != "1" && !string.Equals(envValue, "true", StringComparison.OrdinalIgnoreCase))
                {
                    Console.Error.WriteLine("JYYJ_AI_AUTONOMOUS 仅支持 0/1/false/true，当前：" + autonomousEnv);
                    return 2;
                }
            }
            // M29：外部接口面总开关（与侧边栏面板共用 connectors.json）——关闭即拒绝启动，不创建 Excel 实例。
            // 审查 A8：闸门放在横幅/日志参数之前，避免"已关闭"时 stderr 仍打启动横幅造成误判。
            try
            {
                var connectorSettings = new ExcelAiAssistant.Core.Configuration.ConnectorSettingsStore().Load(out var settingsState);
                if (settingsState == "invalid" || settingsState == "denied")
                {
                    Console.Error.WriteLine("[WARN] connectors.json 状态=" + settingsState + "，已按默认开启处理；"
                        + "如需可靠关闭 MCP，请修复该文件或删除后重新设置。");
                }

                if (!connectorSettings.McpEnabled)
                {
                    Console.Error.WriteLine("MCP 入口已在侧边栏「连接」面板关闭（connectors.json: mcp_enabled=false）。"
                        + "请在面板中重新开启，或直接修改 %LocalAppData%\\XaiAssistant\\connectors.json 后重试。");
                    return 6;
                }
            }
            catch (Exception settingsException)
            {
                // 读取失败不阻断：按默认开启处理（与 Store.Load 容错口径一致）
                Console.Error.WriteLine("[WARN] 读取 connectors.json 失败，按默认开启处理：" + settingsException.Message);
            }

            // M28 审查（第三轮）：默认带外硬门 stderr；file 需显式开启；支持 --confirm-mode=x 形式；
            // 缺值/非法值一律启动失败（不得静默降级为软门）。
            ConfirmMode = "stderr";
            string requestedMode = null;
            for (var modeIndex = 0; modeIndex < args.Length; modeIndex++)
            {
                if (args[modeIndex] == "--confirm-mode")
                {
                    if (modeIndex + 1 >= args.Length || args[modeIndex + 1].StartsWith("--", StringComparison.Ordinal))
                    {
                        Console.Error.WriteLine("--confirm-mode 缺少值（需要 file|stderr）");
                        return 2;
                    }
                    requestedMode = args[modeIndex + 1];
                    modeIndex++;
                }
                else if (args[modeIndex].StartsWith("--confirm-mode=", StringComparison.Ordinal))
                {
                    requestedMode = args[modeIndex].Substring("--confirm-mode=".Length);
                }
            }

            if (requestedMode != null)
            {
                if (string.Equals(requestedMode, "stderr", StringComparison.OrdinalIgnoreCase)) { ConfirmMode = "stderr"; }
                else if (string.Equals(requestedMode, "file", StringComparison.OrdinalIgnoreCase)) { ConfirmMode = "file"; }
                else
                {
                    Console.Error.WriteLine("--confirm-mode 仅支持 file|stderr，当前：" + requestedMode);
                    return 2;
                }
            }
            // M28 审查：启动即提示确认门安全等级（file=软门可被可读文件的宿主自答；legacy=不安全兼容）
            if (string.Equals(ConfirmMode, "file", StringComparison.Ordinal))
            {
                Console.Error.WriteLine("[WARN] 确认门 file 模式为软门：具备文件读取能力的主机可自答确认码；带外硬门请用 --confirm-mode stderr。");
            }

            if (LegacyConfirm)
            {
                Console.Error.WriteLine("[WARN] --legacy-confirm 为不安全兼容选项：confirmed=true 可被 AI 自答，仅用于兼容旧客户端。");
            }

            if (AutonomousMode)
            {
                Console.Error.WriteLine("[AI-AUTONOMOUS] 自主模式：Dangerous 级工具不再要求人工确认码，由 AI 直接执行。");
                Console.Error.WriteLine("[AI-AUTONOMOUS] 补偿控制：全部操作写入 operations-mcp.db 审计库、可 undo_excel_operation 回滚、自建内存工作簿不碰用户文件。");
                Console.Error.WriteLine("[AI-AUTONOMOUS] 需要人工确认门请以 --require-confirmation 启动，或设环境变量 JYYJ_AI_AUTONOMOUS=0。");
            }

            var logPath = (string)null;
            for (var i = 0; i < args.Length - 1; i++)
            {
                if (args[i] == "--log") { logPath = args[i + 1]; }
                if (args[i] == "--memory-root") { _memoryRoot = args[i + 1]; }
            }

            // 日志绝不上 stdout——stdout 是 MCP 协议通道，污染即协议崩
            if (logPath != null)
            {
                // M29 审查（S7/S3'）：--log 曾以截断模式打开任意路径（可覆盖用户 .bat/.ps1）。
                // 改为：只接受**完全限定**的本地固定盘路径、拒绝 UNC/重解析点（含父目录链）、只追加不截断。
                try
                {
                    var safe = !string.IsNullOrWhiteSpace(logPath)
                        && System.IO.Path.IsPathRooted(logPath)
                        && !logPath.StartsWith("\\\\", StringComparison.Ordinal)
                        && logPath.IndexOf('\0') < 0;

                    var full = safe ? System.IO.Path.GetFullPath(logPath) : null;
                    if (safe)
                    {
                        var root = System.IO.Path.GetPathRoot(full);
                        safe = !string.IsNullOrWhiteSpace(root)
                            && new System.IO.DriveInfo(root).DriveType == System.IO.DriveType.Fixed;
                    }

                    // 父目录链上任何 reparse point 一律拒绝（防 junction 把"合法路径"重定向到敏感目录）
                    if (safe)
                    {
                        var dir = System.IO.Path.GetDirectoryName(full);
                        while (!string.IsNullOrWhiteSpace(dir))
                        {
                            if (System.IO.Directory.Exists(dir)
                                && (System.IO.File.GetAttributes(dir) & FileAttributes.ReparsePoint) != 0)
                            {
                                safe = false;
                                break;
                            }

                            var parent = System.IO.Path.GetDirectoryName(dir);
                            if (string.Equals(parent, dir, StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }

                            dir = parent;
                        }
                    }

                    if (safe && File.Exists(full))
                    {
                        safe = (File.GetAttributes(full) & FileAttributes.ReparsePoint) == 0;
                    }

                    if (safe)
                    {
                        _log = new StreamWriter(full, append: true, new UTF8Encoding(false)) { AutoFlush = true };
                        Console.Error.WriteLine("[INFO] 日志追加写入：" + full);
                    }
                    else
                    {
                        Console.Error.WriteLine("[WARN] --log 路径不合法（需完全限定的本地固定盘绝对路径，拒绝 UNC/重解析点），已禁用文件日志：" + logPath);
                    }
                }
                catch (Exception logException)
                {
                    Console.Error.WriteLine("[WARN] --log 打开失败，已禁用文件日志：" + logException.Message);
                }
            }

            // M29：探测模式——只回 initialize，不预热 Excel、不挂看门狗，供侧边栏「测试连接」安全握手。
            // 审查 A5/S8：常规启动会无条件启动预热线程（15-25s 后才挂 Job），探测若在此期间 Kill 会留下孤儿 EXCEL。
            var probeOnly = args.Contains("--probe");
            _probeOnly = probeOnly;

            // M26 懒初始化：后台线程预热 Excel（冷启动 15-25s），stdio 循环立即启动响应 MCP 握手——
            // 修复「宿主等待握手超时 → host-tools-pending 连不上」的根因。
            if (probeOnly)
            {
                // 探测模式无 stdin 输入时 20s 自动退出，防探测端异常未关闭导致僵尸进程
                var probeTimer = new System.Threading.Thread(() =>
                {
                    System.Threading.Thread.Sleep(20000);
                    Environment.Exit(0);
                });
                probeTimer.IsBackground = true;
                probeTimer.Start();
                Log("MCP probe 模式：不预热 Excel");
            }
            else
            {
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
                StartWorker(); // R20：执行线程先起，读循环只做分派
                StartParentWatchdog(); // M28：宿主硬杀兜底（下方 finally 只覆盖温和退出）
                Log("MCP server：stdio 循环启动，Excel 后台预热中…");
            }

            // R1 修复：stdio 主循环任何异常（客户端断连 IOException/畸形输入）都不得跳过 Cleanup——
            // 泄漏 Visible=false 的 EXCEL.EXE 是红线级故障。
            try
            {
                RunStdioLoop();
            }
            finally
            {
                // R20 回归修复：stdin 关闭（管道/一次性脚本）时读循环先退出，
                // 已入队的 tools/call 若不等就成了"静默丢结果"。先排空队列再 Cleanup，
                // 上限 60s——超过就放弃等待，避免 Cleanup 永远挂住。
                try
                {
                    _workQueue.CompleteAdding();
                    _workerThread?.Join(TimeSpan.FromSeconds(60));
                }
                catch { }
                // 审查 B1：probe 模式从未 Set _readyEvent，走通用 Shutdown 会白等 120s 才被定时器打断。
                // 探测路径直接落盘退出，不进入 Cleanup（probe 从不创建 Excel，无资源可回收）。
                if (probeOnly)
                {
                    try { _log?.Flush(); _log?.Dispose(); } catch { }
                }
                else
                {
                    Shutdown("stdio 结束");
                }
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

            // M28 审查（第三轮 #7）：激活前快照宿主进程 PID 集合——Hwnd 归因得到的 PID 若激活前已存在，
            // 即为复用实例（不依赖时钟粒度的归属强证据）；快照失败退化为仅 StartTime 判定。
            var hostPidsBefore = SnapshotHostPids();
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
                Log("PID 归因失败: " + pidException.Message);
            }

            // M28 审查（审查者）：COM 激活可能复用用户正在编辑的实例（WPS 单实例）——后续
            // Visible=false/Workbooks.Add/Quit 都会操作用户工作簿。只有「PID 可归因且启动时间晚于
            // server」的自建实例才允许继续；否则整体初始化失败（不碰用户实例）。
            if (_spawnedPid <= 0)
            {
                _initError = "无法归因 COM 实例 PID（Hwnd 不可用）。为避免误操作用户工作簿，server 拒绝初始化；请确认本机 Excel 可独立启动。";
                _readyEvent.Set();
                return;
            }

            if (hostPidsBefore != null && hostPidsBefore.Contains(_spawnedPid))
            {
                _initError = "COM 激活复用了本 server 启动前已存在的宿主进程（PID=" + _spawnedPid + "，激活前快照命中）。" +
                             "为避免关闭/改写用户正在编辑的工作簿，server 拒绝初始化；请先关闭该 Excel/WPS 实例后重试。";
                Log("初始化失败：" + _initError);
                _readyEvent.Set();
                return;
            }

            try
            {
                using (var probe = System.Diagnostics.Process.GetProcessById(_spawnedPid))
                {
                    if (probe.StartTime.ToUniversalTime() < _serverStartUtc)
                    {
                        _initError = "COM 激活复用了已存在的实例（PID=" + _spawnedPid + "，启动早于 server）。" +
                                     "为避免关闭/改写用户正在编辑的工作簿，server 拒绝连接；请先关闭该 Excel/WPS 实例后重试。";
                        _readyEvent.Set();
                        return;
                    }
                }
            }
            catch (Exception probeException)
            {
                _initError = "COM 实例 PID 启动时间核验失败：" + probeException.Message +
                             "；为避免误操作用户工作簿，server 拒绝初始化。";
                _readyEvent.Set();
                return;
            }

            _ownsInstance = true;
            TryAttachExcelJob(_spawnedPid); // M28：硬杀兜底——Excel 挂 KILL_ON_JOB_CLOSE，server 任何方式退出都不留孤儿
            if (!_jobAttached)
            {
                // M28 审查（第三轮）：Job 建不起来就没有硬杀兜底——不能只"拒绝初始化"把 Excel 和僵尸进程留下：
                // 先回收自有实例，再退出让宿主重启（错误里带失败阶段 + Win32Error，便于用户处置）。
                var detail = string.IsNullOrEmpty(JobError) ? "未知原因" : JobError;
                _initError = "Excel 的 KILL_ON_JOB_CLOSE Job 建立失败（" + detail +
                             "）：无法保证 server 被硬杀时回收 EXCEL。已回收本次自建实例并退出，请重启宿主后重试。";
                Log("初始化失败（" + detail + "），回收自建 Excel 后退出");
                try { Cleanup(); } catch (Exception cleanupException) { Log("初始化失败回收异常: " + cleanupException.Message); }
                Environment.Exit(3);
                return;
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
            SweepStaleConfirmationFiles(); // M27R3：启动即扫（旧命名 <tool>-<CODE>.txt 孤儿与 >24h 文件），不等首调

            if (_readOnlyMode)
            {
                foreach (var tool in _registry.Tools.Where(t => t.RiskLevel != RiskLevel.ReadOnly).ToList())
                {
                    _registry.Remove(tool.Name);
                }
            }

            // M36-2：工具面分层。197 个工具的 schema 一次性塞给客户端≈上万 token，
            // 而 90% 的会话只用到"看-写-验-图"这条主干。JYYJ_TOOL_PROFILE=core 只暴露高频子集。
            if (ToolProfile == "core")
            {
                foreach (var tool in _registry.Tools.Where(t => !CoreProfileTools.Contains(t.Name)).ToList())
                {
                    _registry.Remove(tool.Name);
                    ProfileHiddenTools.Add(tool.Name);
                }
            }

            Log("Excel 预热完成：工具数=" + _registry.Tools.Count + "，模式=" + (_readOnlyMode ? "只读" : "读写")
                + "，工具面=" + ToolProfile + "，Excel PID=" + _spawnedPid);
            _readyEvent.Set();
        }

        /// <summary>M28：把预热出的 Excel 进程挂到本 server 的 Job（KILL_ON_JOB_CLOSE）。
        /// 内核语义：job 句柄因进程退出（含 TerminateProcess 硬杀）而关闭时，job 内所有进程立即终止——
        /// 这是「MCP server 被硬杀留孤儿 EXCEL」的根治手段；M28（第三轮）：挂接失败不再降级——初始化失败、回收实例后退出。</summary>
        private static void TryAttachExcelJob(int excelPid)
        {
            if (excelPid <= 0) { JobError = "PID 无效"; return; }
            // M28 审查（AI 使用方预警）：挂 Job 前必须做「启动时间晚于 server」校验——WPS 单实例
            // 场景下 COM 激活可能复用用户正在编辑的实例（Hwnd 归因会拿到用户 PID）；把用户实例挂进
            // KILL_ON_JOB_CLOSE 会让 server 退出误杀用户工作簿。与 Cleanup 同口径，不符即跳过（由初始化致命门拒绝启动）。
            try
            {
                using (var probe = System.Diagnostics.Process.GetProcessById(excelPid))
                {
                    if (probe.StartTime.ToUniversalTime() < _serverStartUtc)
                    {
                        JobError = "启动时间早于 server（疑似用户实例，防误杀）";
                        Log("Job 跳过：PID=" + excelPid + " " + JobError);
                        return;
                    }
                }
            }
            catch (Exception probeException)
            {
                JobError = "启动时间核验失败（" + probeException.GetType().Name + "）";
                Log("Job 跳过：PID=" + excelPid + " " + JobError);
                return;
            }

            var job = IntPtr.Zero;
            try
            {
                job = Win32Native.CreateJobObject(IntPtr.Zero, null);
                if (job == IntPtr.Zero)
                {
                    JobError = "CreateJobObject Win32Error=" + Marshal.GetLastWin32Error();
                    Log("Job 创建失败 " + JobError);
                    return;
                }

                var info = new Win32Native.JOBOBJECT_EXTENDED_LIMIT_INFORMATION();
                info.BasicLimitInformation.LimitFlags = Win32Native.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
                var length = Marshal.SizeOf(typeof(Win32Native.JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
                var buffer = Marshal.AllocHGlobal(length);
                try
                {
                    Marshal.StructureToPtr(info, buffer, false);
                    if (!Win32Native.SetInformationJobObject(job, Win32Native.JobObjectExtendedLimitInformation, buffer, (uint)length))
                    {
                        JobError = "SetInformationJobObject Win32Error=" + Marshal.GetLastWin32Error();
                        Log("Job SetInformation 失败 " + JobError);
                        Win32Native.CloseHandle(job);
                        return;
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }

                using (var excel = System.Diagnostics.Process.GetProcessById(excelPid))
                {
                    if (!Win32Native.AssignProcessToJobObject(job, excel.Handle))
                    {
                        JobError = "AssignProcessToJobObject Win32Error=" + Marshal.GetLastWin32Error();
                        Log("Job Assign 失败 " + JobError);
                        Win32Native.CloseHandle(job);
                        return;
                    }
                }

                _jobHandle = job; // 全程持有：server 退出（任意方式）→ 句柄关闭 → Excel 被终止
                _jobAttached = true;
                Log("Excel 已挂入 KILL_ON_JOB_CLOSE Job（硬杀兜底生效）");
            }
            catch (Exception ex)
            {
                JobError = "Job 附加异常：" + ex.Message;
                Log(JobError);
                if (job != IntPtr.Zero) { try { Win32Native.CloseHandle(job); } catch { } }
            }
        }

        /// <summary>M28 审查（第三轮 #7）：激活前宿主进程（EXCEL/et）PID 快照；失败返回 null（退化为 StartTime 判定）。</summary>
        private static HashSet<int> SnapshotHostPids()
        {
            try
            {
                var set = new HashSet<int>();
                foreach (var name in new[] { "EXCEL", "et", "wps" })
                {
                    foreach (var p in System.Diagnostics.Process.GetProcessesByName(name))
                    {
                        try { set.Add(p.Id); } catch { } finally { p.Dispose(); }
                    }
                }

                return set;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>M28：统一关停（stdio finally 与父进程看门狗共用）。
        /// 必须先等后台预热结束再清理——否则客户端连上即断/父进程早死时 _app 尚 null、
        /// SpawnedPids 为空，Cleanup 不 Quit 不 Kill，预热线程随后建出的 EXCEL 成无主孤儿（红线）。</summary>
        private static void Shutdown(string reason)
        {
            if (System.Threading.Interlocked.Exchange(ref _shutdownStarted, 1) != 0)
            {
                // 另一路径已接管关停：等它完成（上限覆盖最慢预热+清理），防 Main 提前返回打断清理
                try { _shutdownDone.Wait(TimeSpan.FromSeconds(150)); } catch { }
                return;
            }

            Log("清理开始：" + reason);
            try
            {
                // M28 审查：等待必须覆盖最长预热路径（tools/list 窗口 120s）——冷启动时若提前放弃，
                // 激活中的 EXCEL 尚未挂 Job（TryAttachExcelJob 在 CreateInstance 之后）会成孤儿（红线）。
                try { _readyEvent.Wait(TimeSpan.FromSeconds(120)); } catch { }
                Cleanup();
            }
            catch (Exception cleanupException)
            {
                Log("Cleanup 异常: " + cleanupException.Message);
            }
            finally
            {
                _shutdownDone.Set();
            }
        }

        /// <summary>M28 父进程看门狗：宿主被硬杀（TerminateProcess / 关闭宿主进程）时 finally 不会执行，
        /// MCP server 与预热出的 EXCEL.EXE 成孤儿（EXCEL 泄漏是红线）。每 2s 校验父进程；
        /// PID + 启动时间双重比对防 PID 复用误判；父进程消失 → Shutdown 后退出。</summary>
        private static void StartParentWatchdog()
        {
            int parentPid;
            DateTime parentStart;
            try
            {
                parentPid = GetParentProcessId();
                if (parentPid <= 0 || parentPid == System.Diagnostics.Process.GetCurrentProcess().Id)
                {
                    Log("父进程看门狗：未识别到有效父进程（服务/手动启动），跳过");
                    return;
                }

                parentStart = System.Diagnostics.Process.GetProcessById(parentPid).StartTime;
            }
            catch (Exception ex)
            {
                Log("父进程看门狗启动失败（跳过，不影响主流程）: " + ex.Message);
                return;
            }

            var thread = new System.Threading.Thread(() =>
            {
                while (true)
                {
                    try { System.Threading.Thread.Sleep(2000); } catch { }
                    var parentGone = false;
                    try
                    {
                        var parent = System.Diagnostics.Process.GetProcessById(parentPid);
                        parentGone = parent.HasExited || parent.StartTime != parentStart;
                    }
                    catch (ArgumentException)
                    {
                        parentGone = true; // 进程不存在
                    }
                    catch
                    {
                        continue; // 查询瞬时异常：下一轮再判，不误杀
                    }

                    if (!parentGone) { continue; }
                    Log("父进程 " + parentPid + " 已退出，看门狗接管清理");
                    Shutdown("父进程已退出（看门狗）");
                    Environment.Exit(0);
                    return;
                }
            });
            thread.IsBackground = true;
            thread.Name = "parent-watchdog";
            thread.Start();
            Log("父进程看门狗已启动（parent PID=" + parentPid + "）");
        }

        /// <summary>取当前进程的父进程 PID（Toolhelp 快照；失败返回 0）。</summary>
        private static int GetParentProcessId()
        {
            var snapshot = Win32Native.CreateToolhelp32Snapshot(Win32Native.TH32CS_SNAPPROCESS, IntPtr.Zero);
            if (snapshot == IntPtr.Zero || snapshot == new IntPtr(-1)) { return 0; }
            try
            {
                var entry = new Win32Native.PROCESSENTRY32();
                entry.dwSize = (uint)Marshal.SizeOf(typeof(Win32Native.PROCESSENTRY32));
                if (!Win32Native.Process32First(snapshot, ref entry)) { return 0; }
                var selfId = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                do
                {
                    if (entry.th32ProcessID == selfId) { return (int)entry.th32ParentProcessID; }
                }
                while (Win32Native.Process32Next(snapshot, ref entry));
                return 0;
            }
            finally
            {
                Win32Native.CloseHandle(snapshot);
            }
        }

        /// <summary>强杀本进程新建的 EXCEL.EXE；用户既有实例绝不动。</summary>
        private static void Cleanup()
        {
            // 先优雅退出（DisplayAlerts=false 下无对话框风险），再对拒绝退出的实例兜底强杀。
            // M28 审查：只对自有实例 Quit——初始化复用用户实例时 _ownsInstance=false，绝不能 Quit。
            if (_ownsInstance && _app != null && Marshal.IsComObject(_app))
            {
                // N10（P3）：Quit 是同步 COM 调用，Excel 卡模态框/崩溃时永不返回——挂后台线程 + 15s
                // 看门狗：超时不再等 Quit，直接落到下面的按 PID 强杀路径（强杀是兜底语义）。
                var quitApp = _app;
                var quitDone = new System.Threading.ManualResetEventSlim(false);
                var quitThread = new System.Threading.Thread(() =>
                {
                    try { quitApp.Quit(); } catch { }
                    finally { quitDone.Set(); }
                });
                quitThread.IsBackground = true;
                try { quitThread.Start(); } catch { }
                try { quitDone.Wait(TimeSpan.FromSeconds(15)); } catch { }
            }

            // R2-1 修复：按 PID 直取强杀（原 GetProcessesByName("EXCEL") 在 WPS 宿主下永不命中 = 死代码）。
            // M28（第三轮 #7）：Kill 前再加"激活前快照"强证据——归因 PID 命中激活前集合即为复用实例，绝不 Kill。
            var snapshotAtExit = SnapshotHostPids();
            foreach (var pid in SpawnedPids)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(pid);
                    var procName = "?";
                    try { procName = p.ProcessName; } catch { } // FIX（M28）：Kill 后再读 ProcessName 抛 ArgumentException，被误记成"未执行"
                    if (snapshotAtExit != null && snapshotAtExit.Contains(pid) && !_ownsInstance)
                    {
                        Log("跳过 PID=" + pid + "（激活前快照命中：复用实例，防误杀）");
                        p.Dispose();
                        continue;
                    }

                    if (p.StartTime.ToUniversalTime() >= _serverStartUtc) // FIX(M27R2)：StartTime 是 Local，直接与 Utc 比会因 Kind 被忽略而恒真
                    {
                        // 用户拍板（2026-11-04）：疑似复用实例只正常关闭、绝不强杀。
                        // 激活前快照之后出现的实例无法区分“本次新建”与“用户空档期新建后被单实例宿主复用”，
                        // 强杀可能连带用户未保存的工作簿。COM Quit 已执行，这里只做优雅关闭尝试，失败仅记录。
                        var closeRequested = false;
                        try { closeRequested = p.CloseMainWindow(); } catch { closeRequested = false; }
                        Log(closeRequested
                            ? "已请求关闭 spawn 进程 PID=" + pid + " name=" + procName + "（不强杀）"
                            : "跳过强杀 PID=" + pid + "（归属证据不足，仅 Quit，防误杀用户实例）");
                    }
                    else
                    {
                        Log("跳过 PID=" + pid + "（启动时间早于 server，疑似宿主语义漂移，防误杀）");
                    }
                    p.Dispose();
                }
                catch (ArgumentException)
                {
                    Log("PID=" + pid + " 已退出（无需强杀）"); // M28：Quit 已生效时属正常路径，不算失败
                }
                catch (InvalidOperationException)
                {
                    Log("PID=" + pid + " 已退出（无需强杀）");
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

                        // （空行跳过逻辑见下方有界循环）
                        // M27R2：有界跳过空行（原 do/while 在「无空行」宽容帧下会把正文首行当头部吞掉）
                        for (int blank = 0; blank < 2; blank++)
                        {
                            var peekChar = reader.Peek();
                            if (peekChar < 0) { break; }
                            if (peekChar == '\r') { reader.Read(); continue; }
                            if (peekChar == '\n') { reader.Read(); break; }
                            break; // 非空行：认为正文立即开始，不再吞行
                        }
                        // M27R2：非破坏性按 UTF-8 字节读帧——先 Peek 预估该字符字节数，
                        // 若会越过 len 则不消费（留给下一帧），彻底消除「多吞下一帧字节 → 流永久错位」。
                        var sb = new System.Text.StringBuilder();
                        int frameBytes = 0;
                        while (frameBytes < len)
                        {
                            int peek = reader.Peek();
                            if (peek < 0) { break; }
                            char c = (char)peek;
                            if (char.IsHighSurrogate(c))
                            {
                                if (frameBytes + 4 > len) { break; } // 代理对整体放不下：留给下一帧
                                reader.Read();
                                sb.Append(c);
                                int lo = reader.Peek();
                                if (lo >= 0 && char.IsLowSurrogate((char)lo)) { reader.Read(); sb.Append((char)lo); frameBytes += 4; }
                                else { frameBytes += 3; } // 孤立高代理（.NET 编码为 3 字节 U+FFFD）
                            }
                            else
                            {
                                int cb = c <= 0x7F ? 1 : (c <= 0x7FF ? 2 : 3);
                                if (frameBytes + cb > len) { break; } // 该字符属于下一帧：不消费
                                reader.Read();
                                sb.Append(c);
                                frameBytes += cb;
                            }
                        }

                        // 重同步：若声明长度与实际消费字节不符（畸形帧），按剩余字节数丢弃，
                        // 避免残余字节被外层 ReadLine 当成下一条消息（此前会连环污染后续所有帧）。
                        int residual = len - frameBytes;
                        for (int d = 0; d < residual && residual < 1024 * 1024; d++)
                        {
                            if (reader.Read() < 0) { break; }
                        }

                        HandleMessage(sb.ToString(), stdout); // FIX(M27R2 回归)：读满正文后必须派发，否则整帧被吞
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
            // R23：method 类型必须先判定——直接 (string) 强转对象/数组会抛 ArgumentException 穿出读循环，
            // 把整个 server 打停（畸形请求不该升级成进程级故障）。
            var methodToken = msg["method"];
            if (methodToken != null && methodToken.Type != JTokenType.String && methodToken.Type != JTokenType.Null)
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = msg["id"],
                    ["error"] = new JObject { ["code"] = -32600, ["message"] = "Invalid Request: method 必须是字符串" }
                });
                return;
            }
            var method = methodToken == null || methodToken.Type == JTokenType.Null ? null : (string)methodToken;

            // 通知（无 id 或显式 notifications/*）：initialized 等不回复
            if (id == null || id.Type == JTokenType.Null || (method != null && method.StartsWith("notifications/", StringComparison.Ordinal)))
            {
                Log("通知: " + method);
                return;
            }

            // 审查 B2：探测模式只做握手——其余方法直接回 -32601，避免 tools/list 在未预热时白等 120s。
            if (_probeOnly && method != null
                && method != "initialize" && method != "ping"
                && !method.StartsWith("notifications/", StringComparison.Ordinal))
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["error"] = new JObject
                    {
                        ["code"] = -32601,
                        ["message"] = "probe 模式仅支持 initialize/ping（--probe 用于连通性自检，不加载 Excel）"
                    }
                });
                return;
            }

            try
            {
                switch (method)
                {
                    case "initialize":
                        HandleInitialize(id, msg, stdout);
                        break;
                    case "ping":
                        Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject() });
                        break;
                    case "tools/list":
                        HandleToolsList(id, stdout);
                        break;
                    case "tools/call":
                        // R20：长任务下移到 worker，读循环继续应答 ping/取消
                        _workQueue.Add(() =>
                        {
                            try { HandleToolsCall(id, msg, stdout); }
                            catch (Exception ex)
                            {
                                Log("tools/call 异常: " + ex);
                                Emit(stdout, new JObject
                                {
                                    ["jsonrpc"] = "2.0",
                                    ["id"] = id,
                                    ["error"] = new JObject { ["code"] = -32603, ["message"] = "Internal error: " + ex.Message }
                                });
                            }
                        });
                        break;
                    case "resources/list":
                        HandleResourcesList(id, stdout);
                        break;
                    case "resources/read":
                        HandleResourcesRead(id, msg, stdout);
                        break;
                    case "prompts/list":
                        HandlePromptsList(id, stdout);
                        break;
                    case "prompts/get":
                        HandlePromptsGet(id, msg, stdout);
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

        /// <summary>本 server 实现过的协议版本（新→旧）。协商规则见 MCP 生命周期规范：
        /// 客户端请求版本若在列表内则回同一版本，否则回 server 支持的最高稳定版本。</summary>
        private static readonly string[] SupportedProtocolVersions = { "2025-06-18", "2025-03-26", "2024-11-05" };

        private static string NegotiateProtocolVersion(JObject msg)
        {
            var requested = msg?["params"]?["protocolVersion"]?.Type == JTokenType.String
                ? (string)msg["params"]["protocolVersion"]
                : null;
            if (!string.IsNullOrWhiteSpace(requested))
            {
                foreach (var v in SupportedProtocolVersions)
                {
                    if (string.Equals(v, requested, StringComparison.Ordinal)) { return v; }
                }
                Log("initialize：客户端请求协议版本 " + requested + " 不受支持，回退 " + SupportedProtocolVersions[SupportedProtocolVersions.Length - 1]);
            }
            // 无请求版本时保持历史行为（Claude Desktop 长期使用 2024-11-05）
            return "2024-11-05";
        }

        private static void HandleInitialize(JToken id, JObject msg, Stream stdout)
        {
            Emit(stdout, new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["result"] = new JObject
                {
                    ["protocolVersion"] = NegotiateProtocolVersion(msg),
                    ["capabilities"] = new JObject
                    {
                        ["tools"] = new JObject { ["listChanged"] = false },
                        ["resources"] = new JObject { ["listChanged"] = false, ["subscribe"] = false },
                        ["prompts"] = new JObject { ["listChanged"] = false },
                        ["experimental"] = new JObject
                        {
                            ["aiAutonomous"] = AutonomousMode,
                            ["screenshot"] = _allowWrite,
                            ["selfCheck"] = true,
                            ["correction"] = true
                        }
                    },
                    ["serverInfo"] = new JObject
                    {
                        ["name"] = "jyyj-excel",
                        ["version"] = typeof(Program).Assembly.GetName().Version.ToString()
                    },
                    // MCP 2025-03 起 initialize 可返回 instructions（服务端给模型的常驻提示）。
                    // 这里解决一个真实断点：只读模式下 tools/list 只含 ReadOnly 工具，
                    // 模型看不到写工具会误判"这服务只能读"，于是在这里明说原因与开启方式。
                    ["instructions"] = BuildInstructions()
                }
            });
        }

        private static string BuildInstructions()
        {
            var readOnly = _registry != null && _registry.Tools.All(t => t.RiskLevel == RiskLevel.ReadOnly);
            if (_probeOnly)
            {
                return "JYYJ助手（探测模式）：本进程仅做连通性握手，未加载 Excel，任何工具调用都会返回 -32601。正式使用请去掉 --probe。";
            }
            if (_registry == null)
            {
                // 正常启动时 initialize 往往先于预热完成到达（stdio 先就绪、Excel 冷启动 15-25s）。
                // 此时不能断言写模式可用，只能如实说明"仍在预热"，并提示参数决定写权限。
                return "JYYJ助手：Excel 正在后台预热（冷启动约 15-25 秒），工具列表稍后可用。"
                     + (_allowWrite ? "【写模式已开启】" : "【当前为只读模式，仅暴露只读工具；加 --allow-write 并重启可用写入】")
                     + " 预热期间 tools/call 会返回 EXCEL_WARMING，稍后重试即可。";
            }
            var sb = new System.Text.StringBuilder();
            sb.Append("JYYJ助手：通过工具操作当前 Excel/WPS 工作簿。");
            if (ToolProfile == "core")
            {
                // 审查（AI 使用方视角）：core 面只有 54 个工具，模型很容易判定"没有这个能力"而放弃。
                // 必须主动说明这是精简面，并给出切到全量的方法。
                sb.Append("【当前为 core 精简工具面（" + _registry.Tools.Count + " 个高频工具）】："
                    + "未在此列表中的能力并未缺失，只是未暴露；如任务需要（如 VBA/Python 脚本、条件格式、透视表细项、外部数据源、"
                    + "草稿模式等），请提示用户把环境变量 JYYJ_TOOL_PROFILE 设为 all 后重启 MCP 服务。");
            }
            sb.Append("推荐工作闭环：先规划步骤，逐步执行；每次写入后用 run_self_check 复核（纯只读：只做 JSON 回读，不碰剪贴板/不落盘）；");

            if (!readOnly)
            {
                sb.Append("自检不通过时用 apply_correction 提交修正阶梯（≤5 轮，首个已验证通过即停）——以 verified 字段为准，verified=false 不得当作已修正。");
            }
            if (readOnly)
            {
                sb.Append("【当前为只读模式】工具列表只暴露只读工具（写入/结构变更/危险级均未列出）。");
                sb.Append("看图请用 render_range_image（只读渲染：不碰剪贴板、不插图表，只读模式可用，近似还原外观）；"
                          + "take_screenshot 属写入级（像素级外观，需 --allow-write）。");
                sb.Append("同理，connect_ai_client（把 JYYJ 接到 DSH/WorkBuddy 等客户端）属写入级，只读模式下不会出现在工具列表；")
                  .Append("接入流程是 connect_ai_client(apply=true) → 重启目标客户端 → test_ai_client_connection。");
            }
            else
            {
                sb.Append("【写模式已开启】写入与结构变更工具可用；危险级工具默认在 AI 自主模式下直接执行，");
                sb.Append("可用 --require-confirmation 恢复人工确认门。");
                sb.Append("需要「看见」结果时调用 take_screenshot（写入级，像素级外观；会占用剪贴板并临时插入图表对象，但不改单元格值/格式）。");
            }
            return sb.ToString();
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

        /// <summary>M26：等待 Excel 后台预热完成；失败抛人话异常。
        /// M28 审查（AI 使用方）：窗口 45s——必须低于常见宿主 60s 请求超时；超时返回可重试的
        /// EXCEL_WARMING（预热在后台继续，本次调用**未执行任何工具**，重试不会造成重复写入）。</summary>
        private static void EnsureReady(int timeoutMs = 45000)
        {
            if (!_readyEvent.Wait(timeoutMs))
            {
                throw new InvalidOperationException(
                    "EXCEL_WARMING：Excel 仍在后台预热（已等待 " + (timeoutMs / 1000) + " 秒）。" +
                    "本次调用未执行任何工具；请稍后重试同一调用（宿主超时重试不会造成重复写入）。" +
                    "若持续失败，请确认本机 Excel 可正常手动启动（COM 未被占用）。");
            }
            if (!string.IsNullOrEmpty(_initError))
            {
                throw new InvalidOperationException("Excel 初始化失败：" + _initError);
            }
        }

        private static void HandleToolsList(JToken id, Stream stdout)
        {
            // M28 审查（第三轮）：tools/list 是客户端初始化后的首个请求，窗口放宽到 120s（冷启动）；
            // 仍超时则返回带 data.error_code=EXCEL_WARMING/retry_after_seconds 的机读错误。
            try { EnsureReady(120000); }
            catch (Exception readyEx)
            {
                var error = new JObject { ["code"] = -32000, ["message"] = readyEx.Message };
                if (readyEx.Message.StartsWith("EXCEL_WARMING", StringComparison.Ordinal))
                {
                    error["data"] = new JObject
                    {
                        ["error_code"] = "EXCEL_WARMING",
                        ["retry_after_seconds"] = 15,
                        ["tool_executed"] = false
                    };
                }
                Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = error });
                return;
            }

            var tools = new JArray();
            foreach (var tool in _registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
            {
                tools.Add(new JObject
                {
                    ["name"] = tool.Name,
                    ["description"] = RiskHint(tool) + tool.Description,
                    ["inputSchema"] = ToJsonSchema(tool.Parameters),
                    // MCP 2025-06 起客户端可按 annotations 自动放行只读调用、对破坏性调用强制人工确认。
                    // 纯附加字段：不认识它的客户端直接忽略，不改变既有行为。
                    ["annotations"] = new JObject
                    {
                        ["readOnlyHint"] = tool.RiskLevel == RiskLevel.ReadOnly,
                        ["destructiveHint"] = tool.RiskLevel == RiskLevel.Dangerous,
                        ["idempotentHint"] = tool.RiskLevel == RiskLevel.ReadOnly,
                        ["openWorldHint"] = false
                    }
                });
            }

            Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["tools"] = tools } });
            Log("tools/list 返回 " + tools.Count + " 个工具");
        }

        // ===================== prompts / resources（协议能力补齐） =====================

        private sealed class PromptSpec
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string Text { get; set; }
        }

        private static List<PromptSpec> ScenePrompts()
        {
            return new List<PromptSpec>
            {
                new PromptSpec
                {
                    Name = "report_monthly",
                    Description = "月度经营分析报告（汇总→图表→结论先行）",
                    Text = "请对当前工作簿生成月度经营分析报告：1) 先 get_current_excel_context / get_excel_sheet_list 确认数据表与区域；2) create_excel_pivot_summary 按月汇总关键指标；3) create_excel_chart 生成趋势图；4) verify_pivot 与 verify_chart 回读自检；5) 按「结论先行 + 图表分析 + 数据分析 + 计算依据 + 自检结果」输出，数字必须来自工具结果。"
                },
                new PromptSpec
                {
                    Name = "clean_data_audit",
                    Description = "数据清洗体检（空行/空值/文本型数字/重复/日期）",
                    Text = "请对当前表做清洗体检但先不改数据：1) analyze_column_profile 逐列画像；2) 汇总空行、空值率、文本型数字、重复记录、非法日期；3) 给出建议清洗动作清单（每项写明将调用哪个工具、影响范围）；4) 等我确认后再执行，执行后必须 verify_data_clean 自检。"
                },
                new PromptSpec
                {
                    Name = "dashboard_summary",
                    Description = "一屏 KPI 看板（4 张卡片 + 趋势）",
                    Text = "请基于当前数据生成 KPI 速览看板：1) 确认指标列与维度列；2) create_dashboard_kpi_brief 生成 4 张卡片；3) 需要图表时用 create_excel_chart 或 render_dashboard；4) 用对应 verify_* 回读自检；5) 说明每张卡片的计算口径（分子/分母/筛选条件）。"
                },
                new PromptSpec
                {
                    Name = "formula_audit",
                    Description = "公式审计（错误值/硬编码/可读性）",
                    Text = "请审计当前区域的公式：1) 用 preview_excel_range 或读取工具取公式；2) explain_formula 解释复杂嵌套；3) 标出 #REF!/#DIV/0!/#VALUE! 与硬编码常量、易失函数；4) 给出改写建议并说明影响范围；5) 改写后 verify_formula 自检。"
                },
                new PromptSpec
                {
                    Name = "outlier_scan",
                    Description = "异常值与异常行排查（只读，不改数据）",
                    Text = "请对当前数据做异常排查（只读）：1) analyze_column_profile + analyze_range_statistics 给出分布与边界；2) 标记疑似异常（超 3σ / IQR / 业务阈值）；3) 逐条给出判断依据与需人工确认项；4) 不写入任何单元格，等我确认清洗口径。"
                },
                new PromptSpec
                {
                    Name = "sheet_inventory",
                    Description = "工作簿结构盘点（表/行列/公式/图表）",
                    Text = "请盘点当前工作簿结构：1) get_excel_sheet_list 列出全部工作表；2) 逐表统计已用区域行列、公式格数、图表/图片/智能表数量；3) 标出空表、隐藏表、命名不规范的表；4) 输出一张结构清单表（不修改原数据）。"
                }
            };
        }

        private static List<PromptSpec> WorkflowPrompts()
        {
            var prompts = new List<PromptSpec>();
            var index = 0;
            foreach (var template in WorkflowStore.BuiltInWorkflows())
            {
                index++;
                prompts.Add(new PromptSpec
                {
                    Name = "workflow_" + index.ToString("00"),
                    Description = "内置流程：" + template.Name
                        + (string.IsNullOrWhiteSpace(template.Description) ? "" : " — " + template.Description),
                    Text = "请执行内置流程「" + template.Name + "」：调用 run_workflow，name 传「" + template.Name
                        + "」，按返回的步骤逐步执行；每步完成后用对应 verify_* 工具回读自检，失败时按错误码修正后重试，不要跳过自检。"
                });
            }

            return prompts;
        }

        private static List<PromptSpec> AllPrompts()
        {
            var all = ScenePrompts();
            all.AddRange(WorkflowPrompts());
            return all;
        }

        private static void HandlePromptsList(JToken id, Stream stdout)
        {
            var prompts = new JArray();
            foreach (var spec in AllPrompts())
            {
                prompts.Add(new JObject
                {
                    ["name"] = spec.Name,
                    ["description"] = spec.Description,
                    ["arguments"] = new JArray()
                });
            }

            Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["prompts"] = prompts } });
            Log("prompts/list 返回 " + prompts.Count + " 条");
        }

        private static void HandlePromptsGet(JToken id, JObject msg, Stream stdout)
        {
            var name = (string)msg["params"]?["name"];
            var spec = AllPrompts().FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
            if (spec == null)
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["error"] = new JObject { ["code"] = -32602, ["message"] = "未知 prompt：" + name }
                });
                return;
            }

            Emit(stdout, new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["result"] = new JObject
                {
                    ["description"] = spec.Description,
                    ["messages"] = new JArray
                    {
                        new JObject
                        {
                            ["role"] = "user",
                            ["content"] = new JObject { ["type"] = "text", ["text"] = spec.Text }
                        }
                    }
                }
            });
        }

        /// <summary>R21：活动工作簿实时快照（表名 + 已用区域）。COM 异常向上抛，由调用方降级为 JSON 错误体。</summary>
        private static string BuildActiveWorkbookJson()
        {
            dynamic wb = _app.ActiveWorkbook;
            var sheets = new JArray();
            if (wb != null)
            {
                foreach (object shObj in wb.Worksheets)
                {
                    dynamic sh = shObj;
                    var used = "";
                    try { used = (string)sh.UsedRange.Address; } catch { }
                    sheets.Add(new JObject { ["name"] = (string)sh.Name, ["used_range"] = used });
                }
            }
            var result = new JObject
            {
                ["ok"] = true,
                ["name"] = wb == null ? null : (string)wb.Name,
                ["path"] = wb == null ? null : Convert.ToString(wb.FullName),
                ["saved"] = wb == null ? false : (bool)wb.Saved,
                ["sheet_count"] = sheets.Count,
                ["sheets"] = sheets,
                ["active_sheet"] = wb == null ? null : (string)_app.ActiveSheet.Name
            };
            return result.ToString(Formatting.None);
        }

        /// <summary>R21：当前选区快照（地址/行列数/前 5×5 样本）。只读，不写任何单元格。</summary>
        private static string BuildCurrentSelectionJson()
        {
            dynamic sh = _app.ActiveSheet;
            dynamic sel = _app.Selection;
            var rows = 0; var cols = 0; var address = "";
            try { rows = (int)sel.Rows.Count; cols = (int)sel.Columns.Count; } catch { }
            try { address = (string)sel.Address; } catch { }

            var sample = new JArray();
            try
            {
                dynamic block = sel.Resize[Math.Min(5, Math.Max(rows, 1)), Math.Min(5, Math.Max(cols, 1))];
                object raw = block.Value2;
                if (raw is object[,] grid)
                {
                    var rMax = Math.Min(grid.GetLength(0), 5);
                    var cMax = Math.Min(grid.GetLength(1), 5);
                    for (var r = 1; r <= rMax; r++)
                    {
                        var line = new JArray();
                        for (var c = 1; c <= cMax; c++) { line.Add(grid[r, c] == null ? null : JToken.FromObject(grid[r, c])); }
                        sample.Add(line);
                    }
                }
                else { sample.Add(raw == null ? null : JToken.FromObject(raw)); }
            }
            catch { /* 图表区等非单元格选区：样本留空 */ }

            var result = new JObject
            {
                ["ok"] = true,
                ["sheet"] = sh == null ? null : (string)sh.Name,
                ["address"] = address,
                ["rows"] = rows,
                ["columns"] = cols,
                ["sample_rows"] = sample,
                ["note"] = "最多返回前 5 行 × 5 列样本；完整数据请用 preview_excel_sheet 或 read_excel_range。"
            };
            return result.ToString(Formatting.None);
        }

        private static void HandleResourcesList(JToken id, Stream stdout)
        {
            var resources = new JArray
            {
                new JObject { ["uri"] = "jyyj://tools/catalog", ["name"] = "工具目录", ["description"] = "当前可用工具（名称/风险级/描述）", ["mimeType"] = "application/json" },
                // R21（P1-1）：实时状态资源——模型开局一次 read 就能拿到工作簿/选区上下文，
                // 省掉一次 tools/call 往返，也避免"先猜后改"。
                new JObject { ["uri"] = "jyyj://workbook/active", ["name"] = "当前工作簿", ["description"] = "活动工作簿路径与各工作表名/已用区域（实时）", ["mimeType"] = "application/json" },
                new JObject { ["uri"] = "jyyj://selection/current", ["name"] = "当前选区", ["description"] = "活动工作表与当前选区的地址、行列数、前几行样本（实时）", ["mimeType"] = "application/json" },
                new JObject { ["uri"] = "jyyj://workflows/catalog", ["name"] = "流程模板目录", ["description"] = "内置流程模板清单（名称/描述/步骤数）", ["mimeType"] = "application/json" },
                new JObject { ["uri"] = "jyyj://help/errors", ["name"] = "错误码与恢复", ["description"] = "高频错误码的含义与恢复路径", ["mimeType"] = "text/markdown" }
            };

            Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = new JObject { ["resources"] = resources } });
        }

        private static void HandleResourcesRead(JToken id, JObject msg, Stream stdout)
        {
            var uri = (string)msg["params"]?["uri"];
            string text;
            string mime = "application/json";
            switch (uri)
            {
                // R21（P1-1）：实时工作簿/选区资源。与 tools/call 的区别是"读状态不该占用写队列"，
                // 且模型可在第一轮就拿到上下文。COM 异常一律降级为 JSON 错误体，不抛协议错误。
                case "jyyj://workbook/active":
                case "jyyj://selection/current":
                    try
                    {
                        EnsureReady(30000);
                        text = uri.EndsWith("/workbook/active", StringComparison.Ordinal)
                            ? BuildActiveWorkbookJson()
                            : BuildCurrentSelectionJson();
                    }
                    catch (Exception liveEx)
                    {
                        text = Newtonsoft.Json.JsonConvert.SerializeObject(new
                        {
                            ok = false,
                            error_code = liveEx.Message.StartsWith("EXCEL_WARMING", StringComparison.Ordinal) ? "EXCEL_WARMING" : "RESOURCE_UNAVAILABLE",
                            message = liveEx.Message
                        }, Newtonsoft.Json.Formatting.None);
                    }
                    break;
                case "jyyj://tools/catalog":
                    try { EnsureReady(120000); }
                    catch (Exception readyEx)
                    {
                        Emit(stdout, new JObject { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JObject { ["code"] = -32000, ["message"] = readyEx.Message } });
                        return;
                    }

                    text = new JArray(_registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                        .Select(t => new JObject
                        {
                            ["name"] = t.Name,
                            ["risk"] = t.RiskLevel.ToString(),
                            ["description"] = t.Description
                        })).ToString();
                    break;
                case "jyyj://workflows/catalog":
                    text = new JArray(WorkflowStore.BuiltInWorkflows().Select(t => new JObject
                    {
                        ["name"] = t.Name,
                        ["description"] = t.Description,
                        ["steps"] = t.Steps == null ? 0 : t.Steps.Count
                    })).ToString();
                    break;
                case "jyyj://help/errors":
                    mime = "text/markdown";
                    text = "# 高频错误码与恢复路径\n\n"
                        + "- `SHEET_PROTECTED`：目标表被保护。先 `verify_protection` 确认，经用户同意后 `protect_sheet_watermark`（action=unprotect）解除，再重放写入。\n"
                        + "- `INVALID_SHEET`：表名不存在。先 `get_excel_sheet_list` 核对精确拼写，不要回退活动表。\n"
                        + "- `CONSERVATION_VIOLATED`：排序/写回行数守恒不成立，写前已拦截。核对 region 参数后重试。\n"
                        + "- `NOT_ALLOWLISTED` / `CLI_DENIED`：脚本不在白名单，请用户在设置里加白，不要尝试路径绕过。\n"
                        + "- `BLOCKED_PRIVATE_ADDR`：fetch_url 目标是内网/回环地址，属安全拦截，不提供绕过。\n"
                        + "- `INVALID_ARGS`：参数不合法。tools/list 的 inputSchema 是唯一参数说明，按其修正。\n"
                        + "- `OVERWRITE_DENIED` / `BACKUP_FAILED`：目标文件已存在或备份失败，显式 `overwrite=true` 或换路径。\n"
                        + "- verify_* 的 `ok` 表示工具执行成功，判定结论看 `verdict`（ok/no_expected/mismatch…）。\n"
                        + "\n## 截图工具（take_screenshot）\n\n"
                        + "- `EMPTY_RANGE`：目标工作表无内容（UsedRange 仅 A1 且为空）。先 `get_excel_sheet_list` 核对表名，或改 `target=range` 并提供具体 address。\n"
                        + "- `RANGE_TOO_LARGE`：区域超过 200000 格上限。用 address 缩小到具体区域后重试。\n"
                        + "- `INVALID_TARGET`：target 拼写错误。仅支持 sheet/range/active_window。\n"
                        + "- `INVALID_FORMAT`：format 仅支持 png/jpeg。\n"
                        + "- `INVALID_PATH`：输出路径不合法（含 .. 段、UNC/网络盘等）。改用本地固定盘符路径，或省略 file_path 走临时目录。\n"
                        + "- `SCREENSHOT_FAILED`：截图执行失败（COM 异常/表被保护等）。看消息内 HRESULT 与提示；无头模式改 target=sheet。\n"
                        + "\n## 自检/修正工具（run_self_check / apply_correction）\n\n"
                        + "- `NO_VERIFY_PLAN`：该主工具无确定性自检计划（豁免清单工具属设计如此，消息里会给出豁免理由）。按业务意图显式调用对应 verify_* 工具。\n"
                        + "- `INVALID_ARGS`（run_self_check）：主工具实参被平铺在顶层。正确形态是 `{\"tool\":\"...\",\"args\":{...}}`，把实参移入 args 对象。\n"
                        + "- `SELF_CHECK_FAILED`：至少一条自检计划未通过。看 checks 内各条 verdict/data 定位差异格后修正。\n"
                        + "- `SELF_CHECK_SKIPPED`：verify 工具未注册或非只读被跳过。确认 verify 工具在当前模式可用。\n"
                        + "- `CORRECTION_FAILED`：所有修正尝试均未通过。检查各 attempt 的 main_error_code / verdict 后调整参数或换工具，勿盲目加轮数。\n"
                        + "- `CONFIRMATION_REQUIRED`（apply_correction 内）：不执行 Dangerous 子工具。需单独调用该工具并走确认门。\n"
                        + "- `UNKNOWN_TOOL`：工具未注册（只读模式下写工具已下线）。`--allow-write` 启动可开放。\n"
                        + "\n## 自主模式\n\n"
                        + "- Dangerous 分两类：结构/数据操作类（可 undo）在自主模式下直接放行；代码执行类（run_vba_snippet/run_python_*/run_cli_command/run_exe_command/install_python_packages）与路径敏类（open_excel_external_book/replace_external_links）仍需确认门。\n"
                        + "- 需要 VBA/Python 执行时：请用户以 `--require-confirmation` 关闭自主模式后走确认码流程，或用户在设置里加白。\n";
                    break;
                default:
                    Emit(stdout, new JObject
                    {
                        ["jsonrpc"] = "2.0",
                        ["id"] = id,
                        ["error"] = new JObject { ["code"] = -32602, ["message"] = "未知资源：" + uri }
                    });
                    return;
            }

            Emit(stdout, new JObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id,
                ["result"] = new JObject
                {
                    ["contents"] = new JArray
                    {
                        new JObject { ["uri"] = uri, ["mimeType"] = mime, ["text"] = text }
                    }
                }
            });
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
            // R22：tools/call 只等 5s——超时立刻回 EXCEL_WARMING（tool_executed=false），让模型退避重试，
            // 而不是把客户端 60s 超时线耗光。tools/list 仍等满 120s（客户端启动必经，不能返回空清单）。
            try { EnsureReady(5000); }
            catch (Exception readyEx)
            {
                // M28 审查（第三轮）：预热超时改为工具级机读结果（isError + _meta），
                // 明确 tool_executed=false，模型可安全重试；初始化硬失败仍走协议错误。
                if (readyEx.Message.StartsWith("EXCEL_WARMING", StringComparison.Ordinal))
                {
                    var warm = new JObject
                    {
                        ["ok"] = false,
                        ["error_code"] = "EXCEL_WARMING",
                        ["retry_after_seconds"] = 15,
                        ["tool_executed"] = false,
                        ["message"] = readyEx.Message
                    };
                    Emit(stdout, new JObject
                    {
                        ["jsonrpc"] = "2.0",
                        ["id"] = id,
                        ["result"] = new JObject
                        {
                            ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = warm.ToString(Newtonsoft.Json.Formatting.None) } },
                            ["isError"] = true,
                            ["_meta"] = new JObject { ["error_code"] = "EXCEL_WARMING", ["tool_executed"] = false }
                        }
                    });
                    return;
                }

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

            if (ProfileHiddenTools.Contains(name))
            {
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = new JObject
                        {
                            ["ok"] = false,
                            ["error_code"] = "PROFILE_FILTERED",
                            ["tool_executed"] = false,
                            ["message"] = "工具 " + name + " 未在当前工具面（JYYJ_TOOL_PROFILE=core）中暴露。"
                                + "如确需使用，请把环境变量设为 all 后重启 MCP 服务。"
                        }.ToString(Formatting.None) } },
                        ["isError"] = true,
                        ["_meta"] = new JObject { ["error_code"] = "PROFILE_FILTERED", ["tool_executed"] = false }
                    }
                });
                Log("tools/call 被工具面过滤: " + name);
                return;
            }

            // M36：allow_duplicate 是 server 级旁路开关，不能透传给工具（工具 schema 不认这个键，
            // 严格校验的工具会因此报 INVALID_ARGS）。先摘掉再走后续流程。
            var allowDuplicate = IsDuplicateAllowed(args);
            if (args != null) { args.Remove("allow_duplicate"); }

            // M36：写操作去重（只读工具允许重复；args 里显式 allow_duplicate=true 可绕过）
            if (tool.RiskLevel != RiskLevel.ReadOnly && !allowDuplicate)
            {
                var dedupKey = DedupKey(name, args);
                lock (RecentWritesLock)
                {
                    foreach (var stale in RecentWrites.Where(kv => (DateTime.UtcNow - kv.Value).TotalSeconds > WriteDedupWindowSeconds)
                                                      .Select(kv => kv.Key).ToList())
                    {
                        RecentWrites.Remove(stale);
                    }
                    if (RecentWrites.TryGetValue(dedupKey, out var at))
                    {
                        var seconds = (int)(DateTime.UtcNow - at).TotalSeconds;
                        var blocked = new JObject
                        {
                            ["ok"] = false,
                            ["error_code"] = "DUPLICATE_WRITE_SUPPRESSED",
                            ["tool_executed"] = false,
                            ["message"] = "相同工具与参数在 " + seconds + " 秒前已执行过，本次未重复写入。",
                            ["next_action"] = ExcelAiAssistant.Core.Errors.ErrorCatalog.NextActionOf("DUPLICATE_WRITE_SUPPRESSED")
                        };
                        Emit(stdout, new JObject
                        {
                            ["jsonrpc"] = "2.0",
                            ["id"] = id,
                            ["result"] = new JObject
                            {
                                ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = blocked.ToString(Formatting.None) } },
                                ["isError"] = true,
                                ["_meta"] = new JObject { ["error_code"] = "DUPLICATE_WRITE_SUPPRESSED", ["tool_executed"] = false }
                            }
                        });
                        Log("tools/call 去重拦截: " + name);
                        return;
                    }
                }
            }

            // Dangerous 级二次确认（需求/MCP接入.md §3）：首次调用返回确认载荷，confirmed=true 才放行
            // P0-3 确认门强制力：Dangerous 首调生成一次性确认码文件（人工读码），回传 confirmed=<code> 才放行
            // M29：AI 自主模式——Dangerous 级由 AI 直接执行，不再要求人工确认码；
            // 白名单外（代码执行/路径敏类）保留确认门（审查 S2：放行这些=默认任意本地文件读写）。
            // 审计无条件写 stderr（审查 S6：只落 --log 时默认配置零落盘）。
            if (tool.RiskLevel == RiskLevel.Dangerous && AutonomousMode && !RequiresConfirmationInAutonomous(name))
            {
                Console.Error.WriteLine("[AI-AUTONOMOUS] Dangerous 放行: " + name
                    + " argsHash=" + ConfirmationArgsHash(args).Substring(0, 12));
            }

            if (tool.RiskLevel == RiskLevel.Dangerous
                && (!AutonomousMode || RequiresConfirmationInAutonomous(name))
                && !IsConfirmedWithCode(name, args))
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
                var content = new JArray { new JObject { ["type"] = "text", ["text"] = text } };
                AppendImageContent(content, result);
                Emit(stdout, new JObject
                {
                    ["jsonrpc"] = "2.0",
                    ["id"] = id,
                    ["result"] = new JObject
                    {
                        ["content"] = content,
                        ["isError"] = !result.Ok
                    }
                });
                if (tool.RiskLevel != RiskLevel.ReadOnly && (result.Ok || result.ErrorCode == "BATCH_WRITE_PARTIAL"))
                {
                    // 只登记"确实写成"的调用：失败（含缺参数/瞬态错误）必须允许原样重试，
                    // 否则一次 SHEET_PROTECTED 之后的正确重试会被误判成重复写入。部分写入仍要防重。
                    lock (RecentWritesLock) { RecentWrites[DedupKey(name, args)] = DateTime.UtcNow; }
                }
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
                        ["content"] = new JArray { new JObject { ["type"] = "text", ["text"] = Newtonsoft.Json.JsonConvert.SerializeObject(new
                        {
                            ok = false,
                            error_code = "TOOL_EXCEPTION",
                            tool_executed = true,
                            message = ExcelAiAssistant.Core.Errors.ErrorCatalog.MessageOf("TOOL_EXCEPTION"),
                            next_action = ExcelAiAssistant.Core.Errors.ErrorCatalog.NextActionOf("TOOL_EXCEPTION"),
                            data = "工具执行异常: " + ex.Message
                        }, Newtonsoft.Json.Formatting.None) } },
                        ["isError"] = true
                    }
                });
            }
        }

        /// <summary>未知工具人话指引：只读模式下明确说明模式限制并列出可用只读工具族，防止 AI 反复重试写工具。</summary>
        internal static string BuildUnknownToolText(string name)
        {
            // FIX（M27R3）：未知工具返回**结构化 JSON**（原纯文本破坏「content[0].text 恒为 JSON」契约），
            // 同时保留 F2 语义：只读模式下说明模式限制、列出可用只读工具、明示请勿重试写操作。
            // M36-2：被工具面（core）隐藏的工具要给出可行动的提示，不能报成"未知工具"
            if (ProfileHiddenTools.Contains(name))
            {
                return Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error_code = "PROFILE_FILTERED",
                    data = "工具 " + name + " 存在于服务端，但未在当前工具面（JYYJ_TOOL_PROFILE=core）中暴露，"
                         + "因此 tools/list 看不到它。如确需使用，请把环境变量 JYYJ_TOOL_PROFILE 设为 all 后重启 MCP 服务。",
                    message = ExcelAiAssistant.Core.Errors.ErrorCatalog.MessageOf("PROFILE_FILTERED"),
                    next_action = ExcelAiAssistant.Core.Errors.ErrorCatalog.NextActionOf("PROFILE_FILTERED"),
                    tool_profile = ToolProfile
                }, Newtonsoft.Json.Formatting.None);
            }

            var msg = "未知工具：" + name + "。";
            var suggestions = SuggestToolNames(name);
            if (_readOnlyMode)
            {
                msg += "当前为只读模式（未加 --allow-write），写工具不可用；可用只读工具：" + string.Join("、", AvailableToolNames()) + "。";
                msg += "请勿重试写操作，可改用上述只读工具或提示用户以 --allow-write 启动。";
            }
            else if (suggestions.Count > 0)
            {
                msg += "相近工具建议：" + string.Join("、", suggestions) + "。";
            }
            msg += " 可调用 tools/list 获取完整工具清单。";
            return Newtonsoft.Json.JsonConvert.SerializeObject(new
            {
                ok = false,
                error_code = "UNKNOWN_TOOL",
                message = ExcelAiAssistant.Core.Errors.ErrorCatalog.MessageOf("UNKNOWN_TOOL"),
                next_action = ExcelAiAssistant.Core.Errors.ErrorCatalog.NextActionOf("UNKNOWN_TOOL"),
                data = msg,
                read_only_mode = _readOnlyMode,
                available_tools = _readOnlyMode ? AvailableToolNames() : null,
                suggestions = suggestions
            }, Newtonsoft.Json.Formatting.None);
        }

        /// <summary>当前注册表内全部工具名（只读模式下即只读工具清单）。</summary>
        private static System.Collections.Generic.List<string> AvailableToolNames()
        {
            var names = new System.Collections.Generic.List<string>();
            if (_registry == null) { return names; }
            foreach (var tool in _registry.Tools) { names.Add(tool.Name); }
            return names;
        }

        /// <summary>M27R3：按编辑距离给未知工具名找最相近的 3 个已注册工具（无匹配则返回空）。</summary>
        private static System.Collections.Generic.List<string> SuggestToolNames(string name)
        {
            var result = new System.Collections.Generic.List<string>();
            if (string.IsNullOrWhiteSpace(name) || _registry == null) { return result; }
            var scored = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>();
            foreach (var tool in _registry.Tools)
            {
                var d = Levenshtein(name.ToLowerInvariant(), tool.Name.ToLowerInvariant());
                if (d <= Math.Max(2, name.Length / 3)) { scored.Add(new System.Collections.Generic.KeyValuePair<int, string>(d, tool.Name)); }
            }
            scored.Sort((a, b) => a.Key.CompareTo(b.Key));
            for (var k = 0; k < scored.Count && k < 3; k++) { result.Add(scored[k].Value); }
            return result;
        }

        private static int Levenshtein(string a, string b)
        {
            var n = a.Length; var m = b.Length;
            var prev = new int[m + 1]; var cur = new int[m + 1];
            for (var j = 0; j <= m; j++) { prev[j] = j; }
            for (var i2 = 1; i2 <= n; i2++)
            {
                cur[0] = i2;
                for (var j2 = 1; j2 <= m; j2++)
                {
                    var cost = a[i2 - 1] == b[j2 - 1] ? 0 : 1;
                    cur[j2] = Math.Min(Math.Min(cur[j2 - 1] + 1, prev[j2] + 1), prev[j2 - 1] + cost);
                }
                var tmp = prev; prev = cur; cur = tmp;
            }
            return prev[m];
        }

        /// <summary>P0-3：确认码校验——比对 PendingConfirmations 中的一次性码，命中即消费（防重放）。
        /// 兼容旧 confirmed=true 语义仅限 --legacy-confirm 启动参数（默认关闭）。</summary>
        /// <summary>M27R3：上一次确认校验的失败原因（none / no-record / expired / mismatch:N / exhausted）。</summary>
        internal static string ConfirmationFailure = "none";

        internal static bool IsConfirmedWithCode(string toolName, Dictionary<string, object> args)
        {
            ConfirmationFailure = "none";
            if (!args.TryGetValue("confirmed", out var v) || v == null) { ConfirmationFailure = "no-record"; return false; }
            // #15（M28）：兼容旧 confirmed=true 语义——只有显式 --legacy-confirm 才认布尔 true；
            // 默认关闭，防「AI 自答式确认」回潮。
            if (LegacyConfirm && ((v is bool b && b) || string.Equals(Convert.ToString(v), "true", StringComparison.OrdinalIgnoreCase)))
            {
                ConfirmationFailure = "none";
                return true;
            }
            var supplied = Convert.ToString(v);
            if (!PendingConfirmations.TryGetValue(toolName, out var pending)) { ConfirmationFailure = "no-record"; return false; }

            // M27R2：TTL（10 分钟）——过期即作废并要求重新发起，压缩穷举窗口
            if (DateTime.UtcNow > pending.ExpiryUtc)
            {
                PendingConfirmations.TryRemove(toolName, out _);
                ConfirmationFailure = "expired";
                return false;
            }

            // M28 审查（第三轮）：码必须绑定参数集——人工确认的是 A 参数，不能用同一码执行 B 参数
            if (!string.Equals(ConfirmationArgsHash(args), pending.ArgsHash, StringComparison.Ordinal))
            {
                ConfirmationFailure = "args-mismatch";
                return false;
            }

            if (string.Equals(supplied, pending.Code, StringComparison.OrdinalIgnoreCase))
            {
                // M27R3：TryRemove 原子消费（并发同码重放只有一路赢）
                var consumed = PendingConfirmations.TryRemove(toolName, out _);
                ConfirmationFailure = consumed ? "none" : "no-record";
                return consumed;
            }

            // M27R2/R3：连续错码 ≥5 次即作废（Interlocked 原子计数，且不被重发重置）
            var fails = System.Threading.Interlocked.Increment(ref pending.Failures);
            if (fails >= 5)
            {
                PendingConfirmations.TryRemove(toolName, out _);
                ConfirmationFailure = "exhausted";
            }
            else
            {
                ConfirmationFailure = "mismatch:" + (5 - fails);
            }
            return false;
        }

        /// <summary>
        /// M29：把截图类工具结果里的 base64 图片作为 MCP image 内容块附加到 content 数组——
        /// 这是多模态模型"自己看结果做视觉自检"的通道；纯文本客户端会忽略 image 块，不影响兼容。
        /// 仅在 ok 且 Data 含 base64 + mime_type 时附加，大小上限由工具层（3MB）兜底。
        /// </summary>
        // M35（审查建议）：每会话图片配额，防止 AI 连续截图把客户端上下文撑爆。
        // 8 张 / 16MB（按 base64 解算后的真实字节估算：base64 长度 × 3/4）。
        // M36（审查建议）：写操作去重。客户端超时重试会重发同一条 tools/call（id 可能变、args 不变），
        // 之前会真的再写一遍。这里按「工具名 + 参数哈希」记最近执行的写操作，命中即拒绝执行。
        private static string DedupKey(string tool, Dictionary<string, object> args)
        {
            var json = args == null ? "{}" : JObject.FromObject(args).ToString(Formatting.None);
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(tool + "|" + json));
                return tool + ":" + BitConverter.ToString(bytes).Replace("-", "").Substring(0, 16);
            }
        }

        private static bool IsDuplicateAllowed(Dictionary<string, object> args)
        {
            if (args == null || !args.TryGetValue("allow_duplicate", out var v) || v == null) { return false; }
            if (v is bool b) { return b; }
            return string.Equals(Convert.ToString(v), "true", StringComparison.OrdinalIgnoreCase);
        }

        private static readonly Dictionary<string, DateTime> RecentWrites =
            new Dictionary<string, DateTime>(StringComparer.Ordinal);
        private static readonly object RecentWritesLock = new object();
        private const int WriteDedupWindowSeconds = 600;

        private const int MaxSessionImages = 8;
        private const long MaxSessionImageBytes = 16L * 1024 * 1024;
        private static int _sessionImageCount;
        private static long _sessionImageBytes;

        internal static void AppendImageContent(JArray content, ToolResult result)
        {
            // 审查 R28：不能用 result.Ok 门控——自检失败恰恰是最需要看图的时候
            if (content == null || result == null || result.Data == null)
            {
                return;
            }

            try
            {
                if (!(JToken.FromObject(result.Data) is JObject data))
                {
                    return;
                }

                var base64 = data.Value<string>("base64");
                if (string.IsNullOrWhiteSpace(base64))
                {
                    return;
                }

                // 配额检查：超限不投图，改投一条可行动提示（同一 worker 串行，无需加锁）
                var approxBytes = (long)(base64.Length * 3L / 4L);
                if (_sessionImageCount >= MaxSessionImages || _sessionImageBytes + approxBytes > MaxSessionImageBytes)
                {
                    content.Add(new JObject
                    {
                        ["type"] = "text",
                        ["text"] = "[图片已省略] 本会话图片配额已用尽（上限 " + MaxSessionImages + " 张 / "
                                 + (MaxSessionImageBytes / 1024 / 1024) + "MB）。请读取结果里的 file_path 查看，"
                                 + "或继续基于已返回的数值判定/文字说明作答，不要重复截图。"
                    });
                    return;
                }
                _sessionImageCount++;
                _sessionImageBytes += approxBytes;

                var mime = data.Value<string>("mime_type");
                if (string.IsNullOrWhiteSpace(mime))
                {
                    mime = "image/png";
                }

                content.Add(new JObject
                {
                    ["type"] = "image",
                    ["data"] = base64,
                    ["mimeType"] = mime
                });
            }
            catch (Newtonsoft.Json.JsonException)
            {
                // Data 非 JSON 形态：没有可附加的图片，保持纯文本结果。
            }
        }

        /// <summary>ToolResult → MCP text 载荷：snake_case 数据 + 警告 + 受影响区域 + 错误码全量透传。</summary>
        internal static string BuildResultText(ToolResult result) => BuildResultText(result, null);

        internal static string BuildResultText(ToolResult result, ToolDescriptor tool)
        {
            var payload = new JObject();
            payload["ok"] = result.Ok;
            if (result.Data != null)
            {
                try
                {
                    var token = result.Data is string s ? JToken.Parse(s) : JToken.FromObject(result.Data);
                    // M29 审查（U8）：base64 图片已由 AppendImageContent 以 image 内容块单独投递，
                    // 文本载荷不再重复序列化（否则多模态模型收两份、纯文本模型上下文被 base64 噪声塞满并触发截断）。
                    if (token is JObject dataObj && dataObj.Remove("base64"))
                    {
                        token = dataObj; // JObject.Remove 原位删除，token 即清理后对象
                    }

                    payload["data"] = token;
                }
                catch { payload["data"] = Convert.ToString(result.Data); }
            }

            if (!string.IsNullOrEmpty(result.Warning)) { payload["warning"] = result.Warning; }
            if (result.AffectedRanges != null && result.AffectedRanges.Count > 0)
            {
                payload["affected_ranges"] = JArray.FromObject(result.AffectedRanges);
            }

            if (!string.IsNullOrEmpty(result.ErrorCode))
            {
                payload["error_code"] = result.ErrorCode;
                // M36-3：错误码目录富化——补人话与下一步动作，AI 不用再猜 NO_VERIFY_PLAN 之类是什么意思
                if (ExcelAiAssistant.Core.Errors.ErrorCatalog.TryGet(result.ErrorCode, out var entry))
                {
                    // data 可能是 JValue（字符串/数字），必须先按 JObject 判型再取字段，否则 Newtonsoft 会抛
                    // "Cannot access child value on JValue"（真机踩到过）
                    var dataObj = payload["data"] as JObject;
                    if (dataObj == null || dataObj["message"] == null) { payload["message"] = entry.Message; }
                    if (dataObj == null || dataObj["next_action"] == null) { payload["next_action"] = entry.NextAction; }
                }
            }
            if (result.Truncated) { payload["truncated"] = true; }

            var text = payload.ToString(Formatting.Indented);
            // R4 修复：传输层限幅不得产出非法 JSON（原 Substring 拦腰砍断 JSON 结构 + 可能劈开 UTF-16 代理对）。
            // 结构化降级：重建精简载荷，data 换前缀预览，永远合法可解析。
            // N8（P3）：限幅改 UTF-8 字节口径——中文每字最多 3 字节，原按字符数限 512K 时线上实际可达 1MB+。
            const int LimitBytes = 512 * 1024;
            if (System.Text.Encoding.UTF8.GetByteCount(text) > LimitBytes)
            {
                var preview = BuildResultText(payload, tool);
                return preview;
            }
            return text;
        }

        /// <summary>R4：超限时重建合法 JSON 载荷——data 降为前缀预览 + truncated 标记。hint 按工具 schema 动态生成（Y2-2：
        /// 全局 hint 对 verify_chart 等无 max_cells/address 的工具是无 效指引；图片类工具给专属建议）。</summary>
        internal static string BuildResultText(JObject payload) => BuildResultText(payload, null);

        private static string BuildResultText(JObject payload, ToolDescriptor tool)
        {
            const int PreviewLimitBytes = 200 * 1024; // N8：UTF-8 字节口径
            var slim = new JObject
            {
                ["ok"] = payload["ok"],
                ["transport_truncated"] = true,
                ["hint"] = BuildTruncationHint(tool)
            };
            if (payload["error_code"] != null) { slim["error_code"] = payload["error_code"]; }
            // Y2-1 修复：warning 与 data 同规格限幅；N8：均按 UTF-8 字节
            if (payload["warning"] != null)
            {
                var warningText = payload["warning"].ToString(Formatting.None);
                if (System.Text.Encoding.UTF8.GetByteCount(warningText) > PreviewLimitBytes)
                {
                    slim["warning"] = TrimToUtf8Bytes(warningText, PreviewLimitBytes) + "...[已截断]";
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
                if (System.Text.Encoding.UTF8.GetByteCount(dataText) > PreviewLimitBytes)
                {
                    slim["data_preview"] = TrimToUtf8Bytes(dataText, PreviewLimitBytes) + "...[已截断]";
                }
                else
                {
                    slim["data"] = data;
                }
            }

            return slim.ToString(Formatting.Indented);
        }

        /// <summary>按工具 schema 生成截断后的自愈指引（Y2-2）：全局 hint 对无 max_cells/address 的工具无效。</summary>
        // R20：出站串行化——worker 线程与读循环（ping/initialize）会同时写 stdout，
        // 不锁会出现两帧字节交错，客户端直接 JSON 解析失败。
        private static void Emit(Stream stdout, JObject message)
        {
            var bytes = Encoding.UTF8.GetBytes(message.ToString(Formatting.None) + "\n");
            lock (_stdoutLock)
            {
                stdout.Write(bytes, 0, bytes.Length);
                stdout.Flush();
            }
        }
        /// <summary>Dangerous 拦截文案（Y2-6）：参数回显截断 500 字符防上下文白烧/敏感值复述；声明 MCP 模式不可撤销。</summary>
        /// <summary>P0-3 确认门强制力：生成一次性确认码文件（人工可读），返回拦截文案。
        /// 软门口径：码只存在于文件中，人工读码后由客户端回传 confirmed=<code>；具备文件读取能力的宿主仍可自答，硬门用 --confirm-mode stderr（默认）。</summary>
        /// <summary>M27R3：清扫待确认目录——旧命名（文件名内嵌码）一律删，其余 >24h 删。启动与首调各跑一次。</summary>
        internal static void SweepStaleConfirmationFiles()
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "pending-confirmations", "mcp");
                if (!Directory.Exists(dir)) { return; }
                var cutoff = DateTime.UtcNow.AddHours(-24);
                foreach (var stale in Directory.GetFiles(dir, "*.txt"))
                {
                    var name = Path.GetFileName(stale);
                    var legacy = name.EndsWith("-pending.txt", StringComparison.OrdinalIgnoreCase) == false;
                    try
                    {
                        if (legacy || File.GetLastWriteTimeUtc(stale) < cutoff) { File.Delete(stale); }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>M28 审查（第三轮）：确认码绑定参数集（排序键 + JSON 值 → SHA256），防"确认 A 执行 B"。</summary>
        internal static string ConfirmationArgsHash(Dictionary<string, object> args)
        {
            var parts = new List<string>();
            if (args != null)
            {
                foreach (var kv in args.Where(p => p.Key != "confirmed").OrderBy(p => p.Key, StringComparer.Ordinal))
                {
                    parts.Add(kv.Key + "=" + Newtonsoft.Json.JsonConvert.SerializeObject(kv.Value));
                }
            }

            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(string.Join("&", parts)));
                return System.BitConverter.ToString(bytes).Replace("-", string.Empty);
            }
        }

        /// <summary>M28 审查（AI 使用方）：人话之外附一行无码 JSON，便于模型机读分支（与 CLI 字段同名）。</summary>
        internal static string ConfirmationMachineLine(string name)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(new
            {
                ok = false,
                error_code = "CONFIRMATION_REQUIRED",
                requires_confirmation = true,
                tool = name,
                next_action = "ask_user_to_read_confirmation",
                confirmed_arg = "confirmed"
            });
        }

        internal static string BuildConfirmationText(string name, Dictionary<string, object> args)
        {
            // M27R3（闭环实测抓到的真缺陷）：码不匹配/已耗尽时**绝不重发新码**——
            // 否则每次错码都会重置 Failures，「连错 5 次作废」形同虚设（实测 5 次错码后正确码仍能通过）。
            // M28 审查（第三轮）：参数与确认码不匹配——不重发新码，要求按原参数重新发起
            if (ConfirmationFailure == "args-mismatch")
            {
                return Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error_code = "CONFIRMATION_ARGS_MISMATCH",
                    data = "确认码与本次调用的参数集不一致：人工确认的是原参数，旧码不能用于执行新参数。请按原参数重新发起该危险操作。"
                }, Newtonsoft.Json.Formatting.None);
            }

            if (ConfirmationFailure != null && (ConfirmationFailure.StartsWith("mismatch") || ConfirmationFailure == "exhausted"))
            {
                var exhausted = ConfirmationFailure == "exhausted";
                var remaining = exhausted ? 0 : int.Parse(ConfirmationFailure.Substring(9));
                return Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error_code = exhausted ? "CONFIRMATION_EXHAUSTED" : "CONFIRMATION_CODE_INVALID",
                    data = exhausted
                        ? "确认码连续错误已达上限，本次待确认记录已作废。请重新发起该危险操作以生成新的确认文件。"
                        : ("确认码不正确，还可尝试 " + remaining + " 次；连续 5 次错误将作废本次待确认记录。"),
                    remaining_attempts = remaining
                }, Newtonsoft.Json.Formatting.None);

            }

            // M27R3：累计错码≥5 的待确认记录不再重发新码（防「4 次错码 + 重发」无限刷）
            if (PendingConfirmations.TryGetValue(name, out var existingGate) && existingGate.Failures >= 5)
            {
                PendingConfirmations.TryRemove(name, out _);
                return Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    ok = false,
                    error_code = "CONFIRMATION_EXHAUSTED",
                    data = "该工具确认码连续错误已达上限，本次待确认记录已作废。请稍后重新发起该危险操作（建议先人工核对待办内容）。",
                    remaining_attempts = 0
                }, Newtonsoft.Json.Formatting.None);
            }

            var code = (Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")).Substring(0, 12).ToUpperInvariant(); // M27R2：8→12 hex（48bit），压缩穷举空间
            PendingConfirmations.AddOrUpdate(name,
                _ => new PendingConfirmation { Code = code, ExpiryUtc = DateTime.UtcNow.AddMinutes(10), Failures = 0, ArgsHash = ConfirmationArgsHash(args) },
                (_, old) => { old.Code = code; old.ExpiryUtc = DateTime.UtcNow.AddMinutes(10); old.ArgsHash = ConfirmationArgsHash(args); return old; }); // M27R3：保留 Failures（跨重发累计）；第三轮：绑定参数哈希
            var argsText = JsonConvert.SerializeObject(args);
            if (argsText.Length > 500) { argsText = argsText.Substring(0, 500) + "...[截断]"; }

            // M28 审查（第三轮）：MCP 与 CLI 分目录，避免同名工具确认文件互相覆盖/误读
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "pending-confirmations", "mcp");
            Directory.CreateDirectory(dir);
            // F5（M27 审查）：清理 24h 前的历史确认文件（含 M26 的 <tool>-<CODE>.txt 旧命名）——
            // 否则码文件与旧格式孤儿无限堆积，且旧码可能被误读。
            try
            {
                var staleCutoff = DateTime.UtcNow.AddHours(-24);
                foreach (var stale in Directory.GetFiles(dir, "*.txt"))
                {
                    try { if (File.GetLastWriteTimeUtc(stale) < staleCutoff) { File.Delete(stale); } } catch { }
                }
            }
            catch { }
            // M28 带外硬门：--confirm-mode stderr 时码只写 server 的 stderr（终端/宿主日志），
            // 不落盘、不进响应、不写 --log 文件——未把 stderr 转给模型的宿主上 AI 无法自答。
            if (string.Equals(ConfirmMode, "stderr", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    Console.Error.WriteLine("[JYYJ-MCP-CONFIRM] tool=" + name + " code=" + code + " args=" + argsText);
                    Console.Error.Flush();
                }
                catch { }

                return "该工具为 Dangerous 级（" + name + "），已拦截等待人工确认。" +
                       "确认码已输出到 server 的 stderr（启动它的终端/宿主日志中的 [JYYJ-MCP-CONFIRM] 行），不落盘、不出现在本响应中。\n" +
                       "请转告用户：从终端/宿主日志读取该行，用其中的 code 重新调用（confirmed=<码>）。确认码一次性，使用后作废。" +
                       "当前参数：" + argsText + "\n" + ConfirmationMachineLine(name);
            }

            // D1 收口（M27）：文件名不得内嵌确认码——响应里给的是文件路径，
            // 原名 <tool>-<CODE>.txt 会把码随路径泄露给 AI（可自答）。码只在文件内容。
            var file = Path.Combine(dir, name + "-pending.txt");
            File.WriteAllText(file,
                "危险操作确认\n" +
                "工具: " + name + "\n" +
                "确认码: " + code + "\n" +
                "时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\n" +
                "参数: " + argsText + "\n" +
                "\n确认方式: 人工核对影响范围后，在 arguments 中携带 \"confirmed\": \"" + code + "\" 重新调用。\n" +
                "该确认码一次性，使用后作废。");

            return "该工具为 Dangerous 级（" + name + "），已拦截等待人工确认。" +
                   "确认文件已写入待确认目录（%AppData%\\ExcelAiAssistant\\pending-confirmations\\mcp，文件名 " + name + "-pending.txt）——本响应不提供可直接读取的完整路径。\n" +
                   "请将此确认文件路径原样转告用户：用户核对后用文件中的确认码重新调用（confirmed=<用户提供的码>）。本响应不携带确认码。" +
                   "注意：确认码一次性；写操作已接入审计库，执行后可用 undo_excel_operation 回滚。" +
                   "当前参数：" + argsText + "\n" + ConfirmationMachineLine(name);
        }
        /// <summary>按工具 schema 生成截断后的自愈指引（Y2-2）：全局 hint 对无 max_cells/address 的工具无效。</summary>
        /// <summary>N8：按 UTF-8 字节预算截断字符串，不劈开代理对（回退到上一完整字符）。</summary>
        internal static string TrimToUtf8Bytes(string s, int maxBytes)
        {
            if (s == null) return "";
            if (System.Text.Encoding.UTF8.GetByteCount(s) <= maxBytes) return s;
            int chars = 0, bytes = 0;
            while (chars < s.Length)
            {
                char c = s[chars];
                int cb;
                if (char.IsHighSurrogate(c) && chars + 1 < s.Length && char.IsLowSurrogate(s[chars + 1])) cb = 4;
                else if (char.IsSurrogate(c)) cb = 3; // 孤立代理：.NET UTF8 编码为 U+FFFD（实测 3 字节），按 3 计才不超预算
                else if (c <= 0x7F) cb = 1;
                else if (c <= 0x7FF) cb = 2;
                else cb = 3;
                if (bytes + cb > maxBytes) break;
                bytes += cb;
                chars += (cb == 4 ? 2 : 1);
            }
            return s.Substring(0, chars);
        }

        private static string BuildTruncationHint(ToolDescriptor tool)
        {
            if (tool == null) { return "结果超限已降级为预览；请缩小查询范围或用分页参数（max_cells/offset 等）重取完整数据。"; }
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
    /// <summary>Win32 互操作（MCP server 自用：PID 归因 + 父进程看门狗）。</summary>
    internal static class Win32Native
    {
        [DllImport("user32")]
        public static extern uint GetWindowThreadProcessId(int hWnd, uint[] processId);

        public const uint TH32CS_SNAPPROCESS = 0x00000002;

        [DllImport("kernel32", SetLastError = true)]
        public static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, IntPtr th32ProcessID);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct PROCESSENTRY32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeFile;
        }

        // —— M28 Job Object：server 硬杀时内核终止其 Excel ——
        public const int JobObjectExtendedLimitInformation = 9;
        public const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;

        [StructLayout(LayoutKind.Sequential)]
        public struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string lpName);

        [DllImport("kernel32", SetLastError = true)]
        public static extern bool SetInformationJobObject(IntPtr hJob, int jobObjectInfoClass, IntPtr lpJobObjectInfo, uint cbJobObjectInfoLength);

        [DllImport("kernel32", SetLastError = true)]
        public static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);
    }
}