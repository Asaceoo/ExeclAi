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
        private static bool _allowWrite;
        private static bool _ownsInstance; // M28 审查：仅「自建且启动时间晚于 server」的实例允许 Visible/Add/Quit
        // M28：确认门硬门选项——LegacyConfirm=兼容旧 confirmed=true（仅显式 --legacy-confirm 开启）；
        // ConfirmMode=file（默认，人工读确认文件）| stderr（带外：码只进 server stderr，不落盘、不进响应）
        internal static bool LegacyConfirm;
        internal static string ConfirmMode = "stderr"; // M28 审查（第三轮）：默认带外硬门（file 需显式开启）
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
            StartParentWatchdog(); // M28：宿主硬杀兜底（下方 finally 只覆盖温和退出）
            Log("MCP server：stdio 循环启动，Excel 后台预热中…");

            // R1 修复：stdio 主循环任何异常（客户端断连 IOException/畸形输入）都不得跳过 Cleanup——
            // 泄漏 Visible=false 的 EXCEL.EXE 是红线级故障。
            try
            {
                RunStdioLoop();
            }
            finally
            {
                Shutdown("stdio 结束");
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

            Log("Excel 预热完成：工具数=" + _registry.Tools.Count + "，模式=" + (_readOnlyMode ? "只读" : "读写") + "，Excel PID=" + _spawnedPid);
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
                        p.Kill();
                        Log("已强杀 spawn 进程 PID=" + pid + " name=" + procName);
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
            // FIX（M27R3）：未知工具返回**结构化 JSON**（原纯文本破坏「content[0].text 恒为 JSON」契约），
            // 同时保留 F2 语义：只读模式下说明模式限制、列出可用只读工具、明示请勿重试写操作。
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
        private static void Emit(Stream stdout, JObject message)
        {
            var bytes = Encoding.UTF8.GetBytes(message.ToString(Formatting.None) + "\n");
            stdout.Write(bytes, 0, bytes.Length);
            stdout.Flush();
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