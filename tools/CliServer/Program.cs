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

namespace ExcelAiAssistant.Cli
{
    /// <summary>
    /// JYYJ助手 CLI（命令行直调工具面）：
    /// ExcelAiCli list                          —— 列出全部工具（名称/风险级/描述）
    /// ExcelAiCli exec &lt;tool&gt; [--key value ...] —— 执行工具，参数键值对
    /// ExcelAiCli exec &lt;tool&gt; --json '{...}'    —— 执行工具，JSON 参数体
    /// ExcelAiCli schema &lt;tool&gt;                 —— 查看工具参数 schema
    ///
    /// 安全语义与 MCP server 同源：默认只读（--allow-write 开放写工具）、Dangerous 需 --confirmed、
    /// 自建内存工作簿不碰用户文件、会话结束按 PID 强杀自建实例。
    /// 【已知缺口】与 MCP server 相同：未传 OperationRepository → 写操作无审计留痕与回滚（N3）。
    /// </summary>
    internal static class Program
    {
        private static ToolRegistry _registry;
        private static ExcelToolHost _host;
        private static OperationRepository _repo; // P0-1：审计/回滚落库
        private static object _app; // object 而非 dynamic：Cleanup 阶段 RCW 已断连，dynamic 的 null 比较也会走 DLR 重绑定抛 0x800706BA
        private static bool _ownsInstance; // M28 审查：仅自建且启动时间晚于 CLI 的实例允许 Visible/Add/Quit
        private static bool _jobAttached;  // M28 审查：Job 建立成功才允许运行（否则硬杀兜底缺失）
        private static string JobError;    // M28 审查（第三轮）：Job 失败阶段 + Win32Error（供启动门可执行报错）
        private static readonly List<int> SpawnedPids = new List<int>();
        private static readonly DateTime ServerStartUtc = DateTime.UtcNow;
        private static IntPtr _jobHandle; // M28：Excel 所属 Job（KILL_ON_JOB_CLOSE），防硬杀留孤儿

        private static int Main(string[] args)
        {
            if (args.Length == 0 || args[0] == "help" || args[0] == "--help" || args[0] == "/?")
            {
                PrintHelp();
                return 0;
            }

            bool allowWrite = args.Contains("--allow-write");
            var command = args[0];

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Console.Error.WriteLine("未找到 Excel.Application COM 组件（本机未安装 Excel？）");
                    return 5; // M28 审查：help 文档口径 5=COM 组件缺失，原实现误返回 2
                }

                // M28 审查（第三轮 #7）：激活前快照宿主 PID——归因到的 PID 若激活前已存在即为复用实例
                var hostPidsBefore = SnapshotHostPids();
                dynamic appLocal = Activator.CreateInstance(excelType);
                _app = appLocal;
                // PID 归因（MCP server R2 同款）：从自身 Hwnd 取确定 PID
                try
                {
                    var hwnd = (int)appLocal.Hwnd;
                    var pid = new uint[1];
                    Win32Native.GetWindowThreadProcessId(hwnd, pid);
                    if (pid[0] != 0) { SpawnedPids.Add((int)pid[0]); }
                }
                catch { }

                // M28 审查（审查者）：复用用户实例时后续 Visible=false/Workbooks.Add/Quit 会动用户工作簿——
                // 必须整体拒绝启动；只有「PID 可归因且启动时间晚于 CLI」的自有实例才继续。
                if (SpawnedPids.Count == 0)
                {
                    try { Marshal.FinalReleaseComObject(appLocal); } catch { }
                    _app = null;
                    Console.Error.WriteLine("无法归因 COM 实例 PID（Hwnd 不可用）；为避免误操作用户工作簿，拒绝启动。");
                    return 3;
                }

                if (hostPidsBefore != null && hostPidsBefore.Contains(SpawnedPids[0]))
                {
                    try { Marshal.FinalReleaseComObject(appLocal); } catch { }
                    _app = null;
                    Console.Error.WriteLine("COM 激活复用了本进程启动前已存在的宿主实例（PID=" + SpawnedPids[0] +
                        "，激活前快照命中）。为避免关闭用户正在编辑的工作簿，拒绝启动；请先关闭该 Excel/WPS 实例后重试。");
                    return 3;
                }

                try
                {
                    using (var probe = System.Diagnostics.Process.GetProcessById(SpawnedPids[0]))
                    {
                        if (probe.StartTime.ToUniversalTime() < ServerStartUtc)
                        {
                            try { Marshal.FinalReleaseComObject(appLocal); } catch { }
                            _app = null;
                            Console.Error.WriteLine("COM 激活复用了已存在的实例（PID=" + SpawnedPids[0] +
                                "，启动早于 CLI）。为避免关闭用户正在编辑的工作簿，拒绝启动；请先关闭该 Excel/WPS 实例后重试。");
                            return 3;
                        }
                    }
                }
                catch (Exception probeException)
                {
                    try { Marshal.FinalReleaseComObject(appLocal); } catch { }
                    _app = null;
                    Console.Error.WriteLine("COM 实例启动时间核验失败：" + probeException.Message + "；为避免误操作用户工作簿，拒绝启动。");
                    return 3;
                }

                _ownsInstance = true;
                appLocal.Visible = false;
                try { appLocal.DisplayAlerts = false; } catch { }
                try { appLocal.AutomationSecurity = 1; } catch { }
                appLocal.Workbooks.Add();
                TryAttachExcelJob(SpawnedPids.Count > 0 ? SpawnedPids[0] : 0); // M28：硬杀兜底，防 Ctrl+C/崩溃留孤儿 EXCEL
                if (!_jobAttached)
                {
                    // M28 审查：无 Job 则硬杀时必留孤儿 EXCEL——拒绝启动而不是降级（报错带失败阶段）
                    var detail = string.IsNullOrEmpty(JobError) ? "未知原因" : JobError;
                    Console.Error.WriteLine("Excel 的 KILL_ON_JOB_CLOSE Job 建立失败（" + detail +
                        "）：无法保证 Ctrl+C/崩溃时回收 EXCEL，拒绝启动。");
                    Cleanup();
                    return 3;
                }

                // P0-1：审计/回滚接入——独立 operations-cli.db
                var repoDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant");
                _repo = new OperationRepository(Path.Combine(repoDir, "operations-cli.db"));
                _host = new ExcelToolHost(appLocal, _repo, null, "cli-" + Guid.NewGuid().ToString("N").Substring(0, 8));
                _registry = _host.CreateRegistry();
                // M26：外部 AI 智能体桥接（workbuddy 记忆读写）
                AgentBridgeToolRegistration.Register(_registry, null);

                if (!allowWrite)
                {
                    foreach (var tool in _registry.Tools.Where(t => t.RiskLevel != RiskLevel.ReadOnly).ToList())
                    {
                        _registry.Remove(tool.Name);
                    }
                }
                else
                {
                    Console.Error.WriteLine("[写模式] 操作已审计入 operations-cli.db，可用 undo_excel_operation 回滚。");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("启动失败: " + ex.Message);
                Cleanup();
                return 3;
            }

            int exitCode;
            try
            {
                switch (command)
                {
                    case "list":
                        exitCode = CmdList();
                        break;
                    case "schema":
                        exitCode = args.Length < 2 ? Fail("用法：ExcelAiCli schema <tool>") : CmdSchema(args[1]);
                        break;
                    case "exec":
                        exitCode = args.Length < 2 ? Fail("用法：ExcelAiCli exec <tool> [--key value ... | --json '{...}']") : CmdExec(args, allowWrite);
                        break;
                    default:
                        exitCode = Fail("未知命令：" + command + "（可用：list / schema / exec / help）");
                        break;
                }
            }
            finally
            {
                Cleanup();
            }
            // 修复：工具执行链里大量 dynamic COM CallSite，进程退出时 finalizer 对已断连 RCW
            // 的重绑定会抛 0x800706BA 未捕获异常（污染退出码/输出）。显式断链 + Exit 强制收尾。
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Environment.Exit(exitCode);
            return exitCode;
        }

        private static int CmdList()
        {
            foreach (var tool in _registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase))
            {
                Console.WriteLine(tool.Name + "\t" + tool.RiskLevel + "\t" + tool.Description.Replace("\r", " ").Replace("\n", " ").Replace("\t", " "));
            }
            Console.Error.WriteLine("共 " + _registry.Tools.Count + " 个工具");
            return 0;
        }

        private static int CmdSchema(string name)
        {
            if (!_registry.TryGet(name, out var tool))
            {
                return Fail("未知工具：" + name + (IsReadOnlyMode() ? "（当前只读模式，写工具已下线）" : ""));
            }
            // Y3-3（与 MCP R3 同构）：剥属性定义层 name 路由键，业务参数 name 保留
            Console.WriteLine(JsonConvert.SerializeObject(CleanSchema(tool.Parameters ?? new Dictionary<string, object>()), Formatting.Indented));
            return 0;
        }

        private static int CmdExec(string[] args, bool allowWrite)
        {
            var name = args[1];
            if (!_registry.TryGet(name, out var tool))
            {
                return Fail("未知工具：" + name + (IsReadOnlyMode() ? "（当前只读模式，写工具已下线；--allow-write 可开放；运行 ExcelAiCli list 查看全部可用工具）" : ""));
            }

            // 参数解析：--json '{...}' 或 --key value 序列
            var dict = new Dictionary<string, object>();
            var globalSwitches = new HashSet<string> { "--confirmed", "--allow-write", "--legacy-confirm", "--confirm-code" }; // Y3-2：全局开关不进参数解析（吞键/污染键）
            var valueIdx = new HashSet<int>(); // Y3-4：value 占用位记录，--confirmed 判定只认 key 位
            for (var i = 2; i < args.Length; i++)
            {
                if (args[i] == "--confirm-code" && i + 1 < args.Length)
                {
                    valueIdx.Add(i + 1); // 确认码的值不是工具参数
                    i++;
                    continue;
                }

                if (globalSwitches.Contains(args[i])) { continue; }
                if (args[i] == "--json" && i + 1 < args.Length)
                {
                    // R3-1/Y3-5（与 MCP :236 同构）：MaxDepth 防深嵌套 SO 打穿 finally；
                    // DateParseHandling.None 保 ISO 日期串原样进工具层；畸形 JSON 回 exit 2 人话。
                    try
                    {
                        using var stringReader = new StringReader(args[i + 1]);
                        using var jsonReader = new JsonTextReader(stringReader) { DateParseHandling = DateParseHandling.None, MaxDepth = 64 };
                        dict = new JsonSerializer().Deserialize<Dictionary<string, object>>(jsonReader) ?? new Dictionary<string, object>();
                    }
                    catch (Exception jsonException)
                    {
                        Console.Error.WriteLine("参数 JSON 解析失败: " + jsonException.Message);
                            return 2; // FIX(M27R3 评审)：原 Environment.Exit(2) 会跳过 Main finally 的 Cleanup，泄漏已启动的 EXCEL 进程
                    }
                    break;
                }
                if (args[i].StartsWith("--"))
                {
                    // M28 审查：未知/缺值的 --key 与游离位置参数一律 fail-closed（原实现静默忽略，
                    // 工具会在缺参状态下执行，错误被推到工具层甚至误写）
                    if (i + 1 >= args.Length)
                    {
                        Console.Error.WriteLine("参数缺少值：" + args[i]);
                        return 2;
                    }
                    var key = args[i].Substring(2);
                    var value = args[i + 1];
                    valueIdx.Add(i + 1);
                    // Y3-1 数值无损转换守卫：TryParse 后回读比对原串，不等即保持字符串
                    //（修 001→1 / +7→7 / 1,234→1234 / Infinity→∞ / 9.2E+18 科学计数毁值）
                    if (long.TryParse(value, out var l) && Convert.ToString(l) == value)
                    {
                        dict[key] = l;
                    }
                    else if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var d) && Convert.ToString(d) == value)
                    {
                        dict[key] = d;
                    }
                    else
                    {
                        dict[key] = value;
                    }
                    i++;
                }
                else
                {
                    Console.Error.WriteLine("无法识别的参数：" + args[i] + "（参数请用 --key value 或 --json）");
                    return 2;
                }
            }

            // Dangerous 二次确认（M28 审查：与 MCP 同源的确认码门，裸 --confirmed 不再放行）
            var confirmed = false;
            string confirmCode = null;
            var legacyConfirm = args.Contains("--legacy-confirm");
            for (var i = 2; i < args.Length; i++)
            {
                if (valueIdx.Contains(i)) { continue; }
                if (args[i] == "--confirmed") { confirmed = true; }
                if (args[i] == "--confirm-code" && i + 1 < args.Length) { confirmCode = args[i + 1]; }
            }

            if (tool.RiskLevel == RiskLevel.Dangerous)
            {
                if (legacyConfirm && confirmed)
                {
                    Console.Error.WriteLine("[legacy-confirm] 已按旧 confirmed=true 语义放行（仅显式开关生效）。");
                }
                else if (confirmCode != null)
                {
                    var codeVerdict = ConsumePendingConfirmation(name, confirmCode, dict);
                    if (codeVerdict != 0)
                    {
                        Console.WriteLine(JsonConvert.SerializeObject(new
                        {
                            ok = false,
                            error_code = codeVerdict == 3 ? "CONFIRMATION_ARGS_MISMATCH" : "CONFIRMATION_CODE_INVALID",
                            requires_confirmation = true,
                            hint = codeVerdict == 2
                                ? "确认文件不存在或已过期（10 分钟）。请先不带 --confirm-code 执行一次以生成确认文件。"
                                : (codeVerdict == 3
                                    ? "确认码与本次参数集不一致：人工确认的是原参数，旧码不能用于新参数。请按原参数重新发起。"
                                    : "确认码不正确（一次性，错误不重发）。请核对确认文件中的码后重试。"),
                            tool = name
                        }, Formatting.Indented));
                        return 4;
                    }
                }
                else
                {
                    var code = WritePendingConfirmation(name, dict);
                    Console.Error.WriteLine("确认文件：" + PendingConfirmationPath(name));
                    Console.WriteLine(JsonConvert.SerializeObject(new
                    {
                        ok = false,
                        error_code = "CONFIRMATION_REQUIRED",
                        requires_confirmation = true,
                        hint = "该工具为 Dangerous 级。请让用户读取确认文件中的一次性确认码，再用 --confirm-code <码> 重新执行（--confirmed 布尔已不放行）。",
                        tool = name,
                        args = dict
                    }, Formatting.Indented));
                    return 4;
                }
            }

            // 协议层参数不透传工具层（与 MCP :783 同构）：--json 里夹带的 confirmed 不得进工具参数校验
            dict.Remove("confirmed");

            ToolResult result;
            try
            {
                result = tool.Execute(dict);
            }
            catch (Exception ex)
            {
                result = new ToolResult { Ok = false, ErrorCode = "CLI_EXECUTE_ERROR", Data = ex.Message };
            }

            object payload;
            try
            {
                payload = new
                {
                    ok = result.Ok,
                    data = result.Data,
                    warning = result.Warning,
                    affected_ranges = result.AffectedRanges,
                    error_code = result.ErrorCode,
                    truncated = result.Truncated
                };
            }
            catch
            {
                // Y3-8（与 MCP 同位兜底对齐）：Data 含不可序列化残留 RCW 时降级字符串
                payload = new { ok = result.Ok, data = Convert.ToString(result.Data), error_code = result.ErrorCode, serialization_fallback = true };
            }

            Console.WriteLine(JsonConvert.SerializeObject(payload, Formatting.Indented));
            return result.Ok ? 0 : 1;
        }

        /// <summary>Y3-3：schema 清洗——剥属性定义对象内的 name 路由键（与 MCP CleanSchemaNode 同构简化版）。</summary>
        private static Dictionary<string, object> CleanSchema(Dictionary<string, object> node)
        {
            var result = new Dictionary<string, object>();
            var isContainer = node.ContainsKey("properties");
            foreach (var kv in node)
            {
                if (kv.Key == "properties" && kv.Value is Dictionary<string, object> props)
                {
                    var cleanedProps = new Dictionary<string, object>();
                    foreach (var pk in props)
                    {
                        if (pk.Value is Dictionary<string, object> def)
                        {
                            var cleanedDef = new Dictionary<string, object>();
                            foreach (var dk in def)
                            {
                                if (dk.Key == "name") { continue; } // 属性定义层路由键剥除
                                cleanedDef[dk.Key] = dk.Value;
                            }
                            cleanedProps[pk.Key] = cleanedDef;
                        }
                        else
                        {
                            cleanedProps[pk.Key] = pk.Value;
                        }
                    }
                    result[kv.Key] = cleanedProps;
                }
                else if (!isContainer || kv.Key != "name")
                {
                    result[kv.Key] = kv.Value;
                }
            }
            return result;
        }

        private static bool IsReadOnlyMode() => _registry != null && _registry.Tools.All(t => t.RiskLevel == RiskLevel.ReadOnly);

        private static int Fail(string message)
        {
            Console.Error.WriteLine(message);
            return 2;
        }

        private static void PrintHelp()
        {
            Console.WriteLine(@"JYYJ助手 CLI —— 命令行直调工具面（工具数以 list 输出为准；默认只读，--allow-write 开放写工具）
用法：
  ExcelAiCli list                              列出全部可用工具
  ExcelAiCli schema <tool>                     查看工具参数 schema
  ExcelAiCli exec <tool> [--key value ...]     执行工具
  ExcelAiCli exec <tool> --json '{""a"":1}'     JSON 参数体执行
选项：
  --allow-write       开放写工具（默认只读）
  --confirm-code <码>  Dangerous 级一次性确认码：先不带该参数执行一次，从确认文件读出码，再带码执行（码绑定参数集、10 分钟有效、使用即作废）
  --legacy-confirm    仅旧脚本兼容：配合 --confirmed 使用（不安全，AI 可自答，默认拒绝裸 --confirmed）
示例：
  ExcelAiCli list
  ExcelAiCli exec get_current_excel_context
  ExcelAiCli exec preview_excel_range --address A1:C10
  ExcelAiCli exec write_excel_range --sheet Sheet1 --address A1:B2 --json <二维数组JSON> --allow-write
  ExcelAiCli exec delete_excel_worksheet --sheet Sheet1 --allow-write            # 首次：生成确认文件并返回待确认
  ExcelAiCli exec delete_excel_worksheet --sheet Sheet1 --allow-write --confirm-code <码>  # 二次：人工读码后执行
退出码：0=成功 1=工具执行失败 2=用法/参数错误 3=启动失败 4=Dangerous 待确认/确认码无效 5=COM 组件缺失");
        }

        /// <summary>M28 审查（第三轮 #7）：激活前宿主进程（EXCEL/et/wps）PID 快照；失败返回 null（退化 StartTime 判定）。</summary>
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

        /// <summary>M28：CLI 被硬杀（Ctrl+C/任务管理器/崩溃）时 finally 不执行——Excel 挂 Job 由内核回收。
        /// M28 审查（第三轮）：失败原因写入 JobError（阶段 + Win32Error），启动门据此给出可执行报错。</summary>
        private static void TryAttachExcelJob(int excelPid)
        {
            if (excelPid <= 0) { JobError = "PID 无效"; return; }
            // M28 审查：与 Cleanup 同口径的启动时间守卫——防 WPS 单实例复用用户实例后被 Job 误杀
            try
            {
                using (var probe = System.Diagnostics.Process.GetProcessById(excelPid))
                {
                    if (probe.StartTime.ToUniversalTime() < ServerStartUtc)
                    {
                        JobError = "启动时间早于 CLI（疑似用户实例）";
                        return;
                    }
                }
            }
            catch (Exception probeException)
            {
                JobError = "启动时间核验失败（" + probeException.GetType().Name + "）";
                return;
            }

            var job = IntPtr.Zero;
            try
            {
                job = Win32Native.CreateJobObject(IntPtr.Zero, null);
                if (job == IntPtr.Zero) { JobError = "CreateJobObject Win32Error=" + Marshal.GetLastWin32Error(); return; }
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
                        Win32Native.CloseHandle(job);
                        return;
                    }
                }

                _jobHandle = job;
                _jobAttached = true;
            }
            catch (Exception jobException)
            {
                JobError = "Job 附加异常：" + jobException.Message;
                if (job != IntPtr.Zero) { try { Win32Native.CloseHandle(job); } catch { } }
            }
        }

        // ===================== M28 确认码门（与 MCP 同源语义） =====================

        private static string PendingConfirmationDir()
        {
            // M28 审查（第三轮）：CLI 与 MCP 分目录，避免同名工具确认文件互相覆盖
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "pending-confirmations", "cli");
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static string PendingConfirmationPath(string tool) => Path.Combine(PendingConfirmationDir(), tool + "-pending.txt");

        /// <summary>生成一次性确认码并写确认文件（码只落文件，不回显到响应）。</summary>
        private static string WritePendingConfirmation(string tool, Dictionary<string, object> args)
        {
            var code = (Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")).Substring(0, 12).ToUpperInvariant();
            var argsText = JsonConvert.SerializeObject(args);
            File.WriteAllText(PendingConfirmationPath(tool),
                "危险操作确认" + Environment.NewLine +
                "工具: " + tool + Environment.NewLine +
                "确认码: " + code + Environment.NewLine +
                "参数摘要: " + ConfirmationArgsHash(args) + Environment.NewLine +
                "时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine +
                "参数: " + argsText + Environment.NewLine +
                Environment.NewLine +
                "确认方式: 人工核对影响范围后执行：ExcelAiCli exec " + tool + " ... --confirm-code <确认码> --allow-write" + Environment.NewLine +
                "该确认码一次性，10 分钟内有效。");
            return code;
        }

        /// <summary>校验并消费确认码：0=通过；1=码不符；2=文件缺失/过期。</summary>
        private static string ConfirmationArgsHash(Dictionary<string, object> args)
        {
            var parts = new List<string>();
            foreach (var kv in args.Where(p => p.Key != "confirmed").OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                parts.Add(kv.Key + "=" + JsonConvert.SerializeObject(kv.Value));
            }

            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(string.Join("&", parts)));
                return BitConverter.ToString(bytes).Replace("-", string.Empty);
            }
        }

        private static int ConsumePendingConfirmation(string tool, string supplied, Dictionary<string, object> args)
        {
            var file = PendingConfirmationPath(tool);
            if (!File.Exists(file))
            {
                return 2;
            }

            if (DateTime.UtcNow - File.GetLastWriteTimeUtc(file) > TimeSpan.FromMinutes(10))
            {
                try { File.Delete(file); } catch { }
                return 2;
            }

            string expected = null;
            string expectedArgsHash = null;
            foreach (var line in File.ReadAllLines(file))
            {
                var codeIndex = line.IndexOf("确认码:", StringComparison.Ordinal);
                if (codeIndex >= 0)
                {
                    expected = line.Substring(codeIndex + 4).Trim();
                }

                var hashIndex = line.IndexOf("参数摘要:", StringComparison.Ordinal);
                if (hashIndex >= 0)
                {
                    expectedArgsHash = line.Substring(hashIndex + 5).Trim();
                }
            }

            if (string.IsNullOrEmpty(expected))
            {
                return 2;
            }

            // M28 审查（第三轮）：确认码绑定参数集——参数变了必须重新发起
            if (!string.IsNullOrEmpty(expectedArgsHash) &&
                !string.Equals(expectedArgsHash, ConfirmationArgsHash(args), StringComparison.Ordinal))
            {
                return 3;
            }

            if (!string.Equals(expected, (supplied ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            try { File.Delete(file); } catch { }
            return 0;
        }

        private static void Cleanup()
        {
            // 修复：先杀进程后 Quit 会让 COM 断连，dynamic Quit 的 CallSite 重绑定抛 0x800706BA。
            // 改反射调用（无 CallSite 缓存）+ MCP server 同款顺序：先优雅 Quit → 按 PID 强杀 → 释放 RCW。
            // M28 审查：只对自有实例 Quit（复用用户实例时 _ownsInstance=false）。
            if (_ownsInstance && _app != null && Marshal.IsComObject(_app))
            {
                try { _app.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, _app, null); }
                catch { }
            }

            // M28（第三轮 #7）：Kill 前加"激活前快照"强证据——复用实例绝不 Kill（与 MCP Cleanup 同口径）
            var snapshotAtExit = SnapshotHostPids();
            foreach (var pid in SpawnedPids)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(pid);
                    if (snapshotAtExit != null && snapshotAtExit.Contains(pid) && !_ownsInstance)
                    {
                        p.Dispose();
                        continue;
                    }

                    // FIX(M27R3)：Local 与 Utc 直接比会因 Kind 被忽略而恒真（UTC+8 实测 103/103 恒真）
                    if (p.StartTime.ToUniversalTime() >= ServerStartUtc)
                    {
                        p.Kill();
                    }
                    p.Dispose();
                }
                catch { }
            }

            if (_app != null && Marshal.IsComObject(_app))
            {
                try { Marshal.FinalReleaseComObject(_app); } catch { }
                _app = null;
            }
        }
    }

    /// <summary>Win32 互操作（PID 归因 + Job 兜底）。</summary>
    internal static class Win32Native
    {
        [DllImport("user32")]
        public static extern uint GetWindowThreadProcessId(int hWnd, uint[] processId);

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

        [DllImport("kernel32", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hObject);
    }
}
