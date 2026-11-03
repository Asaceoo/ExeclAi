using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ExcelAiAssistant.Core.Models;
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
        private static object _app; // object 而非 dynamic：Cleanup 阶段 RCW 已断连，dynamic 的 null 比较也会走 DLR 重绑定抛 0x800706BA
        private static readonly List<int> SpawnedPids = new List<int>();
        private static readonly DateTime ServerStartUtc = DateTime.UtcNow;

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
                    return 2;
                }

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

                appLocal.Visible = false;
                try { appLocal.DisplayAlerts = false; } catch { }
                try { appLocal.AutomationSecurity = 1; } catch { }
                appLocal.Workbooks.Add();

                _host = new ExcelToolHost(appLocal);
                _registry = _host.CreateRegistry();

                if (!allowWrite)
                {
                    foreach (var tool in _registry.Tools.Where(t => t.RiskLevel != RiskLevel.ReadOnly).ToList())
                    {
                        _registry.Remove(tool.Name);
                    }
                }
                else
                {
                    Console.Error.WriteLine("[写模式] 写操作无操作审计/回滚（N3 已知缺口），重要文件请自行备份。");
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
            var globalSwitches = new HashSet<string> { "--confirmed", "--allow-write" }; // Y3-2：全局开关不进参数解析（吞键/污染键）
            var valueIdx = new HashSet<int>(); // Y3-4：value 占用位记录，--confirmed 判定只认 key 位
            for (var i = 2; i < args.Length; i++)
            {
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
                        Environment.Exit(2);
                    }
                    break;
                }
                if (args[i].StartsWith("--") && i + 1 < args.Length)
                {
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
            }

            // Dangerous 二次确认（MCP 同语义）：--confirmed 才放行
            var confirmed = false;
            for (var i = 2; i < args.Length; i++)
            {
                if (args[i] == "--confirmed" && !valueIdx.Contains(i)) { confirmed = true; break; }
            }

            if (tool.RiskLevel == RiskLevel.Dangerous && !confirmed)
            {
                Console.WriteLine(JsonConvert.SerializeObject(new
                {
                    ok = false,
                    requires_confirmation = true,
                    hint = "该工具为 Dangerous 级。确认影响范围后追加 --confirmed 重新执行。",
                    tool = name,
                    args = dict
                }, Formatting.Indented));
                return 4;
            }

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
            Console.WriteLine(@"JYYJ助手 CLI —— 命令行直调 156 工具面
用法：
  ExcelAiCli list                              列出全部可用工具
  ExcelAiCli schema <tool>                     查看工具参数 schema
  ExcelAiCli exec <tool> [--key value ...]     执行工具
  ExcelAiCli exec <tool> --json '{""a"":1}'     JSON 参数体执行
选项：
  --allow-write       开放写工具（默认只读）
  --confirmed         Dangerous 级工具确认放行
示例：
  ExcelAiCli list
  ExcelAiCli exec get_current_excel_context
  ExcelAiCli exec preview_excel_range --address A1:C10
  ExcelAiCli exec write_excel_range --sheet Sheet1 --address A1:B2 --json <二维数组JSON> --allow-write
  ExcelAiCli exec delete_excel_worksheet --sheet Sheet1 --confirmed --allow-write
退出码：0=成功 1=工具执行失败 2=用法/参数错误 3=启动失败 4=Dangerous 待确认 5=COM 组件缺失");
        }

        private static void Cleanup()
        {
            // 修复：先杀进程后 Quit 会让 COM 断连，dynamic Quit 的 CallSite 重绑定抛 0x800706BA。
            // 改反射调用（无 CallSite 缓存）+ MCP server 同款顺序：先优雅 Quit → 按 PID 强杀 → 释放 RCW。
            if (_app != null && Marshal.IsComObject(_app))
            {
                try { _app.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, _app, null); }
                catch { }
            }

            foreach (var pid in SpawnedPids)
            {
                try
                {
                    var p = System.Diagnostics.Process.GetProcessById(pid);
                    if (p.StartTime >= ServerStartUtc)
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

    /// <summary>Win32 互操作（PID 归因）。</summary>
    internal static class Win32Native
    {
        [DllImport("user32")]
        public static extern uint GetWindowThreadProcessId(int hWnd, uint[] processId);
    }
}
