// 真机 harness：本机真实 Excel 上直接驱动 run_vba_snippet（用户报告的 AutoFitRows 场景），
// 验证修复后的行为：(1) 不崩溃不回传晦涩原始异常；(2) 返回结构化 ToolResult；
// (3) 宏环境受限时被分类为可操作的错误码（VBA_MACROS_DISABLED / VBA_RUN_ERROR）。
// 内置看门狗：避免独立 Excel 的 VBE 模态状态卡死整个进程。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyVbaFix
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyVbaFix real-machine run_vba_snippet ==");

            // 纯逻辑：把用户实录的“宏被禁用”异常消息喂给已部署的 ClassifyVbaRunError，
            // 证明该场景会被分类为 VBA_MACROS_DISABLED（不再返回晦涩 VBA_RUN_ERROR）。
            CheckClassifier("EN 宏禁用", "Cannot run the macro 'JYYJHelper.AutoFitRows'. The macro may not be available in this workbook or all macros may be disabled.",
                expected: "VBA_MACROS_DISABLED");
            CheckClassifier("ZH 宏禁用", "无法运行“JYYJHelper.AutoFitRowHeight”宏。可能是因为该宏在此工作簿中不可用，或者所有的宏都被禁用。",
                expected: "VBA_MACROS_DISABLED");
            CheckClassifier("普通运行时错误", "VBA 运行时错误 438：对象不支持此属性或方法",
                expected: "VBA_RUN_ERROR");

            VerifyAutoFit2dGuard(); // 场景2 纯逻辑护栏（无 Excel 依赖）
            VerifyPythonReal();     // 场景3 真机 run_python_code（无 Excel 依赖，复用同一注册表链路）

            // Excel COM 是 STA。必须让创建 app 与执行用例跑在同一个 STA 线程上，才能像真实 Addin 一样
            // 由线程泵送 VBA 错误弹窗（否则 hidden Excel 弹窗卡死跨线程 COM → 假性超时）。整体加硬看门狗。
            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(120000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 120s（可能 VBE 模态弹窗仍未被抑制，属需修复的缺陷）");
            }

            Console.WriteLine(_failed == 0 ? "== VerifyVbaFix ALL PASS ==" : "== VerifyVbaFix FAIL count=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

        private static void RunRealCases()
        {
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                wb = app.Workbooks.Add();

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true; // 开启受控执行，放行 run_vba_snippet
                var registry = host.CreateRegistry();

                // 场景1：用户报告的成功类写法（Rows/Columns 分别设置，无 2D .AutoFit）→ 应成功。
                // 证实在真实 Excel 上探针+宏路径能正常返回（out-of-proc 下仅成功路径可验证：
                // 任何 VBA 运行时错误在外部进程自动化 Excel 都会被 Excel 模态化，On Error 探针也压不住，
                // 故错误路径改由静态护栏(场景2)纯逻辑验证，而非真的去运行错误宏）。
                RunCase(registry, "场景1 正常 AutoFitRows(应成功)",
                    "Sub AutoFitRows()\n    Rows.AutoFit\nEnd Sub",
                    expectOk: true);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  harness 异常: " + ex.Message);
            }
            finally
            {
                try { if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) { app.Quit(); } } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
            }
        }

        /// <summary>场景2 纯逻辑：2D .Range(...).AutoFit 是运行必抛 1004 的确定缺陷写法，执行前硬闸门
        /// PreRunVbaGate 必须判阻断（error，引导改写），而 Rows/Columns 的 AutoFit / RowHeight=0 /
        /// 矩形 Range 经 Rows/Columns/EntireRow 中介调 AutoFit 等“安全改写”不得误伤。
        /// 之所以不真跑错误宏，是因此类错误在 out-of-proc 自动化 Excel 必然弹模态、无法验证；
        /// 闸门即生产实际拦截路径，用它验证即为“生产行为”验证。</summary>
        private static void VerifyAutoFit2dGuard()
        {
            var bad = new[] {
                "Sub T()\n    ws.Range(\"G12:K18\").AutoFit\nEnd Sub",
                "Sub T()\n    ActiveSheet.Range(\"A1:B10\").AutoFit\nEnd Sub",
                "Sub T()\n    Range(\"G12:K18\").AutoFit\nEnd Sub",
                "Sub T()\n    ws.Range(\"G12:K18\").Offset(2,3).AutoFit\nEnd Sub"
            };
            var good = new[] {
                "Sub T()\n    ws.Rows(\"12:18\").AutoFit\nEnd Sub",
                "Sub T()\n    ws.Columns(\"G:K\").AutoFit\nEnd Sub",
                "Sub T()\n    ws.Rows(\"12:18\").RowHeight = 0\nEnd Sub",
                "Sub T()\n    ws.Columns(\"G:K\").ColumnWidth = 0\nEnd Sub",
                "Sub T()\n    ws.Range(\"G12:K18\").Font.Bold = True\nEnd Sub",
                "Sub T()\n    ws.Range(\"G12:K18\").Rows(\"12:18\").AutoFit\nEnd Sub",
                "Sub T()\n    ws.Range(\"G12:K18\").EntireRow.AutoFit\nEnd Sub",
                "Sub T()\n    ws.Range(\"G12\").AutoFit\nEnd Sub"
            };

            var failed = false;
            foreach (var c in bad)
            {
                var gate = ScriptRunnerService.PreRunVbaGate(c);
                if (gate == null)
                {
                    Console.WriteLine($"FAIL  [护栏] 应拦截却放行: {OneLine(c)}");
                    failed = true;
                }
            }

            foreach (var c in good)
            {
                var gate = ScriptRunnerService.PreRunVbaGate(c);
                if (gate != null)
                {
                    Console.WriteLine($"FAIL  [护栏] 应放行却误伤: {OneLine(c)}");
                    failed = true;
                }
            }

            if (failed)
            {
                _failed++;
                Console.WriteLine("FAIL  场景2 2D-AutoFit 执行前闸门校验不通过");
            }
            else
            {
                Console.WriteLine("PASS  场景2 执行前闸门正确拦截矩形 AutoFit / 放行安全改写");
            }
        }

        private static string OneLine(string code)
        {
            return string.Join(" ⏎ ", code.Split('\n')).Trim();
        }

        /// <summary>诊断：完全手写一遍 worker 启动+写请求+超时同步读，模拟手动 PowerShell 场景，
        /// 判断是「Process 管道重定向在该进程失效」还是「PythonWorkerHost 内部编排」问题。</summary>
        private static void DiagManualWorker()
        {
            try
            {
                var worker = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xai_python_worker.py");
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = "-u \"" + worker + "\"",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = System.Text.Encoding.UTF8,
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };
                psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
                psi.EnvironmentVariables["PYTHONUTF8"] = "1";
                var p = System.Diagnostics.Process.Start(psi);
                var code = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("print('manual-observ-1')"));
                p.StandardInput.WriteLine("{\"id\":\"m1\",\"code\":\"" + code + "\",\"out\":65536}");
                p.StandardInput.Flush();
                var readTask = p.StandardOutput.ReadLineAsync();
                if (!readTask.Wait(6000))
                {
                    Console.WriteLine("[diag] ManualWorker → 6s 内无回应（管道读不通）");
                }
                else
                {
                    Console.WriteLine("[diag] ManualWorker → 回应: " + readTask.Result);
                }

                p.Kill();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[diag] ManualWorker → 异常 " + ex.Message);
            }
        }

        /// <summary>场景3 真机 run_python_code：经同一 ExcelToolHost.CreateRegistry 注册链路，在本机真实
        /// Python 常驻 worker 上验证——(a) 中文输出经 UTF-8 往返不乱码；(b) 失控大循环打印被输出上限截断且
        /// 快速返回（不撑爆内存、不满时超时）；(c) 交互式 input() 被禁用守卫拦截并快速失败（不再 30s 卡死）。
        /// 用看门狗兜底，避免 worker 异常把整个 harness 吊死。</summary>
        private static void VerifyPythonReal()
        {
            var host = new ExcelToolHost(new object());
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();

            // 诊断：直接驱动 PythonWorkerHost.TryExecute，与工具注册链路隔离，定位超时是发生在
            // worker 层还是 run_python_code 调度层。
            var dOk = PythonWorkerHost.TryExecute("print('direct-ok-观察')", "python", 8000, 65536,
                null, 1048576, null,
                out var d1, out var d2, out var d3, out _, out _, out var d4);
            Console.WriteLine($"[diag] TryExecute 直接调用 → okExec={dOk}, ok={d1}, exit={d4}, out='{Short(d2)}'");

            DiagManualWorker();

            PyRun(registry, "中文UTF8往返",
                "print('销量', sum([1,2,3]))\nprint('北京 上海 广州')",
                expect: true,
                assert: o => o.Contains("北京 上海 广州") && o.Contains("6"),
                extra: "中文输出不应乱码");

            PyRun(registry, "大循环打印截断",
                "for i in range(20000):\n    print(i)",
                expect: true,
                assert: o => o.Contains("（已截断）") && o.Length <= 8192 + 32,
                extra: "失控打印应被截断且快速返回");

            var sw = System.Diagnostics.Stopwatch.StartNew();
            PyRun(registry, "input()禁用不超时",
                "x = input('p')",
                expect: false,
                assert: o => o.Contains("input"),
                extra: "input() 应被禁用守卫拦截（快速返回而非 30s 超时）");
            sw.Stop();
            if (sw.ElapsedMilliseconds > 5000)
            {
                _failed++;
                Console.WriteLine($"FAIL  input()禁用 返回偏慢：{sw.ElapsedMilliseconds}ms（期望 <<5s）");
            }
        }

        private static void PyRun(ToolRegistry registry, string label, string code, bool expect, Func<string, bool> assert, string extra)
        {
            try
            {
                if (!registry.TryGet("run_python_code", out var tool))
                {
                    throw new InvalidOperationException("run_python_code 未注册");
                }

                var started = System.Diagnostics.Stopwatch.StartNew();
                var result = tool.Execute(new Dictionary<string, object>
                {
                    ["code"] = code,
                    ["timeout_ms"] = 15000,
                    ["output_limit_kb"] = 8
                });
                started.Stop();

                var err = $"{result.ErrorCode}: {Short(result.Data as string)}".Trim();
                var text = result.Ok ? PropStr(result) : (result.Data as string ?? "");
                if (result.Ok == expect && assert(text))
                {
                    Console.WriteLine($"PASS  [Py] {label} → {err} [{started.ElapsedMilliseconds}ms]");
                    return;
                }

                _failed++;
                Console.WriteLine($"FAIL  [Py] {label} → ok={result.Ok}（期望 {expect}），{err}；{extra} [{started.ElapsedMilliseconds}ms]");
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"FAIL  [Py] {label} → 异常 {ex.Message}；{extra}");
            }
        }

        private static string PropStr(ToolResult result)
        {
            try
            {
                var t = result.Data.GetType();
                var p = t.GetProperty("output");
                return p == null ? "?" : Convert.ToString(p.GetValue(result.Data, null));
            }
            catch
            {
                return "?";
            }
        }

        // 用例在 STA 线程上与 Excel 同线程执行（见 Main），模拟真实 Addin 环境；此处直接同步调用。
        private static void RunCase(ToolRegistry registry, string label, string snippet, bool expectOk)
        {
            ToolResult result;
            try
            {
                result = DoRun(registry, new Dictionary<string, object> { ["snippet"] = snippet });
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"FAIL  {label} → 异常 {ex.Message}");
                return;
            }

            if (result == null)
            {
                _failed++;
                Console.WriteLine($"FAIL  {label} → 返回 null");
                return;
            }

            var ok = result.Ok;
            var detail = ok
                ? (result.Data == null ? "ok" : Newtonsoft.Json.JsonConvert.SerializeObject(result.Data))
                : $"{result.ErrorCode}: {Short(result.Data as string)}";

            if (ok == expectOk)
            {
                Console.WriteLine($"PASS  {label} → {detail}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  {label} → expected ok={expectOk}, actual ok={ok}, {detail}");
            }
        }

        private static ToolResult DoRun(ToolRegistry registry, Dictionary<string, object> args)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                throw new InvalidOperationException("run_vba_snippet 未注册");
            }

            return tool.Execute(args);
        }

        private static void CheckClassifier(string label, string message, string expected)
        {
            var actual = ScriptRunnerService.ClassifyVbaRunError(new Exception(message));
            var pass = string.Equals(actual, expected, StringComparison.Ordinal);
            if (pass)
            {
                Console.WriteLine($"PASS  分类[{label}] -> {actual}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  分类[{label}] -> {actual}（期望 {expected}）");
            }
        }

        private static string Short(string s)
        {
            if (s == null)
            {
                return "(null)";
            }

            return s.Length <= 700 ? s : s.Substring(0, 700) + "…";
        }
    }
}