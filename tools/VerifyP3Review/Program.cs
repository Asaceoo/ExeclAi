// P3 审查（r4 修复项）真机 harness：在本机真实 Excel + 真实 Python 上针对性验证四轮对抗性审查后落地修复的行为。
//  P3R-1 PreRunVbaGate 裸代码放行：无 Sub 包裹的裸代码 run_vba_snippet 可执行（写 A1），致命错误仍被拦截。
//  P3R-2 state_clear=true 无条件回 state_reset=true（persist 与否都兑现 D4 契约）；worker 重建后首次 persist 也回 true；
//         未清除且状态保留时回 false。
//  P3R-3 persist 命名空间跨工具共享：run_python_code persist 写变量 → run_python_excel persist 读到并写回 Excel 单元格，
//         验证两工具共享同一 worker（描述承诺的「变量互通」）。
//  P3R-4 InstallDeps 自检失败继续下一镜像：真机难以复现「pip 成功但 import 失败」，以代码审查+单测为准（此处仅打印状态）。
// Excel COM 是 STA：创建 app 与执行 Excel 用例必须同 STA 线程；整体加硬看门狗防 VBE 模态卡死。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyP3Review
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyP3Review real-machine (r4 fixes) ==");

            VerifyPythonStateReset(); // P3R-2 纯 Python（无 Excel 依赖）

            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(90000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 90s（可能 VBE 模态弹窗仍未被抑制）");
            }

            Console.WriteLine();
            Console.WriteLine("P3R-4  InstallDeps 镜像兜底：已代码审查确认（pip 成功但 import 自检失败 → 记录原因并继续下一镜像）；真机难以复现，由单测覆盖。");
            Console.WriteLine(_failed == 0 ? "== VerifyP3Review ALL PASS ==" : "== VerifyP3Review FAIL count=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

        // ==================== P3R-2 Python state_clear → state_reset ====================

        private static void VerifyPythonStateReset()
        {
            var host = new ExcelToolHost(new object());
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();

            // a：state_clear=true + persist=false → state_reset=true（修复核心：persist=false 也须兑现）
            Check(registry, "P3R-2a state_clear=true,persist=false → state_reset=true",
                code: "print('hi')", persist: false, stateClear: true, expectStateReset: true,
                expectOutput: _ => true);

            // b：state_clear=true + persist=true → state_reset=true
            Check(registry, "P3R-2b state_clear=true,persist=true → state_reset=true",
                code: "print('hi')", persist: true, stateClear: true, expectStateReset: true,
                expectOutput: _ => true);

            // c：2b 已重建 persist 命名空间，此处无 clear 写变量不触发重建 → state_reset=false 且变量保留
            Check(registry, "P3R-2c persist 写变量(无 clear) → state_reset=false",
                code: "shared_r4 = 88", persist: true, stateClear: false, expectStateReset: false,
                expectOutput: _ => true);
            Check(registry, "P3R-2c' persist 读变量 + 状态保留 → state_reset=false",
                code: "xai_result({'shared_r4': shared_r4})", persist: true, stateClear: false,
                expectStateReset: false,
                expectOutput: _ => true,
                expectResult: r => r is JObject jo && (int)jo["shared_r4"] == 88);

            // d：worker 重建后首次 persist → state_reset=true（状态丢失感知）
            PythonWorkerHost.ResetWorker();
            Check(registry, "P3R-2d worker 重建后首次 persist → state_reset=true",
                code: "print('rebuilt')", persist: true, stateClear: false, expectStateReset: true,
                expectOutput: _ => true);

            // e：state_clear 确实清空 xai_state（写→清→读为空）
            Check(registry, "P3R-2e 写 xai_state",
                code: "xai_state['k'] = 5", persist: false, stateClear: false, expectStateReset: false,
                expectOutput: _ => true);
            Check(registry, "P3R-2e' state_clear 后 xai_state 为空",
                code: "print('state=', xai_state)", persist: false, stateClear: true,
                expectStateReset: true,
                expectOutput: o => o.Contains("state= {}"));
        }

        // ==================== Excel（STA）用例：P3R-1 / P3R-3 ====================

        private static void RunRealCases()
        {
            dynamic app = null;
            dynamic wb = null;
            int excelPid = 0;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                excelPid = ExcelPidOf(app); // Quit 前抓 PID，供 finally 兜底强杀防僵尸
                wb = app.Workbooks.Add();

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                VerifyBareCodeGate(registry, wb);    // P3R-1
                VerifyCrossToolPersist(registry, wb); // P3R-3
                VerifyWorkflowChain(registry, wb);    // P3R-5 D3：run_workflow 真机串联
                VerifyVbaJsonRoundTrip(registry, wb); // P3R-6 开放问题3：序列化器特殊字符往返
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  harness 异常: " + ex.Message);
            }
            finally
            {
                // 清理：先关簿、Quit，再释放全部 RCW + 强制回收，最后按 PID 兜底强杀，
                // 防「Quit 后子 RCW 未放导致 Excel 进程僵尸化」的残留（项目记忆踩坑）。
                try { if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) { app.Quit(); } } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (wb != null && Marshal.IsComObject(wb)) Marshal.FinalReleaseComObject(wb); } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                KillExcelProcess(excelPid);
            }
        }

        private static int ExcelPidOf(dynamic app)
        {
            try
            {
                var hwnd = (IntPtr)Convert.ToInt64(app.Hwnd);
                if (hwnd == IntPtr.Zero) return 0;
                int pid;
                GetWindowThreadProcessId(hwnd, out pid);
                return pid;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>兜底强杀：Quit/RCW 释放后 Excel 进程仍存活（僵尸）时按 PID 强制终止。</summary>
        private static void KillExcelProcess(int pid)
        {
            if (pid <= 0) return;
            for (var i = 0; i < 5; i++)
            {
                try
                {
                    using (var p = System.Diagnostics.Process.GetProcessById(pid))
                    {
                        if (p.HasExited) return;
                        p.Kill();
                    }
                }
                catch
                {
                    return; // 进程已消失或无权访问
                }
                Thread.Sleep(200);
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        /// <summary>P3R-1：PreRunVbaGate 必须放行「无 Sub 包裹的裸代码」（自动包裹 RunSnippet），
        /// 但含致命错误（如对多行多列矩形调 .AutoFit）的裸代码仍须被拦截。</summary>
        private static void VerifyBareCodeGate(ToolRegistry registry, dynamic wb)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                _failed++;
                Console.WriteLine("FAIL  P3R-1 run_vba_snippet 未注册");
                return;
            }

            // a：裸代码（无过程声明）应被放行并执行 → A1=42
            var bare = tool.Execute(new Dictionary<string, object>
            {
                ["snippet"] = "Range(\"A1\").Value = 42"
            });
            if (bare.Ok)
            {
                dynamic ws = wb.Worksheets.Item[1];
                var v = Convert.ToDouble(ws.Cells.Item[1, 1].Value2);
                if (v == 42)
                {
                    Console.WriteLine("PASS  P3R-1a 裸代码放行并执行 → A1=" + v);
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  P3R-1a 裸代码执行了但 A1={v}（期望 42）");
                }
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-1a 裸代码被误拦: {bare.ErrorCode}: {Short(bare.Data as string)}");
            }

            // b：致命错误裸代码（矩形 AutoFit）必须被拦截（VBA_INVALID）
            var fatal = tool.Execute(new Dictionary<string, object>
            {
                ["snippet"] = "Range(\"A2:B3\").AutoFit"
            });
            if (!fatal.Ok && fatal.ErrorCode == "VBA_INVALID")
            {
                Console.WriteLine("PASS  P3R-1b 致命错误裸代码仍被拦截（VBA_INVALID）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-1b 致命错误裸代码未被拦截: ok={fatal.Ok} code={fatal.ErrorCode}: {Short(fatal.Data as string)}");
            }
        }

        /// <summary>P3R-3：run_python_code persist 写入变量 → run_python_excel persist 读取并写回 Excel 单元格，
        /// 验证两工具共享同一常驻 worker 的 persist 命名空间（描述承诺「变量互通」）。</summary>
        private static void VerifyCrossToolPersist(ToolRegistry registry, dynamic wb)
        {
            if (!registry.TryGet("run_python_code", out var pyTool) ||
                !registry.TryGet("run_python_excel", out var xlTool))
            {
                _failed++;
                Console.WriteLine("FAIL  P3R-3 run_python_code / run_python_excel 未注册");
                return;
            }

            var py = pyTool.Execute(new Dictionary<string, object>
            {
                ["code"] = "cross_tool_var = 99",
                ["persist"] = true,
                ["timeout_ms"] = 15000
            });
            if (!py.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-3 run_python_code persist 写入: {py.ErrorCode}: {Short(py.Data as string)}");
                return;
            }

            // 走 run_python_excel：读同一 worker 命名空间里的 cross_tool_var，写回 Excel D1。
            // 注意：不声明 write_ranges——PythonExcelService 的 Snapshot 会在 finally 对解析出的 sheet 调
            // FinalReleaseComObject，销毁按对象标识缓存的共享 sheet RCW（harness 随后读单元格必炸「COM 对象与其
            // 基础 RCW 分开」，见项目记忆）。本用例目标仅验证跨工具 persist 共享 + xlwings 真机写回，避开快照路径。
            var xl = xlTool.Execute(new Dictionary<string, object>
            {
                ["code"] = "import xlwings as xw\nsht = xw.books.active.sheets[0]\nsht.range(\"D1\").value = cross_tool_var\nxai_result({'cross_tool_var': cross_tool_var})",
                ["persist"] = true,
                ["timeout_ms"] = 30000
            });
            if (!xl.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-3 run_python_excel 读共享变量: {xl.ErrorCode}: {Short(xl.Data as string)}");
                return;
            }

            dynamic ws = wb.Worksheets.Item[1];
            var v = Convert.ToDouble(ws.Cells.Item[1, 4].Value2);
            if (v == 99)
            {
                Console.WriteLine("PASS  P3R-3 persist 命名空间跨工具共享（run_python_code → run_python_excel）→ D1=" + v);
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-3 共享变量写回 D1={v}（期望 99）");
            }
        }

        /// <summary>P3R-5（需求手册 3.4 D3 验收）：「Python 算→VBA 写」经 run_workflow 编排在真机闭环——
        /// step0 run_python_code 用 xai_result 返回结构化 result，step1 run_vba_snippet 用 {step0.result}
        /// 占位符把 result 注入 args 通道，XaiArg 读出 total 写回 Excel E1。</summary>
        private static void VerifyWorkflowChain(ToolRegistry registry, dynamic wb)
        {
            if (!registry.TryGet("run_python_code", out _) || !registry.TryGet("run_vba_snippet", out _))
            {
                _failed++;
                Console.WriteLine("FAIL  P3R-5 run_python_code / run_vba_snippet 未注册");
                return;
            }

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "Python算→VBA写(真机)",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "run_python_code",
                        Description = "Python 计算",
                        Arguments = new Dictionary<string, object>
                        {
                            { "code", "xai_result({'total': 42, 'avg': 21.5})" }
                        }
                    },
                    new WorkflowStep
                    {
                        Tool = "run_vba_snippet",
                        Description = "VBA 写回",
                        Arguments = new Dictionary<string, object>
                        {
                            { "snippet", "Sub F()\n Range(\"E1\").Value = XaiArg(\"total\")\nEnd Sub" },
                            { "args", "{step0.result}" }
                        }
                    }
                }
            };

            try
            {
                var result = executor.RunAsync(template, null, null, 0, CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (result.Completed)
                {
                    dynamic ws = wb.Worksheets.Item[1];
                    var v = Convert.ToDouble(ws.Cells.Item[1, 5].Value2);
                    if (v == 42)
                    {
                        Console.WriteLine("PASS  P3R-5 run_workflow 真机串联（Python 算→VBA 写）→ E1=" + v);
                    }
                    else
                    {
                        _failed++;
                        Console.WriteLine($"FAIL  P3R-5 run_workflow 串联写回 E1={v}（期望 42）");
                    }
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  P3R-5 run_workflow 未完成: FailedIndex={result.FailedIndex} " +
                        $"(step={result.StepAt(result.FailedIndex ?? 0)?.Tool}, " +
                        $"{result.StepAt(result.FailedIndex ?? 0)?.Result?.ErrorCode}: " +
                        $"{Short(result.StepAt(result.FailedIndex ?? 0)?.Result?.Data as string)})");
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-5 run_workflow 真机串联异常: {ex.Message}");
            }
        }

        /// <summary>P3R-6（需求手册开放问题3 + 验收总纲2）：「VBA XaiReturn → C# 结构化 result」的特殊字符
        /// 完整往返——中文/emoji/双引号/反斜杠/换行/Tab 经自研序列化器 XaiJson/XaiEsc 编码，C# 侧解析回 JToken
        /// 逐项比对，验证转义正确性（往返一致、不丢字符、类型不串）。</summary>
        private static void VerifyVbaJsonRoundTrip(ToolRegistry registry, dynamic wb)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                _failed++;
                Console.WriteLine("FAIL  P3R-6 run_vba_snippet 未注册");
                return;
            }

            // 诊断分支：先验证「VBA 直接把 emoji 写入单元格 → C# 读回」是否保留（隔离注入层 vs 序列化层）。
            var diagSnippet =
                "Sub T()\n" +
                "    Range(\"F2\").Value = \"emoji\U0001F600直写\"\n" +
                "End Sub";
            var diag = tool.Execute(new Dictionary<string, object> { ["snippet"] = diagSnippet });
            if (diag.Ok)
            {
                dynamic ws = wb.Worksheets.Item[1];
                var f2 = Convert.ToString(ws.Cells.Item[2, 6].Value2);
                var f2ok = Show(f2) == "emoji\U0001F600直写" ? "保留" : "已损坏";
                Console.WriteLine($"  >> 诊断: VBA 直写 emoji 后 C# 读回 F2={Show(f2)}（{f2ok}）");
            }
            else
            {
                Console.WriteLine($"  >> 诊断: VBA 直写 emoji 失败: {diag.ErrorCode}: {Short(diag.Data as string)}");
            }

            var snippet =
                "Sub T()\n" +
                "    XaiReturn Array(\"中文\", \"emoji\U0001F600\", \"引号\"\"测试\", \"反斜杠\\路径\", " +
                "\"行一\" & vbLf & \"行二\", \"Tab\" & vbTab & \"x\")\n" +
                "End Sub";

            var result = tool.Execute(new Dictionary<string, object> { ["snippet"] = snippet });
            if (!result.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-6 XaiReturn 执行: {result.ErrorCode}: {Short(result.Data as string)}");
                return;
            }

            var token = PyResult(result) as JArray;
            if (token == null || token.Count != 6)
            {
                _failed++;
                Console.WriteLine($"FAIL  P3R-6 result 结构: {(PyResult(result)?.ToString() ?? "(null)")}（期望 6 元素数组）");
                return;
            }

            // 严格比对仅针对系统代码页可表示字符（中文/引号/反斜杠/换行/Tab）——这直接验证开放问题3的
            // 「序列化器转义正确性」。emoji 属非 BMP 字符，VBE 模块源码在中文系统按 ANSI/GBK 存储，注入层
            // 即被替换为 ?（已在 run_vba_snippet 描述文档化），与序列化器无关（诊断分支已证实直写同样损坏）。
            string[] expect =
            {
                "中文", null, "引号\"测试", "反斜杠\\路径", "行一\n行二", "Tab\tx"
            };
            var allOk = true;
            for (var i = 0; i < 6; i++)
            {
                if (expect[i] == null) continue; // emoji 槽位：走下方文档化损坏确认
                var got = Convert.ToString(token[i]);
                if (got != expect[i])
                {
                    allOk = false;
                    _failed++;
                    Console.WriteLine($"FAIL  P3R-6 元素[{i}] 往返失真: 期望={Show(expect[i])} 实得={Show(got)}");
                }
            }

            if (!allOk)
            {
                return;
            }

            var emojiGot = Convert.ToString(token[1]);
            var emojiCorrupted = emojiGot != "emoji\U0001F600";
            Console.WriteLine($"  >> P3R-6 emoji 注入层损坏确认（文档化限制）: 实得={Show(emojiGot)}，"
                + (emojiCorrupted ? "非 BMP 字符被 ANSI/GBK 替换为 ?，与描述一致，非序列化器缺陷" : "意外保留"));
            if (allOk && emojiCorrupted)
            {
                Console.WriteLine("PASS  P3R-6 XaiReturn 特殊字符往返：中文/引号/反斜杠/换行/Tab 全部一致（序列化器转义正确），emoji 按文档损坏");
            }
        }

        // ==================== 工具封装 ====================

        private static void Check(ToolRegistry registry, string label, string code, bool persist, bool stateClear,
            bool expectStateReset, Func<string, bool> expectOutput, Func<JToken, bool> expectResult = null)
        {
            try
            {
                if (!registry.TryGet("run_python_code", out var tool))
                {
                    throw new InvalidOperationException("run_python_code 未注册");
                }

                var args = new Dictionary<string, object>
                {
                    ["code"] = code,
                    ["timeout_ms"] = 15000,
                    ["output_limit_kb"] = 8
                };
                if (persist) args["persist"] = true;
                if (stateClear) args["state_clear"] = true;

                var result = tool.Execute(args);
                var text = result.Ok ? PropStr(result) : (result.Data as string ?? "");
                var stateReset = result.Ok && StateResetOf(result);
                var resultVal = result.Ok ? PyResult(result) : null;

                var ok = result.Ok && stateReset == expectStateReset && expectOutput(text)
                    && (expectResult == null || (resultVal != null && expectResult(resultVal)));
                if (ok)
                {
                    Console.WriteLine($"PASS  {label} (state_reset={stateReset})");
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  {label} → ok={result.Ok} state_reset={stateReset}（期望 {expectStateReset}），" +
                        $"{result.ErrorCode}: {Short(result.Data as string)}");
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"FAIL  {label} → 异常 {ex.Message}");
            }
        }

        private static bool StateResetOf(ToolResult result)
        {
            try
            {
                var t = result.Data.GetType();
                var p = t.GetProperty("state_reset");
                return p != null && Convert.ToBoolean(p.GetValue(result.Data, null));
            }
            catch
            {
                return false;
            }
        }

        private static JToken PyResult(ToolResult result)
        {
            try
            {
                var t = result.Data.GetType();
                var p = t.GetProperty("result");
                return p?.GetValue(result.Data, null) as JToken;
            }
            catch
            {
                return null;
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

        private static string Show(string s)
        {
            if (s == null) return "(null)";
            return s.Replace("\\", "\\\\").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        private static string Short(string s)
        {
            if (s == null) return "(null)";
            return s.Length <= 700 ? s : s.Substring(0, 700) + "…";
        }
    }
}
