// P3 真机 harness：在本机真实 Excel + 真实 Python 上验证「脚本编排与状态」（2.0.0.0 需求手册 3.4）：
//  D1 Python 持久状态：persist 命名空间跨调用、xai_state 常驻、非 persist 隔离、state_clear 显式清空、不落盘。
//  D2 VBA 持久状态：XaiStateSet/XaiStateGet 跨调用保留、state_clear 显式清空。
//  D3 多步串联：Python 算（结构化 result）→ VBA 写（args 通道）真实打通。
//  D4 状态护栏：显式清除生效；状态仅存内存不落盘（无状态文件残留）；超时重建后 state_reset 提示可感知。
// Excel COM 是 STA：创建 app 与执行 Excel 用例必须同 STA 线程；整体加硬看门狗防 VBE 模态卡死。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyP3
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyP3 real-machine 编排与状态 ==");

            VerifyPythonState();  // D1/D4 纯 Python（无 Excel 依赖）

            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(60000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 180s（可能 VBE 模态弹窗仍未被抑制，属需修复的缺陷）");
            }

            Console.WriteLine(_failed == 0 ? "== VerifyP3 ALL PASS ==" : "== VerifyP3 FAIL count=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

        // ==================== P3-D1/D4 纯 Python 状态 ====================

        private static void VerifyPythonState()
        {
            var host = new ExcelToolHost(new object());
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();

            // D1-a：xai_state 跨调用常驻（persist 与否都提供），首次先 state_clear 归零再写。
            PyRunOpt(registry, "D1 xai_state 跨调用常驻",
                "xai_state['x'] = 1", persist: false, stateClear: true, expect: true,
                assert: _ => true, resultAssert: null, extra: "写状态");
            PyRunOpt(registry, "D1 xai_state 跨调用读取",
                "print('x=', xai_state.get('x'))", persist: false, stateClear: false, expect: true,
                assert: o => o.Contains("x= 1"), resultAssert: null, extra: "未开 persist 也应读到 x=1");

            // D1-b：persist 命名空间跨调用复用；首调用带 state_clear 归零避免历史残留。
            PyRunOpt(registry, "D1 persist 命名空间复用(写)",
                "y = 42", persist: true, stateClear: true, expect: true,
                assert: _ => true, resultAssert: null, extra: "persist 写 y");
            PyRunOpt(registry, "D1 persist 命名空间复用(读)",
                "xai_result({'y': y})", persist: true, stateClear: false, expect: true,
                assert: _ => true, resultAssert: r => r is ToolResult tr && PyResult(tr) is JObject jo && (int)jo["y"] == 42,
                extra: "persist 第二次应读到 y=42 且经 xai_result 带回结构化 result");

            // D1-c：非 persist 每次全新命名空间，y 不泄漏。
            PyRunOpt(registry, "D1 非persist 隔离(写)",
                "y = 5", persist: false, stateClear: false, expect: true,
                assert: _ => true, resultAssert: null, extra: "非 persist 写 y");
            PyRunOpt(registry, "D1 非persist 隔离(读)",
                "print('y_in_globals=', 'y' in globals())", persist: false, stateClear: false, expect: true,
                assert: o => o.Contains("y_in_globals= False"), resultAssert: null, extra: "非 persist 第二次不应看到 y");

            // D4-a：state_clear 显式清空持久状态（命名空间 + xai_state 一并归零）。
            PyRunOpt(registry, "D4 state_clear 显式清空(写)",
                "y = 5\nxai_state['k'] = 9", persist: true, stateClear: false, expect: true,
                assert: _ => true, resultAssert: null, extra: "persist 写入 y 与 xai_state.k");
            PyRunOpt(registry, "D4 state_clear 显式清空(读)",
                "print('y_in=', 'y' in globals(), 'state=', xai_state)", persist: true, stateClear: true, expect: true,
                assert: o => o.Contains("y_in= False") && o.Contains("state= {}"), resultAssert: null,
                extra: "state_clear 后 y 与 xai_state 都应为空");

            // D4-b：状态仅存内存，不落盘——harness 运行前后状态文件应不存在。
            var stateFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xai_python_state.json");
            if (System.IO.File.Exists(stateFile))
            {
                _failed++;
                Console.WriteLine("FAIL  D4 状态文件不应落盘: " + stateFile);
            }
            else
            {
                Console.WriteLine("PASS  D4 Python 状态不落盘（无状态文件残留）");
            }
        }

        // ==================== Excel（STA）用例：D2 / D3 / D4 ====================

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
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                // 诊断：先验证「连续多次平凡 run_vba_snippet 是否第二次就挂」（排除状态模块干扰）。
                // 平凡调用#1 是本工作簿真正的首个 run_vba_snippet：应新建状态模块 → state_reset=true。
                var diag1 = VbaRunWithResult(registry, "DIAG 平凡调用#1",
                    "Sub T()\n Range(\"A1\").Value = 1\nEnd Sub", null, true, "诊断");
                Console.WriteLine("  >> 平凡调用#1 返回，准备平凡调用#2");
                var diag2 = VbaRunWithResult(registry, "DIAG 平凡调用#2",
                    "Sub T()\n Range(\"A2\").Value = 2\nEnd Sub", null, true, "诊断");
                Console.WriteLine("  >> 平凡调用#2 返回，准备状态用例");
                VerifyStateResetSemantics(diag1, diag2); // D2/D4：首个调用 state_reset=true，随后 false

                VerifyVbaState(registry, wb);  // D2 + D4(c)
                VerifyPythonToVbaChain(registry, wb); // D3
                VerifyVbaModuleHygiene(wb); // D2：每轮运行模块无残留
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

        /// <summary>P3-D2/D4(c)：VBA 持久状态——独立模块 JYYJXaiState 的模块级变量跨 run_vba_snippet 调用保留；
        /// XaiStateSet 写入一次，下次调用 XaiStateGet 仍可读到；state_clear=true 后读取为空。</summary>
        private static void VerifyVbaState(ToolRegistry registry, dynamic wb)
        {
            // 先清空，保证本用例从零状态开始，不依赖历史上残留的 gXaiState。
            VbaRun(registry, "D2 先清空状态", "Sub T()\n XaiStateClear\nEnd Sub", null, true, "前置清空");

            VbaRun(registry, "D2 写入状态",
                "Sub T()\n XaiStateSet \"k\", 7\nEnd Sub", null, true, "XaiStateSet k=7");

            // 下次调用读回并写入单元格 B1，证明跨调用保留。
            VbaRun(registry, "D2 跨调用读回状态",
                "Sub T()\n Range(\"B1\").Value = XaiStateGet(\"k\")\nEnd Sub", null, true,
                "XaiStateGet 应读到 k=7 并写入 B1");

            dynamic ws = wb.Worksheets.Item[1];
            var v = Convert.ToDouble(ws.Cells.Item[1, 2].Value2);
            if (v == 7)
            {
                Console.WriteLine("PASS  D2 VBA 持久状态跨调用保留 → B1=" + v);
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  D2 VBA 持久状态跨调用保留 → B1={v}（期望 7）");
            }

            // D4(c)：state_clear=true 显式清空 VBA 状态后，读回应为空（Null → 0）。写入 B2、回读 B2。
            VbaRun(registry, "D4 VBA state_clear 清空",
                "Sub T()\n Range(\"B2\").Value = XaiStateGet(\"k\")\nEnd Sub", null, true, "先执行，随后 state_clear 重跑",
                stateClear: true);
            var v2 = Convert.ToDouble(ws.Cells.Item[2, 2].Value2);
            _ = ws;
            if (v2 == 0)
            {
                Console.WriteLine("PASS  D4 VBA state_clear 显式清空 → B1 归零");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  D4 VBA state_clear 显式清空 → B1={v2}（期望 0）");
            }
        }

        /// <summary>P3-D2/D4：state_reset 语义——工作簿首个 run_vba_snippet（新建状态模块）回 state_reset=true，
        /// 后续调用（状态模块已存在）回 false。这是「超时/升级重建后状态被重置」的可感知提示的 VBA 侧实现。</summary>
        private static void VerifyStateResetSemantics(ToolResult first, ToolResult second)
        {
            var r1ok = first != null && StateResetOf(first);
            var r2ok = second != null && !StateResetOf(second);
            if (r1ok && r2ok)
            {
                Console.WriteLine("PASS  D2/D4 state_reset 语义：首个调用=true（状态模块新建），随后=false（状态保留）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  D2/D4 state_reset 语义：首个调用应 true 实得 {first != null && StateResetOf(first)}，" +
                    $"后续调用应 false 实得 {second != null && StateResetOf(second)}");
            }
        }

        /// <summary>P3-D2 模块卫生：经过多次 run_vba_snippet 后，工作簿里不得残留任何 JYYJRun_*/JYYJHelper 运行模块
        /// （否则每轮泄漏一个模块、工程越来越脏）；状态模块 JYYJXaiState 应恰好存在一个。</summary>
        private static void VerifyVbaModuleHygiene(dynamic wb)
        {
            var runResidues = 0;
            var stateCount = 0;
            var total = 0;
            try
            {
                dynamic vbProject = wb.VBProject;
                dynamic components = vbProject.VBComponents;
                var count = Convert.ToInt32(components.Count);
                total = count;
                for (var i = 1; i <= count; i++)
                {
                    dynamic c = components.Item(i);
                    var name = Convert.ToString(c.Name);
                    if (name.StartsWith("JYYJRun_", StringComparison.Ordinal) || name == "JYYJHelper")
                    {
                        runResidues++;
                    }

                    if (name == "JYYJXaiState")
                    {
                        stateCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  D2 模块卫生检查异常: " + ex.Message);
                return;
            }

            if (runResidues == 0 && stateCount == 1)
            {
                Console.WriteLine($"PASS  D2 模块卫生：无运行模块残留，JYYJXaiState 恰 1 个（共 {total} 组件）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  D2 模块卫生：运行模块残留 {runResidues} 个，状态模块 {stateCount} 个（期望 0/1，共 {total} 组件）");
            }
        }

        /// <summary>P3-D3：Python 算（结构化 xai_result）→ 经 args 通道把 result 注入 run_vba_snippet →
        /// VBA 用 XaiArg 读出并写回单元格，回读验证「多步数据经 3.2 通道传递」在真机闭环。</summary>
        private static void VerifyPythonToVbaChain(ToolRegistry registry, dynamic wb)
        {
            if (!registry.TryGet("run_python_code", out var pyTool))
            {
                _failed++;
                Console.WriteLine("FAIL  D3 run_python_code 未注册");
                return;
            }

            var pyRes = pyTool.Execute(new Dictionary<string, object>
            {
                ["code"] = "xai_result({'total': 42, 'avg': 21.5, 'items': [1, 2, 3]})",
                ["timeout_ms"] = 15000
            });
            if (!pyRes.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  D3 Python 计算: {pyRes.ErrorCode}: {Short(pyRes.Data as string)}");
                return;
            }

            var result = PyResult(pyRes);
            if (result == null)
            {
                _failed++;
                Console.WriteLine("FAIL  D3 Python 未返回结构化 result");
                return;
            }

            VbaRun(registry, "D3 VBA 写回(Python result)",
                "Sub T()\n Range(\"C1\").Value = XaiArg(\"total\")\n Range(\"C2\").Value = XaiArg(\"avg\")\nEnd Sub",
                result, true, "把 Python 算出的 total/avg 写入 C1/C2");

            dynamic ws = wb.Worksheets.Item[1];
            var total = Convert.ToDouble(ws.Cells.Item[1, 3].Value2);
            var avg = Convert.ToDouble(ws.Cells.Item[2, 3].Value2);
            if (total == 42 && Math.Abs(avg - 21.5) < 1e-9)
            {
                Console.WriteLine($"PASS  D3 Python 算→VBA 写：C1={total}, C2={avg}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  D3 Python 算→VBA 写：C1={total}, C2={avg}（期望 42 / 21.5）");
            }
        }

        // ==================== 工具封装 ====================

        private static void PyRunOpt(ToolRegistry registry, string label, string code, bool persist, bool stateClear,
            bool expect, Func<string, bool> assert, Func<ToolResult, bool> resultAssert, string extra)
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

                var sw = System.Diagnostics.Stopwatch.StartNew();
                var result = tool.Execute(args);
                sw.Stop();

                var text = result.Ok ? PropStr(result) : (result.Data as string ?? "");
                var ok = result.Ok == expect && assert(text) && (resultAssert == null || resultAssert(result));
                if (ok)
                {
                    Console.WriteLine($"PASS  [Py] {label} [{sw.ElapsedMilliseconds}ms]");
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  [Py] {label} → ok={result.Ok}（期望 {expect}），{result.ErrorCode}: {Short(result.Data as string)}；{extra} [{sw.ElapsedMilliseconds}ms]");
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"FAIL  [Py] {label} → 异常 {ex.Message}；{extra}");
            }
        }

        private static void VbaRun(ToolRegistry registry, string label, string snippet, JToken args, bool expectOk, string extra, bool stateClear = false)
        {
            VbaRunInner(registry, label, snippet, args, expectOk, extra, stateClear, out _);
        }

        /// <summary>执行 run_vba_snippet 并在成功路径返回其 ToolResult（供调用方检查 state_reset 等字段）。</summary>
        private static ToolResult VbaRunWithResult(ToolRegistry registry, string label, string snippet, JToken args, bool expectOk, string extra, bool stateClear = false)
        {
            return VbaRunInner(registry, label, snippet, args, expectOk, extra, stateClear, out _);
        }

        private static ToolResult VbaRunInner(ToolRegistry registry, string label, string snippet, JToken args, bool expectOk, string extra, bool stateClear, out bool failed)
        {
            failed = false;
            try
            {
                if (!registry.TryGet("run_vba_snippet", out var tool))
                {
                    throw new InvalidOperationException("run_vba_snippet 未注册");
                }

                var callArgs = new Dictionary<string, object> { ["snippet"] = snippet };
                if (args != null) callArgs["args"] = args;
                if (stateClear) callArgs["state_clear"] = true;

                var result = tool.Execute(callArgs);
                var ok = result.Ok == expectOk;
                if (ok)
                {
                    Console.WriteLine("PASS  " + label);
                }
                else
                {
                    failed = true;
                    _failed++;
                    Console.WriteLine($"FAIL  {label} → ok={result.Ok}（期望 {expectOk}），{result.ErrorCode}: {Short(result.Data as string)}；{extra}");
                }

                return result;
            }
            catch (Exception ex)
            {
                failed = true;
                _failed++;
                Console.WriteLine($"FAIL  {label} → 异常 {ex.Message}；{extra}");
                return null;
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

        private static string Short(string s)
        {
            if (s == null) return "(null)";
            return s.Length <= 700 ? s : s.Substring(0, 700) + "…";
        }
    }
}
