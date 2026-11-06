using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyPythonPanel
{
    /// <summary>
    /// PT 面板（Python 脚本面板，第 6 个面板）真机验证 harness。
    /// 覆盖三条新链路：
    ///  VPS  预置模板播种（真实默认库，幂等增量，py_* 带 python 标签）
    ///  VPV  ValidatePython 的 py_compile 命令语义（合法/非法代码）
    ///  VPC  RunPythonCode 委托的 run_python_code 纯计算路径（xai_result 结构化 + persist）
    ///  VPE  RunPythonExcel 委托的 run_python_excel 路径（xlwings 真机写回）
    /// Excel 用例跑在 STA 线程，退出时按 PID 兜底强杀防僵尸。
    /// </summary>
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyPythonPanel real-machine (PT 第6面板) ==");

            VerifyPresetSeeding();
            VerifyValidateCommand();
            VerifyCalcExecution();

            var runner = new Thread(RunExcelCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(90000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness Excel 用例超时 90s");
            }

            Console.WriteLine();
            Console.WriteLine(_failed == 0 ? "== VerifyPythonPanel ALL PASS ==" : "== VerifyPythonPanel FAIL count=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

        // ===== VPS 预置模板播种（真机默认库：与 ListCustomTools→SeedPresets 同路径）=====
        private static void VerifyPresetSeeding()
        {
            var store = new CustomToolStore(CustomToolStore.DefaultRoot());
            var before = store.LoadAll().Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            store.SeedPresets();
            var all = store.LoadAll();
            var pyPresets = all.Where(t => t.Name.StartsWith("py_", StringComparison.Ordinal)).ToList();
            if (pyPresets.Count == 4
                && pyPresets.All(p => p.Tags != null && p.Tags.Any(t => string.Equals(t, "python", StringComparison.OrdinalIgnoreCase))))
            {
                Console.WriteLine("PASS  VPS-1 py_* 预置已播种 4 个且均带 python 标签：" + string.Join(",", pyPresets.Select(p => p.Name)));
            }
            else
            {
                _failed++;
                Console.WriteLine("FAIL  VPS-1 py_* 预置应播种 4 个且带 python 标签，实得 "
                    + pyPresets.Count + " 个：" + string.Join(",", pyPresets.Select(p => p.Name)));
            }

            var missing = before.Where(n => !all.Any(t => string.Equals(t.Name, n, StringComparison.OrdinalIgnoreCase))).ToList();
            if (missing.Count == 0)
            {
                Console.WriteLine($"PASS  VPS-2 幂等播种不覆盖/不丢失已有条目（播种前 {before.Count} 条仍在）");
            }
            else
            {
                _failed++;
                Console.WriteLine("FAIL  VPS-2 播种后原有条目缺失: " + string.Join(",", missing));
            }
        }

        // ===== VPV ValidatePython 命令语义（python -m py_compile，不执行）=====
        private static void VerifyValidateCommand()
        {
            var dir = Path.Combine(Path.GetTempPath(), "xai_vpp_check_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var good = Path.Combine(dir, "good.py");
                File.WriteAllText(good, "total = sum(v for v in [1,2,3] if v > 0)\nprint(total)", new System.Text.UTF8Encoding(false));
                var bad = Path.Combine(dir, "bad.py");
                File.WriteAllText(bad, "def broken(:\n    print(1)\n", new System.Text.UTF8Encoding(false));

                var goodExit = RunPyCompile(good, dir, out var goodErr);
                var badExit = RunPyCompile(bad, dir, out var badErr);
                if (goodExit == 0 && badExit != 0 && badErr.IndexOf("SyntaxError", StringComparison.Ordinal) >= 0)
                {
                    Console.WriteLine("PASS  VPV-1 py_compile 合法代码通过 / 非法代码报 SyntaxError（校验语义正确）");
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPV-1 py_compile 语义: goodExit={goodExit} badExit={badExit} badErr={Short(badErr)}");
                }
            }
            finally
            {
                try { if (Directory.Exists(dir)) Directory.Delete(dir, true); } catch { }
            }
        }

        private static int RunPyCompile(string file, string workDir, out string err)
        {
            err = "";
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = "-m py_compile " + Path.GetFileName(file),
                    WorkingDirectory = workDir,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                var se = proc.StandardError.ReadToEnd();
                var so = proc.StandardOutput.ReadToEnd();
                if (!proc.WaitForExit(15000))
                {
                    try { proc.Kill(); } catch { }
                    return -1;
                }
                err = string.IsNullOrWhiteSpace(se) ? so : se;
                return proc.ExitCode;
            }
            catch (Exception ex)
            {
                err = ex.Message;
                return -2;
            }
        }

        // ===== VPC RunPythonCode 委托路径：run_python_code 纯计算（xai_result + persist）=====
        private static void VerifyCalcExecution()
        {
            var host = new ExcelToolHost(new object());
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();
            if (!registry.TryGet("run_python_code", out var tool))
            {
                _failed++;
                Console.WriteLine("FAIL  VPC-0 run_python_code 未注册");
                return;
            }

            var r = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "xai_result({'total': 42, 'avg': 21.5, 'label': '面板计算'})",
                ["timeout_ms"] = 20000,
                ["output_limit_kb"] = 8
            });
            var result = PyResult(r);
            if (r.Ok && result is JObject jo && (int)jo["total"] == 42 && Math.Abs((double)jo["avg"] - 21.5) < 1e-9)
            {
                Console.WriteLine("PASS  VPC-1 run_python_code xai_result 结构化返回（total=42, avg=21.5）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-1 run_python_code 结构化返回: ok={r.Ok} {r.ErrorCode}: {Short(r.Data?.ToString())}");
            }

            var p1 = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "pp_total = 100", ["persist"] = true, ["timeout_ms"] = 15000, ["output_limit_kb"] = 8
            });
            var p2 = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "xai_result({'pp_total': pp_total})", ["persist"] = true, ["timeout_ms"] = 15000, ["output_limit_kb"] = 8
            });
            var p2r = PyResult(p2);
            if (p1.Ok && p2.Ok && p2r is JObject jo2 && (int)jo2["pp_total"] == 100)
            {
                Console.WriteLine("PASS  VPC-2 persist 跨调用命名空间保留（pp_total=100）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-2 persist 状态保留: p1={p1.Ok} p2={p2.Ok} {Short(p2.Data?.ToString())}");
            }
        }

        // ===== VPE RunPythonExcel 委托路径：run_python_excel（xlwings 真机写回）=====
        private static void RunExcelCases()
        {
            dynamic app = null;
            dynamic wb = null;
            int excelPid = 0;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                excelPid = ExcelPidOf(app);
                wb = app.Workbooks.Add();

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();
                if (!registry.TryGet("run_python_excel", out var tool))
                {
                    _failed++;
                    Console.WriteLine("FAIL  VPE-0 run_python_excel 未注册");
                    return;
                }

                var code = "import xlwings as xw\n"
                    + "sht = xw.books.active.sheets[0]\n"
                    + "sht.range(\"A1\").value = 42\n"
                    + "xai_result({'written': sht.range(\"A1\").value})";
                var r = tool.Execute(new Dictionary<string, object>
                {
                    ["code"] = code, ["timeout_ms"] = 30000, ["output_limit_kb"] = 8
                });
                dynamic ws = wb.Worksheets.Item[1];
                var v = Convert.ToDouble(ws.Cells.Item[1, 1].Value2);
                if (r.Ok && v == 42)
                {
                    Console.WriteLine("PASS  VPE-1 run_python_excel xlwings 写回 A1=42（Excel 读写模式路径可用）");
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPE-1 run_python_excel 写回: ok={r.Ok} A1={v} {r.ErrorCode}: {Short(r.Data?.ToString())}");
                }

                // VPS-3 py_sum_range 预置：在真实多表工作簿上取「活动表」而非第一张表（sheets.active 生效）
                var preset = VbaPresetCatalog.All.FirstOrDefault(t => t.Name == "py_sum_range");
                if (preset == null)
                {
                    _failed++;
                    Console.WriteLine("FAIL  VPS-3 py_sum_range 预置模板缺失");
                }
                else
                {
                    dynamic ws2 = wb.Worksheets.Add();
                    ws2.Activate();
                    for (var i = 1; i <= 10; i++) ws2.Cells.Item[i, 1].Value2 = i;
                    var r2 = tool.Execute(new Dictionary<string, object>
                    {
                        ["code"] = preset.Code, ["timeout_ms"] = 30000, ["output_limit_kb"] = 8
                    });
                    var r2j = PyResult(r2);
                    if (r2.Ok && r2j is JObject jo3 && (int)jo3["count"] == 10 && (int)jo3["sum"] == 55)
                    {
                        Console.WriteLine("PASS  VPS-3 py_sum_range 预置在活动表求和=55（sheets.active 生效）");
                    }
                    else
                    {
                        _failed++;
                        Console.WriteLine($"FAIL  VPS-3 py_sum_range 活动表求和: ok={r2.Ok} {Short(r2.Data?.ToString())}");
                    }
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  Excel 用例异常: " + ex.Message);
            }
            finally
            {
                try { if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) app.Quit(); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (wb != null && Marshal.IsComObject(wb)) Marshal.FinalReleaseComObject(wb); } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                KillExcelProcess(excelPid);
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
            catch { return 0; }
        }

        private static void KillExcelProcess(int pid)
        {
            if (pid <= 0) return;
            for (var i = 0; i < 5; i++)
            {
                try
                {
                    using (var p = Process.GetProcessById(pid))
                    {
                        if (p.HasExited) return;
                        p.Kill();
                    }
                }
                catch { return; }
                Thread.Sleep(200);
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        private static string Short(string s)
        {
            if (s == null) return "(null)";
            return s.Length <= 700 ? s : s.Substring(0, 700) + "…";
        }
    }
}
