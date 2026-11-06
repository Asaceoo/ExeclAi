// 真机验证 run_python_excel 的 write_ranges 快照+回滚（开放问题 #2「Python 直写 Excel 撤销策略」决策依据）。
//   VWR-1 成功路径：声明 write_ranges 后 xlwings 写回目标区 → ok=true、write_ranges_snapshotted=1、目标区值正确。
//   VWR-2 失败回滚：声明 write_ranges 后 Python 抛异常 → PYTHON_RUN_ERROR + 「已自动回滚」，目标区恢复原值。
// COM 无状态铁律：本 harness 绝不跨工具调用持有 wb/ws RCW（Snapshot/ResolveSheet 内部 FinalRelease 共享 RCW
// 会连带销毁持有引用 → COM 分离）。所有读/写单元格一律经 app 重解析、用完即放。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyWriteRanges
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyWriteRanges real-machine ==");

            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(120000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 120s");
            }

            Console.WriteLine(_failed == 0 ? "== VerifyWriteRanges ALL PASS ==" : "== VerifyWriteRanges FAIL count=" + _failed);

            PythonWorkerHost.ResetWorker(); // 杀常驻 worker，防进程残留
            return _failed == 0 ? 0 : 1;
        }

        private static void RunRealCases()
        {
            dynamic app = null;
            int excelPid = 0;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                excelPid = ExcelPidOf(app);
                app.Workbooks.Add();

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                if (!registry.TryGet("run_python_excel", out var xlTool))
                {
                    _failed++;
                    Console.WriteLine("FAIL  run_python_excel 未注册");
                    return;
                }

                VerifyInstanceDiagnostics(xlTool); // VWR-0：只读诊断 xlwings 实例附着
                VerifySuccessPath(xlTool, app, excelPid);   // VWR-1（声明 write_ranges）
                VerifyNoWriteRanges(xlTool, app, excelPid); // VWR-1b：无 write_ranges 写入可见性对照
                VerifyRollbackPath(xlTool, app, excelPid);  // VWR-2
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  harness 异常: " + ex.Message);
            }
            finally
            {
                try { dynamic wb = app != null ? app.ActiveWorkbook : null; if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) { app.Quit(); } } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                KillExcelProcess(excelPid);
            }
        }

        /// <summary>VWR-0：只读诊断——列出 xlwings 可见实例、active 实例/簿，不写任何单元格。
        /// 用于确认「多 Excel 实例并存时 xlwings 附着目标」是否与 C# application 一致。</summary>
        private static void VerifyInstanceDiagnostics(ToolDescriptor xlTool)
        {
            var res = xlTool.Execute(new Dictionary<string, object>
            {
                ["code"] = "import xlwings as xw\napps = [(a.pid, a.hwnd, a.books.active.name if a.books else None) for a in xw.apps]\n"
                           + "info = {'apps': apps, 'active_app_pid': xw.apps.active.pid, 'active_app_hwnd': xw.apps.active.hwnd, 'active_book': xw.books.active.name}"
                           + "\nxai_result(info)",
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-0 诊断失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            try
            {
                var data = JObject.FromObject(res.Data);
                Console.WriteLine("INFO  VWR-0 xlwings 实例诊断: " + data);
            }
            catch
            {
                Console.WriteLine("INFO  VWR-0 xlwings 实例诊断: " + Short(res.Data as string));
            }
        }

        /// <summary>VWR-1：声明 write_ranges → 成功后目标区值正确、返回快照计数。</summary>
        private static void VerifySuccessPath(ToolDescriptor tool, dynamic app, int excelPid)
        {
            WriteCell(app, 6, 1, 123); // F1 初始值
            Console.WriteLine("INFO  VWR-1 前: C# ExcelPid=" + ExcelPidOf(app) + " Hwnd=" + AppHwndOf(app)
                + " F1=" + Convert.ToString(ReadCell(app, 6, 1)));

            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "import xlwings as xw\napp = xw.apps[" + excelPid + "]\nsht = app.books.active.sheets[0]\nsht.range(\"F1\").value = 99\n"
                           + "xai_result({'readback': sht.range('F1').value, 'sheet': sht.name, 'book': app.books.active.name, 'app_pid': app.pid})",
                ["write_ranges"] = new object[] { "Sheet1!F1" },
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-1 成功路径失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            var snapshotted = 0;
            var diag = "";
            try
            {
                var data = JObject.FromObject(res.Data);
                snapshotted = (int)(data["write_ranges_snapshotted"] ?? 0);
                if (data["result"] is JObject r) diag = " result=" + r;
            }
            catch { }

            var v = ReadCell(app, 6, 1);
            Console.WriteLine("INFO  VWR-1 后: C# ExcelPid=" + ExcelPidOf(app) + " Hwnd=" + AppHwndOf(app)
                + " | " + EnumerateWorkbooks(app));
            if (Convert.ToDouble(v) == 99 && snapshotted == 1)
            {
                Console.WriteLine($"PASS  VWR-1 成功路径 write_ranges 快照执行（snapshotted={snapshotted}）→ F1=" + Convert.ToDouble(v));
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-1 成功路径异常: F1={v} snapshotted={snapshotted}（期望 99 / 1）{diag}");
            }
        }

        /// <summary>VWR-1b：对照实验——不声明 write_ranges，xlwings 写 F1=99 后 C# 能否读到。
        /// 用于隔离「write_ranges 的 Snapshot(FinalReleaseComObject) 是否破坏 C# 与 Excel 连接」。</summary>
        private static void VerifyNoWriteRanges(ToolDescriptor tool, dynamic app, int excelPid)
        {
            WriteCell(app, 6, 1, 123);
            Console.WriteLine("INFO  VWR-1b 前: C# F1=" + Convert.ToString(ReadCell(app, 6, 1)));

            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "import xlwings as xw\napp = xw.apps[" + excelPid + "]\nsht = app.books.active.sheets[0]\nsht.range(\"F1\").value = 99\n"
                           + "xai_result({'readback': sht.range('F1').value, 'app_pid': app.pid})",
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-1b 执行失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            var v = ReadCell(app, 6, 1);
            if (Convert.ToDouble(v) == 99)
            {
                Console.WriteLine("PASS  VWR-1b 无 write_ranges：C# 读到 xlwings 写入 F1=" + Convert.ToDouble(v));
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-1b 无 write_ranges：C# 读 F1={v}（期望 99）——xlwings 写入对 C# 不可见");
            }
        }

        /// <summary>VWR-2：声明 write_ranges → Python 抛异常 → 自动回滚，目标区恢复原值。</summary>
        private static void VerifyRollbackPath(ToolDescriptor tool, dynamic app, int excelPid)
        {
            WriteCell(app, 6, 1, 123); // F1 初始值

            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = "import xlwings as xw\napp = xw.apps[" + excelPid + "]\nsht = app.books.active.sheets[0]\nsht.range(\"F1\").value = 999\nraise RuntimeError('boom')",
                ["write_ranges"] = new object[] { "Sheet1!F1" },
                ["timeout_ms"] = 30000
            });

            if (res.Ok)
            {
                _failed++;
                Console.WriteLine("FAIL  VWR-2 应失败却成功");
                return;
            }

            var msg = Convert.ToString(res.Data);
            var v = ReadCell(app, 6, 1);
            if (res.ErrorCode == "PYTHON_RUN_ERROR" && msg.Contains("已自动回滚") && Convert.ToDouble(v) == 123)
            {
                Console.WriteLine($"PASS  VWR-2 失败自动回滚（{res.ErrorCode}）→ F1=" + Convert.ToDouble(v) + "（恢复初始值）");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VWR-2 回滚异常: code={res.ErrorCode} F1={v}（期望 PYTHON_RUN_ERROR / 123）msg={Short(msg)}");
            }
        }

        // ==================== 无状态 COM 辅助 ====================

        private static string EnumerateWorkbooks(dynamic app)
        {
            var sb = new System.Text.StringBuilder();
            dynamic wbs = null;
            dynamic wb = null;
            try
            {
                sb.Append("ActiveWorkbook=");
                try { sb.Append(Convert.ToString(app.ActiveWorkbook.Name)); } catch (Exception ex) { sb.Append("<err:" + ex.Message + ">"); }
                wbs = app.Workbooks;
                var count = Convert.ToInt32(wbs.Count);
                sb.Append(" count=" + count);
                for (var i = 1; i <= count; i++)
                {
                    wb = wbs.Item[i];
                    dynamic ws = null;
                    dynamic cell = null;
                    try
                    {
                        ws = wb.Worksheets.Item[1];
                        cell = ws.Range["F1"];
                        sb.Append(" | [" + Convert.ToString(wb.Name) + "] S1!F1=" + Convert.ToString(cell.Value2));
                    }
                    catch (Exception ex)
                    {
                        sb.Append(" | [" + Convert.ToString(wb.Name) + "] <read err:" + ex.Message + ">");
                    }
                    finally
                    {
                        Release((object)cell);
                        Release((object)ws);
                        Release((object)wb);
                    }
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                return "<enum err:" + ex.Message + ">";
            }
            finally
            {
                Release((object)wbs);
            }
        }

        private static int AppHwndOf(dynamic app)
        {
            try
            {
                return Convert.ToInt32(app.Hwnd);
            }
            catch
            {
                return 0;
            }
        }

        private static object ReadCell(dynamic app, int row, int col)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic rng = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[1];
                rng = ws.Range["F1"];
                return rng.Value2;
            }
            finally
            {
                Release((object)rng);
                Release((object)ws);
                Release((object)wb);
            }
        }

        private static void WriteCell(dynamic app, int row, int col, object value)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic rng = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[1];
                rng = ws.Range["F1"];
                rng.Value2 = value;
            }
            finally
            {
                Release((object)rng);
                Release((object)ws);
                Release((object)wb);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.FinalReleaseComObject(value); } catch { }
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
                    return;
                }
                Thread.Sleep(200);
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        private static string Short(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length > 160 ? text.Substring(0, 160) : text;
        }
    }
}
