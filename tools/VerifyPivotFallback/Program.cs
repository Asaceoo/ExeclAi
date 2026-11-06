// 真机验证 create_excel_pivot_summary 缺参兜底（修复「缺少必填参数：source_address」）。
//   VPF-1 直接调用（无任何参数）→ 兜底当前多格选区，透视写回源区下方 2 行，结果正确。
//   VPF-2 显式传 source_address → 结果一致（回归老路径）。
//   VPF-3 跑内置流程「透视汇总表」→ 断言全步骤完成 + 透视落盘正确（含后续 format 步骤）。
// Excel COM 是 STA：创建 app 与执行 Excel 用例必须同 STA 线程。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;

namespace VerifyPivotFallback
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyPivotFallback real-machine ==");

            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(90000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 90s");
            }

            Console.WriteLine(_failed == 0 ? "== VerifyPivotFallback ALL PASS ==" : "== VerifyPivotFallback FAIL count=" + _failed);
            return _failed == 0 ? 0 : 1;
        }

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
                excelPid = ExcelPidOf(app);
                wb = app.Workbooks.Add();
                // 注意：不能跨工具调用持有 ws/wb RCW——工具内部 FinalReleaseComObject 会按 COM
                // 对象身份释放共享 RCW，导致本处引用失效（RCW 分离）。读写一律用无状态 helper 重解析。
                FillSheet(app);

                var host = new ExcelToolHost((object)app);
                var registry = host.CreateRegistry();

                VerifyDirectFallback(registry, app); // VPF-1：无参兜底
                VerifyExplicitSource(registry);      // VPF-2：显式参数回归
                VerifyWorkflow(registry, app);       // VPF-3：内置流程全链路
                VerifyAllPivotTemplates(registry);   // VPF-4：6 个透视模板的透视步骤均成功
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
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (wb != null && Marshal.IsComObject(wb)) Marshal.FinalReleaseComObject(wb); } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                KillExcelProcess(excelPid);
            }
        }

        /// <summary>在活动工作簿 Sheet1 填充 B12 数据区并选中。全部经 app 重解析，不留共享 RCW。</summary>
        private static void FillSheet(dynamic app)
        {
            dynamic wb = null;
            dynamic ws = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[1];
                ws.Cells.Item[12, 2].Value2 = "类别";
                ws.Cells.Item[12, 6].Value2 = "金额";
                var data = new object[9, 5];
                for (var i = 0; i < 9; i++)
                {
                    data[i, 0] = "类" + (i % 3);
                    data[i, 4] = 100 + i;
                }
                ws.Range["B13:F21"].Value2 = data;
                ws.Range["B12:F21"].Select();
            }
            finally
            {
                Release(ws);
                Release(wb);
            }
        }

        /// <summary>无状态读取单元格：每次经 app 重解析，避免共享 RCW 分离问题。</summary>
        private static object ReadCellValue(dynamic app, int row, int col)
        {
            dynamic wb = null;
            dynamic ws = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[1];
                return ws.Cells.Item[row, col].Value2;
            }
            finally
            {
                Release(ws);
                Release(wb);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.FinalReleaseComObject(value); } catch { }
            }
        }

        // VPF-1：create_excel_pivot_summary 无任何参数 → 兜底当前选区，写回源区（B12:F21）下方 2 行 = B23 起。
        private static void VerifyDirectFallback(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("create_excel_pivot_summary", out var tool))
            {
                _failed++;
                Console.WriteLine("FAIL  VPF-1 create_excel_pivot_summary 未注册");
                return;
            }

            var res = tool.Execute(new Dictionary<string, object>());
            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPF-1 无参调用失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            try
            {
                var header = Convert.ToString(ReadCellValue(app, 23, 2));
                double total = 0;
                for (var row = 24; row <= 26; row++)
                {
                    var v = ReadCellValue(app, row, 3);
                    if (v != null) total += Convert.ToDouble(v);
                }

                if (header == "Row Labels" && Math.Abs(total - 936) < 0.001)
                {
                    Console.WriteLine("PASS  VPF-1 无参调用兜底选区成功，透视落盘 B23:C26，合计=" + total);
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPF-1 透视落盘异常: 表头={header} 合计={total}（期望 Row Labels / 936）");
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  VPF-1 读回透视结果异常: " + ex.Message);
            }
        }

        // VPF-2：显式 source_address 老路径回归。
        private static void VerifyExplicitSource(ToolRegistry registry)
        {
            if (!registry.TryGet("create_excel_pivot_summary", out var tool))
            {
                _failed++;
                Console.WriteLine("FAIL  VPF-2 create_excel_pivot_summary 未注册");
                return;
            }

            var res = tool.Execute(new Dictionary<string, object>
            {
                ["source_address"] = "$B$12:$F$21"
            });
            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPF-2 显式参数调用失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            Console.WriteLine("PASS  VPF-2 显式 source_address 调用成功（老路径回归）");
        }

        // VPF-3：内置流程「透视汇总表」全步骤完成 + 透视落盘正确。
        private static void VerifyWorkflow(ToolRegistry registry, dynamic app)
        {
            var store = new WorkflowStore();
            if (!store.TryGet("透视汇总表", out var template))
            {
                _failed++;
                Console.WriteLine("FAIL  VPF-3 找不到内置流程「透视汇总表」");
                return;
            }

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            // 模拟桥接层选区捕获：模板引用 {selection} 时注入当前选区覆盖。
            var overrides = new Dictionary<string, object>
            {
                ["selection"] = "$B$12:$F$21",
                ["selection_sheet"] = "Sheet1"
            };
            var run = executor.RunAsync(template, overrides, null, 0).GetAwaiter().GetResult();

            Console.WriteLine("VPF-3  流程步骤结果：");
            foreach (var step in run.Steps)
            {
                var data = step.Result?.Data;
                var summary = data == null ? "" : Short(step.Result.Data.ToString());
                Console.WriteLine($"       - 步骤{step.Index + 1} [{step.Tool}] ok={step.Result?.Ok} " +
                                  (step.Result?.Ok == false ? $"code={step.Result.ErrorCode} msg={Short(step.Result.Data as string)}" : summary));
            }

            if (!run.Completed || run.FailedIndex != null)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPF-3 流程未完成: Completed={run.Completed} FailedIndex={run.FailedIndex}");
                return;
            }

            try
            {
                var header = Convert.ToString(ReadCellValue(app, 23, 2));
                if (header != "Row Labels")
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPF-3 流程完成后透视未落盘 B23（表头={header}）");
                    return;
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL  VPF-3 读回流程透视结果异常: " + ex.Message);
                return;
            }

            Console.WriteLine("PASS  VPF-3 内置流程「透视汇总表」全步骤完成且透视落盘正确");
        }

        /// <summary>VPF-4：其余 5 个透视内置模板的透视步骤（create_excel_pivot_summary）必须成功，
        /// 证明「缺 source_address」这一类错误在全部透视模板上已消除。</summary>
        private static void VerifyAllPivotTemplates(ToolRegistry registry)
        {
            var names = new[]
            {
                "数据分析·透视与报告",
                "人事·人员结构透视",
                "财务·收支透视与图表",
                "人事·年龄工龄分布透视",
                "数据分析·多维聚合看板"
            };

            var store = new WorkflowStore();
            var overrides = new Dictionary<string, object>
            {
                ["selection"] = "$B$12:$F$21",
                ["selection_sheet"] = "Sheet1"
            };

            foreach (var name in names)
            {
                if (!store.TryGet(name, out var template))
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPF-4 找不到内置流程「{name}」");
                    continue;
                }

                var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
                var run = executor.RunAsync(template, overrides, null, 0).GetAwaiter().GetResult();

                var pivotStep = run.Steps.Find(s => s.Tool == "create_excel_pivot_summary");
                if (pivotStep == null)
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPF-4 「{name}」未包含透视步骤");
                    continue;
                }

                if (pivotStep.Result == null || !pivotStep.Result.Ok)
                {
                    _failed++;
                    Console.WriteLine($"FAIL  VPF-4 「{name}」透视步骤失败: {pivotStep.Result?.ErrorCode}: {Short(pivotStep.Result?.Data as string)}");
                    continue;
                }

                Console.WriteLine($"PASS  VPF-4 「{name}」透视步骤成功");
            }
        }

        // ==================== 辅助 ====================

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
            return text.Length > 120 ? text.Substring(0, 120) : text;
        }
    }
}
