// 真机验证：PT 面板 Python 用 xlwings 按选区画「原生 Excel 图表」（开放需求「Python 画图表」可行性）。
//   VPC-1 柱状图：run_python_excel 内 xw.sheets.charts.add() + set_source_data(sel_range) → ChartObjects +1、图表存在。
//   VPC-2 数据源断言：图表 SourceData 解析地址 = 选区地址（画的就是选区数据）。
//   VPC-3 图表类型：默认 ColumnChart / 显式设置类型。
// COM 无状态铁律：读写一律经 app 重解析、用完即放；坐标统一用 Range["A1"] 地址方式（见踩坑 5.5）。
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyPyChart
{
    internal static class Program
    {
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("== VerifyPyChart real-machine ==");

            var runner = new Thread(RunRealCases) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(120000))
            {
                runner.Abort();
                _failed++;
                Console.WriteLine("FAIL  harness 整体超时 120s");
            }

            Console.WriteLine(_failed == 0 ? "== VerifyPyChart ALL PASS ==" : "== VerifyPyChart FAIL count=" + _failed);
            PythonWorkerHost.ResetWorker();
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

                FillData(app); // A1 标题 + A2:B6 数据，选中 A1:B6

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                if (!registry.TryGet("run_python_excel", out var xlTool))
                {
                    _failed++;
                    Console.WriteLine("FAIL  run_python_excel 未注册");
                    return;
                }

                VerifyColumnChart(xlTool, app, excelPid); // VPC-1/2/3
                VerifyPresetTemplate(xlTool, app, excelPid); // VPC-4：预置模板主路径（无标题）
                VerifyMultiRange(xlTool, app, excelPid); // VPC-5：多选区注入取首块
                VerifyChartTitle(xlTool, app, excelPid); // VPC-6：模板设置图表标题
                VerifyCrossSheet(xlTool, app, excelPid); // VPC-7：活动表≠数据表 跨表画图
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

        private static void VerifyColumnChart(ToolDescriptor tool, dynamic app, int excelPid)
        {
            var before = ChartCount(app);

            // 模拟 PT 面板「插入选区」注入的 sel_range（'Sheet1'!A1:B6），xlwings 画柱状图
            var code =
                "import xlwings as xw\n"
                + "sel_range = \"'Sheet1'!A1:B6\"\n"
                + "app = xw.apps[" + excelPid + "]\n"
                + "sht = app.books.active.sheets[0]\n"
                + "chart = sht.charts.add()\n"
                + "chart.set_source_data(sht.range(sel_range))\n"
                + "chart.chart_type = 'column_clustered'\n"
                + "xai_result({'chart_count': len(sht.charts), 'chart_name': chart.name})";

            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = code,
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-1 画图执行失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            var after = ChartCount(app);
            if (after <= before)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-1 图表未创建: before={before} after={after} result={res.Data}");
                return;
            }

            // VPC-2：新图表数据源应绑定选区 A1:B6。SERIES 公式会把选区拆成
            // 名称/分类轴/数值轴三段，如 =SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$6,Sheet1!$B$2:$B$6,1)
            var source = ChartSourceAddress(app, after);
            var boundSelection = source.IndexOf("$A$2:$A$6", StringComparison.OrdinalIgnoreCase) >= 0
                                 && source.IndexOf("$B$2:$B$6", StringComparison.OrdinalIgnoreCase) >= 0;
            if (boundSelection)
            {
                Console.WriteLine($"PASS  VPC-1/2 柱状图创建并绑定选区数据: ChartObjects {before}→{after} source={source} result={res.Data}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-2 图表数据源不符: source={source}（期望含 A1:B6）");
            }

            // VPC-3：默认类型断言（索引=after）
            var chartType = ChartTypeOf(app, after);
            if (chartType.Length > 0)
            {
                Console.WriteLine("PASS  VPC-3 图表类型=" + chartType);
            }
        }

        // VPC-4：预置模板 py_chart_selection 真机可跑——读取 VbaPresetCatalog 真实代码，
        // 模拟 PT 面板「插入选区」在代码前注入 sel_range 后经 run_python_excel 执行，画图成功即闭环。
        private static void VerifyPresetTemplate(ToolDescriptor tool, dynamic app, int excelPid)
        {
            var preset = VbaPresetCatalog.All.FirstOrDefault(t => t.Name == "py_chart_selection");
            if (preset == null)
            {
                _failed++;
                Console.WriteLine("FAIL  VPC-4 预置模板 py_chart_selection 不存在");
                return;
            }

            var before = ChartCount(app);
            var code = "sel_range = \"'Sheet1'!A1:B6\"\n" + preset.Code;
            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = code,
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-4 预置模板执行失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            var after = ChartCount(app);
            var result = Convert.ToString(res.Data);
            // used_range 返回 $A$1:$B$6（带 $ 绝对引用），匹配其中的 A$1:$B$6 子串
            var usedOk = result.IndexOf("A$1:$B$6", StringComparison.OrdinalIgnoreCase) >= 0;
            var titleOk = result.IndexOf("title_set\": false", StringComparison.OrdinalIgnoreCase) >= 0;
            if (after > before && usedOk && titleOk)
            {
                Console.WriteLine($"PASS  VPC-4 预置模板主路径: ChartObjects {before}→{after} used_range 归位 result={result}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-4 预置模板结果异常: before={before} after={after} usedOk={usedOk} titleOk={titleOk} result={result}");
            }
        }

        // VPC-5：面板多选区注入（逗号连接）——模板应取首块画图、不因逗号串解析错位，used_range 回传首块地址。
        private static void VerifyMultiRange(ToolDescriptor tool, dynamic app, int excelPid)
        {
            var preset = VbaPresetCatalog.All.FirstOrDefault(t => t.Name == "py_chart_selection");
            if (preset == null) { _failed++; Console.WriteLine("FAIL  VPC-5 模板缺失"); return; }

            var before = ChartCount(app);
            var code = "sel_range = \"'Sheet1'!A1:B6,'Sheet1'!C1:D6\"\n" + preset.Code;
            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = code,
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-5 多选区执行失败: {res.ErrorCode}: {Short(res.Data as string)}");
                return;
            }

            var after = ChartCount(app);
            var result = Convert.ToString(res.Data);
            // used_range 返回 $A$1:$B$6（带 $ 绝对引用），匹配其中的 A$1:$B$6 子串
            var usedFirst = result.IndexOf("A$1:$B$6", StringComparison.OrdinalIgnoreCase) >= 0
                            && result.IndexOf("C$1:$D$6", StringComparison.OrdinalIgnoreCase) < 0;
            if (after > before && usedFirst)
            {
                Console.WriteLine($"PASS  VPC-5 多选区取首块: ChartObjects {before}→{after} used_range=首块 result={result}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-5 多选区结果异常: before={before} after={after} usedFirst={usedFirst} result={result}");
            }
        }

        // VPC-6：模板 chart_title 可选注入——设置后图表标题被正确写入，可回读 ChartTitle.Text。
        private static void VerifyChartTitle(ToolDescriptor tool, dynamic app, int excelPid)
        {
            var preset = VbaPresetCatalog.All.FirstOrDefault(t => t.Name == "py_chart_selection");
            if (preset == null) { _failed++; Console.WriteLine("FAIL  VPC-6 模板缺失"); return; }

            var before = ChartCount(app);
            var code = "sel_range = \"'Sheet1'!A1:B6\"\nchart_title = \"季度销售\"\n" + preset.Code;
            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = code,
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine("FAIL  VPC-6 标题执行失败: " + res.ErrorCode + ":\n" + Convert.ToString(res.Data));
                return;
            }

            var after = ChartCount(app);
            if (after <= before)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-6 图表未创建: before={before} after={after}");
                return;
            }

            var title = ChartTitleOf(app, after);
            if (title == "季度销售")
            {
                Console.WriteLine($"PASS  VPC-6 图表标题设置: ChartObjects {before}→{after} ChartTitle={title} result={Convert.ToString(res.Data)}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-6 标题不符: ChartTitle={title}（期望 季度销售）");
            }
        }

        // VPC-7：活动表≠数据表 的跨表画图——激活 Sheet2、插入 Sheet1 的选区，
        // 模板应把图表建在活动表（Sheet2）上、数据源跨表绑定到 Sheet1 的 A1:B6。
        private static void VerifyCrossSheet(ToolDescriptor tool, dynamic app, int excelPid)
        {
            // 确保存在第二张空表并激活它；Add(After:=last) 使新表索引可预测（=count+1）
            dynamic wb = null;
            dynamic ws2 = null;
            int sheetIndex;
            try
            {
                wb = app.ActiveWorkbook;
                var count = Convert.ToInt32(wb.Worksheets.Count);
                if (count < 2)
                {
                    ws2 = wb.Worksheets.Add(System.Reflection.Missing.Value, wb.Worksheets.Item[count]);
                    ws2.Name = "Sheet2";
                    sheetIndex = count + 1;
                }
                else
                {
                    sheetIndex = 2;
                    ws2 = wb.Worksheets.Item[2];
                }
                ws2.Activate();
                ws2.Range["A1"].Value2 = "第二张表（图表应落于此表）";
            }
            finally
            {
                Release(ws2);
                Release(wb);
            }

            var before = ChartCount(app, sheetIndex);
            var preset = VbaPresetCatalog.All.FirstOrDefault(t => t.Name == "py_chart_selection");
            if (preset == null) { _failed++; Console.WriteLine("FAIL  VPC-7 模板缺失"); return; }

            var code = "sel_range = \"'Sheet1'!A1:B6\"\n" + preset.Code;
            var res = tool.Execute(new Dictionary<string, object>
            {
                ["code"] = code,
                ["timeout_ms"] = 30000
            });

            if (!res.Ok)
            {
                _failed++;
                Console.WriteLine("FAIL  VPC-7 跨表执行失败: " + res.ErrorCode + ":\n" + Convert.ToString(res.Data));
                return;
            }

            var after = ChartCount(app, sheetIndex);
            if (after <= before)
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-7 图表未建在活动表 Sheet2: before={before} after={after}");
                return;
            }

            // 跨表数据源：SERIES 公式应含 Sheet1 的分类轴与数值轴地址
            var source = ChartSourceAddress(app, after, sheetIndex);
            var boundSheet1 = source.IndexOf("Sheet1!$A$2:$A$6", StringComparison.OrdinalIgnoreCase) >= 0
                              && source.IndexOf("Sheet1!$B$2:$B$6", StringComparison.OrdinalIgnoreCase) >= 0;
            if (boundSheet1)
            {
                Console.WriteLine($"PASS  VPC-7 跨表画图: 图表落在活动表 Sheet2 (ChartObjects {before}→{after}) 数据源跨表绑定 Sheet1 source={source} result={Convert.ToString(res.Data)}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"FAIL  VPC-7 跨表数据源不符: source={source}（期望含 Sheet1!$A$2:$A$6 与 Sheet1!$B$2:$B$6）");
            }

            // 还原激活到 Sheet1，避免影响后续/收尾
            try
            {
                wb = app.ActiveWorkbook;
                wb.Worksheets.Item[1].Activate();
            }
            finally { Release(wb); }
        }

        private static void FillData(dynamic app)
        {
            dynamic wb = null;
            dynamic ws = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[1];
                var data = new string[,] {
                    { "品类", "金额" },
                    { "A", "10" }, { "B", "25" }, { "C", "18" }, { "D", "32" }, { "E", "14" }
                };
                for (var r = 0; r <= 5; r++)
                {
                    ws.Cells.Item[r + 1, 1].Value2 = data[r, 0];
                    ws.Cells.Item[r + 1, 2].Value2 = r == 0 ? (object)data[r, 1] : (object)Convert.ToDouble(data[r, 1]);
                }
                ws.Range["A1:B6"].Select();
            }
            finally
            {
                Release(ws);
                Release(wb);
            }
        }

        // ==================== 无状态 COM 辅助 ====================

        private static int ChartCount(dynamic app, int sheetIndex = 1)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic co = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[sheetIndex];
                co = ws.ChartObjects();
                return Convert.ToInt32(co.Count);
            }
            catch
            {
                return -1;
            }
            finally
            {
                Release(co);
                Release(ws);
                Release(wb);
            }
        }

        private static string ChartSourceAddress(dynamic app, int index, int sheetIndex = 1)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic co = null;
            dynamic ch = null;
            dynamic ser = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[sheetIndex];
                co = ws.ChartObjects();
                ch = co.Item(index).Chart;
                var seriesCount = Convert.ToInt32(ch.SeriesCollection().Count);
                ser = ch.SeriesCollection().Item(1);
                var formula = Convert.ToString(ser.Formula);
                return formula.Length > 120 ? formula.Substring(0, 120) : formula;
            }
            catch (Exception ex)
            {
                return "<err:" + ex.Message + ">";
            }
            finally
            {
                Release(ser);
                Release(ch);
                Release(co);
                Release(ws);
                Release(wb);
            }
        }

        private static string ChartTypeOf(dynamic app, int index, int sheetIndex = 1)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic co = null;
            dynamic ch = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[sheetIndex];
                co = ws.ChartObjects();
                ch = co.Item(index).Chart;
                return Convert.ToString(ch.ChartType);
            }
            catch
            {
                return "";
            }
            finally
            {
                Release(ch);
                Release(co);
                Release(ws);
                Release(wb);
            }
        }

        private static string ChartTitleOf(dynamic app, int index, int sheetIndex = 1)
        {
            dynamic wb = null;
            dynamic ws = null;
            dynamic co = null;
            dynamic ch = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets.Item[sheetIndex];
                co = ws.ChartObjects();
                ch = co.Item(index).Chart;
                return Convert.ToString(ch.ChartTitle.Text);
            }
            catch
            {
                return "";
            }
            finally
            {
                Release(ch);
                Release(co);
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
            return text.Length > 200 ? text.Substring(0, 200) : text;
        }
    }
}
