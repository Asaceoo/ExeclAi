using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using ExcelAiAssistant.Core.Models;

namespace SlicerProbe
{
    /// <summary>
    /// 产品场景判别实验（expw 模式，2026-09-12）：在用户文件副本上手工 AddChart2+SetSourceData，
    /// 二分定位"判别实验成功/产品真失败"的触发条件：
    /// W1 现状复刻（A4:B10）；W2 删除上方说明行/合并单元格后重试；W3 数据搬新表 A1 起重试。
    /// </summary>
    internal static class ExpWorkbookProbe
    {
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static void Release(object o)
        {
            if (o == null) { return; }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(o); } catch { }
        }

        private static (int pts, string f) ReadSeries(object chartObj)
        {
            dynamic chart = chartObj;
            try
            {
                dynamic sc = chart.SeriesCollection();
                if (Convert.ToInt32(sc.Count) < 1) { return (-1, "NOSERIES"); }
                dynamic ser = sc.Item(1);
                int pts = Convert.ToInt32(ser.Points().Count);
                string f;
                try { f = Convert.ToString(ser.Formula); } catch { f = "NOFORMULA"; }
                if (f.Length > 100) { f = f.Substring(0, 100); }
                return (pts, f);
            }
            catch (Exception ex) { return (-2, "EXC:" + ex.Message); }
        }

        private static void ExportPng(dynamic wb, dynamic co, string path)
        {
            try
            {
                dynamic windows = wb.Windows;
                if (Convert.ToInt32(windows.Count) > 0) { windows.Item(1).Visible = true; }
                co.Activate();
                co.Chart.Export(path, "PNG");
            }
            catch (Exception ex) { Console.WriteLine("  [export EXC] " + ex.Message); }
        }

        private static void RunCase(dynamic ws, string tag, string addr, string pngDir)
        {
            dynamic shp = ws.Shapes.AddChart2(-1, 119, 100, 100, 420, 280);
            string shpName = Convert.ToString(shp.Name);
            dynamic co = ws.ChartObjects().Item(shpName);
            dynamic chart = co.Chart;
            string err = "";
            dynamic rng = ws.Range[addr];
            try { chart.SetSourceData(rng, 2); } catch { err = "E_FAIL"; }
            Release((object)rng);
            System.Threading.Thread.Sleep(800);
            var rs = ReadSeries((object)chart);
            ExportPng(ws.Parent, co, System.IO.Path.Combine(pngDir, "expw_" + tag + ".png"));
            Console.WriteLine($"W {tag}: pts={rs.pts} f={rs.f} {err}");
            shp.Delete();
            Release((object)co); Release((object)chart); Release((object)shp);
        }

        public static int Run(string sourcePath, string pngDir)
        {
            Console.WriteLine("== 产品场景判别实验 ==");
            string tempCopy = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "EXPW_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
            System.IO.File.Copy(sourcePath, tempCopy, true);
            dynamic excel = null;
            try
            {
                var before = System.Diagnostics.Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); }
                    p.Dispose();
                }
                dynamic wb = excel.Workbooks.Open(tempCopy);
                dynamic ws = wb.Worksheets.Item["高级07_瀑布图"];

                // W0 sheet 状态诊断
                dynamic coAll = ws.ChartObjects();
                int coN = Convert.ToInt32(coAll.Count);
                dynamic loAll = ws.ListObjects();
                int loN = Convert.ToInt32(loAll.Count);
                dynamic names = ws.Names;
                int nmN = Convert.ToInt32(names.Count);
                Console.WriteLine($"W0 chartobjects={coN} listobjects={loN} names={nmN}");
                for (int i = 1; i <= coN; i++)
                {
                    dynamic c = coAll.Item(i);
                    Console.WriteLine($"  co[{i}] name={Convert.ToString(c.Name)} type={Convert.ToInt32(c.Chart.ChartType)}");
                    Release((object)c);
                }
                Release((object)coAll); Release((object)loAll); Release((object)names);
                dynamic uRange = ws.UsedRange;
                Console.WriteLine($"  used={Convert.ToString(uRange.Address)}");
                Release((object)uRange);

                // W1 现状复刻：数据 A4:B10，上方 3 行说明（含合并单元格）
                RunCase(ws, "w1_asis", "A4:B10", pngDir);

                // W2 删除上方说明行（1:3 整行删除，合并单元格一并清除），数据上移到 A1:B7
                dynamic rows13 = ws.Range["1:3"];
                rows13.Delete();
                Release((object)rows13);
                RunCase(ws, "w2_nosay", "A1:B7", pngDir);

                // W4 重命名 sheet 为纯 ASCII 再试（二分 sheet 名因素）
                ws.Name = "EXPW4";
                RunCase(ws, "w4_rename", "A1:B7", pngDir);

                // W5 读回坏绑定的实际内容（绑到了哪几行？）
                {
                    dynamic shp = ws.Shapes.AddChart2(-1, 119, 100, 100, 420, 280);
                    string shpName = Convert.ToString(shp.Name);
                    dynamic co = ws.ChartObjects().Item(shpName);
                    dynamic chart = co.Chart;
                    try { chart.SetSourceData(ws.Range["A1:B7"], 2); } catch { }
                    System.Threading.Thread.Sleep(800);
                    dynamic ser = chart.SeriesCollection().Item(1);
                    string xv = "";
                    try
                    {
                        object[] xa = ser.XValues as object[];
                        if (xa != null) { xv = "X=[" + string.Join(",", xa) + "]"; }
                        else { xv = "X=null"; }
                    }
                    catch (Exception ex5) { xv = "X-EXC:" + ex5.Message; }
                    Console.WriteLine("W5 badbind " + xv);
                    shp.Delete();
                    Release((object)ser); Release((object)co); Release((object)chart); Release((object)shp);
                }

                // W6 两段式：先对新 sheet 区域 SetSourceData（成功），再改绑原 sheet 区域
                {
                    dynamic shp = ws.Shapes.AddChart2(-1, 119, 100, 100, 420, 280);
                    string shpName = Convert.ToString(shp.Name);
                    dynamic co = ws.ChartObjects().Item(shpName);
                    dynamic chart = co.Chart;
                    dynamic wsTmp = wb.Worksheets.Add();  // 会话内新建（SetSourceData 已验证在此类 sheet 上生效）
                    wsTmp.Name = "EXPWTMP_" + Guid.NewGuid().ToString("N").Substring(0, 4);
                    wsTmp.Cells.Item[1, 1].Value2 = "item";
                    wsTmp.Cells.Item[1, 2].Value2 = "val";
                    var tmpVals = new object[] { "a", 45.0, "b", 30.0, "c", 18.0, "d", 12.0, "e", 6.0, "f", 4.0 };
                    for (int r = 0; r < 6; r++)
                    {
                        wsTmp.Cells.Item[r + 2, 1].Value2 = tmpVals[r * 2];
                        wsTmp.Cells.Item[r + 2, 2].Value2 = tmpVals[r * 2 + 1];
                    }
                    string step1 = "ok";
                    try { chart.SetSourceData(wsTmp.Range["A1:B7"], 2); } catch (Exception ex6) { step1 = "EXC:" + ex6.Message; }
                    System.Threading.Thread.Sleep(600);
                    var rs1 = ReadSeries((object)chart);
                    string step2 = "ok";
                    try { chart.SetSourceData(ws.Range["A1:B7"], 2); } catch (Exception ex6) { step2 = "EXC:" + ex6.Message; }
                    System.Threading.Thread.Sleep(600);
                    var rs2 = ReadSeries((object)chart);
                    Console.WriteLine($"W6 twostep: step1={step1} pts1={rs1.pts} | step2={step2} pts2={rs2.pts}");
                    shp.Delete();
                    Release((object)co); Release((object)chart); Release((object)shp);
                }

                // W7 修复路径验证：临时 sheet 建图+绑定成功 → Cut/Paste 回用户 sheet → 读回+导图
                {
                    dynamic wsTmp7 = wb.Worksheets.Add();
                    wsTmp7.Name = "EXPW7TMP";
                    wsTmp7.Cells.Item[1, 1].Value2 = "item";
                    wsTmp7.Cells.Item[1, 2].Value2 = "val";
                    var v7 = new object[] { "a", 1000.0, "b", 400.0, "c", -650.0, "d", -80.0, "e", -40.0, "f", 630.0 };
                    for (int r = 0; r < 6; r++)
                    {
                        wsTmp7.Cells.Item[r + 2, 1].Value2 = v7[r * 2];
                        wsTmp7.Cells.Item[r + 2, 2].Value2 = v7[r * 2 + 1];
                    }
                    dynamic shp7 = wsTmp7.Shapes.AddChart2(-1, 119, 100, 100, 420, 280);
                    string n7 = Convert.ToString(shp7.Name);
                    dynamic co7 = wsTmp7.ChartObjects().Item(n7);
                    dynamic ch7 = co7.Chart;
                    try { ch7.SetSourceData(wsTmp7.Range["A1:B7"], 2); } catch { }
                    System.Threading.Thread.Sleep(800);
                    var rs7 = ReadSeries((object)ch7);
                    // 移回用户 sheet
                    co7.Cut();
                    dynamic wsDst = wb.Worksheets.Item["EXPW4"];  // 原"高级07"已改名
                    wsDst.Activate();
                    wsDst.Paste();
                    System.Threading.Thread.Sleep(400);
                    dynamic coAll2 = wsDst.ChartObjects();
                    int cn = Convert.ToInt32(coAll2.Count);
                    dynamic coNew = coAll2.Item(cn);
                    coNew.Left = 120; coNew.Top = 130;
                    System.Threading.Thread.Sleep(600);
                    var rs8 = ReadSeries((object)coNew.Chart);
                    ExportPng(wb, coNew, System.IO.Path.Combine(pngDir, "expw_w7_pasted.png"));
                    Console.WriteLine($"W7 tmp pts={rs7.pts} | pasted pts={rs8.pts}");
                    // W8：粘贴回后把源改绑回用户 sheet 区域（成功则临时 sheet 可删，实现最干净）
                    string w8 = "ok";
                    try { coNew.Chart.SetSourceData(wsDst.Range["A1:B7"], 2); } catch (Exception ex8) { w8 = "EXC:" + ex8.Message; }
                    System.Threading.Thread.Sleep(800);
                    var rs9 = ReadSeries((object)coNew.Chart);
                    ExportPng(wb, coNew, System.IO.Path.Combine(pngDir, "expw_w8_rebind.png"));
                    Console.WriteLine($"W8 rebind: {w8} pts={rs9.pts}");
                    Release((object)coNew); Release((object)coAll); Release((object)ch7); Release((object)co7); Release((object)shp7);
                }

                // W9 干净上下文单跑产品 waterfall（区分"产品 native 代码坏"vs"前置用例污染"）
                {
                    var dashboard = new ExcelDashboardService(excel);
                    var pivotChart = new ExcelPivotChartService(excel, null, null);
                    var registry = new ToolRegistry();
                    Func<Dictionary<string, object>, ToolResult> stub = _ =>
                        new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                    ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                    var r = registry.TryGet("create_chart_advanced", out var desc)
                        ? desc.Execute(new Dictionary<string, object>
                        {
                            ["type"] = "waterfall", ["source_address"] = "A4:B10", ["sheet"] = "高级07_瀑布图"
                        })
                        : new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL" };
                    Console.WriteLine("W9 product waterfall: ok=" + r.Ok + " code=" + r.ErrorCode + " data=" + Convert.ToString(r.Data));
                    // 读回 pts + 导 PNG
                    dynamic ws9 = wb.Worksheets.Item["高级07_瀑布图"];
                    dynamic co9 = ws9.ChartObjects();
                    if (Convert.ToInt32(co9.Count) > 0)
                    {
                        dynamic ch9 = co9.Item(Convert.ToInt32(co9.Count)).Chart;
                        int pts9 = Convert.ToInt32(ch9.SeriesCollection().Item(1).Points().Count);
                        dynamic wbWin = wb.Windows;
                        if (Convert.ToInt32(wbWin.Count) > 0) { wbWin.Item(1).Visible = true; }
                        co9.Item(Convert.ToInt32(co9.Count)).Activate();
                        ch9.Export(System.IO.Path.Combine(pngDir, "expw_w9_product.png"), "PNG");
                        Console.WriteLine("W9 pts=" + pts9);
                        Release((object)ch9); Release((object)wbWin);
                    }
                    Release((object)co9); Release((object)ws9);
                }

                // W3 数据搬到全新 sheet A1 起（排除 sheet 状态/名字影响）
                dynamic ws2 = wb.Worksheets.Add();
                ws2.Name = "EXPW3";
                ws2.Cells.Item[1, 1].Value2 = "item";
                ws2.Cells.Item[1, 2].Value2 = "val";
                var vals = new object[] { "a", 1000.0, "b", 400.0, "c", -650.0, "d", -80.0, "e", -40.0, "f", 630.0 };
                for (int r = 0; r < 6; r++)
                {
                    ws2.Cells.Item[r + 2, 1].Value2 = vals[r * 2];
                    ws2.Cells.Item[r + 2, 2].Value2 = vals[r * 2 + 1];
                }
                RunCase(ws2, "w3_newsheet", "A1:B7", pngDir);

                Console.WriteLine("== 判别完成 ==");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal] " + ex.GetType().Name + " " + ex.Message);
                return 2;
            }
            finally
            {
                foreach (var process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    if (SpawnedPids.Contains(process.Id)) { process.Kill(); }
                    process.Dispose();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { System.IO.File.Delete(tempCopy); } catch { }
            }
        }
    }
}
