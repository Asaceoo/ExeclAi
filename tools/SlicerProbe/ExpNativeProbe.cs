using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core;
using ExcelAiAssistant.Core.Tools;

namespace SlicerProbe
{
    /// <summary>
    /// 原生新图表（waterfall/treemap/sunburst/pareto/histogram/boxwhisker）数据绑定判别实验（2026-09-12）：
    /// 用户实证 Series.Values 逐点赋值路径渲染全坏（pts 错乱、Formula 空）。
    /// 实验：E1 SetSourceData(含表头, xlColumns)；E2 删默认系列后 SetSourceData；E3 Series.Formula 直写 SERIES 公式。
    /// 每实验读回 pts/Formula 并 Export PNG 供视觉核验。
    /// </summary>
    internal static class ExpNativeProbe
    {
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static void Set(dynamic ws, int r, int c, object v)
        {
            dynamic cell = ws.Cells.Item[r, c];
            cell.Value2 = v;
            Release(cell);
        }

        private static void Release(object o)
        {
            if (o == null) { return; }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(o); } catch { }
        }

        private static (int pts, string f) ReadSeries(object chartObj)  // object 入参：dynamic 元组传染边界（踩坑纪律）
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
                if (f.Length > 120) { f = f.Substring(0, 120); }
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
                dynamic chart = co.Chart;
                chart.Export(path, "PNG");
            }
            catch (Exception ex) { Console.WriteLine("  [export EXC] " + ex.Message); }
        }

        public static int Run(string outDir)
        {
            Console.WriteLine("== 原生图表绑定判别实验 ==");
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
                dynamic wb = excel.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];

                // 数据：A:B = item|val（6 行，pareto/waterfall/histogram/boxwhisker）；D:E = cat|val（7 行，treemap/sunburst）
                Set(ws, 1, 1, "item"); Set(ws, 1, 2, "val");
                var d = new object[] { "a", 45.0, "b", 30.0, "c", 18.0, "d", 12.0, "e", 6.0, "f", 4.0 };
                for (int r = 0; r < 6; r++) { Set(ws, r + 2, 1, d[r * 2]); Set(ws, r + 2, 2, d[r * 2 + 1]); }
                Set(ws, 1, 4, "cat"); Set(ws, 1, 5, "val");
                var d2 = new object[] { "x1", 220.0, "x2", 150.0, "x3", 180.0, "x4", 60.0, "x5", 180.0, "x6", 120.0, "x7", 90.0 };
                for (int r = 0; r < 7; r++) { Set(ws, r + 2, 4, d2[r * 2]); Set(ws, r + 2, 5, d2[r * 2 + 1]); }

                dynamic rngAB = ws.Range["A1:B7"];
                dynamic rngDE = ws.Range["D1:E8"];

                var cases = new[]
                {
                    new[] { "waterfall", "119", "A" }, new[] { "pareto", "122", "A" },
                    new[] { "histogram", "118", "A" }, new[] { "boxwhisker", "121", "A" },
                    new[] { "treemap", "117", "D" }, new[] { "sunburst", "120", "D" },
                };

                foreach (var c in cases)
                {
                    string name = c[0]; int ctype = int.Parse(c[1], CultureInfo.InvariantCulture);
                    dynamic rng = c[2] == "A" ? rngAB : rngDE;

                    // E1: SetSourceData(含表头, xlColumns=2)
                    dynamic shp = ws.Shapes.AddChart2(-1, ctype, 100, 100, 420, 280);
                    string shpName = Convert.ToString(shp.Name);
                    dynamic co = ws.ChartObjects().Item(shpName);
                    dynamic chart = co.Chart;
                    string e1 = "ok"; string e1f = "";
                    try { chart.SetSourceData(rng, 2); } catch (Exception ex) { e1 = "EXC:" + ex.Message; }
                    System.Threading.Thread.Sleep(400);
                    var rs = ReadSeries((object)chart);
                    ExportPng(wb, co, System.IO.Path.Combine(outDir, "exp_" + name + "_setsrc.png"));
                    Console.WriteLine($"E1 {name}: setsrc={e1} pts={rs.pts} f={rs.f} {e1f}");
                    shp.Delete();
                    Release(co); Release(chart); Release(shp);

                    // E3: Series.Formula 直写
                    shp = ws.Shapes.AddChart2(-1, ctype, 100, 100, 420, 280);
                    shpName = Convert.ToString(shp.Name);
                    co = ws.ChartObjects().Item(shpName);
                    chart = co.Chart;
                    string e3 = "ok"; string addr = c[2] == "A" ? "A2:A7" : "D2:D8"; string vaddr = c[2] == "A" ? "B2:B7" : "E2:E8";
                    try { chart.SeriesCollection().Item(1).Formula = $"=SERIES(Sheet1!$B$1,Sheet1!${addr[0]}${addr.Substring(1)},Sheet1!${vaddr[0]}${vaddr.Substring(1)},1)"; }
                    catch (Exception ex) { e3 = "EXC:" + ex.Message; }
                    System.Threading.Thread.Sleep(400);
                    rs = ReadSeries(chart);
                    ExportPng(wb, co, System.IO.Path.Combine(outDir, "exp_" + name + "_formula.png"));
                    Console.WriteLine($"E3 {name}: formula={e3} pts={rs.pts} f={rs.f}");
                    shp.Delete();
                    Release(co); Release(chart); Release(shp);
                }

                Console.WriteLine("== 判别实验完成 ==");
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
            }
        }
    }
}
