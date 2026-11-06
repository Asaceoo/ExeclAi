using System;
using System.Linq;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// 基础图表全类型真机测试（charts12 模式）：create_excel_chart 全部 12 种 chart_type 逐一渲染断言，产物保留。
    /// 用法：SlicerProbe.exe charts12 "源.xlsx" "产物.xlsx"
    /// </summary>
    internal static class ChartAllProbe
    {
        private static int _passed;
        private static int _failed;
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static void Release(object o)
        {
            if (o != null && System.Runtime.InteropServices.Marshal.IsComObject(o))
            {
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(o); } catch { }
            }
        }

        private static ToolResult Exec(ToolRegistry registry, string tool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }
            try { return descriptor.Execute(args); }
            catch (Exception ex) { return new ToolResult { Ok = false, ErrorCode = "HOST_EXCEPTION", Data = ex.Message }; }
        }

        private static Dictionary<string, object> Args(params (string k, object v)[] pairs)
        {
            var map = new Dictionary<string, object>();
            foreach (var (k, v) in pairs) { map[k] = v; }
            return map;
        }

        /// <summary>对照实验：干净工作簿单字段透视 @Z1 vs @第200列，验证 200 列位置是否导致行区不全。</summary>
        public static int RunSlicerProbe()
        {
            dynamic excel = null;
            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false; excel.DisplayAlerts = false;
                dynamic wb = excel.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                string[] regions = { "华东", "华北", "华南" };
                for (int r = 2; r <= 19; r++)
                {
                    ws.Cells.Item[r, 1].Value2 = regions[(r - 2) % 3];
                    ws.Cells.Item[r, 2].Value2 = r * 10;
                }
                ws.Cells.Item[1, 1].Value2 = "地区"; ws.Cells.Item[1, 2].Value2 = "金额";
                dynamic cache = wb.PivotCaches().Create(SourceType: 1, SourceData: "Sheet1!R1C1:R19C2");

                foreach (var (col, tag) in new[] { (26, "Z1"), (200, "GR200") })
                {
                    dynamic pt = ws.PivotTables().Add(PivotCache: cache, TableDestination: ws.Cells.Item[1, col]);
                    pt.Name = "P_" + tag;
                    pt.PivotFields("地区").Orientation = 1;
                    dynamic tr = pt.TableRange1;
                    string items = "";
                    for (int i = 2; i <= Convert.ToInt32(tr.Rows.Count); i++)
                    {
                        dynamic c = tr.Cells.Item[i, 1];
                        items += Convert.ToString(c.Value2) + "|";
                        Release((object)c);
                    }
                    Console.WriteLine("[cmp] @" + tag + " 行区: " + items + " (rows=" + Convert.ToInt32(tr.Rows.Count) + ")");
                    Release((object)tr); Release((object)pt);
                }

                // 复刻 md1：扩量至 201 行（混入其他用例数据形态）→ 建透视 → 切片器全流程
                for (int r = 20; r <= 201; r++)
                {
                    ws.Cells.Item[r, 1].Value2 = "垃圾" + r;
                    ws.Cells.Item[r, 4].Value2 = r;
                }
                dynamic cache2 = wb.PivotCaches().Create(SourceType: 1, SourceData: "Sheet1!R1C1:R201C4");
                dynamic ptB = ws.PivotTables().Add(PivotCache: cache2, TableDestination: ws.Cells.Item[1, 240]);
                ptB.Name = "P_Wide";
                ptB.PivotFields("地区").Orientation = 1;
                dynamic scB = wb.SlicerCaches().Add2(ptB, "地区");
                dynamic slB = scB.Slicers().Add(SlicerDestination: ws, Top: 6, Left: 900, Width: 144, Height: 192);
                // 只选华东
                dynamic itemsB = scB.SlicerItems();
                for (int i = 1; i <= Convert.ToInt32(itemsB.Count); i++)
                {
                    dynamic it = itemsB.Item(i);
                    it.Selected = Convert.ToString(it.Name) == "华东";
                    Release((object)it);
                }
                Release((object)itemsB);
                excel.CalculateFullRebuild();
                dynamic trB = ptB.TableRange1;
                string itemsBStr = "";
                for (int i = 2; i <= Convert.ToInt32(trB.Rows.Count); i++)
                {
                    dynamic c = trB.Cells.Item[i, 1];
                    itemsBStr += Convert.ToString(c.Value2) + "|";
                    Release((object)c);
                }
                Console.WriteLine("[cmp] Wide切片后行区: " + itemsBStr + " (rows=" + Convert.ToInt32(trB.Rows.Count) + ")");
                Release((object)trB); Release((object)ptB); Release((object)scB); Release((object)slB);
                return 0;
            }
            catch (Exception ex) { Console.WriteLine("[fatal] " + ex.Message); return 3; }
            finally { try { excel.Quit(); } catch { } }
        }

        /// <summary>运行时暴力探测 XlChartType 权威值（项目全晚绑定，无 interop 程序集可查）。</summary>
        public static int RunEnumProbe()
        {
            dynamic excel = null;
            try
            {
                var before = System.Diagnostics.Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                { if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); } p.Dispose(); }
                dynamic wb = excel.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                ws.Range["A1"].Value2 = "类"; ws.Range["B1"].Value2 = "值";
                ws.Range["A2"].Value2 = "a"; ws.Range["B2"].Value2 = 1;
                dynamic co = ws.ChartObjects().Add(10, 10, 260, 180);
                dynamic ch = co.Chart;
                ch.SetSourceData(ws.Range["A1:B2"]);
                int[] cands = { -4121, -4120, -4122, -4151, -4152, -4153, -4154, -4155 };
                foreach (var v in cands)
                {
                    try
                    {
                        ch.ChartType = v;
                        Console.WriteLine("[enum] " + v + " OK, 读回=" + Convert.ToInt32(ch.ChartType));
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("[enum] " + v + " E_FAIL(无效)");
                    }
                }
                return 0;
            }
            catch (Exception ex) { Console.WriteLine("[fatal] " + ex.Message); return 3; }
            finally
            {
                foreach (var pid in SpawnedPids)
                { try { var pr = System.Diagnostics.Process.GetProcessById(pid); pr.Kill(); } catch { } }
            }
        }

        public static int Run(string sourcePath, string outPath)
        {
            Console.WriteLine("== 基础图表全类型真机测试：" + sourcePath + " ==");
            if (!System.IO.File.Exists(sourcePath))
            {
                Console.WriteLine("FILE_NOT_FOUND: " + sourcePath);
                return 2;
            }
            System.IO.File.Copy(sourcePath, outPath, true);
            dynamic excel = null;
            dynamic workbook = null;
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
                workbook = excel.Workbooks.Open(outPath);
                var pivotChart = new ExcelPivotChartService(excel, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard: null);

                var cases = new (string type, string sheet, string addr, string title)[]
                {
                    ("ColumnClustered",  "月度销售", "A1:C13",  "月度销售柱形图"),
                    ("Line",             "月度销售", "A1:C13",  "月度销售折线图"),
                    ("Area",             "月度销售", "A1:C13",  "月度销售面积图"),
                    ("BarClustered",     "地区排行", "A1:B6",   "地区排行条形图"),
                    ("Pie",              "品类构成", "A1:B6",   "品类构成饼图"),
                    ("Doughnut",         "品类构成", "A1:B6",   "品类构成环形图"),
                    ("DoughnutExploded", "品类构成", "A1:B6",   "品类构成分离环形"),
                    ("Scatter",          "XY观测",   "A1:B25",  "广告-销售额散点图"),
                    ("Histogram",        "响应时长", "A1:A101", "响应时长直方图"),
                    ("Radar",            "服务能力", "A1:C7",   "服务能力雷达图"),
                    ("RadarFilled",      "服务能力", "A1:C7",   "服务能力填充雷达"),
                    ("RadarMarkers",     "服务能力", "A1:C7",   "服务能力标记雷达"),
                };

                foreach (var (type, sheet, addr, title) in cases)
                {
                    var r = Exec(registry, "create_excel_chart", Args(
                        ("source_address", addr), ("sheet", sheet),
                        ("chart_type", type), ("title", title)));
                    string detail;
                    if (r.Ok)
                    {
                        var data = Newtonsoft.Json.Linq.JObject.FromObject(r.Data);
                        detail = (string)data["chart_name"] + " @ " + (string)data["anchor"];
                    }
                    else
                    {
                        detail = r.ErrorCode + " " + Convert.ToString(r.Data);
                    }
                    if (r.Ok) { _passed++; Console.WriteLine("[PASS] " + type + " → " + detail); }
                    else { _failed++; Console.WriteLine("[FAIL] " + type + " → " + detail); }
                }

                // D5 补测试缺口：create_doughnut 的分离型变体（-4121 臆造值修复后回归）
                var dExp = Exec(registry, "create_doughnut", Args(
                    ("source_address", "A1:B6"), ("sheet", "品类构成"),
                    ("chart_type", "doughnutexploded"), ("title", "品类构成分离环形")));
                if (dExp.Ok) { _passed++; Console.WriteLine("[PASS] create_doughnut doughnutexploded → created"); }
                else { _failed++; Console.WriteLine("[FAIL] create_doughnut doughnutexploded → " + dExp.ErrorCode + " " + Convert.ToString(dExp.Data)); }

                workbook = excel.ActiveWorkbook;  // 图表工具可能分离 RCW——重取再保存
                workbook.Save();
                Console.WriteLine("[saved] " + outPath);
                Console.WriteLine("== 汇总：通过 " + _passed + "，失败 " + _failed + " ==");
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception exTop)
            {
                Console.WriteLine("[fatal] " + exTop.GetType().Name + " " + exTop.Message);
                return 3;
            }
            finally
            {
                foreach (var pid in SpawnedPids)
                {
                    try { var pr = System.Diagnostics.Process.GetProcessById(pid); pr.Kill(); } catch { }
                }
            }
        }
    }
}
