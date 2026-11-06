using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    internal static partial class SuperBoardProbe
    {
        /// <summary>
        /// SB1 超级看板真机全链验收（需求-超级看板.md P1 AC-1~AC-4）：
        /// 生成（含 extra_kpis/利润列/透视分析开关）→ 三层结构 → 守恒对账 → 指标切换守恒（AC-3）
        /// → 维度筛选联动（AC-2）→ 透视分析隔离（AC-4）→ 截图（AC-1 视觉闭环输入）。
        /// </summary>
        public static int RunFull(string outPath)
        {
            Console.WriteLine("== SB1 超级看板真机全链 ==");
            dynamic excel = null;
            dynamic workbook = null;
            try
            {
                var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); }
                    p.Dispose();
                }
                workbook = excel.Workbooks.Add();
                var appLocal = excel;
                appLocalStatic = excel;

                // ---- 测试明细：月份×品类×地区×金额×数量×利润 24 行 ----
                dynamic ws = workbook.Worksheets.Item[1];
                ws.Name = "明细";
                string[] headers = { "月份", "品类", "地区", "金额", "数量", "利润" };
                for (int c = 0; c < 6; c++) { ws.Cells.Item[1, c + 1].Value2 = headers[c]; }
                var rnd = new Random(42);
                string[] months = { "1月", "2月", "3月", "4月" };
                string[] cats = { "产品1", "产品2" };
                string[] regions = { "北京", "上海" };
                double amountSum = 0, profitSum = 0;
                int r2 = 2;
                foreach (var m in months)
                    foreach (var cat in cats)
                        foreach (var rg in regions)
                        {
                            double amt = 1000 * r2 + rnd.Next(100);
                            double qty = 10 + (r2 % 7);
                            double profit = amt * 0.4;
                            ws.Cells.Item[r2, 1].Value2 = m;
                            ws.Cells.Item[r2, 2].Value2 = cat;
                            ws.Cells.Item[r2, 3].Value2 = rg;
                            ws.Cells.Item[r2, 4].Value2 = amt;
                            ws.Cells.Item[r2, 5].Value2 = qty;
                            ws.Cells.Item[r2, 6].Value2 = profit;
                            amountSum += amt; profitSum += profit;
                            r2++;
                        }
                int detailLast = r2 - 1;
                Console.WriteLine("[diag] 明细 " + (detailLast - 1) + " 行，金额和=" + amountSum + " 利润和=" + profitSum);

                var dashboard = new ExcelDashboardService(appLocal);
                var pivotChart = new ExcelPivotChartService(appLocal, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, appLocal, null, null);

                // ---- T1 生成（extra_kpis + 利润列 + 透视分析开关）----
                var extra = Newtonsoft.Json.JsonConvert.SerializeObject(new object[]
                {
                    new Dictionary<string, object> { ["name"] = "百万金额", ["formula"] = "销售总额/1000000", ["format"] = "0.00" }
                });
                var genArgs = new Dictionary<string, object>
                {
                    ["source_address"] = "A1:F" + detailLast,
                    ["sheet"] = "明细",
                    ["dashboard_sheet"] = "超级看板",
                    ["label_column"] = 3,      // 地区列（明细 A=月份 B=品类 C=地区）
                    ["category_column"] = 2,   // 品类列
                    ["month_column"] = 1,      // 月份列
                    ["value_column"] = 4,      // 金额列
                    ["quantity_column"] = 5,   // 数量列
                    ["profit_column"] = 6,     // 利润列
                    ["extra_kpis"] = extra,
                    ["pivot_analysis"] = "true"
                };
                var gen = ExecSuper(registry, "render_super_dashboard", genArgs);
                string genJson = Newtonsoft.Json.JsonConvert.SerializeObject(gen.Data);
                Expect(gen.Ok && Convert.ToString(gen.Data).Contains("created = True"),
                    "T1 生成", gen.Ok ? genJson.Substring(0, Math.Min(300, genJson.Length)) : gen.ErrorCode + " " + gen.Data);

                // ---- T2 三层结构 ----
                bool layers = true;
                foreach (var name in new[] { "超级数据表", "超级看板", "_超级看板_calc", "透视分析-超级看板" })
                {
                    bool exists = false;
                    foreach (dynamic s in workbook.Worksheets) { if (Convert.ToString(s.Name) == name) { exists = true; } }
                    layers &= exists;
                }
                dynamic board = workbook.Worksheets.Item["超级看板"];
                int chartCount = Convert.ToInt32(board.ChartObjects().Count);
                int shapeCount = Convert.ToInt32(board.Shapes.Count);
                string viewCell = Convert.ToString(board.Range["H4"].Value2);
                Expect(layers && chartCount == 4 && shapeCount >= 16 && viewCell == "销售视图",
                    "T2 三层结构+四宫格+KPI 形状", "charts=" + chartCount + " shapes=" + shapeCount + " view=" + viewCell);

                // ---- T3 守恒对账（staging KPI vs 源明细直和；AC-2 基础）----
                dynamic staging = workbook.Worksheets.Item["_超级看板_calc"];
                double tTotal = ToD(staging.Range["T2"].Value2);
                double uTotal = ToD(staging.Range["U2"].Value2);
                double zSales = ToD(staging.Range["Z2"].Value2);
                Expect(Math.Abs(tTotal - amountSum) < 0.01 && Math.Abs(uTotal - profitSum) < 0.01 && Math.Abs(zSales - amountSum) < 0.01,
                    "T3 守恒对账", "T2=" + tTotal + "(期望" + amountSum + ") U2=" + uTotal + "(期望" + profitSum + ") Z2=" + zSales);

                // ---- T4 指标切换守恒（AC-3）----
                board.Range["H4"].Value2 = "利润视图";
                double zProfit = ToD(staging.Range["Z2"].Value2);
                string titleAfter = ToS(staging.Range["AA2"].Value2);
                board.Range["H4"].Value2 = "自定义视图";
                double zExtra = ToD(staging.Range["Z2"].Value2);
                string titleExtra = ToS(staging.Range["AA2"].Value2);
                board.Range["H4"].Value2 = "销售视图";
                Expect(Math.Abs(zProfit - profitSum) < 0.01 && titleAfter == "利润总额"
                    && Math.Abs(zExtra - amountSum / 1000000.0) < 0.0001 && titleExtra == "百万金额",
                    "T4 指标切换守恒", "zProfit=" + zProfit + " titleAfter=" + titleAfter + " zExtra=" + zExtra + " titleExtra=" + titleExtra);

                // ---- T5 维度筛选联动（AC-2）：品类=产品1 后，D2=「1月×产品1」金额、Z2=品类总额 ----
                board.Range["C4"].Value2 = "产品1";
                double catMonthSum = 0, catSum = 0;
                for (int rr = 2; rr <= detailLast; rr++)
                {
                    if (Convert.ToString(ws.Cells.Item[rr, 2].Value2) == "产品1")
                    {
                        catSum += ToD(ws.Cells.Item[rr, 4].Value2);
                        if (Convert.ToString(ws.Cells.Item[rr, 1].Value2) == "1月") { catMonthSum += ToD(ws.Cells.Item[rr, 4].Value2); }
                    }
                }
                double dAfter = ToD(staging.Range["D2"].Value2);
                double zFiltered = ToD(staging.Range["Z2"].Value2);
                Console.WriteLine("[diag] T5 D2公式=" + Convert.ToString(staging.Range["D2"].Formula)
                    + " | stagingC2=" + ToS(staging.Range["C2"].Value2)
                    + " | 看板C2=" + ToS(board.Range["C4"].Value2)
                    + " | 看板E2=" + ToS(board.Range["E2"].Value2)
                    + " | D2=" + dAfter + " D3=" + ToD(staging.Range["D3"].Value2));
                board.Range["C4"].Value2 = "全部";
                Expect(Math.Abs(dAfter - catMonthSum) < 0.01 && Math.Abs(zFiltered - catSum) < 0.01,
                    "T5 维度筛选联动", "1月×产品1 直和=" + catMonthSum + " D2=" + dAfter + "；产品1 直和=" + catSum + " Z2=" + zFiltered);

                // ---- T6 透视分析隔离（AC-4）----
                workbook.Worksheets.Item["透视分析-超级看板"].Range["B2"].Value2 = "被篡改";
                double zIsolated = ToD(staging.Range["Z2"].Value2);
                Expect(Math.Abs(zIsolated - amountSum) < 0.01,
                    "T6 透视分析隔离", "篡改透视分析后 Z2=" + zIsolated + "（期望仍 " + amountSum + "）");

                // ---- T7 截图（AC-1 视觉闭环输入）----
                var shot = ExecSuper(registry, "screenshot_dashboard", new Dictionary<string, object>
                {
                    ["dashboard_sheet"] = "超级看板"
                });
                string shotJson = Newtonsoft.Json.JsonConvert.SerializeObject(shot.Data ?? new Newtonsoft.Json.Linq.JObject());
                bool shotOk = shot.Ok && shotJson.Contains(".png");
                Expect(shotOk, "T7 截图", shot.Ok ? shotJson.Substring(0, Math.Min(220, shotJson.Length)) : shot.ErrorCode + " " + shot.Data);

                Console.WriteLine("== SB1 结果：通过 " + _passed + "，失败 " + _failed + " ==");
                workbook.SaveAs(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SB1_super_board.xlsx"));
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][SB1] " + ex.GetType().Name + " " + ex.Message);
                return 3;
            }
            finally
            {
                foreach (var pid in SpawnedPids)
                {
                    try { Process.GetProcessById(pid)?.Kill(); } catch { }
                }
                SpawnedPids.Clear();
            }
        }

        private static ToolResult ExecSuper(ToolRegistry registry, string tool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }
            try { return descriptor.Execute(args); }
            catch (Exception ex) { return new ToolResult { Ok = false, ErrorCode = "HOST_EXCEPTION", Data = ex.GetType().Name + " " + ex.Message }; }
        }

        private static double ToD(object v)
        {
            return v == null ? double.NaN : Convert.ToDouble(v);
        }

        private static string ToS(object v)
        {
            return v == null ? "" : Convert.ToString(v);
        }
    }
}
