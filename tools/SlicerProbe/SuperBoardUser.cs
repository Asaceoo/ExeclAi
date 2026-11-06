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
        /// SB2 用户真实文件超级看板验收（测试2026 (4).xlsx）：
        /// 明细 B10:S27（B 日期/F 品类/I 数量/K 金额/M 利润/O 销售员），月份用日期列。
        /// 断言：生成/结构/守恒（源直和）/指标切换/维度筛选/隔离/截图。
        /// </summary>
        public static int RunUserFile(string sourcePath)
        {
            Console.WriteLine("== SB2 用户文件超级看板：" + sourcePath + " ==");
            if (!System.IO.File.Exists(sourcePath)) { Console.WriteLine("FILE_NOT_FOUND"); return 2; }
            string workPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "SB2_user_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".xlsx");
            System.IO.File.Copy(sourcePath, workPath, true);

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
                workbook = excel.Workbooks.Open(workPath);
                appLocalStatic = excel;
                dynamic ws = workbook.Worksheets.Item["Sheet1"];

                // 用户明细直和基准（B11:S27 数据区；K=金额列11、M=利润列13）
                double amountSum = 0, profitSum = 0;
                var catAmount = new Dictionary<string, double>();
                for (int r = 11; r <= 27; r++)
                {
                    double amt = Convert.ToDouble(ws.Cells.Item[r, 11].Value2);
                    double pf = Convert.ToDouble(ws.Cells.Item[r, 13].Value2);
                    string cat = Convert.ToString(ws.Cells.Item[r, 6].Value2);
                    amountSum += amt; profitSum += pf;
                    if (!catAmount.ContainsKey(cat)) { catAmount[cat] = 0; }
                    catAmount[cat] += amt;
                }
                string firstCat = catAmount.Keys.OrderBy(k => k).First();
                Console.WriteLine("[diag] 金额和=" + amountSum + " 利润和=" + profitSum + " 品类数=" + catAmount.Count + " 首品类=" + firstCat + " 直和=" + catAmount[firstCat]);

                var dashboard = new ExcelDashboardService(appLocalStatic);
                var pivotChart = new ExcelPivotChartService(appLocalStatic, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, appLocalStatic, null, null);

                // ---- U1 生成（月份用日期列 B；利润列 M=相对12；含 extra_kpis + 透视分析）----
                var extra = Newtonsoft.Json.JsonConvert.SerializeObject(new object[]
                {
                    new Dictionary<string, object> { ["name"] = "百万金额", ["formula"] = "销售总额/1000000", ["format"] = "0.00" }
                });
                var gen = ExecSuper(registry, "render_super_dashboard", new Dictionary<string, object>
                {
                    ["source_address"] = "B10:S27",
                    ["sheet"] = "Sheet1",
                    ["dashboard_sheet"] = "超级看板",
                    ["label_column"] = 14,       // O 销售员（相对 B 列起第 14 列）
                    ["category_column"] = 5,     // F 产品类别
                    ["month_column"] = 1,        // B 日期（月份维度）
                    ["value_column"] = 10,       // K 金额
                    ["quantity_column"] = 8,     // I 数量
                    ["profit_column"] = 12,      // M 利润
                    ["extra_kpis"] = extra,
                    ["pivot_analysis"] = "true"
                });
                string genJson = Newtonsoft.Json.JsonConvert.SerializeObject(gen.Data);
                Expect(gen.Ok, "U1 用户文件生成", gen.Ok ? genJson.Substring(0, Math.Min(260, genJson.Length)) : gen.ErrorCode + " " + gen.Data);
                if (!gen.Ok) { Console.WriteLine("== SB2 提前结束 =="); return 1; }

                // ---- U2 结构 ----
                dynamic board = workbook.Worksheets.Item["超级看板"];
                dynamic staging = workbook.Worksheets.Item["_超级看板_calc"];
                bool layers = true;
                foreach (var name in new[] { "超级数据表", "超级看板", "_超级看板_calc", "透视分析-超级看板" })
                {
                    bool exists = false;
                    foreach (dynamic s in workbook.Worksheets) { if (Convert.ToString(s.Name) == name) { exists = true; } }
                    layers &= exists;
                }
                int chartCount = Convert.ToInt32(board.ChartObjects().Count);
                int shapeCount = Convert.ToInt32(board.Shapes.Count);
                Expect(layers && chartCount == 4 && shapeCount >= 16,
                    "U2 三层结构+四宫格+形状", "charts=" + chartCount + " shapes=" + shapeCount);

                // ---- U3 守恒（源直和 vs staging）----
                double tTotal = ToD(staging.Range["T2"].Value2);
                double uTotal = ToD(staging.Range["U2"].Value2);
                double zSales = ToD(staging.Range["Z2"].Value2);
                Expect(Math.Abs(tTotal - amountSum) < 0.01 && Math.Abs(uTotal - profitSum) < 0.01 && Math.Abs(zSales - amountSum) < 0.01,
                    "U3 守恒对账", "T2=" + tTotal + "(期望" + amountSum + ") U2=" + uTotal + "(期望" + profitSum + ") Z2=" + zSales);

                // ---- U4 指标切换（利润视图/自定义视图）----
                board.Range["H4"].Value2 = "利润视图";
                double zProfit = ToD(staging.Range["Z2"].Value2);
                string titleAfter = ToS(staging.Range["AA2"].Value2);
                board.Range["H4"].Value2 = "自定义视图";
                double zExtra = ToD(staging.Range["Z2"].Value2);
                string titleExtra = ToS(staging.Range["AA2"].Value2);
                board.Range["H4"].Value2 = "销售视图";
                Expect(Math.Abs(zProfit - profitSum) < 0.01 && titleAfter == "利润总额"
                    && Math.Abs(zExtra - amountSum / 1000000.0) < 0.0001 && titleExtra == "百万金额",
                    "U4 指标切换守恒", "zProfit=" + zProfit + " titleAfter=" + titleAfter + " zExtra=" + zExtra + " titleExtra=" + titleExtra);

                // ---- U5 维度筛选（品类=首品类，Z2=该品类总额）----
                board.Range["C4"].Value2 = firstCat;
                double zFiltered = ToD(staging.Range["Z2"].Value2);
                board.Range["C4"].Value2 = "全部";
                Expect(Math.Abs(zFiltered - catAmount[firstCat]) < 0.01,
                    "U5 维度筛选联动", firstCat + " 直和=" + catAmount[firstCat] + " Z2=" + zFiltered);

                // ---- U6 透视分析隔离 ----
                workbook.Worksheets.Item["透视分析-超级看板"].Range["B2"].Value2 = "被篡改";
                double zIso = ToD(staging.Range["Z2"].Value2);
                Expect(Math.Abs(zIso - amountSum) < 0.01, "U6 透视分析隔离", "篡改后 Z2=" + zIso);

                // ---- U7 截图 ----
                var shot = ExecSuper(registry, "screenshot_dashboard", new Dictionary<string, object>
                {
                    ["dashboard_sheet"] = "超级看板"
                });
                string shotJson = Newtonsoft.Json.JsonConvert.SerializeObject(shot.Data ?? new Newtonsoft.Json.Linq.JObject());
                Expect(shot.Ok && shotJson.Contains(".png"), "U7 截图", shot.Ok ? shotJson.Substring(0, Math.Min(200, shotJson.Length)) : shot.ErrorCode + " " + shot.Data);

                Console.WriteLine("== SB2 结果：通过 " + _passed + "，失败 " + _failed + " ==");
                workbook.SaveAs(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SB2_user_super_board.xlsx"));
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][SB2] " + ex.GetType().Name + " " + ex.Message);
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
    }
}
