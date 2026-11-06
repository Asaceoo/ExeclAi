using System;
using System.Linq;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// 全功能真机测试（full 模式）：对测试数据工作簿跑六类看板 + 全部技巧型/基础图表，
    /// 产物保留（副本生成），每项带 PASS/FAIL 断言与 deep 核验。
    /// 用法：SlicerProbe.exe full "源.xlsx" "产物.xlsx"
    /// </summary>
    internal static class FullFeatureProbe
    {
        private static int _passed;
        private static int _failed;
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static void Expect(bool ok, string name, string detail)
        {
            if (ok)
            {
                _passed++;
                Console.WriteLine("[PASS] " + name + (string.IsNullOrEmpty(detail) ? "" : " → " + detail));
            }
            else
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name + " → " + detail);
            }
        }

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

        private static string DeepFailDetail(string deepJson)
        {
            if (deepJson.Contains("\"verdict\":\"ok\"")) { return "ok"; }
            var sb = new System.Text.StringBuilder();
            int from = 0;
            while (true)
            {
                int p = deepJson.IndexOf("\"Status\":\"fail\"", from, StringComparison.Ordinal);
                if (p < 0) { break; }
                int itemStart = deepJson.LastIndexOf("\"Item\":\"", p, StringComparison.Ordinal);
                sb.Append('[').Append(deepJson, itemStart + 8, Math.Min(150, deepJson.Length - itemStart - 8)).Append("] ");
                from = p + 10;
                if (sb.Length > 420) { break; }
            }
            return sb.Length > 0 ? sb.ToString() : deepJson.Substring(0, Math.Min(200, deepJson.Length));
        }

        private static Dictionary<string, object> Args(params (string k, object v)[] pairs)
        {
            var map = new Dictionary<string, object>();
            foreach (var (k, v) in pairs) { map[k] = v; }
            return map;
        }

        public static int Run(string sourcePath, string outPath)
        {
            Console.WriteLine("== 全功能真机测试：" + sourcePath + " ==");
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
                var dashboard = new ExcelDashboardService(excel);
                var pivotChart = new ExcelPivotChartService(excel, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, excel, null, null);

                string det = Convert.ToString(workbook.Worksheets.Item["销售明细"].Name);
                string fin = "财务损益", tgt = "目标达成", prj = "项目进度", fun = "转化漏斗",
                       bug = "缺陷统计", tree = "多级构成", demo = "人群对比", dept = "部门达成", pnl = "月度损益";

                // ===== 看板组（6）=====
                var ov = Exec(registry, "render_dashboard", Args(
                    ("source_address", "A1:I181"), ("sheet", det),
                    ("month_column", 1), ("label_column", 4), ("category_column", 5),
                    ("value_column", 9), ("quantity_column", 7),
                    ("dashboard_sheet", "销售经营看板"), ("theme", "orange")));
                var ovS = Convert.ToString(ov.Data);
                Expect(ov.Ok, "B1 render_dashboard sales_overview", ov.Ok ? "months/categories 见返回" : ov.ErrorCode + " " + ovS);
                var shot = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", "销售经营看板")));
                string shotDetail;
                bool shotOk = false;
                if (shot.Ok)
                {
                    var dShot = Newtonsoft.Json.Linq.JObject.FromObject(shot.Data);
                    string pngPath = (string)dShot["screenshot_path"];
                    double ratio = (double)dShot["non_blank_ratio"];
                    shotOk = System.IO.File.Exists(pngPath) && ratio > 0.01;
                    shotDetail = "[" + (string)dShot["route"] + "] " + System.IO.Path.GetFileName(pngPath) + " 非白=" + ratio.ToString("0.0%");
                }
                else { shotDetail = shot.ErrorCode + " " + Convert.ToString(shot.Data); }
                Expect(shotOk, "B1-shot screenshot_dashboard 看板截图", shotDetail);

                var deep1 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "销售经营看板")));
                var d1 = Newtonsoft.Json.JsonConvert.SerializeObject(deep1.Data);
                Expect(deep1.Ok && d1.Contains("\"verdict\":\"ok\""), "B1-deep 销售经营看板", d1.Contains("verdict") ? "verdict=" + (d1.Contains("\"verdict\":\"ok\"") ? "ok" : "mismatch") : "no-json");

                var kpi = Exec(registry, "create_dashboard_kpi_brief", Args(
                    ("source_address", "A1:I181"), ("sheet", det),
                    ("label_column", 4), ("value_column", 9),
                    ("dashboard_sheet", "KPI速览看板"), ("theme", "blue")));
                Expect(kpi.Ok, "B2 create_dashboard_kpi_brief", kpi.Ok ? "created" : kpi.ErrorCode + " " + Convert.ToString(kpi.Data));
                var deep2 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "KPI速览看板")));
                var d2 = Newtonsoft.Json.JsonConvert.SerializeObject(deep2.Data);
                Expect(deep2.Ok && d2.Contains("\"verdict\":\"ok\""), "B2-deep KPI速览看板", DeepFailDetail(d2));

                var pnlB = Exec(registry, "render_pnl_brief", Args(
                    ("source_address", "A1:C9"), ("sheet", fin),
                    ("dashboard_sheet", "财务损益看板"), ("theme", "blue")));
                Expect(pnlB.Ok, "B3 render_pnl_brief", pnlB.Ok ? "created" : pnlB.ErrorCode + " " + Convert.ToString(pnlB.Data));
                var deep3 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "财务损益看板")));
                var d3 = Newtonsoft.Json.JsonConvert.SerializeObject(deep3.Data);
                Expect(deep3.Ok && d3.Contains("\"verdict\":\"ok\""), "B3-deep 财务损益看板", d3.Contains("\"verdict\":\"ok\"") ? "ok" : d3.Length > 200 ? d3.Substring(0, 200) : d3);

                var tgtB = Exec(registry, "render_target_track", Args(
                    ("source_address", "A1:D6"), ("sheet", tgt),
                    ("dashboard_sheet", "目标达成看板"), ("theme", "graygreen")));
                Expect(tgtB.Ok, "B4 render_target_track", tgtB.Ok ? "created" : tgtB.ErrorCode + " " + Convert.ToString(tgtB.Data));
                var deep4 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "目标达成看板")));
                var d4 = Newtonsoft.Json.JsonConvert.SerializeObject(deep4.Data);
                Expect(deep4.Ok && d4.Contains("\"verdict\":\"ok\""), "B4-deep 目标达成看板", d4.Contains("\"verdict\":\"ok\"") ? "ok" : d4.Length > 200 ? d4.Substring(0, 200) : d4);

                var dimB = Exec(registry, "render_dim_cross", Args(
                    ("source_address", "A1:I181"), ("sheet", det),
                    ("row_dim_column", 5), ("col_dim_column", 1), ("value_column", 9),
                    ("dashboard_sheet", "多维交叉看板"), ("theme", "blue")));
                Expect(dimB.Ok, "B5 render_dim_cross", dimB.Ok ? "created" : dimB.ErrorCode + " " + Convert.ToString(dimB.Data));
                var deep5 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "多维交叉看板")));
                var d5 = Newtonsoft.Json.JsonConvert.SerializeObject(deep5.Data);
                Expect(deep5.Ok && d5.Contains("\"verdict\":\"ok\""), "B5-deep 多维交叉看板", DeepFailDetail(d5));

                var rkB = Exec(registry, "render_rank_dynamic", Args(
                    ("source_address", "A1:I181"), ("sheet", det),
                    ("category_column", 4), ("time_column", 1), ("value_column", 9), ("top_n", 5),
                    ("dashboard_sheet", "排名动态看板"), ("theme", "blue")));
                Expect(rkB.Ok, "B6 render_rank_dynamic", rkB.Ok ? "created" : rkB.ErrorCode + " " + Convert.ToString(rkB.Data));
                var deep6 = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "排名动态看板")));
                var d6 = Newtonsoft.Json.JsonConvert.SerializeObject(deep6.Data);
                Expect(deep6.Ok && d6.Contains("\"verdict\":\"ok\""), "B6-deep 排名动态看板", d6.Contains("\"verdict\":\"ok\"") ? "ok" : d6.Length > 200 ? d6.Substring(0, 200) : d6);

                // ===== 技巧型图表组（11）=====
                var cGantt = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C11"), ("sheet", prj), ("type", "gantt"), ("title", "项目甘特图")));
                Expect(cGantt.Ok, "C1 gantt 甘特图", cGantt.Ok ? "created" : cGantt.ErrorCode + " " + Convert.ToString(cGantt.Data));
                var cFun = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:B7"), ("sheet", fun), ("type", "funnel"), ("title", "转化漏斗")));
                Expect(cFun.Ok, "C2 funnel 漏斗图", cFun.Ok ? "created" : cFun.ErrorCode + " " + Convert.ToString(cFun.Data));
                var cPar = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:B8"), ("sheet", bug), ("type", "pareto"), ("title", "缺陷排列图")));
                Expect(cPar.Ok, "C3 pareto 排列图", cPar.Ok ? "created" : cPar.ErrorCode + " " + Convert.ToString(cPar.Data));
                var cHist = Exec(registry, "create_chart_advanced", Args(("source_address", "A10:A130"), ("sheet", bug), ("type", "histogram"), ("title", "响应时长分布")));
                Expect(cHist.Ok, "C4 histogram 直方图", cHist.Ok ? "created" : cHist.ErrorCode + " " + Convert.ToString(cHist.Data));
                var cTree = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C10"), ("sheet", tree), ("type", "treemap"), ("title", "销售构成树状图")));
                Expect(cTree.Ok, "C5 treemap 树状图", cTree.Ok ? "created" : cTree.ErrorCode + " " + Convert.ToString(cTree.Data));
                var cSun = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C10"), ("sheet", tree), ("type", "sunburst"), ("title", "销售构成旭日图")));
                Expect(cSun.Ok, "C6 sunburst 旭日图", cSun.Ok ? "created" : cSun.ErrorCode + " " + Convert.ToString(cSun.Data));
                var cBuf = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C7"), ("sheet", demo), ("type", "butterfly"), ("title", "用户性别对比")));
                Expect(cBuf.Ok, "C7 butterfly 蝴蝶图", cBuf.Ok ? "created" : cBuf.ErrorCode + " " + Convert.ToString(cBuf.Data));
                var cBul = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C7"), ("sheet", dept), ("type", "bullet"), ("title", "部门达成子弹图")));
                Expect(cBul.Ok, "C8 bullet 子弹图", cBul.Ok ? "created" : cBul.ErrorCode + " " + Convert.ToString(cBul.Data));
                var cSli = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:C7"), ("sheet", dept), ("type", "slider"), ("title", "部门达成滑珠图")));
                Expect(cSli.Ok, "C9 slider 滑珠图", cSli.Ok ? "created" : cSli.ErrorCode + " " + Convert.ToString(cSli.Data));
                var cWaf = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:D7"), ("sheet", dept), ("type", "waffle"), ("title", "完成率华夫图")));
                Expect(cWaf.Ok, "C10 waffle 华夫图", cWaf.Ok ? "created" : cWaf.ErrorCode + " " + Convert.ToString(cWaf.Data));
                var cWf = Exec(registry, "create_chart_advanced", Args(("source_address", "A1:B8"), ("sheet", pnl), ("type", "waterfall"), ("title", "月度损益瀑布图")));
                Expect(cWf.Ok, "C11 waterfall 瀑布图", cWf.Ok ? "created" : cWf.ErrorCode + " " + Convert.ToString(cWf.Data));

                // ===== 基础图表组（4）=====
                var bChart = Exec(registry, "create_excel_chart", Args(("source_address", "A1:B8"), ("sheet", bug), ("chart_type", "ColumnClustered"), ("title", "缺陷分布柱形图")));
                Expect(bChart.Ok, "D1 create_excel_chart 柱形图", bChart.Ok ? "created" : bChart.ErrorCode + " " + Convert.ToString(bChart.Data));
                var bCombo = Exec(registry, "create_combo_chart", Args(("source_address", "A1:C7"), ("sheet", dept), ("secondary_series", "2"), ("title", "实际vs目标双轴")));
                Expect(bCombo.Ok, "D2 create_combo_chart 双轴", bCombo.Ok ? "created" : bCombo.ErrorCode + " " + Convert.ToString(bCombo.Data));
                var bDough = Exec(registry, "create_doughnut", Args(("source_address", "A1:B7"), ("sheet", fun), ("title", "转化占比环形")));
                Expect(bDough.Ok, "D3 create_doughnut 环形", bDough.Ok ? "created" : bDough.ErrorCode + " " + Convert.ToString(bDough.Data));
                var bTrunc = Exec(registry, "create_truncated_column", Args(("source_address", "A1:B8"), ("sheet", bug), ("title", "缺陷截断柱")));
                Expect(bTrunc.Ok, "D4 create_truncated_column 截断柱", bTrunc.Ok ? "created" : bTrunc.ErrorCode + " " + Convert.ToString(bTrunc.Data));

                // ===== F 组：多切片器联动（field_column="4,5"：地区+类别 同时筛选）=====
                var multi = Exec(registry, "set_linkage", Args(
                    ("mode", "native"), ("source_address", "A1:I181"), ("sheet", det),
                    ("field_column", "4,5"), ("value_column", 9),
                    ("month_column", 1), ("category_column", 5), ("quantity_column", 7),
                    ("dashboard_sheet", "销售经营看板")));
                var multiStr = Convert.ToString(multi.Data);
                Expect(multi.Ok, "F1 set_linkage 多切片器(地区+类别)", multi.Ok ? multiStr.Substring(0, Math.Min(150, multiStr.Length)) : multi.ErrorCode + " " + multiStr);
                workbook = excel.ActiveWorkbook;  // linkage 失败回滚可能分离 RCW——重取
                var deepM = Exec(registry, "verify_dashboard_deep", Args(("dashboard_sheet", "销售经营看板")));
                var dmStr = Newtonsoft.Json.JsonConvert.SerializeObject(deepM.Data);
                // 多维掩码下 KPI/显示值守恒必须仍成立
                Expect(deepM.Ok && dmStr.Contains("\"verdict\":\"ok\""), "F1-deep 多切片器联动后深度核验", DeepFailDetail(dmStr));
                // 值级断言：多维掩码失效时 KPI 会"一致地全 0"（deep 判 ok 但数据是死的）——销售总额必须 >0
                // diag3：两个维度透视的行区实际值（定位多维掩码失效）
                try
                {
                    dynamic stgD3 = workbook.Worksheets.Item["_销售经营看板_calc"];
                    for (int pi = 0; pi < 2; pi++)
                    {
                        dynamic pt = stgD3.PivotTables("_linkage_pivot_销售经营看板_" + pi);
                        dynamic tr = pt.TableRange1;
                        Console.WriteLine("[diag3] pivot" + pi + " addr=" + Convert.ToString(tr.Address));
                        for (int rr = 1; rr <= 3; rr++)
                        {
                            for (int cc = 1; cc <= 2; cc++)
                            {
                                dynamic c3 = tr.Cells.Item[rr, cc];
                                Console.WriteLine("  [diag3] R" + rr + "C" + cc + " = " + Convert.ToString(c3.Value2));
                                Release((object)c3);
                            }
                        }
                        Release((object)tr); Release((object)pt);
                    }
                    Release((object)stgD3);
                }
                catch (Exception exD3) { Console.WriteLine("[diag3] 失败: " + exD3.Message); }

                dynamic kpiB5 = workbook.Worksheets.Item["销售经营看板"].Range["B5"];
                double kpiVal = Convert.ToDouble(kpiB5.Value2);
                Release((object)kpiB5);
                Expect(kpiVal > 0, "F1-value 多切片器下销售总额>0", "B5=" + kpiVal.ToString("N0"));
                var shotM = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", "销售经营看板")));
                string shotMDetail;
                bool shotMOk = false;
                if (shotM.Ok)
                {
                    var dSM = Newtonsoft.Json.Linq.JObject.FromObject(shotM.Data);
                    string ppM = (string)dSM["screenshot_path"];
                    double rtM = (double)dSM["non_blank_ratio"];
                    shotMOk = System.IO.File.Exists(ppM) && rtM > 0.01;
                    shotMDetail = "[" + (string)dSM["route"] + "] " + System.IO.Path.GetFileName(ppM) + " 非白=" + rtM.ToString("0.0%");
                }
                else { shotMDetail = shotM.ErrorCode + " " + Convert.ToString(shotM.Data); }
                Expect(shotMOk, "F1-shot 多切片器看板截图", shotMDetail);

                // ===== E 组：自检工具 + 图表级截图 + 空白防御负例 =====
                string barChartName = null, doughChartName = null;
                try { barChartName = (string)Newtonsoft.Json.Linq.JObject.FromObject(bChart.Data)["chart_name"]; } catch { }
                try { doughChartName = (string)Newtonsoft.Json.Linq.JObject.FromObject(bDough.Data)["chart_name"]; } catch { }

                var v1 = Exec(registry, "verify_chart", Args(
                    ("sheet", bug), ("chart_name", barChartName ?? ""),
                    ("expected_chart_type", "ColumnClustered"), ("expected_title", "缺陷分布柱形图")));
                var v1s = Convert.ToString(v1.Data);
                Expect(v1.Ok, "E1 verify_chart 柱形图自检", v1.Ok ? v1s.Substring(0, Math.Min(110, v1s.Length)) : v1.ErrorCode + " " + v1s);

                var v2 = Exec(registry, "verify_chart", Args(
                    ("sheet", fun), ("chart_name", doughChartName ?? ""),
                    ("expected_chart_type", "Doughnut")));
                Expect(v2.Ok, "E2 verify_chart 环形图自检", v2.Ok ? "ok" : v2.ErrorCode + " " + Convert.ToString(v2.Data));

                foreach (var (nm, cn, sh) in new[] { ("E3", barChartName, bug), ("E4", doughChartName, fun) })
                {
                    if (string.IsNullOrEmpty(cn)) { Expect(false, nm + " chart级截图", "chart_name 缺失"); continue; }
                    var sc = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", sh), ("chart_name", cn)));
                    bool scOk = false; string scDetail;
                    if (sc.Ok)
                    {
                        var dS = Newtonsoft.Json.Linq.JObject.FromObject(sc.Data);
                        string pp = (string)dS["screenshot_path"];
                        double rt = (double)dS["non_blank_ratio"];
                        scOk = System.IO.File.Exists(pp) && rt > 0.01;
                        scDetail = "[" + (string)dS["route"] + "] " + System.IO.Path.GetFileName(pp) + " 非白=" + rt.ToString("0.0%");
                    }
                    else { scDetail = sc.ErrorCode + " " + Convert.ToString(sc.Data); }
                    Expect(scOk, nm + " chart级截图 " + cn, scDetail);
                }

                // E6 选区截图（address 模式）：截销售明细 A1:I5（用户"截当前选区数据"场景）
                var selShot = Exec(registry, "screenshot_dashboard", Args(
                    ("sheet", det), ("address", "A1:I5")));
                bool selOk = false; string selDetail;
                if (selShot.Ok)
                {
                    var dSl = Newtonsoft.Json.Linq.JObject.FromObject(selShot.Data);
                    string ppS = (string)dSl["screenshot_path"];
                    double rtS = (double)dSl["non_blank_ratio"];
                    selOk = System.IO.File.Exists(ppS) && rtS > 0.01;
                    selDetail = "[" + (string)dSl["route"] + "] " + System.IO.Path.GetFileName(ppS) + " 非白=" + rtS.ToString("0.0%");
                }
                else { selDetail = selShot.ErrorCode + " " + Convert.ToString(selShot.Data); }
                Expect(selOk, "E6 选区截图(address 模式)", selDetail);

                workbook = excel.ActiveWorkbook;  // screenshot 内部操作会分离 RCW——重取再建空表
                dynamic emptySheet = workbook.Worksheets.Add();
                string emptyName = Convert.ToString(emptySheet.Name);
                Release((object)emptySheet);
                var neg = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", emptyName)));
                // 负例本质 = 不得假成功：空表可能以 EMPTY_BOARD/SCREENSHOT_BLANK/SCREENSHOT_FAILED 任一形态拒绝
                bool negOk = !neg.Ok && (Convert.ToString(neg.ErrorCode) == "EMPTY_BOARD" || Convert.ToString(neg.ErrorCode) == "SCREENSHOT_BLANK" || Convert.ToString(neg.ErrorCode) == "SCREENSHOT_FAILED");
                Expect(negOk, "E5 空白截图防御负例", neg.Ok ? "未拦截（假成功）！" : neg.ErrorCode);

                workbook = excel.ActiveWorkbook;  // 原生图表工具会分离 Open 缓存的 RCW——重取再保存
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
