using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// 图表测试集真机套件（udchart2 模式，2026-09-13）：
    /// 针对新版 16-sheet 图表测试集（经典 6 分类 + 高级 10 类 + 看板 6 示例）。
    /// 特点：表头行不固定（各块独立小标题）、多块并排数据、11 高级图A 自带 1 张既有图表（一并纳入截图遍历）。
    /// 基础图表 9+combo 1、高级 10 类（原生四类 pts 语义+PNG 导出）、看板 5 生成+1 矩阵防呆、
    /// pnl_brief 负记账勾稽回归（净利润=190）；S 组截图三维度全覆盖。
    /// 原件只读，产物副本保留交付。
    /// </summary>
    internal static class ChartTestProbe
    {
        private static int _passed, _failed;
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();
        private static dynamic appStatic;

        private static void Expect(bool ok, string name, string detail)
        {
            if (ok) { _passed++; Console.WriteLine("[PASS] " + name + " → " + detail); }
            else { _failed++; Console.WriteLine("[FAIL] " + name + " → " + detail); }
        }

        private static ToolResult Exec(ToolRegistry registry, string tool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }
            try { return descriptor.Execute(args); }
            catch (Exception ex) { return new ToolResult { Ok = false, ErrorCode = "HARNESS_EXC", Data = ex.GetType().Name + " " + ex.Message }; }
        }

        private static Dictionary<string, object> Args(params (string k, object v)[] pairs)
        {
            var d = new Dictionary<string, object>();
            foreach (var p in pairs) { d[p.k] = p.v; }
            return d;
        }

        private static void Release(object value)
        {
            if (value == null) { return; }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
        }

        private static string ColumnLetter(int col)
        {
            string s = "";
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - 1) / 26; }
            return s;
        }

        public static int Run(string sourcePath)
        {
            Console.WriteLine("== 图表测试集真机套件：" + sourcePath + " ==");
            if (!System.IO.File.Exists(sourcePath))
            {
                Console.WriteLine("FILE_NOT_FOUND: " + sourcePath);
                return 2;
            }
            string tempCopy = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "UDC2_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
            System.IO.File.Copy(sourcePath, tempCopy, true);

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
                workbook = excel.Workbooks.Open(tempCopy);
                appStatic = excel;

                var dashboard = new ExcelDashboardService(appStatic);
                var pivotChart = new ExcelPivotChartService(appStatic, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, appStatic, null, null);

                dynamic Ws(string name)
                {
                    // M-D1 纪律：服务工具 FinalRelease 会分离探针缓存的 workbook/sheet RCW，每步从 Application 重取
                    workbook = excel.Workbooks.Item[1];
                    return workbook.Worksheets.Item[name];
                }

                // ---- 基础图表（本套数据块表头行不固定：01@3、02@3、03@3、04@4、05@4/@27、06@2）----
                var chartCases = new[]
                {
                    new[] { "01 比较类", "A3:D8", "ColumnClustered", "" },
                    new[] { "01 比较类", "A3:D8", "BarClustered", "" },
                    new[] { "01 比较类", "A3:D8", "Radar", "" },
                    new[] { "02 趋势类", "A3:D15", "Line", "" },
                    new[] { "02 趋势类", "A3:D15", "Area", "" },
                    new[] { "03 构成层级类", "A3:B7", "Pie", "" },
                    new[] { "03 构成层级类", "A3:B7", "Doughnut", "" },
                    new[] { "03 构成层级类", "A3:B7", "DoughnutExploded", "" },
                    new[] { "04 分布统计类", "A4:E10", "Scatter", "" },
                    new[] { "05 工程金融类", "A4:D12", "Line", "股价无原生类型（散点/折线替代口径，与 udchart 同）" },
                    new[] { "06 地理交互类", "A2:B10", "ColumnClustered", "地图类中国区已停用，柱形替代" },
                };
                foreach (var c in chartCases)
                {
                    string sheetName = c[0], addr = c[1], chartType = c[2], note = c[3];
                    var r = Exec(registry, "create_excel_chart", Args(
                        ("chart_type", chartType), ("source_address", addr), ("sheet", sheetName)));
                    string rStr = Convert.ToString(r.Data);
                    bool ok = r.Ok && (rStr.Contains("chart_name") || rStr.Contains("created = True"));
                    Expect(ok, "基础 " + sheetName + " " + chartType + (note == "" ? "" : "（" + note + "）"),
                        r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }

                // combo 双轴（05 A27:C33：月份|销量|增长率%，次轴挂系列 2）
                {
                    var r = Exec(registry, "create_combo_chart", Args(
                        ("source_address", "A27:C33"), ("sheet", "05 工程金融类"),
                        ("secondary_series", "2"), ("title", "销量vs增长率双轴")));
                    Expect(r.Ok, "基础 05 工程金融类 combo 双轴", r.Ok ? "created" : r.ErrorCode + " " + Convert.ToString(r.Data));
                }

                // ---- 高级图表 10 类（数据块分散于 11/12 表；原生四类 pts 语义 + PNG 导出）----
                var advCases = new[]
                {
                    new[] { "11 高级图A", "A4:E6", "bullet", "①子弹：目标/分档/实际" },
                    new[] { "11 高级图A", "L15:L15", "waffle", "②华夫：单值完成率 62%" },
                    new[] { "05 工程金融类", "A28:C33", "slider", "③滑块：月份|销量|增长率（类别/实际/目标三列契约形态）" },
                    new[] { "11 高级图A", "A36:B41", "funnel", "④漏斗：阶段-人数" },
                    new[] { "11 高级图A", "A51:C55", "butterfly", "⑤蝴蝶：左右两组含负值" },
                    new[] { "12 高级图B", "A4:C9", "gantt", "⑥甘特：任务-开始-工期" },
                    new[] { "12 高级图B", "A15:B22", "waterfall", "⑦瀑布：收支净额含负" },
                    new[] { "12 高级图B", "H15:I20", "treemap", "⑧树状：分类-值" },
                    new[] { "12 高级图B", "A29:C36", "sunburst", "⑨旭日：大区-子类-值" },
                    new[] { "12 高级图B", "H29:I35", "pareto", "⑩帕累托：缺陷-频次" },
                };
                var scratch = System.Environment.GetEnvironmentVariable("BOX_AGENT_SCRATCH_DIR");
                if (string.IsNullOrEmpty(scratch)) { scratch = "."; }
                foreach (var c in advCases)
                {
                    string sheetName = c[0], addr = c[1], type = c[2], note = c[3];
                    var r = Exec(registry, "create_chart_advanced", Args(
                        ("type", type), ("source_address", addr), ("sheet", sheetName)));
                    string rStr = Convert.ToString(r.Data);
                    bool ok = r.Ok && rStr.Contains("adv_type");

                    // 原生四类视觉闭环：pts 语义断言 + PNG 导出（udchart 同口径）
                    string extra = "";
                    if (type == "waterfall" || type == "treemap" || type == "sunburst" || type == "pareto")
                    {
                        try
                        {
                            var wsAdv = Ws(sheetName);  // 服务工具分离 RCW 后重取
                            System.Threading.Thread.Sleep(1500);
                            dynamic coAdv = wsAdv.ChartObjects();
                            int cn = Convert.ToInt32(coAdv.Count);
                            dynamic chAdv = coAdv.Item(cn).Chart;
                            int pts = Convert.ToInt32(chAdv.SeriesCollection().Item(1).Points().Count);
                            string png = System.IO.Path.Combine(scratch, "udchart2_adv_" + type + ".png");
                            coAdv.Item(cn).Activate();
                            chAdv.Export(png, "PNG");
                            extra = " pts=" + pts + " png=" + (System.IO.File.Exists(png) ? "Y" : "N");
                            ok = ok && pts >= 3;
                            Release((object)chAdv); Release((object)coAdv); Release((object)wsAdv);
                        }
                        catch (Exception exv) { extra = " VERIFY-EXC:" + exv.Message; ok = false; }
                    }
                    Expect(ok, "高级 " + sheetName + " " + type + "（" + note + "）" + extra,
                        r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }

                // 高级图负例：文件自带滑块块为「位置|值」单数值形态——工具按契约人话拒绝（MISSING_TARGET_COLUMN）
                {
                    var rNeg = Exec(registry, "create_chart_advanced", Args(
                        ("type", "slider"), ("source_address", "A28:B30"), ("sheet", "11 高级图A")));
                    bool rejected = !rNeg.Ok && Convert.ToString(rNeg.ErrorCode) == "MISSING_TARGET_COLUMN";
                    Expect(rejected, "高级负例 slider 单数值形态必拦截",
                        rejected ? "MISSING_TARGET_COLUMN 已拦截" : rNeg.ErrorCode + " " + Convert.ToString(rNeg.Data));
                }

                // ---- 看板族（本套为示例看板+原始数据块；pnl 为负记账回归样本）----
                {
                    // kpi_brief：20 表 A4:A33 单列数值明细（30 项）；先按缺省列探测，失败再显式 value_column=1
                    var r = Exec(registry, "create_dashboard_kpi_brief", Args(
                        ("source_address", "A4:A33"), ("sheet", "20 看板-KPI摘要"),
                        ("dashboard_sheet", "看板KPI摘要")));
                    if (!r.Ok)
                    {
                        r = Exec(registry, "create_dashboard_kpi_brief", Args(
                            ("source_address", "A4:A33"), ("sheet", "20 看板-KPI摘要"),
                            ("value_column", 1), ("dashboard_sheet", "看板KPI摘要")));
                    }
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板① KPI 摘要（单列明细）", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    // pnl_brief：23 A3:C11，成本为负记账——净利润勾稽回归（总收入 1000，总成本 -810，净利润应=190）
                    var r = Exec(registry, "render_pnl_brief", Args(
                        ("source_address", "A3:C11"), ("sheet", "23 看板-财务损益"),
                        ("dashboard_sheet", "看板财务损益")));
                    string rStr = Convert.ToString(r.Data);
                    bool ok = r.Ok && rStr.Contains("created = True");
                    string extra = "";
                    if (ok)
                    {
                        try
                        {
                            var wsP = Ws("看板财务损益");
                            dynamic cell = wsP.Range["B16"];   // 净利润 KPI 卡显示值（M-D3 版式）
                            double profit = Convert.ToDouble(cell.Value2);
                            Release((object)cell); Release((object)wsP);
                            extra = " 净利润显示=" + profit.ToString("0.##");
                            ok = ok && Math.Abs(profit - 190.0) < 0.01;
                        }
                        catch (Exception exP) { extra = " PROFIT-READ-EXC:" + exP.Message; ok = false; }
                    }
                    Expect(ok, "看板④ 财务损益（负记账净利润勾稽=190）" + extra,
                        r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }
                {
                    var r = Exec(registry, "render_target_track", Args(
                        ("source_address", "A3:C10"), ("sheet", "24 看板-目标达成"),
                        ("dashboard_sheet", "看板目标达成")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板⑤ 目标达成", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    var r = Exec(registry, "render_rank_dynamic", Args(
                        ("source_address", "A5:C19"), ("sheet", "25 看板-排名动态"),
                        ("dashboard_sheet", "看板排名动态")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板⑥ 排名动态", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    // 22 为交叉矩阵输入——必须被 MATRIX_LAYOUT 防呆拦截（本套文件的看板③为静态示例，无可跑长表）
                    var r = Exec(registry, "render_dim_cross", Args(
                        ("source_address", "A4:F9"), ("sheet", "22 看板-多维交叉"),
                        ("row_dim_column", 1), ("col_dim_column", 2), ("value_column", 3),
                        ("dashboard_sheet", "看板多维交叉")));
                    bool rejected = !r.Ok && Convert.ToString(r.ErrorCode) == "MATRIX_LAYOUT";
                    Expect(rejected, "看板③ 多维交叉（矩阵输入必拦截）",
                        rejected ? "MATRIX_LAYOUT 已拦截" : r.ErrorCode + " " + r.Data);
                }

                // ===== S 组：截图三维度全覆盖（自由截数据 / 自由截全部图表 / 看板全景）=====
                string shotsDir = System.Environment.GetEnvironmentVariable("BOX_AGENT_SCRATCH_DIR");
                if (string.IsNullOrEmpty(shotsDir)) { shotsDir = System.IO.Path.GetTempPath(); }
                shotsDir = System.IO.Path.Combine(shotsDir, "udchart2_shots");
                if (System.IO.Directory.Exists(shotsDir)) { try { System.IO.Directory.Delete(shotsDir, true); } catch { } }
                System.IO.Directory.CreateDirectory(shotsDir);
                var shotPngs = new List<string>();

                // S1 自由截全部图表：遍历所有可见 sheet 的全部 ChartObject（含 11 高级图A 既有子弹图），逐图 chart 级截图
                {
                    int chartTotal = 0, chartShotOk = 0;
                    var chartFails = new List<string>();
                    workbook = excel.Workbooks.Item[1];
                    int sheetCount = Convert.ToInt32(workbook.Worksheets.Count);
                    for (int si = 1; si <= sheetCount; si++)
                    {
                        dynamic wsC = null, cosC = null;
                        try
                        {
                            workbook = excel.Workbooks.Item[1];  // 截图工具内部操作会分离 RCW——每 sheet 重取
                            wsC = workbook.Worksheets.Item[si];
                            int vis = 0;
                            try { vis = Convert.ToInt32(wsC.Visible); } catch { }
                            if (vis != -1) { continue; }  // 隐藏临时表（_advsrc_* 等）无图可截
                            string shName = Convert.ToString(wsC.Name);
                            try { wsC.Activate(); } catch { }  // 隐藏实例非活动 sheet ChartObjects 曾 0x800A03EC
                            cosC = wsC.ChartObjects();
                            int cCount = Convert.ToInt32(cosC.Count);
                            for (int ci = 1; ci <= cCount; ci++)
                            {
                                chartTotal++;
                                string cName;
                                dynamic cItem = null;
                                try { cItem = cosC.Item(ci); cName = Convert.ToString(cItem.Name); }
                                catch { cName = "Chart" + ci; }
                                finally { Release((object)cItem); }
                                var sc = Exec(registry, "screenshot_dashboard", Args(
                                    ("dashboard_sheet", shName), ("chart_name", cName), ("out_dir", shotsDir)));
                                bool ok = false; string detail;
                                if (sc.Ok)
                                {
                                    var dS = Newtonsoft.Json.Linq.JObject.FromObject(sc.Data);
                                    string pp = (string)dS["screenshot_path"];
                                    double rt = (double)dS["non_blank_ratio"];
                                    ok = System.IO.File.Exists(pp) && rt > 0.01;
                                    if (ok) { shotPngs.Add(pp); chartShotOk++; detail = "[" + (string)dS["route"] + "] 非白=" + rt.ToString("0.0%") + " " + System.IO.Path.GetFileName(pp); }
                                    else { detail = "产物无效 非白=" + rt.ToString("0.0%"); }
                                }
                                else { detail = sc.ErrorCode + " " + Convert.ToString(sc.Data); }
                                if (!ok) { chartFails.Add(shName + "/" + cName + " → " + detail); }
                            }
                        }
                        catch (Exception exS1) { chartFails.Add("sheet#" + si + " → EXC " + exS1.Message); }
                        finally { Release((object)cosC); Release((object)wsC); }
                    }
                    Expect(chartTotal > 0 && chartShotOk == chartTotal,
                        "S1 自由截全部图表（" + chartShotOk + "/" + chartTotal + "）",
                        chartFails.Count == 0 ? "全部图表逐图导出成功（含既有图表）" : "失败: " + string.Join(" | ", chartFails));
                }

                // S2 自由截数据（address 矩阵：常规块/单列源/矩阵块/宽说明区/股票块）
                var dataCases = new[]
                {
                    new[] { "01 比较类", "A3:D8", "常规比较块" },
                    new[] { "04 分布统计类", "G4:G23", "直方图源单列前 20" },
                    new[] { "22 看板-多维交叉", "A4:F9", "交叉矩阵块" },
                    new[] { "00 说明", "A1:U5", "宽说明区（含合并文本）" },
                    new[] { "05 工程金融类", "A20:B26", "瀑布源小区域" },
                };
                foreach (var dc in dataCases)
                {
                    var sc = Exec(registry, "screenshot_dashboard", Args(
                        ("sheet", dc[0]), ("address", dc[1]), ("out_dir", shotsDir)));
                    bool ok = false; string detail;
                    if (sc.Ok)
                    {
                        var dS = Newtonsoft.Json.Linq.JObject.FromObject(sc.Data);
                        string pp = (string)dS["screenshot_path"];
                        double rt = (double)dS["non_blank_ratio"];
                        ok = System.IO.File.Exists(pp) && rt > 0.01;
                        if (ok) { shotPngs.Add(pp); detail = "[" + (string)dS["route"] + "] 非白=" + rt.ToString("0.0%") + " " + System.IO.Path.GetFileName(pp); }
                        else { detail = "产物无效 非白=" + rt.ToString("0.0%"); }
                    }
                    else { detail = sc.ErrorCode + " " + Convert.ToString(sc.Data); }
                    Expect(ok, "S2 自由截数据 " + dc[2] + " " + dc[0] + "!" + dc[1], detail);
                }

                // S2f 缺省选区截图（未传参 = 截当前活动选区）
                try
                {
                    dynamic wsSel = Ws("03 构成层级类");
                    wsSel.Activate();
                    wsSel.Range["A3:B7"].Select();
                    Release((object)wsSel);
                    var scSel = Exec(registry, "screenshot_dashboard", Args(("out_dir", shotsDir)));
                    bool ok = false; string detail;
                    if (scSel.Ok)
                    {
                        var dS = Newtonsoft.Json.Linq.JObject.FromObject(scSel.Data);
                        string pp = (string)dS["screenshot_path"];
                        double rt = (double)dS["non_blank_ratio"];
                        ok = System.IO.File.Exists(pp) && rt > 0.01;
                        if (ok) { shotPngs.Add(pp); detail = "[" + (string)dS["route"] + "] 非白=" + rt.ToString("0.0%") + " " + System.IO.Path.GetFileName(pp); }
                        else { detail = "产物无效 非白=" + rt.ToString("0.0%"); }
                    }
                    else { detail = scSel.ErrorCode + " " + Convert.ToString(scSel.Data); }
                    Expect(ok, "S2f 缺省选区截图（选中哪儿截哪儿）", detail);
                }
                catch (Exception exSel) { Expect(false, "S2f 缺省选区截图（选中哪儿截哪儿）", "EXC " + exSel.Message); }

                // S2g 空表截图防御负例（不得假成功）
                {
                    workbook = excel.Workbooks.Item[1];
                    dynamic wsEmpty = workbook.Worksheets.Add();
                    string emptyName = Convert.ToString(wsEmpty.Name);
                    Release((object)wsEmpty);
                    var neg = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", emptyName), ("out_dir", shotsDir)));
                    bool negOk = !neg.Ok && (Convert.ToString(neg.ErrorCode) == "EMPTY_BOARD" || Convert.ToString(neg.ErrorCode) == "SCREENSHOT_BLANK" || Convert.ToString(neg.ErrorCode) == "SCREENSHOT_FAILED");
                    Expect(negOk, "S2g 空表截图防御负例", neg.Ok ? "未拦截（假成功）！" : Convert.ToString(neg.ErrorCode));
                    try
                    {
                        workbook = excel.Workbooks.Item[1];
                        workbook.Worksheets.Item[emptyName].Delete();
                    }
                    catch { }
                }

                // S3 看板全景截图（4 生成看板 + 6 示例看板 sheet）
                foreach (var db in new[] { "看板KPI摘要", "看板财务损益", "看板目标达成", "看板排名动态",
                    "20 看板-KPI摘要", "21 看板-销售概览", "22 看板-多维交叉", "23 看板-财务损益", "24 看板-目标达成", "25 看板-排名动态" })
                {
                    var sc = Exec(registry, "screenshot_dashboard", Args(("dashboard_sheet", db), ("out_dir", shotsDir)));
                    bool ok = false; string detail;
                    if (sc.Ok)
                    {
                        var dS = Newtonsoft.Json.Linq.JObject.FromObject(sc.Data);
                        string pp = (string)dS["screenshot_path"];
                        double rt = (double)dS["non_blank_ratio"];
                        ok = System.IO.File.Exists(pp) && rt > 0.01;
                        if (ok) { shotPngs.Add(pp); detail = "[" + (string)dS["route"] + "] 非白=" + rt.ToString("0.0%") + " " + System.IO.Path.GetFileName(pp); }
                        else { detail = "产物无效 非白=" + rt.ToString("0.0%"); }
                    }
                    else { detail = sc.ErrorCode + " " + Convert.ToString(sc.Data); }
                    Expect(ok, "S3 看板全景截图 " + db, detail);
                }

                // PNG 清单落盘供 AI 读图审查
                try
                {
                    System.IO.File.WriteAllLines(System.IO.Path.Combine(shotsDir, "manifest.txt"), shotPngs);
                    Console.WriteLine("[shots] " + shotPngs.Count + " 张 → " + shotsDir);
                }
                catch { }

                Console.WriteLine("\n== 图表测试集套件结果：通过 " + _passed + "，失败 " + _failed + " ==");

                // 产物保留：另存到任务输出目录交付
                try
                {
                    workbook = excel.Workbooks.Item[1];
                    string outDirSave = System.Environment.GetEnvironmentVariable("BOX_AGENT_OUTPUT_DIR");
                    if (string.IsNullOrEmpty(outDirSave)) { outDirSave = System.IO.Path.GetTempPath(); }
                    string outPath = System.IO.Path.Combine(outDirSave, "图表测试集真机产物.xlsx");
                    if (System.IO.File.Exists(outPath)) { System.IO.File.Delete(outPath); }
                    workbook.SaveAs(outPath, 51);
                    Console.WriteLine("[saved] " + outPath);
                }
                catch (Exception svEx) { Console.WriteLine("[saved][WARN] 产物另存失败：" + svEx.Message); }

                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][UDC2] " + ex.GetType().Name + " " + ex.Message);
                return 2;
            }
            finally
            {
                Console.WriteLine("[cleanup][UDC2] 清理 Excel 进程...");
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
