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
    /// 用户图表生成测试数据全表真机套件（udchart 模式，2026-09-12）：
    /// 针对 24-sheet 全表测试数据（每表 3 行说明 + 第 4 行表头），按 sheet 名逐一映射工具生成 + 断言。
    /// 基础图表 12 类（气泡/股价无原生类型，用散点/折线替代并备注）、高级图表 10 类、看板 6 类、超级数据表链路。
    /// 原件只读，产物副本保留交付。
    /// </summary>
    internal static class UserChartProbe
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
            // 减计数释放（FinalRelease 会断开同对象全部 RCW——跨用例重复取同 sheet 必炸 InvalidComObjectException，踩坑手册 §RCW 纪律）
            if (value == null) { return; }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
        }

        /// <summary>按 sheet 名取数据区地址（固定：第 4 行表头，末行取 UsedRange 实测）。</summary>
        private static string DataAddress(dynamic sheet, out int lastRow)
        {
            dynamic used = sheet.UsedRange;
            int startRow = Convert.ToInt32(used.Row);
            lastRow = startRow + Convert.ToInt32(used.Rows.Count) - 1;
            int colCount = Convert.ToInt32(used.Columns.Count);
            string lastColL = ColumnLetter(Math.Max(1, Convert.ToInt32(used.Column) + colCount - 1));
            Release((object)used);
            return "A4:" + lastColL + lastRow;
        }

        private static string ColumnLetter(int col)
        {
            string s = "";
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - 1) / 26; }
            return s;
        }

        public static int Run(string sourcePath)
        {
            Console.WriteLine("== 用户图表生成全表真机套件：" + sourcePath + " ==");
            if (!System.IO.File.Exists(sourcePath))
            {
                Console.WriteLine("FILE_NOT_FOUND: " + sourcePath);
                return 2;
            }
            string tempCopy = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "UDC_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
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
                    // M-D1 纪律：服务工具 FinalRelease 会分离探针缓存的 workbook/sheet RCW，每步从 Application 重取（ud 套件同款）
                    workbook = excel.Workbooks.Item[1];
                    return workbook.Worksheets.Item[name];
                }

                // ---- 基础图表 12 类型（数据驱动循环；气泡/股价无原生类型用替代并备注）----
                var chartCases = new[]
                {
                    new[] { "基础01_比较_柱形条形", "ColumnClustered", "" },
                    new[] { "基础01_比较_柱形条形", "BarClustered", "" },
                    new[] { "基础02_比较_雷达", "Radar", "" },
                    new[] { "基础02_比较_雷达", "RadarFilled", "" },
                    new[] { "基础02_比较_雷达", "RadarMarkers", "" },
                    new[] { "基础03_趋势_折线面积", "Line", "" },
                    new[] { "基础03_趋势_折线面积", "Area", "" },
                    new[] { "基础04_构成_饼图圆环", "Pie", "" },
                    new[] { "基础04_构成_饼图圆环", "Doughnut", "" },
                    new[] { "基础04_构成_饼图圆环", "DoughnutExploded", "" },
                    new[] { "基础05_分布_散点", "Scatter", "" },
                    new[] { "基础06_分布_气泡", "Scatter", "气泡无原生类型（工具矩阵 12 类不含 bubble），散点替代" },
                    new[] { "基础07_金融_股价", "Line", "股价无原生类型（工具矩阵 12 类不含 stock），折线替代" },
                };
                foreach (var c in chartCases)
                {
                    string sheetName = c[0], chartType = c[1], note = c[2];
                    dynamic ws = Ws(sheetName);
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "create_excel_chart", Args(
                        ("chart_type", chartType), ("source_address", addr), ("sheet", sheetName)));
                    string rStr = Convert.ToString(r.Data);
                    bool ok = r.Ok && (rStr.Contains("chart_name") || rStr.Contains("created = True"));
                    Expect(ok, "基础 " + sheetName.Substring(0, 4) + " " + chartType + (note == "" ? "" : "（" + note + "）"),
                        r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }

                // ---- 高级图表 10 类 ----
                var advCases = new[]
                {
                    new[] { "高级01_子弹图", "bullet", "数值列=2,3" },
                    new[] { "高级02_华夫图", "waffle", "数值列=2" },
                    new[] { "高级03_滑珠图", "slider", "数值列=2,3" },
                    new[] { "高级04_漏斗图", "funnel", "数值列=2" },
                    new[] { "高级05_蝴蝶图", "butterfly", "数值列=2,3" },
                    new[] { "高级06_甘特图", "gantt", "数值列=2,3" },
                    new[] { "高级07_瀑布图", "waterfall", "数值列=2" },
                    new[] { "高级08_树状图", "treemap", "数值列=2" },
                    new[] { "高级09_旭日图", "sunburst", "数值列=4" },
                    new[] { "高级10_排列图", "pareto", "数值列=2" },
                };
                var nativePngs = new List<string>();
                foreach (var c in advCases)
                {
                    string sheetName = c[0], type = c[1], note = c[2];
                    dynamic ws = Ws(sheetName);
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "create_chart_advanced", Args(
                        ("type", type), ("source_address", addr), ("sheet", sheetName)));
                    string rStr = Convert.ToString(r.Data);
                    // 返回契约：auto_check 并入后 Data 为 JObject（缩进 JSON），字段名带引号
                    bool ok = r.Ok && (rStr.Contains("\"adv_type\": \"" + type + "\"") || rStr.Contains("adv_type = " + type));

                    // 原生四类视觉闭环：读回 pts（Series.Values 旧路径会错乱成 1~3 点）+ Export PNG 供看图
                    string extra = "";
                    if (type == "waterfall" || type == "treemap" || type == "sunburst" || type == "pareto")
                    {
                        try
                        {
                            ws = Ws(sheetName);  // 服务工具调用分离了旧 sheet RCW，重取（M-D1 纪律）
                            // 时序假设验证：新图类型 SetSourceData 后内部数据布局异步重建，立即读 pts 可能读到旧状态
                            System.Threading.Thread.Sleep(1500);
                            dynamic coAdv = ws.ChartObjects();
                            int cn = Convert.ToInt32(coAdv.Count);
                            dynamic chAdv = coAdv.Item(cn).Chart;
                            int pts = Convert.ToInt32(chAdv.SeriesCollection().Item(1).Points().Count);
                            string png = System.IO.Path.Combine(
                                System.Environment.GetEnvironmentVariable("BOX_AGENT_SCRATCH_DIR") ?? ".",
                                "adv_" + type + ".png");
                            coAdv.Item(cn).Activate();
                            chAdv.Export(png, "PNG");
                            nativePngs.Add(png);
                            extra = " pts=" + pts + " png=" + (System.IO.File.Exists(png) ? "Y" : "N");
                            bool ptsOk = pts >= 4;  // 样本最少 6~7 行；旧坏路径 pts=1~3
                            ok = ok && ptsOk;
                            Release((object)chAdv); Release((object)coAdv);
                        }
                        catch (Exception exv) { extra = " VERIFY-EXC:" + exv.Message; ok = false; }
                    }
                    Expect(ok, "高级 " + sheetName.Substring(0, 5) + " " + type + "（" + note + "）" + extra,
                        r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }

                // ---- 看板 6 类 ----
                {
                    dynamic ws = Ws("看板01_KPI摘要");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "create_dashboard_kpi_brief", Args(
                        ("source_address", addr), ("sheet", "看板01_KPI摘要"),
                        ("label_column", 1), ("value_column", 3),
                        ("dashboard_sheet", "看板KPI摘要")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板① KPI 摘要", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    dynamic ws = Ws("看板02_销售概览");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "render_dashboard", Args(
                        ("source_address", addr), ("sheet", "看板02_销售概览"),
                        ("label_column", 1), ("category_column", 2), ("month_column", 3),
                        ("value_column", 4), ("quantity_column", 4),
                        ("dashboard_sheet", "看板销售概览")));
                    string rStr = Convert.ToString(r.Data);
                    Expect(r.Ok && rStr.Contains("created = True") && rStr.Contains("selfcheck_verdict = ok"),
                        "看板② 销售概览（4 列退化：金额=销量同列）", r.Ok ? rStr : r.ErrorCode + " " + r.Data);
                }
                {
                    dynamic ws = Ws("看板03_多维交叉");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "render_dim_cross", Args(
                        ("source_address", addr), ("sheet", "看板03_多维交叉"),
                        ("row_dim_column", 1), ("col_dim_column", 2), ("value_column", 3),
                        ("dashboard_sheet", "看板多维交叉")));
                    // 矩阵形输入（产品线|华东|华南|...）必须被防呆拦截——曾实测被当长表产出语义错误看板
                    bool rejected = !r.Ok && Convert.ToString(r.ErrorCode) == "MATRIX_LAYOUT";
                    Expect(rejected, "看板③ 多维交叉（矩阵输入必拦截）",
                        rejected ? "MATRIX_LAYOUT 已拦截" : r.ErrorCode + " " + r.Data);
                }
                {
                    dynamic ws = Ws("看板04_财务损益");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "render_pnl_brief", Args(
                        ("source_address", addr), ("sheet", "看板04_财务损益"),
                        ("dashboard_sheet", "看板财务损益")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板④ 财务损益", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    dynamic ws = Ws("看板05_目标达成");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "render_target_track", Args(
                        ("source_address", addr), ("sheet", "看板05_目标达成"),
                        ("dashboard_sheet", "看板目标达成")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板⑤ 目标达成", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }
                {
                    dynamic ws = Ws("看板06_排名动态");
                    string addr = DataAddress(ws, out int _);
                    Release((object)ws);
                    var r = Exec(registry, "render_rank_dynamic", Args(
                        ("source_address", addr), ("sheet", "看板06_排名动态"),
                        ("dashboard_sheet", "看板排名动态")));
                    Expect(r.Ok && Convert.ToString(r.Data).Contains("created = True"),
                        "看板⑥ 排名动态", r.Ok ? Convert.ToString(r.Data) : r.ErrorCode + " " + r.Data);
                }

                // ---- 超级数据表链路：看板02 明细 → 超级表 → 看板共享 ----
                {
                    dynamic ws = Ws("看板02_销售概览");
                    string addr = DataAddress(ws, out int lastRow);
                    Release((object)ws);
                    var st = Exec(registry, "create_super_table", Args(
                        ("source_address", addr), ("sheet", "看板02_销售概览"),
                        ("label_column", 1), ("category_column", 2), ("month_column", 3),
                        ("value_column", 4), ("quantity_column", 4)));
                    string stStr = Convert.ToString(st.Data);
                    bool stOk = st.Ok && stStr.Contains("template = super_table");
                    Expect(stOk, "超级数据表生成（看板02 明细）", st.Ok ? stStr : st.ErrorCode + " " + st.Data);
                    if (stOk)
                    {
                        string detail = "A1:E" + (lastRow - 3);  // 明细行数=lastRow-4，超级表含表头 A1 起：A1:E(lastRow-3)
                        var rd = Exec(registry, "render_dashboard", Args(
                            ("source_address", detail), ("sheet", "超级数据表"),
                            ("dashboard_sheet", "看板共享超级表")));
                        string rdStr = Convert.ToString(rd.Data);
                        Expect(rd.Ok && rdStr.Contains("created = True") && rdStr.Contains("selfcheck_verdict = ok"),
                            "看板共享超级表数据源（native 切片器）", rd.Ok ? rdStr : rd.ErrorCode + " " + rd.Data);
                    }
                }

                // T33 诊断/修复闭环真机用例（对瀑布图：diagnose → repair → 复检）
                {
                    var dg = Exec(registry, "diagnose_chart", Args(("sheet", "高级07_瀑布图")));
                    string dgStr = Convert.ToString(dg.Data);
                    bool dgOk = dg.Ok && dgStr.Contains("root_cause");
                    Expect(dgOk, "T33a diagnose_chart 全面诊断", dg.Ok ? dgStr : dg.ErrorCode + " " + dg.Data);

                    var rp = Exec(registry, "repair_chart", Args(("sheet", "高级07_瀑布图"), ("source_address", "A4:B10")));
                    string rpStr = Convert.ToString(rp.Data);
                    bool rpOk = rp.Ok && rpStr.Contains("repaired_actions");
                    Expect(rpOk, "T33b repair_chart 读回修复", rp.Ok ? rpStr : rp.ErrorCode + " " + rp.Data);

                    var dg2 = Exec(registry, "diagnose_chart", Args(("sheet", "高级07_瀑布图")));
                    string dg2Str = Convert.ToString(dg2.Data);
                    bool fixedOk = dg2.Ok && dg2Str.Contains("root_cause");
                    Expect(fixedOk, "T33c 修复后复检", dg2.Ok ? dg2Str : dg2.ErrorCode + " " + dg2.Data);
                }

                // ===== S 组：截图三维度全覆盖（2026-09-13 用户实证需求：自由截数据/自由截全部图表/看板全景）=====
                string shotsDir = System.Environment.GetEnvironmentVariable("BOX_AGENT_SCRATCH_DIR");
                if (string.IsNullOrEmpty(shotsDir)) { shotsDir = System.IO.Path.GetTempPath(); }
                shotsDir = System.IO.Path.Combine(shotsDir, "udchart_shots");
                if (System.IO.Directory.Exists(shotsDir)) { try { System.IO.Directory.Delete(shotsDir, true); } catch { } }
                System.IO.Directory.CreateDirectory(shotsDir);
                var shotPngs = new List<string>();

                // S1 自由截全部图表：遍历所有可见 sheet 的全部 ChartObject，逐图 chart 级截图
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
                            workbook = excel.Workbooks.Item[1];  // 截图工具内部操作会分离 RCW——每 sheet 重取（M-D1 纪律）
                            wsC = workbook.Worksheets.Item[si];
                            int vis = 0;
                            try { vis = Convert.ToInt32(wsC.Visible); } catch { }
                            if (vis != -1) { continue; }  // 隐藏临时表（_advsrc_* 数据快照等）无图可截
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
                        chartFails.Count == 0 ? "全部图表逐图导出成功" : "失败: " + string.Join(" | ", chartFails));
                }

                // S2 自由截数据（address 模式矩阵：常规区/小区域/大区域/纯文本/单列）
                var dataCases = new[]
                {
                    new[] { "基础01_比较_柱形条形", "A4:C10", "常规数据区" },
                    new[] { "高级02_华夫图", "A4:B5", "两行小区域" },
                    new[] { "看板02_销售概览", "A4:D22", "全明细大区域" },
                    new[] { "0目录与说明", "A1:D8", "纯文本说明区" },
                    new[] { "基础03_趋势_折线面积", "A4:A16", "单列" },
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

                // S2f 缺省选区截图（未传参 = 截当前活动选区，"选中哪儿截哪儿"）
                try
                {
                    dynamic wsSel = Ws("基础05_分布_散点");
                    wsSel.Activate();
                    wsSel.Range["A4:B8"].Select();
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

                // S3 看板全景截图（看板多维交叉为防呆拦截负例、无产物，不在列）
                foreach (var db in new[] { "看板KPI摘要", "看板销售概览", "看板财务损益", "看板目标达成", "看板排名动态", "看板共享超级表" })
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

                Console.WriteLine("\n== 用户图表全表套件结果：通过 " + _passed + "，失败 " + _failed + " ==");

                // 产物保留：另存到任务输出目录交付
                try
                {
                    workbook = excel.Workbooks.Item[1];  // RCW 重取（同上）
                    string outDirSave = System.Environment.GetEnvironmentVariable("BOX_AGENT_OUTPUT_DIR");
                    if (string.IsNullOrEmpty(outDirSave)) { outDirSave = System.IO.Path.GetTempPath(); }
                    string outPath = System.IO.Path.Combine(outDirSave, "JYYJ图表全表真机产物.xlsx");
                    if (System.IO.File.Exists(outPath)) { System.IO.File.Delete(outPath); }
                    workbook.SaveAs(outPath, 51);
                    Console.WriteLine("[saved] " + outPath);
                }
                catch (Exception svEx) { Console.WriteLine("[saved][WARN] 产物另存失败：" + svEx.Message); }

                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][UDC] " + ex.GetType().Name + " " + ex.Message);
                return 2;
            }
            finally
            {
                Console.WriteLine("[cleanup][UDC] 清理 Excel 进程...");
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
