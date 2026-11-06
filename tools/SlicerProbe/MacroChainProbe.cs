using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// 「📊 数据分析报告」快捷宏工具链真机探针（macrochain 模式，2026-09-13）：
    /// 按宏 prompt 的编排顺序真机执行——口径 → 列画像 → 统计摘要 → 相关性 → 分布 → generate_analysis_report(html 落盘)，
    /// 断言每步 Ok、basis 字段在、HTML 报告落盘且含六章节关键标题。验证宏编排的工具链真实可用。
    /// </summary>
    internal static class MacroChainProbe
    {
        private static int _passed, _failed;
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

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

        public static int Run(string sourcePath, string pythonPath = null)
        {
            Console.WriteLine("== 数据分析报告宏工具链真机探针：" + sourcePath + " ==");
            if (!System.IO.File.Exists(sourcePath))
            {
                Console.WriteLine("FILE_NOT_FOUND: " + sourcePath);
                return 2;
            }
            string tempCopy = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "MACRO_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
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
                dynamic workbook = excel.Workbooks.Open(tempCopy);

                var dashboard = new ExcelDashboardService(excel);
                var pivotChart = new ExcelPivotChartService(excel, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, excel, null, null);

                // M1 口径：无 _口径.md → 返回创建模板（Ok + 模板标记）
                var g = Exec(registry, "read_business_glossary", Args());
                string gStr = Convert.ToString(g.Data);
                Expect(g.Ok, "M1 口径读取（缺文件返模板）", g.Ok ? (gStr.Contains("术语") ? "模板含术语字段" : gStr) : g.ErrorCode + " " + gStr);

                // M2 列画像（01 比较类 A3:D8，列 A=区域标签 B-D=产品数值）
                var p1 = Exec(registry, "analyze_column_profile", Args(
                    ("address", "A3:D8"), ("sheet", "01 比较类"), ("column", 2)));
                string p1Str = Convert.ToString(p1.Data);
                Expect(p1.Ok && p1Str.Contains("basis"), "M2 列画像（含 basis）",
                    p1.Ok ? p1Str : p1.ErrorCode + " " + p1Str);

                // M3 统计摘要（含 basis）
                var s1 = Exec(registry, "analyze_range_statistics", Args(
                    ("address", "A3:D8"), ("sheet", "01 比较类")));
                string s1Str = Convert.ToString(s1.Data);
                Expect(s1.Ok && s1Str.Contains("basis"), "M3 统计摘要（含 basis）",
                    s1.Ok ? s1Str : s1.ErrorCode + " " + s1Str);

                // M4 相关性（02 趋势类 B3:D15 三数值列，成对完整观测）
                var c1 = Exec(registry, "analyze_column_correlation", Args(
                    ("address", "B3:D15"), ("sheet", "02 趋势类")));
                string c1Str = Convert.ToString(c1.Data);
                Expect(c1.Ok && c1Str.Contains("basis"), "M4 相关性（含 basis）",
                    c1.Ok ? c1Str : c1.ErrorCode + " " + c1Str);

                // M5 分布（01 比较类 A 列区域标签）
                var d1 = Exec(registry, "analyze_value_distribution", Args(
                    ("address", "A3:D8"), ("sheet", "01 比较类"), ("column", 1)));
                string d1Str = Convert.ToString(d1.Data);
                Expect(d1.Ok && d1Str.Contains("basis"), "M5 分布（含 basis）",
                    d1.Ok ? d1Str : d1.ErrorCode + " " + d1Str);

                // M7 趋势预测（轻量）：临时 sheet 写 12 个月已知上升序列（100+10·(n-1)），linear 预测第 13 期=220
                dynamic wsTrend = workbook.Worksheets.Add();
                wsTrend.Name = "TREND_TMP";
                for (var rowIdx = 1; rowIdx <= 13; rowIdx++)
                {
                    object label = rowIdx == 1 ? "月份" : (object)((rowIdx - 1) + "月");
                    object amount = rowIdx == 1 ? "销量" : (object)(100.0 + 10.0 * (rowIdx - 2));
                    wsTrend.Cells[rowIdx, 1] = label;
                    wsTrend.Cells[rowIdx, 2] = amount;
                }
                var t1 = Exec(registry, "analyze_trend_forecast", Args(
                    ("address", "A1:B13"), ("sheet", "TREND_TMP"), ("periods", 3)));
                // forecasts 为匿名对象数组，直接 ToString 不展开（假 FAIL 坑）；断言一律用 JSON 序列化口径。
                string t1Str = Newtonsoft.Json.JsonConvert.SerializeObject(t1.Data);
                Expect(t1.Ok && t1Str.Contains("basis") && t1Str.Contains("220") && t1Str.Contains("上升"),
                    "M7 趋势预测（线性外推+220+basis）",
                    t1.Ok ? t1Str : t1.ErrorCode + " " + t1Str);

                // M7b 低样本 fail-closed：仅 3 个数据点须拒算
                var t2 = Exec(registry, "analyze_trend_forecast", Args(
                    ("address", "A1:B4"), ("sheet", "TREND_TMP")));
                string t2Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { t2.Ok, t2.ErrorCode, t2.Data });
                Expect(!t2.Ok && t2Str.Contains("INSUFFICIENT_POINTS"), "M7b 低样本 fail-closed",
                    t2Str);

                // M8 假设检验标准工具：Welch t 已知数据（t≈-6.1237，df=4，p<0.05 显著）
                dynamic wsHyp = workbook.Worksheets.Add();
                wsHyp.Name = "HYP_TMP";
                var hypData = new object[7, 2]
                {
                    { "组别", "数值" },
                    { "X", 10.0 }, { "X", 12.0 }, { "X", 14.0 },
                    { "Y", 20.0 }, { "Y", 22.0 }, { "Y", 24.0 }
                };
                for (var rr = 1; rr <= 7; rr++)
                {
                    wsHyp.Cells[rr, 1] = hypData[rr - 1, 0];
                    wsHyp.Cells[rr, 2] = hypData[rr - 1, 1];
                }
                var h1 = Exec(registry, "analyze_hypothesis_test", Args(
                    ("address", "A1:B7"), ("sheet", "HYP_TMP"), ("test", "two_sample_t"), ("alpha", 0.05)));
                string h1Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { h1.Ok, h1.ErrorCode, h1.Data });
                Expect(h1.Ok && h1Str.Contains("-6.12") && h1Str.Contains("\"df\":4.0") && h1Str.Contains("\"significant\":true") && h1Str.Contains("premise_notes"),
                    "M8 假设检验（Welch t=-6.12/df=4/显著/basis）",
                    h1Str);

                // M9 聚类分群标准工具：3 个远离簇（每簇 3 点），自动定 k=3 且分派正确
                dynamic wsClu = workbook.Worksheets.Add();
                wsClu.Name = "CLU_TMP";
                var cluPts = new[]
                {
                    new[] { 0.0, 0.0 }, new[] { 0.5, 0.3 }, new[] { -0.4, -0.2 },
                    new[] { 20.0, 20.0 }, new[] { 20.5, 20.3 }, new[] { 19.6, 19.8 },
                    new[] { 40.0, 5.0 }, new[] { 40.5, 5.3 }, new[] { 39.6, 4.8 }
                };
                wsClu.Cells[1, 1] = "x";
                wsClu.Cells[1, 2] = "y";
                for (var ci = 0; ci < cluPts.Length; ci++)
                {
                    wsClu.Cells[ci + 2, 1] = cluPts[ci][0];
                    wsClu.Cells[ci + 2, 2] = cluPts[ci][1];
                }
                var u1 = Exec(registry, "analyze_cluster", Args(("address", "A1:B10"), ("sheet", "CLU_TMP")));
                string u1Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { u1.Ok, u1.ErrorCode, u1.Data });
                Expect(u1.Ok && u1Str.Contains("\"k\":3") && (u1Str.Contains("\"best_silhouette\":0.9") || u1Str.Contains("\"best_silhouette\":0.8")) && u1Str.Contains("profiles") && u1Str.Contains("assignments"),
                    "M9 聚类（自动定 k=3/高轮廓/画像）",
                    u1Str);

                // M10 多元线性回归标准工具：已知数据集（statsmodels 锁定值 β=[3.0825,1.6925,1.7925]，R²=0.9998）
                dynamic wsReg = workbook.Worksheets.Add();
                wsReg.Name = "REG_TMP";
                var regX1 = new[] { 1.0, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
                var regX2 = new[] { 2.0, 1, 4, 3, 6, 5, 8, 7, 10, 9 };
                var regNoise = new[] { 0.3, -0.2, 0.5, -0.4, 0.1, -0.3, 0.4, -0.1, 0.2, -0.5 };
                wsReg.Cells[1, 1] = "x1";
                wsReg.Cells[1, 2] = "x2";
                wsReg.Cells[1, 3] = "y";
                for (var ri = 0; ri < 10; ri++)
                {
                    wsReg.Cells[ri + 2, 1] = regX1[ri];
                    wsReg.Cells[ri + 2, 2] = regX2[ri];
                    wsReg.Cells[ri + 2, 3] = 3.0 + 2.0 * regX1[ri] + 1.5 * regX2[ri] + regNoise[ri];
                }
                var g1 = Exec(registry, "analyze_regression", Args(("address", "A1:C11"), ("sheet", "REG_TMP"), ("target_column", 3)));
                string g1Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { g1.Ok, g1.ErrorCode, g1.Data });
                Expect(g1.Ok && g1Str.Contains("\"estimate\":3.0825") && g1Str.Contains("\"estimate\":1.6925") && g1Str.Contains("\"estimate\":1.7925")
                        && g1Str.Contains("\"r2\":0.9998") && g1Str.Contains("coefficients") && g1Str.Contains("vif") && g1Str.Contains("durbin_watson") && g1Str.Contains("diagnostics"),
                    "M10 回归（系数还原/R²/诊断面）",
                    g1Str);

                // M11 文本变体聚类建议（FR-K2）：空白/错字变体自动归组，只建议不合并
                dynamic wsTxt = workbook.Worksheets.Add();
                wsTxt.Name = "TXT_TMP";
                wsTxt.Cells[1, 1] = "客户";
                var txtRows = new[]
                {
                    "华为技术有限公司", "华为技术有限公司", "华为技术有限公司 ",
                    "华为技术有很公司", "小米科技", "小米 科技"
                };
                for (var ti = 0; ti < txtRows.Length; ti++)
                {
                    wsTxt.Cells[ti + 2, 1] = txtRows[ti];
                }
                var x1 = Exec(registry, "analyze_text_variants", Args(("address", "A1:A7"), ("sheet", "TXT_TMP"), ("column", 1)));
                string x1Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { x1.Ok, x1.ErrorCode, x1.Data });
                Expect(x1.Ok && x1Str.Contains("\"group_count\":2") && x1Str.Contains("\"representative\":\"华为技术有限公司\"")
                        && x1Str.Contains("华为技术有很公司") && x1Str.Contains("merge_options") && x1Str.Contains("write_excel_cells_batch"),
                    "M11 文本变体聚类（2 组建议/代表值/合并选项/改值语义提醒）",
                    x1Str);

                // M12 看板洞察卡：结论形状落板 + 内嵌读回自检 + 形状真机存在性复核
                var c2 = Exec(registry, "add_insight_card", Args(
                    ("sheet", "TXT_TMP"), ("cell", "D2"),
                    ("title", "客户名存在 2 组疑似变体"), ("body", "建议合并后统一客户口径，涉及 5 个单元格。"), ("color", "blue")));
                string c2Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { c2.Ok, c2.ErrorCode, c2.Data });
                dynamic wsTxtBack = workbook.Worksheets.Item("TXT_TMP");
                bool shapeFound = false;
                foreach (var shp in wsTxtBack.Shapes)
                {
                    if (Convert.ToString(shp.Name).StartsWith("xai_insight_", StringComparison.Ordinal)) { shapeFound = true; break; }
                }
                Expect(c2.Ok && c2Str.Contains("\"verdict\":\"ok\"") && c2Str.Contains("xai_insight_") && shapeFound && c2Str.Contains("basis"),
                    "M12 洞察卡（读回自检 ok/形状真机存在/basis）",
                    c2Str + " shapeFound=" + shapeFound);

                // M13 大表抽样+超时分级契约真机（蓝图 P1-2）：6 万行块写 → run_python_excel 抽样（n=2000/seed=42）+ timeout 120s 分级
                if (string.IsNullOrEmpty(pythonPath))
                {
                    foreach (var cand in new[]
                    {
                        System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles), "Python312", "python.exe"),
                        System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles), "Python311", "python.exe"),
                        "python"
                    })
                    {
                        if (cand == "python" || System.IO.File.Exists(cand)) { pythonPath = cand; break; }
                    }
                }

                if (string.IsNullOrEmpty(pythonPath))
                {
                    Expect(false, "M13 大表抽样+超时分级（本机未找到 Python，环境缺依赖）", "PYTHON_NOT_FOUND");
                }
                else
                {
                    dynamic wsBig = workbook.Worksheets.Add();
                    wsBig.Name = "BIG_TMP";
                    const int bigRows = 60000;
                    var bigArr = new object[bigRows + 1, 2];
                    bigArr[0, 0] = "id";
                    bigArr[0, 1] = "v";
                    var rndBig = new Random(7);
                    for (var i = 1; i <= bigRows; i++)
                    {
                        bigArr[i, 0] = i;
                        bigArr[i, 1] = rndBig.Next(1, 1000);
                    }
                    wsBig.Range("A1:B" + (bigRows + 1)).Value2 = bigArr;

                    var pyOptions = new ScriptOptions { Enabled = true, PythonPath = pythonPath, TimeoutSeconds = 300 };
                    var pyExcel = new PythonExcelService(excel, pyOptions);
                    var m13 = pyExcel.Run(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["code"] = "import pandas as pd\nsht = xw.books.active.sheets['BIG_TMP']\ndf = sht.range('A1').options(pd.DataFrame, expand='table').value\nsamp = df.sample(n=2000, random_state=42)\nxai_result({'total_rows': int(len(df)), 'sample_n': int(len(samp)), 'seed': 42, 'sample_mean': round(float(samp['v'].mean()), 4)})",
                        ["timeout_ms"] = 120000
                    });
                    string m13Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { m13.Ok, m13.ErrorCode, m13.Data });
                    Expect(m13.Ok && m13Str.Contains("\"total_rows\":60000") && m13Str.Contains("\"sample_n\":2000") && m13Str.Contains("\"seed\":42"),
                        "M13 大表抽样+超时分级（6 万行/sample n=2000/seed=42/timeout 120s）",
                        m13Str.Length > 700 ? m13Str.Substring(0, 700) : m13Str);
                }

                // M14 写入文本保护闭环（v3.8.1.76 审查批次）：COM Value2 写入曾把 18 位身份证重解析为
                // 5.787E+17（精度丢失）、"007"→7（前导零丢失）。锁文本格后身份证/编码必须原样回读，
                // 日期串（不拦，Excel 原生日期化）被重解析时 coerced_cells 必须可见。
                {
                    var writeHost = new ExcelToolHost((object)excel, null, ExcelToolHost.CreateWorkbookHash("macrochain-m14"), "macrochain-m14");
                    var writeRegistry = writeHost.CreateRegistry();
                    dynamic wsText = workbook.Worksheets.Add();
                    wsText.Name = "TEXT_LOCK_TMP";
                    var m14 = Exec(writeRegistry, "write_excel_range", Args(
                        ("address", "A1:C2"), ("sheet", "TEXT_LOCK_TMP"),
                        ("values", new List<object> {
                            new List<object> { "员工编号", "身份证号", "入职日期" },
                            new List<object> { "007", "578712345678901234", "1987-07-04" } })));
                    string m14Str = Newtonsoft.Json.JsonConvert.SerializeObject(new { m14.Ok, m14.ErrorCode, m14.Data });
                    int locked = 0, coerced = -1;
                    string idBack = "<unread>", codeBack = "<unread>";
                    if (m14.Ok)
                    {
                        var m14Data = Newtonsoft.Json.Linq.JObject.FromObject(m14.Data);
                        locked = (int?)m14Data["text_locked_cells"] ?? 0;
                        coerced = (int?)m14Data["coerced_count"] ?? -1;
                        idBack = Convert.ToString(wsText.Range("B2").Value2);
                        codeBack = Convert.ToString(wsText.Range("A2").Value2);
                    }
                    bool idOk = idBack == "578712345678901234";   // 若被重解析为 double 则科学计数法不等
                    bool codeOk = codeBack == "007";
                    Expect(m14.Ok && locked == 2 && coerced == 1 && idOk && codeOk,
                        "M14 写入文本保护（身份证 18 位锁文本不丢精度/前导零保留/日期串重解析上报 coerced）",
                        m14Str.Length > 700 ? m14Str.Substring(0, 700) : m14Str + " locked=" + locked + " coerced=" + coerced + " idBack=" + idBack + " codeBack=" + codeBack);
                }

                // M6 报告落盘（output=html，output_path=工作簿同目录——宏语义）
                string outDir = System.Environment.GetEnvironmentVariable("BOX_AGENT_SCRATCH_DIR");
                if (string.IsNullOrEmpty(outDir)) { outDir = System.IO.Path.GetTempPath(); }
                string htmlPath = System.IO.Path.Combine(outDir, "macrochain_report.html");
                try { if (System.IO.File.Exists(htmlPath)) { System.IO.File.Delete(htmlPath); } } catch { }
                var r = Exec(registry, "generate_analysis_report", Args(
                    ("address", "A3:D15"), ("sheet", "02 趋势类"),
                    ("output", "html"), ("output_path", htmlPath)));
                bool ok = false; string detail;
                if (r.Ok)
                {
                    var dR = Newtonsoft.Json.Linq.JObject.FromObject(r.Data);
                    string fp = (string)dR["file_path"];
                    bool exists = !string.IsNullOrEmpty(fp) && System.IO.File.Exists(fp);
                    long size = exists ? new System.IO.FileInfo(fp).Length : 0;
                    string html = exists ? System.IO.File.ReadAllText(fp) : "";
                    bool hasSections = html.Contains("概览") && html.Contains("趋势") && html.Contains("异常")
                        && html.Contains("相关性") && html.Contains("建议");
                    ok = exists && size > 1024 && hasSections;
                    detail = exists
                        ? "html=" + System.IO.Path.GetFileName(fp) + " " + (size / 1024.0).ToString("0.0") + "KB 六章节标题" + (hasSections ? "齐" : "缺")
                        : r.ErrorCode + " " + Convert.ToString(r.Data);
                }
                else { detail = r.ErrorCode + " " + Convert.ToString(r.Data); }
                Expect(ok, "M6 分析报告 HTML 落盘（六章节）", detail);

                Console.WriteLine("\n== 宏工具链探针结果：通过 " + _passed + "，失败 " + _failed + " ==");
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][MACRO] " + ex.GetType().Name + " " + ex.Message);
                return 2;
            }
            finally
            {
                Console.WriteLine("[cleanup][MACRO] 清理 Excel 进程...");
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
