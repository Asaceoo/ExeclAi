using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyPhaseF
{
    /// <summary>
    /// 自动实机冒烟（第 F 批）：与侧栏桥 CreateRegistry(turnLabel) 完全同一条链路，
    /// 启动真实 Excel 生成散点图，并断言四件事：
    ///  1) 加载到 1.4.1 加载项（版本号）；
    ///  2) 高风险才弹确认：create_excel_chart 属 StructuralChange（新 ConfirmTool 会直达不弹窗）；
    ///  3) 图表执行会写入修改历史，且带「第 N 轮」turn_label（历史不再空白、能知道哪轮哪操作）；
    ///  4) 工具执行会产生可供前端渲染的非空 tool_event 卡片数据。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL PHASE-F REAL-MACHINE CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            CheckVersionLoaded();

            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();

                dynamic setupSheet = wb.Worksheets.Item[1];
                setupSheet.Name = "pf";
                var cells = new object[,]
                {
                    { "x", "y" },
                    { 1D, 10D },
                    { 2D, 20D },
                    { 3D, 30D },
                    { 4D, 25D },
                    { 5D, 40D }
                };
                dynamic setupRange = setupSheet.Range["A1:B6"];
                setupRange.Value2 = cells;
                Release((object)setupRange);
                Release((object)setupSheet);

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyphasef-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost(
                    (object)app,
                    repo,
                    ExcelToolHost.CreateWorkbookHash(workbookPath),
                    "phf-session",
                    "第 2 轮");
                host.ScriptOptions.Enabled = true; // 实机直接跑 run_vba_snippet（复现/定位 0x80070578）
                var registry = host.CreateRegistry();

                CheckChartRisk(registry);
                var executed = ExecuteChart(registry, app);
                CheckHistoryTracked(repo, workbookPath, executed);
                CheckToolEventCard(executed);
                CheckVbaRun(registry, app);
                CheckVbaReusePerf(registry, app);
                CheckWriteAutoFit(registry, app);
                CheckWriteRefit(registry, app);
                CheckFormatTolerant(registry, app);
                CheckSmokeRunner(app);
            }
            finally
            {
                // 无状态：活动簿可能是自检临时簿被关闭后回到的本簿，重解析再用完即放，避免踩到被工具 FinalRelease 过期的 wb RCW。
                try { dynamic active = app.ActiveWorkbook; if (active != null) { active.Close(false); Release((object)active); } }
                catch (Exception ex) { Console.WriteLine("NOTE  workbook-close: " + Safe(ex)); }
                try { if (app != null) { app.ScreenUpdating = true; app.Quit(); Release((object)app); } }
                catch (Exception ex) { Console.WriteLine("NOTE  app-quit: " + Safe(ex)); }
            }
        }

        private static void CheckVersionLoaded()
        {
            try
            {
                var version = typeof(SidebarBridge).Assembly.GetName().Version;
                var info = typeof(SidebarBridge).Assembly
                    .GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false)
                    .Cast<AssemblyInformationalVersionAttribute>()
                    .FirstOrDefault()?.InformationalVersion;
                // 不锁绝对版本号，断言确实加载了非占位(<1.0.0.0)的 Addin 程序集，避免每次发版都改断言。
                Assert("loaded-version",
                    version != null && version.Major >= 1 && version.ToString() != "0.0.0.0",
                    "asm=" + version + " info=" + info);
            }
            catch (Exception ex)
            {
                Fail("loaded-version", "HARNESS: " + Safe(ex));
            }
        }

        private static void CheckChartRisk(ToolRegistry registry)
        {
            try
            {
                var level = registry.GetRiskLevel("create_excel_chart");
                var isStructural = level == RiskLevel.StructuralChange;
                // 新的 ConfirmTool：非 Dangerous 一律直接执行（不弹确认框）
                Assert("confirm-gating-chart-direct", isStructural,
                    "create_excel_chart 风险级=" + level + "（StructuralChange => 直达不弹窗）");
            }
            catch (Exception ex)
            {
                Fail("confirm-gating-chart-direct", "HARNESS: " + Safe(ex));
            }
        }

        private static ToolResult ExecuteChart(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("create_excel_chart", out var tool))
            {
                Fail("chart-execute", "create_excel_chart 未注册");
                return null;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = "pf",
                ["source_address"] = "A1:B6",
                ["chart_type"] = "Scatter",
                ["title"] = "VerifyF 散点图"
            };

            ToolResult res;
            try
            {
                res = tool.Execute(args);
            }
            catch (Exception ex)
            {
                Fail("chart-execute", "EXCEPTION: " + Safe(ex));
                return null;
            }

            var chartCount = CountCharts(app);
            Assert("chart-execute", res != null && res.Ok && chartCount > 0,
                "ok=" + (res?.Ok) + " count=" + chartCount
                + " err=" + (res == null || res.Ok ? "-" : (res.ErrorCode + " " + res.Data)));
            return res;
        }

        private static void CheckVbaRun(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                Fail("vba-execute", "run_vba_snippet 未注册");
                return;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["snippet"] = "Sub TT验证()\n    Range(\"C1\").Value = \"ok\"\nEnd Sub",
                ["macro_name"] = "TT验证"
            };

            ToolResult res;
            try
            {
                res = tool.Execute(args);
            }
            catch (Exception ex)
            {
                Fail("vba-execute", "EXCEPTION: " + Safe(ex));
                return;
            }

            if (res == null || !res.Ok)
            {
                var code = res?.ErrorCode ?? "NULL";
                var msg = res?.Data?.ToString() ?? string.Empty;
                Fail("vba-execute", "ok=false err=" + code + " " + msg);
                return;
            }

            // 回读 C1 确认宏真的改了单元格（存在性 + 生效性双重断言）。
            var c1 = "";
            try
            {
                dynamic ws = app.ActiveWorkbook.Worksheets.Item["pf"];
                c1 = Convert.ToString(ws.Range["C1"].Value2);
                Release((object)ws);
            }
            catch (Exception ex)
            {
                c1 = "<read-fail:" + Safe(ex) + ">";
            }

            Assert("vba-execute", string.Equals(c1, "ok", StringComparison.Ordinal),
                "C1=[" + c1 + "] hr_ok, macro resolved=" + (res.Data?.ToString() ?? "-"));
        }

        private static void CheckVbaReusePerf(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                Fail("vba-reuse", "run_vba_snippet 未注册");
                return;
            }

            // 连续三次 run_vba_snippet：验证常量草稿模块复用（无 Add/Remove 重编译）且速度可控，
            // 并顺带覆盖「清空重写旧代码」路径，杜绝上一轮残留 Option Explicit 导致的重复声明报错。
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var okAll = true;
            var detail = new List<string>();
            for (var i = 0; i < 3; i++)
            {
                var t0 = sw.ElapsedMilliseconds;
                ToolResult res = null;
                try
                {
                    res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["snippet"] = "Sub TT复用" + i + "()\n    Range(\"D1\").Value = \"r" + i + "\"\nEnd Sub",
                        ["macro_name"] = "TT复用" + i
                    });
                }
                catch (Exception ex)
                {
                    okAll = false;
                    detail.Add("run" + i + "=EXCEPTION " + Safe(ex));
                    break;
                }

                var d1 = i == 0 ? "<init>" : ReadD1(app, i);
                var ok = res != null && res.Ok && (i == 0 || d1 == "r" + i);
                okAll &= ok;
                detail.Add("run" + i + "=" + (sw.ElapsedMilliseconds - t0) + "ms" + (ok ? "" : " FAIL(D1=" + d1 + ")"));
            }

            sw.Stop();
            var totalMs = sw.ElapsedMilliseconds;
            // 复用铁证：整个活动工作簿里只应存在 1 个标准模块（JYYJHelper），连续运行不累积。
            var stdModuleCount = CountStdModules(app);
            Assert("vba-reuse", okAll && totalMs < 6000 && stdModuleCount == 1,
                string.Format("3次VBA耗时={0}ms, 标准模块数={1} [{2}]", totalMs, stdModuleCount, string.Join(",", detail)));
        }

        private static int CountStdModules(dynamic app)
        {
            dynamic vbProject = null;
            dynamic components = null;
            try
            {
                vbProject = app.VBE?.ActiveVBProject;
                if (vbProject == null)
                {
                    return -1;
                }

                components = vbProject.VBComponents;
                var count = Convert.ToInt32(components.Count);
                var std = 0;
                for (var i = 1; i <= count; i++)
                {
                    dynamic comp = null;
                    try
                    {
                        comp = components.Item(i);
                        dynamic t = comp.Type;
                        if (Convert.ToInt32(t) == 1) // vbext_ct_StdModule
                        {
                            std++;
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Release((object)comp);
                    }
                }

                return std;
            }
            catch
            {
                return -2;
            }
            finally
            {
                Release((object)components);
                Release((object)vbProject);
            }
        }

        private static string ReadD1(dynamic app, int idx)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item["pf"];
                return Convert.ToString(ws.Range["D1"].Value2);
            }
            catch
            {
                return "<read-fail>";
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static void CheckWriteAutoFit(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("write_excel_range", out var tool))
            {
                Fail("write-autofit", "write_excel_range 未注册");
                return;
            }

            // 用扁平列表写 3 格（G1:I1）：验证「形状自适应重排」，目标 1x3 接收 3x1 应自动匹配。
            var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = "G1:I1",
                ["values"] = new object[] { "w", "x", "y" }
            });

            string g1 = "", i1 = "";
            try
            {
                dynamic ws = app.ActiveWorkbook.Worksheets.Item["pf"];
                g1 = Convert.ToString(ws.Range["G1"].Value2);
                i1 = Convert.ToString(ws.Range["I1"].Value2);
                Release((object)ws);
            }
            catch (Exception ex)
            {
                g1 = "<read-fail:" + Safe(ex) + ">";
            }

            Assert("write-autofit", res.Ok && g1 == "w" && i1 == "y",
                "ok=" + res.Ok + " G1=[" + g1 + "] I1=[" + i1 + "] err=" + (res.Ok ? "-" : res.ErrorCode));
        }

        private static void CheckSmokeRunner(dynamic app)
        {
            try
            {
                var assembly = typeof(SidebarBridge).Assembly;
                var type = assembly.GetType("ExcelAiAssistant.Addin.DiagnosticSmokeRunner", false);
                if (type == null)
                {
                    Fail("smoke-runner", "DiagnosticSmokeRunner 类型未找到（Addin 程序集未加载）");
                    return;
                }

                var runner = Activator.CreateInstance(type, app);
                var resultsObj = type.GetMethod("Run").Invoke(runner, null);

                // 自检结果元素是 Addin 程序集内的匿名类型（内部类），跨程序集 dynamic 绑定会被拒，
                // 故沿用与 SidebarBridge 一致的 JSON 序列化链路读取。
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(resultsObj);
                var arr = Newtonsoft.Json.Linq.JArray.Parse(json);
                var items = new List<string>();
                var okAll = true;
                foreach (var j in arr)
                {
                    var ok = Convert.ToBoolean((bool)j["ok"]);
                    okAll &= ok;
                    items.Add((string)j["name"] + "=" + (ok ? "通过" : "失败") + "(" + (string)j["detail"] + ")");
                }

                Assert("smoke-runner", okAll, string.Join("; ", items));
            }
            catch (Exception ex)
            {
                Fail("smoke-runner", "HARNESS: " + Safe(ex));
            }
        }

        private static void CheckFormatTolerant(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("format_excel_range", out var tool))
            {
                Fail("format-tolerant", "format_excel_range 未注册");
                return;
            }

            // 复现线上「居中，好看点」的传参：颜色/布尔均为字符串（含十六进制与中文）。
            // 修复前 Convert.ToInt32/ToBoolean 强转抛 FormatException；修复后宽容解析应 Ok=true 且不抛。
            var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = "$C$25:$J$41",
                ["back_color"] = "#1F4E78",
                ["font_color"] = "white",
                ["bold"] = "是",
                ["italic"] = "no",
                ["auto_fit_columns"] = "yes"
            });

            var ok = res != null && res.Ok;
            Assert("format-tolerant", ok,
                "ok=" + ok + " data=" + (res == null || res.Data is string ? "-" : "anonymous"));
        }

        private static void CheckWriteRefit(ToolRegistry registry, dynamic app)
        {
            if (!registry.TryGet("write_excel_range", out var tool))
            {
                Fail("write-refit", "write_excel_range 未注册");
                return;
            }

            // 1) 值多于区域 → 截断：5 个扁平值写进 2 格（J1:K1），应只保留前 2 个且整轮成功（不再 SHAPE_MISMATCH）。
            var t1 = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = "J1:K1",
                ["values"] = new object[] { "j1", "k1", "j2", "k2", "j3" }
            });
            string j1 = "", k1 = "";
            ReadCells(app, "pf", new[] { "J1", "K1" }, ref j1, ref k1);
            var truncOk = t1 != null && t1.Ok && j1 == "j1" && k1 == "k1";

            // 2) 值少于区域 → 补齐：2 个值写进 5 格（L1:P1），前 2 填值、后 3 格被清空（区域替换）。
            var t2 = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = "L1:P1",
                ["values"] = new object[] { "l1", "m1" }
            });
            string l1 = "", m1 = "", p1 = "<na>";
            ReadCells(app, "pf", new[] { "L1", "M1", "P1" }, ref l1, ref m1, ref p1);
            var padOk = t2 != null && t2.Ok && l1 == "l1" && m1 == "m1" && string.IsNullOrEmpty(p1);

            // 3) 模型把二维数组字符串化了一层 JSON 字面量 → 应自动解析展开为 2x2，不再整段落进单格。
            var t3 = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = "Q1:R2",
                ["values"] = "[[\"q1\",\"r1\"],[\"q2\",\"r2\"]]"
            });
            string q1 = "", r1 = "", q2 = "", r2 = "";
            ReadCells(app, "pf", new[] { "Q1", "R1", "Q2", "R2" }, ref q1, ref r1, ref q2, ref r2);
            var jsonOk = t3 != null && t3.Ok && q1 == "q1" && r1 == "r1" && q2 == "q2" && r2 == "r2";

            Assert("write-refit", truncOk && padOk && jsonOk,
                string.Format(
                    "trunc(J1=[{0}],K1=[{1}],ok={2}); pad(L1=[{3}],M1=[{4}],P1=[{5}],ok={6}); json(Q1=[{7}],R1=[{8}],Q2=[{9}],R2=[{10}],ok={11})",
                    j1, k1, truncOk, l1, m1, p1, padOk, q1, r1, q2, r2, jsonOk));
        }

        private static void ReadCells(dynamic app, string sheetName, string[] addresses,
            ref string a1, ref string a2)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item[sheetName];
                var v1 = ws.Range[addresses[0]].Value2;
                var v2 = ws.Range[addresses[1]].Value2;
                a1 = Convert.ToString(v1);
                a2 = Convert.ToString(v2);
            }
            catch (Exception ex)
            {
                a1 = "<read-fail:" + Safe(ex) + ">";
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static void ReadCells(dynamic app, string sheetName, string[] addresses,
            ref string a1, ref string a2, ref string a3)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item[sheetName];
                a1 = Convert.ToString(ws.Range[addresses[0]].Value2);
                a2 = Convert.ToString(ws.Range[addresses[1]].Value2);
                a3 = Convert.ToString(ws.Range[addresses[2]].Value2);
            }
            catch (Exception ex)
            {
                a1 = "<read-fail:" + Safe(ex) + ">";
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static void ReadCells(dynamic app, string sheetName, string[] addresses,
            ref string a1, ref string a2, ref string a3, ref string a4)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item[sheetName];
                a1 = Convert.ToString(ws.Range[addresses[0]].Value2);
                a2 = Convert.ToString(ws.Range[addresses[1]].Value2);
                a3 = Convert.ToString(ws.Range[addresses[2]].Value2);
                a4 = Convert.ToString(ws.Range[addresses[3]].Value2);
            }
            catch (Exception ex)
            {
                a1 = "<read-fail:" + Safe(ex) + ">";
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static void CheckHistoryTracked(OperationRepository repo, string workbookPath, ToolResult executed)
        {
            try
            {
                var hash = ExcelToolHost.CreateWorkbookHash(workbookPath);
                var records = repo.List(new OperationQuery { WorkbookHash = hash, Limit = 20 });
                var chartRecord = records.FirstOrDefault(r => r.Tool == "create_excel_chart");
                var tracked = chartRecord != null
                              && chartRecord.UndoKind == "chart-insert"
                              && chartRecord.TurnLabel == "第 2 轮";
                Assert("history-tracked-turnlabel", tracked,
                    "records=" + records.Count
                    + " tool=" + (chartRecord?.Tool ?? "-")
                    + " undo_kind=" + (chartRecord?.UndoKind ?? "-")
                    + " turn_label=[" + (chartRecord?.TurnLabel ?? "-") + "]");
            }
            catch (Exception ex)
            {
                Fail("history-tracked-turnlabel", "HARNESS: " + Safe(ex));
            }
        }

        private static void CheckToolEventCard(ToolResult executed)
        {
            try
            {
                var record = new ToolEventRecord
                {
                    Tool = "create_excel_chart",
                    Args = new Dictionary<string, object>(),
                    Confirmed = true,
                    Result = executed
                };
                var view = ToolEventViewBuilder.Build(record);
                Newtonsoft.Json.Linq.JObject jo = null;
                try { jo = Newtonsoft.Json.Linq.JObject.FromObject(view); } catch { }
                var ok = jo != null
                         && Convert.ToString(jo["tool"]) == "create_excel_chart"
                         && Convert.ToBoolean((bool)jo["ok"])
                         && !string.IsNullOrEmpty(Convert.ToString(jo["summary"]));
                Assert("tool-event-card", ok,
                    "tool=" + (jo?["tool"]) + " ok=" + (jo?["ok"])
                    + " summary=" + (jo?["summary"]));
            }
            catch (Exception ex)
            {
                Fail("tool-event-card", "HARNESS: " + Safe(ex));
            }
        }

        private static int CountCharts(dynamic app)
        {
            try
            {
                dynamic ws = app.ActiveWorkbook.Worksheets.Item["pf"];
                var count = 0;
                foreach (var c in ws.ChartObjects())
                {
                    count++;
                }

                return count;
            }
            catch
            {
                return -1;
            }
        }

        private static object obj(ToolResult e, string key)
        {
            if (e?.Data == null || e.Data is string)
            {
                return null;
            }

            try
            {
                var jo = Newtonsoft.Json.Linq.JObject.FromObject(e.Data);
                return jo[key]?.ToObject<object>();
            }
            catch
            {
                return null;
            }
        }

        private static string Str(ToolResult e, string key)
        {
            return obj(e, key)?.ToString();
        }

        private static string Safe(Exception ex)
        {
            try
            {
                var inner = ex;
                while (inner.InnerException != null) inner = inner.InnerException;
                return inner.GetType().Name + ": " + (inner.Message ?? string.Empty);
            }
            catch
            {
                return "unknown";
            }
        }

        private static void Assert(string name, bool ok, string detail)
        {
            if (ok)
            {
                Pass(name, detail);
            }
            else
            {
                Fail(name, detail);
            }
        }

        private static void Pass(string name, string detail)
        {
            _passed++;
            Console.WriteLine("PASS  " + name + "  [" + detail + "]");
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + " -> " + detail);
            Console.WriteLine("FAIL  " + name + "  [" + detail + "]");
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }
}