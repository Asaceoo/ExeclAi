using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyVerifyChart
{
    /// <summary>
    /// 真机验证 verify_chart 自检能力全链路（与侧栏桥 CreateRegistry 同一条链路）：
    /// 真实 Excel 建柱状图后，逐项断言自检工具对「真生成/类型/系列数/标题/数据源/像素渲染」的判定，
    /// 并故意给错期望值确认自检能报 issue（而不是永远 pass）。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static bool enumFirst = true;
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
                foreach (var f in Failures) Console.WriteLine("FAILED: " + f);
                return 1;
            }

            Console.WriteLine("ALL VERIFY_CHART REAL-MACHINE CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
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
                setupSheet.Name = "vvc";
                var cells = new object[,]
                {
                    { "品类", "金额" },
                    { "A", 10D }, { "B", 25D }, { "C", 18D }, { "D", 32D }, { "E", 14D }, { "F", 22D }, { "G", 9D }
                };
                dynamic setupRange = setupSheet.Range["A1:B8"];
                setupRange.Value2 = cells;
                Release((object)setupRange);
                Release((object)setupSheet);

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyverifychart-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost(
                    (object)app,
                    repo,
                    ExcelToolHost.CreateWorkbookHash(workbookPath),
                    "vvc-session",
                    "自检");
                var registry = host.CreateRegistry();

                // 建一个真实柱状图，标题=月度销售
                CreateChart(registry);

                // ---- 自检能力逐维度验证 ----
                VerifyGenerated(registry);      // VVC-1 真生成 + 三要素 + 渲染图非空
                VerifyTypeMatch(registry);      // VVC-2 类型比对命中
                VerifyTypeMismatch(registry);   // VVC-3 类型不符报 issue
                VerifySeries(registry);         // VVC-4 系列数比对（命中/不符）
                VerifyTitle(registry);          // VVC-5 标题比对（命中/不符）
                VerifyDataSource(registry);     // VVC-6 数据源比对（命中/不符）
                VerifyPixelRender(registry);    // VVC-7 像素级渲染核验
                VerifyFamilyMap();              // VVC-8 类型族映射完整性（静态）
            }
            finally
            {
                try { dynamic active = app.ActiveWorkbook; if (active != null) { active.Close(false); Release((object)active); } }
                catch (Exception ex) { Console.WriteLine("NOTE  workbook-close: " + Safe(ex)); }
                try { if (app != null) { app.ScreenUpdating = true; app.Quit(); Release((object)app); } }
                catch (Exception ex) { Console.WriteLine("NOTE  app-quit: " + Safe(ex)); }
            }
        }

        private static void CreateChart(ToolRegistry registry)
        {
            if (!registry.TryGet("create_excel_chart", out var tool))
            {
                Fail("vvc-chart-create", "create_excel_chart 未注册");
                return;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = "vvc",
                ["source_address"] = "A1:B8",
                ["chart_type"] = "Column",
                ["title"] = "月度销售"
            };
            try
            {
                var res = tool.Execute(args);
                Assert("vvc-chart-create", res != null && res.Ok, "ok=" + (res?.Ok) + " err=" + (res == null || res.Ok ? "-" : res.ErrorCode));
            }
            catch (Exception ex)
            {
                Fail("vvc-chart-create", "EXCEPTION: " + Safe(ex));
            }
        }

        private static ToolResult RunVerify(ToolRegistry registry, Dictionary<string, object> extra)
        {
            if (!registry.TryGet("verify_chart", out var tool))
            {
                Fail("vvc-verify-notfound", "verify_chart 未注册");
                return null;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = "vvc"
            };
            foreach (var kv in extra) args[kv.Key] = kv.Value;
            return tool.Execute(args);
        }

        private static string V(ToolResult r, string key)
        {
            if (r?.Data == null || r.Data is string) return null;
            try
            {
                var jo = Newtonsoft.Json.Linq.JObject.FromObject(r.Data);
                return jo[key]?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static void VerifyGenerated(ToolRegistry registry)
        {
            try
            {
                var r = RunVerify(registry, new Dictionary<string, object>());
                var found = V(r, "chart_found");
                var verdict = V(r, "verdict");
                var image = V(r, "chart_image");
                var hasSelfcheck = !string.IsNullOrEmpty(V(r, "selfcheck"));
                var ok = r != null && r.Ok
                         && found == "True"
                         && verdict == "ok"
                         && !string.IsNullOrEmpty(image)
                         && image.StartsWith("data:image/png;base64,", StringComparison.Ordinal)
                         && hasSelfcheck;
                Assert("vvc-1-generated", ok,
                    "found=" + found + " verdict=" + verdict + " selfcheck=" + hasSelfcheck + " img_len=" + (image?.Length ?? 0));
            }
            catch (Exception ex) { Fail("vvc-1-generated", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifyTypeMatch(ToolRegistry registry)
        {
            try
            {
                var r = RunVerify(registry, new Dictionary<string, object> { ["expected_chart_type"] = "Column" });
                var ok = r != null && r.Ok && V(r, "verdict") == "ok";
                Assert("vvc-2-type-match", ok, "expected=Column verdict=" + V(r, "verdict"));
            }
            catch (Exception ex) { Fail("vvc-2-type-match", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifyTypeMismatch(ToolRegistry registry)
        {
            try
            {
                var r = RunVerify(registry, new Dictionary<string, object> { ["expected_chart_type"] = "Pie" });
                var verdict = V(r, "verdict");
                var issues = V(r, "issues");
                // 类型不符 → verdict 变 mismatch（或 +mismatch），并附「图表类型与预期不符」issue
                var ok = r != null && r.Ok
                         && verdict != null && verdict.Contains("mismatch")
                         && issues != null && issues.Contains("类型");
                Assert("vvc-3-type-mismatch", ok, "expected=Pie verdict=" + verdict + " issues=" + issues);
            }
            catch (Exception ex) { Fail("vvc-3-type-mismatch", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifySeries(ToolRegistry registry)
        {
            try
            {
                var hit = RunVerify(registry, new Dictionary<string, object> { ["expected_series"] = 1 });
                var miss = RunVerify(registry, new Dictionary<string, object> { ["expected_series"] = 5 });
                // 系列数不符按设计只追加「提示性 issue」、不改 verdict（verdict 仍 ok）
                var hitOk = hit != null && hit.Ok && V(hit, "verdict") == "ok";
                var missIssue = miss != null && miss.Ok
                                && V(miss, "verdict") == "ok"
                                && V(miss, "issues") != null && V(miss, "issues").Contains("系列数量与预期不符");
                Assert("vvc-4-series", hitOk && missIssue,
                    "expected=1 verdict=" + V(hit, "verdict")
                    + " issues=" + V(hit, "issues")
                    + "; expected=5 verdict=" + V(miss, "verdict")
                    + " issues=" + V(miss, "issues"));
            }
            catch (Exception ex) { Fail("vvc-4-series", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifyTitle(ToolRegistry registry)
        {
            try
            {
                var hit = RunVerify(registry, new Dictionary<string, object> { ["expected_title"] = "月度销售" });
                var miss = RunVerify(registry, new Dictionary<string, object> { ["expected_title"] = "错误标题" });
                var hitOk = hit != null && hit.Ok && V(hit, "verdict") == "ok";
                var missOk = miss != null && miss.Ok && V(miss, "verdict") != "ok";
                Assert("vvc-5-title", hitOk && missOk,
                    "expected=月度销售 verdict=" + V(hit, "verdict") + "; expected=错误标题 verdict=" + V(miss, "verdict"));
            }
            catch (Exception ex) { Fail("vvc-5-title", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifyDataSource(ToolRegistry registry)
        {
            try
            {
                var hit = RunVerify(registry, new Dictionary<string, object> { ["expected_data_source"] = "A2:B8" });
                var miss = RunVerify(registry, new Dictionary<string, object> { ["expected_data_source"] = "D5:E9" });
                var hitOk = hit != null && hit.Ok && V(hit, "verdict") == "ok";
                var missOk = miss != null && miss.Ok && V(miss, "verdict") != "ok";
                Assert("vvc-6-datasource", hitOk && missOk,
                    "expected=A2:B8 verdict=" + V(hit, "verdict") + "; expected=D5:E9 verdict=" + V(miss, "verdict"));
            }
            catch (Exception ex) { Fail("vvc-6-datasource", "HARNESS: " + Safe(ex)); }
        }

        private static void VerifyPixelRender(ToolRegistry registry)
        {
            try
            {
                var r = RunVerify(registry, new Dictionary<string, object>());
                var image = V(r, "chart_image");
                var png = false;
                try
                {
                    if (image != null && image.StartsWith("data:image/png;base64,", StringComparison.Ordinal))
                    {
                        var bytes = Convert.FromBase64String(image.Substring("data:image/png;base64,".Length));
                        png = bytes.Length > 0 && bytes[0] == 0x89 && bytes[1] == 0x50; // PNG 魔数
                    }
                }
                catch { png = false; }

                Assert("vvc-7-pixel-render", png, "png_magic=" + png + " img_len=" + (image?.Length ?? 0));
            }
            catch (Exception ex) { Fail("vvc-7-pixel-render", "HARNESS: " + Safe(ex)); }
        }

        // 静态断言：ChartTypeMap 覆盖 6 大类型族（Column/Bar/Line/Scatter/Pie/Area）——私有字段走反射读取
        private static void VerifyFamilyMap()
        {
            try
            {
                var allFields = typeof(ExcelChartService).GetFields(BindingFlags.NonPublic | BindingFlags.Static);
                var field = allFields.FirstOrDefault(f => f.Name.EndsWith("ChartTypeMap", StringComparison.Ordinal));
                if (field == null)
                {
                    Fail("vvc-8-family-map", "ChartTypeMap 字段不存在；候选字段=" + string.Join(",", allFields.Select(f => f.Name)));
                    return;
                }

                var raw = field.GetValue(null);
                var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var enumErr = "";
                if (raw is System.Collections.IEnumerable en)
                {
                    try
                    {
                        foreach (var item in en)
                        {
                            var t = item.GetType();
                            // ValueTuple 以公开字段（Item1/Item2/Item3）暴露成员，无公开属性
                            var family = Convert.ToString(t.GetField("Item2")?.GetValue(item)
                                ?? t.GetProperty("Item2")?.GetValue(item, null)
                                ?? t.GetProperty("Family")?.GetValue(item, null));
                            if (enumFirst && string.IsNullOrEmpty(family))
                            {
                                enumErr = "首元素 fields=[" + string.Join(",", t.GetFields().Select(f => f.Name))
                                          + "] val=[" + family + "] type=" + t.FullName;
                                enumFirst = false;
                            }
                            if (!string.IsNullOrEmpty(family)) families.Add(family);
                        }
                    }
                    catch (Exception ex) { enumErr = Safe(ex); }
                }
                else
                {
                    enumErr = "字段值不可枚举, type=" + (raw?.GetType().FullName ?? "null");
                }

                var expected = new[] { "Column", "Bar", "Line", "Scatter", "Pie", "Area" };
                var missing = expected.Where(f => !families.Contains(f)).ToList();
                Assert("vvc-8-family-map", missing.Count == 0,
                    "field=" + field.Name + " type=" + field.FieldType.Name
                    + " families=" + string.Join(",", families.OrderBy(x => x))
                    + " missing=" + string.Join(",", missing)
                    + (enumErr.Length > 0 ? " enum_err=" + enumErr : ""));
            }
            catch (Exception ex) { Fail("vvc-8-family-map", "HARNESS: " + Safe(ex)); }
        }

        private static string Safe(Exception ex)
        {
            try
            {
                var inner = ex;
                while (inner.InnerException != null) inner = inner.InnerException;
                return inner.GetType().Name + ": " + (inner.Message ?? string.Empty);
            }
            catch { return "unknown"; }
        }

        private static void Assert(string name, bool ok, string detail)
        {
            if (ok) Pass(name, detail);
            else Fail(name, detail);
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
