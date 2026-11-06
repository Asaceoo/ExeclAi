using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyReviewFix
{
    /// <summary>
    /// 真机冒烟：本轮「剩余工具对抗性审查 + 按算子校验强化」修复的验证闭环。
    /// 走与生产完全相同的服务类/注册链路，在真实 Excel 里执行并断言。
    /// 覆盖：
    ///   S1/S2  apply_page_setup margins 数组形态与单对象形态（含整数边距）都真正生效
    ///   S3/S4  format_excel_conditional 缺 value / between 缺 value2 → INVALID_VALUE
    ///   S5     condition greater 正常仍可用（回归）
    ///   S6/S7  filter gt 缺数值 / contains 空 value → INVALID_OPERATOR
    ///   S8     filter 列号越界 → INVALID_COLUMN
    ///   S9     filter 正常仍可筛选（回归）
    ///   S10    format_excel_range 非法 bold 记入 skipped（与颜色分支一致）
    ///   S11    export_excel_chart_image 按 width/height 导出后还原图表真实尺寸
    ///   S12    apply_fill_blank_cells 缺 fills → INVALID_FILLS（不再假成功）
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
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            foreach (var f in Failures)
            {
                Console.WriteLine("FAILED: " + f);
            }

            return _failed == 0 ? 0 : 1;
        }

        private static void RunAll()
        {
            dynamic app = null;
            dynamic wb = null;
            ExcelToolHost host = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                ws.Name = "dt";
                Seed(ws);
                // 生产是无状态调用：工具/writer 各自经 ActiveWorkbook/Worksheets 重解析并 FinalRelease，
                // harness 不能跨工具调用持有 ws，否则会被写路径连带销毁而抛「RCW 已分离」。
                MarshalRelease((object)ws);

                var wbPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyreviewfix-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(wbPath), "verifyreviewfix", "第 1 轮");
                var registry = host.CreateRegistry();

                MarginsArrayForm(app, registry);
                MarginsObjectForm(app, registry);

                CondMissingValue(app);
                CondBetweenMissingValue2(app);
                CondValid(app);

                FilterGtMissingValue(registry);
                FilterContainsEmpty(registry);
                FilterColumnOutOfRange(registry);
                FilterValid(registry);

                FormatRangeInvalidBold(registry);
                ChartExportRestoresSize(app);
                FillMissing(app);

                Console.WriteLine();
                Console.WriteLine("========== PASSED=" + _passed + " FAILED=" + _failed + " ==========");
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); MarshalRelease((object)wb); } } catch { }
                try { app.Quit(); } catch { }
                MarshalRelease((object)app);
            }
        }

        private static void Seed(dynamic ws)
        {
            ws.Range["A1:B6"].Value2 = new object[,]
            {
                { "name", "amt" },
                { "Beijing", 10 },
                { "Shanghai", 20 },
                { "Beijing", 5 },
                { "Shenzhen", 30 },
                { "Guangzhou", 15 }
            };
        }

        // S1: schema 声明的数组形态 [{top,left}]（含整数 2/3）应真正生效，而非被静默忽略。
        private static void MarginsArrayForm(dynamic app, ToolRegistry registry)
        {
            var svc = new ExcelWorkbookService((object)app);
            var margins = new JArray
            {
                new JObject { ["top"] = 2, ["left"] = 3 }   // JSON 整数 → 反序列化多为 int/long，须被 ToCmNumber 接受
            };
            var r = svc.ApplyPageSetup(new Dictionary<string, object> { ["sheet"] = "dt", ["margins"] = margins });
            Assert("margins-array-ok", r.Ok, "apply_page_setup margins 数组形态执行成功");
            double topPt = 0, leftPt = 0;
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item["dt"];
                var page = ws.PageSetup;
                topPt = Convert.ToDouble(page.TopMargin);
                leftPt = Convert.ToDouble(page.LeftMargin);
            }
            finally
            {
                MarshalRelease((object)ws);
            }

            Assert("margins-array-top", Math.Abs(topPt - Points(2)) < 0.5, "数组形态 top=2cm 生效（读回 " + topPt.ToString("0.0") + " 磅）");
            Assert("margins-array-left", Math.Abs(leftPt - Points(3)) < 0.5, "数组形态 left=3cm 生效（读回 " + leftPt.ToString("0.0") + " 磅）");
        }

        // S2: 历史单对象形态仍可用（含浮点 bottom）。
        private static void MarginsObjectForm(dynamic app, ToolRegistry registry)
        {
            var svc = new ExcelWorkbookService((object)app);
            var margins = new JObject { ["bottom"] = 0.5, ["right"] = 1.25 };
            var r = svc.ApplyPageSetup(new Dictionary<string, object> { ["sheet"] = "dt", ["margins"] = margins });
            Assert("margins-object-ok", r.Ok, "apply_page_setup margins 单对象形态执行成功");
            double bottomPt = 0, rightPt = 0;
            dynamic ws = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item["dt"];
                var page = ws.PageSetup;
                bottomPt = Convert.ToDouble(page.BottomMargin);
                rightPt = Convert.ToDouble(page.RightMargin);
            }
            finally
            {
                MarshalRelease((object)ws);
            }

            Assert("margins-object-bottom", Math.Abs(bottomPt - Points(0.5)) < 0.5, "单对象 bottom=0.5cm 生效（读回 " + bottomPt.ToString("0.0") + " 磅）");
            Assert("margins-object-right", Math.Abs(rightPt - Points(1.25)) < 0.5, "单对象 right=1.25cm 生效");
        }

        // S3: greater 缺 value → 显式 INVALID_VALUE，不再按 0 高亮全部。
        private static void CondMissingValue(dynamic app)
        {
            var svc = new ExcelDataToolsService((object)app);
            var r = svc.ConditionalFormat(new Dictionary<string, object> { ["address"] = "B2:B6", ["sheet"] = "dt", ["rule"] = "greater" });
            Assert("cond-missing-value", !r.Ok && r.ErrorCode == "INVALID_VALUE", "greater 缺 value → 显式 INVALID_VALUE（code=" + (r.Ok ? "ok" : r.ErrorCode) + "）");
        }

        // S4: between 缺 value2 → 显式 INVALID_VALUE，上界不再静默为 0。
        private static void CondBetweenMissingValue2(dynamic app)
        {
            var svc = new ExcelDataToolsService((object)app);
            var r = svc.ConditionalFormat(new Dictionary<string, object> { ["address"] = "B2:B6", ["sheet"] = "dt", ["rule"] = "between", ["value"] = 1 });
            Assert("cond-betw-missing2", !r.Ok && r.ErrorCode == "INVALID_VALUE", "between 缺 value2 → 显式 INVALID_VALUE");
        }

        // S5: 提供合法 value 时正常仍成功（回归）。
        private static void CondValid(dynamic app)
        {
            var svc = new ExcelDataToolsService((object)app);
            var r = svc.ConditionalFormat(new Dictionary<string, object> { ["address"] = "B2:B6", ["sheet"] = "dt", ["rule"] = "greater", ["value"] = 2 });
            Assert("cond-valid", r.Ok, "greater value=2 → 正常成功（回归）");
        }

        // S6: gt 缺数值 value → INVALID_OPERATOR，不再静默滤出空结果。
        private static void FilterGtMissingValue(ToolRegistry registry)
        {
            var r = Need(registry, "filter_excel_data").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B6", ["sheet"] = "dt", ["column"] = 2, ["operator"] = "gt"
            });
            Assert("filter-gt-novalue", !r.Ok && r.ErrorCode == "INVALID_OPERATOR", "gt 缺数值 value → INVALID_OPERATOR");
        }

        // S7: contains 空 value → INVALID_OPERATOR，不再全表命中。
        private static void FilterContainsEmpty(ToolRegistry registry)
        {
            var r = Need(registry, "filter_excel_data").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B6", ["sheet"] = "dt", ["column"] = 1, ["operator"] = "contains", ["value"] = ""
            });
            Assert("filter-contains-empty", !r.Ok && r.ErrorCode == "INVALID_OPERATOR", "contains 空 value → INVALID_OPERATOR");
        }

        // S8: 列号越界 → INVALID_COLUMN（不再被收敛成 EXCEL_COM_ERROR）。
        private static void FilterColumnOutOfRange(ToolRegistry registry)
        {
            var r = Need(registry, "filter_excel_data").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B6", ["sheet"] = "dt", ["column"] = 9, ["operator"] = "equals", ["value"] = "x"
            });
            Assert("filter-coloob", !r.Ok && r.ErrorCode == "INVALID_COLUMN", "列号=9 越界 → INVALID_COLUMN（code=" + (r.Ok ? "ok" : r.ErrorCode) + "）");
        }

        // S9: 正常筛选仍可用（回归），matched_rows=2。
        private static void FilterValid(ToolRegistry registry)
        {
            var r = Need(registry, "filter_excel_data").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B6", ["sheet"] = "dt", ["column"] = 1, ["operator"] = "equals", ["value"] = "Beijing"
            });
            Assert("filter-valid-ok", r.Ok, "equals Beijing 正常执行（回归）");
            int matched = 0;
            try
            {
                JObject data = JObject.FromObject(r.Data);
                matched = data.Value<int?>("matched_rows") ?? 0;
            }
            catch { }
            Assert("filter-valid-matches", matched == 2, "filter equals Beijing → matched_rows=2（实际=" + matched + "）");
        }

        // S10: 非法 bold 计入 skipped，与非法颜色分支对齐。
        private static void FormatRangeInvalidBold(ToolRegistry registry)
        {
            var r = Need(registry, "format_excel_range").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B6", ["sheet"] = "dt", ["bold"] = "not-a-bool", ["font_color"] = "#FF0000"
            });
            Assert("format-invalid-bold-ok", r.Ok, "format_excel_range 执行成功（bold 非法但整体仍成功）");
            string skipped = null;
            try
            {
                JObject data = JObject.FromObject(r.Data);
                skipped = data.Value<string>("skipped");
            }
            catch { }
            Assert("format-invalid-bold-skipped", skipped != null && skipped.IndexOf("bold", StringComparison.Ordinal) >= 0, "非法 bold 计入 skipped（skipped=" + (skipped ?? "null") + "）");
        }

        // S11: export_excel_chart_image 指定 width/height 导出后，图表真实尺寸被还原。
        private static void ChartExportRestoresSize(dynamic app)
        {
            dynamic ws = null;
            dynamic co = null;
            dynamic ch = null;
            string fp = Path.Combine(Path.GetTempPath(), "verifychart-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item["dt"];
                co = ws.ChartObjects().Add(60, 60, 320, 200);
                ch = co.Chart;
                ch.ChartType = 51; // xlColumnClustered
                ch.SetSourceData(ws.Range["B2:B6"]);

                var origW = Convert.ToDouble(co.Width);
                var origH = Convert.ToDouble(co.Height);

                var svc = new ExcelChartService((object)app);
                var r = svc.ExportChartAsImage(new Dictionary<string, object>
                {
                    ["sheet"] = "dt", ["chart_name"] = Convert.ToString(co.Name),
                    ["file_path"] = fp, ["width"] = 640, ["height"] = 400
                });
                Assert("chart-export-ok", r.Ok && File.Exists(fp), "export_excel_chart_image 导出成功且生成文件");

                var w2 = Convert.ToDouble(co.Width);
                var h2 = Convert.ToDouble(co.Height);
                Assert("chart-export-restore-w", Math.Abs(w2 - origW) < 1, "导出后图表宽度还原为原值（" + origW.ToString("0") + "→" + w2.ToString("0") + "）");
                Assert("chart-export-restore-h", Math.Abs(h2 - origH) < 1, "导出后图表高度还原为原值（" + origH.ToString("0") + "→" + h2.ToString("0") + "）");
            }
            finally
            {
                try { if (File.Exists(fp)) File.Delete(fp); } catch { }
                MarshalRelease((object)ch);
                MarshalRelease((object)co);
                MarshalRelease((object)ws);
            }
        }

        // S12: fills 缺失/为空 → INVALID_FILLS，不再“没有可填入的空白格”假成功。
        // 直接复用 RunAll 已持有的 Excel 实例构造服务（isControlledEnabled 默认 true），缺 fills 应在写前拦截。
        private static void FillMissing(dynamic app)
        {
            var svc = new ExcelAiFillService((object)app);
            var r = svc.ApplyFills(new Dictionary<string, object> { ["address"] = "A1:B6", ["sheet"] = "dt" });
            Assert("fill-missing", !r.Ok && r.ErrorCode == "INVALID_FILLS", "缺 fills → INVALID_FILLS（code=" + (r.Ok ? "ok" : r.ErrorCode) + "）");
        }

        private static ToolDescriptor Need(ToolRegistry registry, string name)
        {
            if (registry.TryGet(name, out var t))
            {
                return t;
            }

            throw new InvalidOperationException("工具未注册：" + name);
        }

        private static double Points(double cm) => cm * 28.3464567;

        private static void Assert(string id, bool condition, string message)
        {
            if (condition) { _passed++; Console.WriteLine("PASS " + id + " — " + message); }
            else { _failed++; Failures.Add(id + ": " + message); Console.WriteLine("FAIL " + id + " — " + message); }
        }

        private static string Safe(Exception ex) => ex?.Message ?? ex?.GetType().FullName ?? "未知异常";

        private static string SafeData(object data)
        {
            try
            {
                return data == null ? "<null>" : Convert.ToString(data);
            }
            catch (Exception ex)
            {
                return "<unreadable: " + ex.Message + ">";
            }
        }

        private static void MarshalRelease(object com)
        {
            if (com != null)
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
                catch { }
            }
        }
    }
}