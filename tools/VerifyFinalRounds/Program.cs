using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyFinalRounds
{
    /// <summary>
    /// 真机验证本轮深审修复点（非法枚举显式报错 + 全量读入护栏 + source_count 实计数 + 写路径 Guard）。
    /// 全部走生产 ExcelToolHost.CreateRegistry 链路在真实 Excel 中执行。harness 无状态：
    /// 每次操作前重新解析 sheet、用完即放，绝不跨工具调用持有 Worksheet RCW。
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
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                ws.Name = "dt";
                Release((object)ws);

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyfr-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(workbookPath), "verifyfr", "第 1 轮");
                var registry = host.CreateRegistry();

                Seed(app);
                CheckPageSetup(registry, "dt");
                CheckChartType(registry, "dt");
                CheckSparklineColor(registry, "dt");
                CheckSortDirection(registry, "dt");
                CheckMarkdownGuard(registry, "dt");
                CheckMergeSources(registry, app);
                CheckPdfGuard(registry, "dt");
                CheckRateChangeGuard(registry, "dt");
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); Release((object)wb); } }
                catch { }
                try { app.Quit(); }
                catch { }
                Release((object)app);
            }
        }

        private static void CheckPageSetup(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "apply_page_setup");

            var bad = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["orientation"] = "diagonal"
            });
            Assert("pagesetup-bad-orientation", !bad.Ok && bad.ErrorCode == "INVALID_ORIENTATION",
                "apply_page_setup 非法 orientation 被显式拒绝(INVALID_ORIENTATION)，实际 " + bad.ErrorCode);

            var badPaper = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["paper_size"] = "duplex"
            });
            Assert("pagesetup-bad-papersize", !badPaper.Ok && badPaper.ErrorCode == "INVALID_PAPER_SIZE",
                "apply_page_setup 非法 paper_size 被显式拒绝(INVALID_PAPER_SIZE)，实际 " + badPaper.ErrorCode);

            var ok = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["orientation"] = "landscape", ["paper_size"] = "a4", ["fit_to_pages"] = 1
            });
            Assert("pagesetup-ok", ok.Ok, "apply_page_setup 合法参数应成功");
        }

        private static void CheckChartType(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "create_excel_chart");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["source_address"] = "A1:B5", ["sheet"] = sheet, ["chart_type"] = "pie3d_moon"
            });
            Assert("chart-bad-type", !res.Ok && res.ErrorCode == "INVALID_CHART_TYPE",
                "create_excel_chart 非法 chart_type 被显式拒绝(INVALID_CHART_TYPE)，实际 " + res.ErrorCode);
        }

        private static void CheckSparklineColor(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "create_excel_sparkline");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["data_range"] = "A1:A5", ["dest_range"] = "C2", ["sheet"] = sheet, ["color"] = "#zzz"
            });
            Assert("sparkline-bad-color", !res.Ok && res.ErrorCode == "INVALID_COLOR",
                "create_excel_sparkline 非法 color 被显式拒绝(INVALID_COLOR)，实际 " + res.ErrorCode);
        }

        private static void CheckSortDirection(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "sort_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["sort_column"] = 2, ["direction"] = "sideways", ["header_rows"] = 1
            });
            Assert("sort-bad-direction", !res.Ok && res.ErrorCode == "INVALID_DIRECTION",
                "sort_range 非法 direction 被显式拒绝(INVALID_DIRECTION)，实际 " + res.ErrorCode);
        }

        private static void CheckMarkdownGuard(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "export_selection_to_markdown");
            // 10 列 × 50000 行 = 500000 > 100000，须在拉取 Value2 前被 TOO_MANY_CELLS 拒绝。
            var res = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["address"] = "A1:J50000"
            });
            Assert("markdown-too-many", !res.Ok && res.ErrorCode == "TOO_MANY_CELLS",
                "export_selection_to_markdown 超大选区被显式拒绝(TOO_MANY_CELLS)，实际 " + res.ErrorCode);
        }

        private static void CheckMergeSources(ToolRegistry registry, dynamic app)
        {
            var tempDir = Path.Combine(Path.GetTempPath(), "verifyfr-src-" + Guid.NewGuid().ToString("N"));
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            var src1 = Path.Combine(tempDir, "src1.xlsx");
            var src2 = Path.Combine(tempDir, "src2.xlsx");
            var srcMissing = Path.Combine(tempDir, "missing.xlsx");
            if (File.Exists(src1)) File.Delete(src1);
            if (File.Exists(src2)) File.Delete(src2);
            if (File.Exists(srcMissing)) File.Delete(srcMissing);

            CreateOneSheetWorkbook(app, src1, "S1");
            CreateOneSheetWorkbook(app, src2, "S2");

            var t = Need(registry, "merge_workbooks");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["source_workbooks"] = new object[] { src1, src2, srcMissing }
            });

            // 行为断言：合并后当前工作簿应为 dt + S1 + S2 共 3 个表（missing.xlsx 不产生表）。
            int sheetCount = 0;
            try
            {
                dynamic wb = app.ActiveWorkbook;
                dynamic sheets = wb.Worksheets;
                sheetCount = Convert.ToInt32(sheets.Count);
                Release((object)sheets);
                Release((object)wb);
            }
            catch { }

            int sourceCount = ResolveInt(res.Data, "source_count");

            Assert("merge-source-count", res.Ok && sourceCount == 2,
                "merge_workbooks source_count 应=2（两个存在源，缺失的 missing.xlsx 不计入），实际 Ok=" + res.Ok + " count=" + sourceCount);

            Assert("merge-sheet-behavior", sheetCount == 3,
                "merge_workbooks 应只合并两个存在源（工作簿 dt+S1+S2=3 表），实际 " + sheetCount + " 表");
        }

        private static void CreateOneSheetWorkbook(dynamic app, string path, string sheetName)
        {
            dynamic wb = null;
            try
            {
                wb = app.Workbooks.Add();
                wb.Worksheets.Item[1].Name = sheetName;
                wb.SaveAs(path);
                wb.Close(false);
                Release((object)wb); wb = null;
            }
            finally
            {
                if (wb != null) { try { wb.Close(false); } catch { } Release((object)wb); }
            }
        }

        private static int ResolveInt(object data, string name)
        {
            if (data == null) return int.MinValue;
            if (data is System.Collections.IDictionary dict && dict.Contains(name))
            {
                var v = dict[name];
                if (v == null) return int.MinValue;
                return Convert.ToInt32(Convert.ToDouble(v));
            }

            var type = data.GetType();
            var prop = type.GetProperty(name);
            if (prop != null)
            {
                var v = prop.GetValue(data, null);
                if (v == null) return int.MinValue;
                return Convert.ToInt32(Convert.ToDouble(v));
            }

            var field = type.GetField(name);
            if (field != null)
            {
                var v = field.GetValue(data);
                if (v == null) return int.MinValue;
                return Convert.ToInt32(Convert.ToDouble(v));
            }

            return int.MinValue;
        }

        private static void CheckPdfGuard(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "export_sheets_to_pdf");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["file_path"] = "C:\\legit\\..\\..\\system\\..\\no.pdf"
            });
            Assert("pdf-path-guard", !res.Ok,
                "export_sheets_to_pdf 越界/穿越路径应被 Guard 拒绝（!Ok），实际 Ok=" + res.Ok + " code=" + res.ErrorCode);
        }

        private static void CheckRateChangeGuard(ToolRegistry registry, string sheet)
        {
            var t = Need(registry, "rate_change_analysis");

            // 超大区域 10×50000=500000>100000，须在拉取 Value2 前被 GuardException→TOO_MANY_CELLS 拒绝。
            var big = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["address"] = "A1:J50000", ["label_field"] = 1, ["value_field"] = 2
            });
            Assert("ratechange-too-many", !big.Ok && big.ErrorCode == "TOO_MANY_CELLS",
                "rate_change_analysis 超大选区被显式拒绝(TOO_MANY_CELLS)，实际 " + big.ErrorCode);

            // 合法小区域不应被护栏误伤。
            var ok = t.Execute(new Dictionary<string, object>
            {
                ["sheet"] = sheet, ["address"] = "A1:B5", ["label_field"] = 1, ["value_field"] = 2
            });
            Assert("ratechange-ok", ok.Ok,
                "rate_change_analysis 合法小区域应成功，实际 Ok=" + ok.Ok + " code=" + ok.ErrorCode);
        }

        private static void Seed(dynamic app)
        {
            try
            {
                dynamic ws = app.ActiveSheet;
                ws.Name = "dt";
                var range = ws.Range["A1:B5"];
                range.Value2 = new object[,]
                {
                    { "Col", "Head" }, { "a", 5D }, { "c", 2D }, { "b", 3D }, { "a", 4D }
                };
                Release((object)range);
                Release((object)ws);
            }
            catch { }
        }

        private static ToolDescriptor Need(ToolRegistry registry, string name)
        {
            if (!registry.TryGet(name, out var tool))
            {
                throw new Exception("注册表缺失工具: " + name);
            }

            return tool;
        }

        private static void Assert(string name, bool condition, string message)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("PASS " + name + "  " + message);
            }
            else
            {
                _failed++;
                Failures.Add(name + ": " + message);
                Console.WriteLine("FAIL " + name + "  " + message);
            }
        }

        private static string Safe(Exception ex)
        {
            return ex == null ? "null" : (ex.GetType().FullName + " :: " + ex.Message);
        }

        private static void Release(object obj)
        {
            if (obj != null)
            {
                try { Marshal.FinalReleaseComObject(obj); }
                catch { }
            }
        }
    }
}