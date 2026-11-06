using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyUtilityTools
{
    /// <summary>
    /// 真机冒烟：第二批 7 个通用数据加工工具（text_clean/text_split/text_merge/date_normalize/
    /// number_pad/fill_formula/merge_sheets）走与生产完全相同的 ExcelToolHost.CreateRegistry(turnLabel)
    /// 链路，在真实 Excel 里执行并回读断言。此前单测/相位冒烟只验证注册，这里第一次真正跑这批工具，
    /// 专用于捕获数组写入越界 / 行列交换 / 参数契约类缺陷。
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
                // 生产是无状态调用：工具与 writer 各自经 ActiveWorkbook/Worksheets 重新解析工作表，
                // 并在 finally 里 FinalRelease 其 COM 句柄。因此 harness 不能跨工具持有 Worksheet/range 句柄。
                AliasSheet(wb, 1, "dt");
                AddSheet(wb, "m1");
                AddSheet(wb, "m2");
                AddSheet(wb, "tgt");

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyut-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(workbookPath), "verifyut", "第 1 轮");
                var registry = host.CreateRegistry();

                SeedClean(app, "dt");
                CheckTextClean(registry, app, "dt");

                SeedSplit(app, "dt");
                CheckTextSplit(registry, app, "dt");

                SeedMerge(app, "dt");
                CheckTextMerge(registry, app, "dt");

                SeedDates(app, "dt");
                CheckDateNormalize(registry, app, "dt");

                SeedNumbers(app, "dt");
                CheckNumberPad(registry, app, "dt");

                ResetSmart(app, "dt");
                CheckFillFormula(registry, app, "dt");

                SeedSheets(app);
                CheckMergeSheets(registry, app, "tgt");
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

        private static ToolDescriptor Need(ToolRegistry registry, string name)
        {
            if (registry.TryGet(name, out var t))
            {
                return t;
            }

            throw new InvalidOperationException("工具未注册：" + name);
        }

        private static void CheckTextClean(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "text_clean").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A3", ["sheet"] = sheet,
                ["trim"] = true, ["case"] = "upper"
            });
            Assert("textclean-ok", res.Ok, "text_clean 执行成功");
            if (!res.Ok) return;
            Assert("textclean-trim-case", Cell(app, "A1") == "HELLO" && Cell(app, "A2") == "WORLD" && Cell(app, "A3") == "MIXED",
                "text_clean 去首尾空格+统一大写");

            var rep = Need(registry, "text_clean").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A3", ["sheet"] = sheet,
                ["find"] = "E", ["replace"] = "?"
            });
            Assert("textclean-replace", rep.Ok && Cell(app, "A1") == "H?LLO", "text_clean 查找替换 E→?");
        }

        private static void CheckTextSplit(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "text_split").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A3", ["sheet"] = sheet,
                ["column"] = 1, ["delimiter"] = ",", ["target_column"] = 2
            });
            Assert("textsplit-ok", res.Ok, "text_split 执行成功");
            if (!res.Ok) return;
            Assert("textsplit-parts", Cell(app, "B1") == "a" && Cell(app, "C1") == "b"
                && Cell(app, "B2") == "c" && Cell(app, "C2") == "d" && Cell(app, "B3") == "e",
                "text_split 拆分段写入 B/C 列（逗号前/后）");
        }

        private static void CheckTextMerge(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "text_merge").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B3", ["sheet"] = sheet,
                ["columns"] = new JArray { 1, 2 }, ["target_column"] = 3, ["delimiter"] = "_"
            });
            Assert("textmerge-ok", res.Ok, "text_merge 执行成功");
            if (!res.Ok) return;
            Assert("textmerge-out", Cell(app, "C1") == "a_1" && Cell(app, "C2") == "b_2", "text_merge 将 A/B 合并到 C，分隔符 _");
        }

        private static void CheckDateNormalize(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "date_normalize").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A3", ["sheet"] = sheet,
                ["column"] = 1, ["input_format"] = "yyyy/M/d", ["output_format"] = "yyyy-MM-dd"
            });
            Assert("date-ok", res.Ok, "date_normalize 执行成功");
            if (!res.Ok)
            {
                Console.WriteLine("  date detail: code=" + res.ErrorCode + " data=" + SafeData(res.Data));
                return;
            }
            Assert("date-out", Cell(app, "A1") == "2024-01-05" && Cell(app, "A2") == "2024-12-31", "date_normalize 输出 yyyy-MM-dd 文本");
        }

        private static void CheckNumberPad(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "number_pad").Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A3", ["sheet"] = sheet,
                ["column"] = 1, ["width"] = 4, ["prefix"] = "NO-"
            });
            Assert("numpad-ok", res.Ok, "number_pad 执行成功");
            if (!res.Ok)
            {
                Console.WriteLine("  numpad detail: code=" + res.ErrorCode + " data=" + SafeData(res.Data));
                return;
            }
            Assert("numpad-out", Cell(app, "A1") == "NO-0001" && Cell(app, "A2") == "NO-0012" && Cell(app, "A3") == "NO-0123",
                "number_pad 补零+前缀 NO-");
        }

        private static void CheckFillFormula(ToolRegistry registry, dynamic app, string sheet)
        {
            var res = Need(registry, "fill_formula").Execute(new Dictionary<string, object>
            {
                ["address"] = "B1:B3", ["sheet"] = sheet,
                ["formula"] = "=ROW()", ["start_row"] = 1
            });
            Assert("fill-ok", res.Ok, "fill_formula 执行成功");
            if (!res.Ok) return;
            Assert("fill-result", Cell(app, "B2") == "2" && Cell(app, "B3") == "3", "fill_formula 批量填入 =ROW() 得到行号 2/3");
        }

        private static void CheckMergeSheets(ToolRegistry registry, dynamic app, string target)
        {
            var res = Need(registry, "merge_sheets").Execute(new Dictionary<string, object>
            {
                ["source_sheets"] = new JArray { "m1", "m2" },
                ["target_sheet"] = target, ["include_headers"] = true
            });
            Assert("merge-ok", res.Ok, "merge_sheets 执行成功");
            if (!res.Ok)
            {
                Console.WriteLine("  merge detail: code=" + res.ErrorCode + " data=" + SafeData(res.Data));
                return;
            }

            Assert("merge-header", SheetCell(app, "tgt", "A1") == "Name", "merge_sheets 带首个源表表头");
            Assert("merge-rows", SheetCell(app, "tgt", "A2") == "alice" && SheetCell(app, "tgt", "A3") == "bob", "merge_sheets 纵向合并 m1+m2 数据");
        }

        // ===================== 数据种子 =====================
        private static void SeedClean(dynamic app, string sheet)
        {
            Set(app, sheet, new object[,] { { "  HELLO  " }, { "WORLD" }, { "  MixED  " } }, "A1:A3");
        }

        private static void SeedSplit(dynamic app, string sheet)
        {
            Set(app, sheet, new object[,] { { "a,b" }, { "c,d" }, { "e" } }, "A1:A3");
        }

        private static void SeedMerge(dynamic app, string sheet)
        {
            Set(app, sheet, new object[,] { { "a", "1" }, { "b", "2" }, { "c", "3" } }, "A1:B3");
        }

        private static void SeedDates(dynamic app, string sheet)
        {
            Set(app, sheet, new object[,] { { "2024/1/5" }, { "2024/12/31" }, { "not-a-date" } }, "A1:A3");
        }

        private static void SeedNumbers(dynamic app, string sheet)
        {
            Set(app, sheet, new object[,] { { 1.0 }, { 12.0 }, { 123.0 } }, "A1:A3");
        }

        private static void ResetSmart(dynamic app, string sheet)
        {
            // 清空 B1:B3，避免残留值影响公式结果判定
            dynamic ws = app.Worksheets[sheet];
            dynamic rng = ws.Range["B1:B3"];
            rng.ClearContents();
            Release((object)rng);
            Release((object)ws);
        }

        private static void SeedSheets(dynamic app)
        {
            // 无状态：与 Set/ResetSmart 一致，直接从 app.Worksheets 解析，不经过 workbook 句柄，
            // 也绝不 FinalRelease ActiveWorkbook（共享根 RCW，会连带销毁 RunAll 持有的 wb）。
            // m1/m2 各 A1 表头 + 一行数据；tgt 留空（合并目标，已创建）
            dynamic ws1 = app.Worksheets["m1"];
            dynamic r1 = ws1.Range["A1:A2"];
            r1.Value2 = new object[,] { { "Name" }, { "alice" } };
            Release((object)r1);
            Release((object)ws1);

            dynamic ws2 = app.Worksheets["m2"];
            dynamic r2 = ws2.Range["A1:A2"];
            r2.Value2 = new object[,] { { "Name" }, { "bob" } };
            Release((object)r2);
            Release((object)ws2);

            dynamic wst = app.Worksheets["tgt"];
            dynamic rt = wst.Range["A1:A5"];
            rt.ClearContents();
            Release((object)rt);
            Release((object)wst);
        }

        private static void Set(dynamic app, string sheetName, object[,] values, string address)
        {
            var ws = app.Worksheets[sheetName];
            var rng = ws.Range[address];
            rng.Value2 = values;
            Release((object)rng);
            Release((object)ws);
        }

        private static void AliasSheet(dynamic wb, int index, string name)
        {
            dynamic ws = wb.Worksheets.Item[index];
            ws.Name = name;
            Release((object)ws);
        }

        private static void AddSheet(dynamic wb, string name)
        {
            dynamic ws = wb.Worksheets.Add();
            ws.Name = name;
            Release((object)ws);
        }

        private static string Cell(dynamic app, string addr)
        {
            dynamic ws = app.Worksheets["dt"];
            try
            {
                var v = ws.Range[addr].Value2;
                return AsStr(v);
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static string SheetCell(dynamic app, string sheetName, string addr)
        {
            dynamic ws = app.Worksheets[sheetName];
            try
            {
                var v = ws.Range[addr].Value2;
                return AsStr(v);
            }
            finally
            {
                Release((object)ws);
            }
        }

        private static string AsStr(object v)
        {
            return v == null ? string.Empty : System.Convert.ToString(v);
        }

        private static void Assert(string id, bool condition, string message)
        {
            if (condition) { _passed++; Console.WriteLine("PASS " + id + " — " + message); }
            else { _failed++; Failures.Add(id + ": " + message); Console.WriteLine("FAIL " + id + " — " + message); }
        }

        private static string Safe(Exception ex) => ex?.Message ?? ex?.GetType().FullName ?? "未知异常";

        private static string SafeData(object data)
        {
            try { return data == null ? "<null>" : Convert.ToString(data); }
            catch (Exception ex) { return "<unreadable: " + ex.Message + ">"; }
        }

        private static void Release(object com)
        {
            if (com != null)
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
                catch { }
            }
        }
    }
}