using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyDataTools
{
    /// <summary>
    /// 真机冒烟：5 个数据加工工具走与生产完全相同的 ExcelToolHost.CreateRegistry(turnLabel) 链路，
    /// 在真实 Excel 里执行并断言结果值。这是第一次真正执行这批工具——
    /// 专用于捕获数组写入越界类缺陷（此前单测/相位冒烟只验证注册）。
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
                // 生产是无状态调用：工具与 writer 会各自通过 ActiveWorkbook/Worksheets 重新解析工作表，
                // 并在 finally 里 FinalRelease 其 COM 句柄（与这里解析出的 ws 是同一 .NET RCW）。
                // 因此 harness 不能跨工具调用持有 ws，否则会被 write 路径连带销毁而抛「RCW 已分离」。
                Release((object)ws); ws = null;

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifydt-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(workbookPath), "verifydt", "第 1 轮");
                var registry = host.CreateRegistry();

                Seed(app);
                DebugSort(app, host);
                Seed(app);
                CheckSort(registry, app, "dt");
                CheckSortStable(registry, app, "dt");
                Seed(app);
                CheckDedup(registry, app, "dt");
                CheckDedupListCols(registry, app, "dt");
                Seed(app);
                CheckFindReplace(registry, app, "dt");
                Seed(app);
                CheckTranspose(registry, app, "dt");
                CheckTransposeDefault(registry, app, "dt");
                Seed(app);
                CheckCondFormat(registry, app, "dt");
                Seed(app);
                CheckPivot(registry, app, "dt");
                CheckPivotBound(registry, app, "dt");
                CheckTransformFormulas(registry, app, "dt");
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

        private static void DebugSort(dynamic app, ExcelToolHost host)
        {
            try
            {
                var probe = app.ActiveWorkbook.Name;
                Console.WriteLine("PROBE app.ActiveWorkbook.Name=" + Convert.ToString(probe));
            }
            catch (Exception ex)
            {
                Console.WriteLine("PROBE app ALIVE? NO — " + ex.GetType().FullName + " :: " + ex.Message);
            }

            ToolResult captured = null;
            var svc = new ExcelDataToolsService((object)app,
                (Func<Dictionary<string, object>, object[,], string, string, ToolResult>)(
                    (arg, vals, mode, tool) => { captured = new ToolResult { Ok = true, Data = new { shape = vals.GetLength(0) + "x" + vals.GetLength(1) } }; return captured; }));
            var args = new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = "dt",
                ["sort_column"] = 2, ["direction"] = "asc", ["header_rows"] = 1
            };
            try
            {
                var r = svc.SortRange(args);
                Console.WriteLine("DIRECT sort Ok=" + r.Ok + " code=" + r.ErrorCode + " data=" + SafeData(r.Data) + " capturedShape=" + (captured?.Data == null ? "null" : Convert.ToString(captured.Data)));
            }
            catch (Exception ex)
            {
                Console.WriteLine("DIRECT EX: " + ex.GetType().FullName + " :: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

            ProbeHostSort(host, app);
            ProbeWriteGeneric(host);
            ProbeServiceWithRealWriter(host, app);
            ProbeWriteArray(app);
        }

        private static void ProbeServiceWithRealWriter(ExcelToolHost host, dynamic app)
        {
            try
            {
                dynamic ss = app.ActiveSheet;
                var ra = ss.Range["A1:B5"];
                ra.Value2 = new object[,] { { "Col", "B" }, { "a", 3D }, { "b", 1D }, { "a", 2D }, { "c", 5D } };
                Release((object)ra);

                var svc = new ExcelDataToolsService((object)app,
                    (Func<Dictionary<string, object>, object[,], string, string, ToolResult>)(
                        (arg, vals, mode, tool) => host.WriteGenericRange(arg, vals, mode, tool)));
                var args = new Dictionary<string, object>
                {
                    ["address"] = "A1:B5", ["sheet"] = "dt",
                    ["sort_column"] = 2, ["direction"] = "asc", ["header_rows"] = 1
                };
                var r = svc.SortRange(args);
                Console.WriteLine("SVC+REALWRITER sort Ok=" + r.Ok + " code=" + r.ErrorCode + " data=" + SafeData(r.Data));
            }
            catch (Exception ex)
            {
                Console.WriteLine("SVC+REALWRITER EX: " + ex.GetType().FullName + " :: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void ProbeWriteGeneric(ExcelToolHost host)
        {
            try
            {
                var arg = new Dictionary<string, object>
                {
                    ["address"] = "A1:B5", ["sheet"] = "dt"
                };
                var matrix = new object[,]
                {
                    { "Col", "B" }, { "b", 1D }, { "a", 2D }, { "a", 3D }, { "c", 5D }
                };
                var res = host.WriteGenericRange(arg, matrix, "value", "probe_write");
                Console.WriteLine("WRITEGENERIC Ok=" + res.Ok + " code=" + res.ErrorCode + " data=" + SafeData(res.Data));
            }
            catch (Exception ex)
            {
                Console.WriteLine("WRITEGENERIC EX: " + ex.GetType().FullName + " :: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void ProbeHostSort(ExcelToolHost host, dynamic app)
        {
            try
            {
                dynamic ss = app.ActiveSheet;
                var ra = ss.Range["A1:B5"];
                ra.Value2 = new object[,] { { "Col", "B" }, { "a", 3D }, { "b", 1D }, { "a", 2D }, { "c", 5D } };
                Release((object)ra);

                var registry = host.CreateRegistry();
                if (registry.TryGet("sort_range", out var t))
                {
                    var res = t.Execute(new Dictionary<string, object>
                    {
                        ["address"] = "A1:B5", ["sheet"] = "dt",
                        ["sort_column"] = 2, ["direction"] = "asc", ["header_rows"] = 1
                    });
                    Console.WriteLine("HOST sort Ok=" + res.Ok + " code=" + res.ErrorCode + " data=" + SafeData(res.Data));
                }
                else
                {
                    Console.WriteLine("HOST sort_range NOT REGISTERED");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("HOST EX: " + ex.GetType().FullName + " :: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void ProbeWriteArray(dynamic app)
        {
            try
            {
                dynamic ss = app.ActiveSheet;
                var zeroBased = new object[2, 2] { { "z1", "z2" }, { "i1", "i2" } };
                var ra = ss.Range["A1:B2"];
                ra.Value2 = zeroBased;
                Console.WriteLine("0BASED WRITE ok: A1=" + Convert.ToString(ss.Range["A1"].Value2));
                Release((object)ra);
            }
            catch (Exception ex)
            {
                Console.WriteLine("0BASED WRITE EX: " + ex.GetType().FullName + " :: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        // A1:B5 = 表头 "Col","B" + 数据行 (a,3)(b,1)(a,2)(c,5)
        private static void Seed(dynamic app)
        {
            var seed = new object[,]
            {
                { "Col", "B" },
                { "a", 3D }, { "b", 1D }, { "a", 2D }, { "c", 5D }
            };
            dynamic ws = app.ActiveSheet;
            var rng = ws.Range["A1:B5"];
            rng.Value2 = seed;
            Release((object)rng);
            Release((object)ws);
        }

        private static ToolDescriptor Need(ToolRegistry registry, string name)
        {
            if (registry.TryGet(name, out var t))
            {
                return t;
            }

            throw new InvalidOperationException("工具未注册：" + name);
        }

        private static void CheckSort(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "sort_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["sort_column"] = 2, ["direction"] = "asc", ["header_rows"] = 1
            });
            Assert("sort-ok", res.Ok, "sort_range 执行成功（未越界）");
            if (!res.Ok)
            {
                Console.WriteLine("  sort detail: code=" + res.ErrorCode + " data=" + SafeData(res.Data));
                return;
            }
            // 按 B 列升序：数据行 B 应为 1,2,3,5；表头保持 Col/B
            Assert("sort-header", Cell(app, "A1") == "Col", "sort 表头保持");
            Assert("sort-b", Cell(app, "B2") == "1" && Cell(app, "B3") == "2" && Cell(app, "B4") == "3" && Cell(app, "B5") == "5", "sort 按 B 升序 1,2,3,5");
            Assert("sort-a", Cell(app, "A2") == "b" && Cell(app, "A5") == "c", "sort 关联 A 列同步重排");
        }

        // 稳定性闭环：同键行在排序后保持原始相对顺序（与 Excel 原生排序一致）。
        private static void CheckSortStable(ToolRegistry registry, dynamic app, string sheet)
        {
            var seed = new object[,]
            {
                { "Col", "B" },
                { "d", 1D }, { "b", 2D }, { "a", 2D }, { "c", 2D }
            };
            dynamic ws = app.ActiveSheet;
            var rng = ws.Range["A1:B5"];
            rng.Value2 = seed;
            Release((object)rng);
            Release((object)ws);

            var t = Need(registry, "sort_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["sort_column"] = 2, ["direction"] = "asc", ["header_rows"] = 1
            });
            Assert("sortstable-ok", res.Ok, "sort_range 稳定性场景执行成功");
            if (!res.Ok) return;
            // 同键（B=2）行 b/a/c 应保持原始相对顺序 b,a,c
            Assert("sortstable-order", Cell(app, "A3") == "b" && Cell(app, "A4") == "a" && Cell(app, "A5") == "c", "同键行 b,a,c 保持原始相对顺序（稳定排序）");
        }

        private static void CheckDedup(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "dedup_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["columns"] = new Newtonsoft.Json.Linq.JArray { 1 }, ["keep"] = "first", ["header_rows"] = 1
            });
            Assert("dedup-ok", res.Ok, "dedup_range 执行成功（未越界）");
            if (!res.Ok) return;
            // 按第1列去重 keep=first：a,b,c，尾部补空
            Assert("dedup-header", Cell(app, "A1") == "Col", "dedup 表头保持");
            Assert("dedup-data", Cell(app, "A2") == "a" && Cell(app, "A3") == "b" && Cell(app, "A4") == "c", "dedup 顺序 a,b,c");
            Assert("dedup-pad", Cell(app, "A5") == "", "dedup 尾部补空");
        }

        // columns 多形态闭环：以 object[] 而非 JArray 传列号，验证防御性解析不静默退化为整行判重。
        private static void CheckDedupListCols(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "dedup_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["columns"] = new object[] { 1 }, ["keep"] = "first", ["header_rows"] = 1
            });
            Assert("deduplist-ok", res.Ok, "dedup_range（object[] columns）执行成功");
            if (!res.Ok) return;
            // 按第1列去重 keep=first：a,b,c（仅当 columns 真正生效，而非整行判重导致全保留）
            Assert("deduplist-data", Cell(app, "A2") == "a" && Cell(app, "A3") == "b" && Cell(app, "A4") == "c", "columns(object[])=1 判重得 a,b,c");
        }

        private static void CheckFindReplace(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "find_replace_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:A5", ["sheet"] = sheet,
                ["find"] = "a", ["replace"] = "X", ["match_case"] = false
            });
            Assert("findreplace-ok", res.Ok, "find_replace_range 执行成功");
            if (!res.Ok) return;
            Assert("findreplace-header", Cell(app, "A1") == "Col", "find_replace 表头未受污染");
            Assert("findreplace-data", Cell(app, "A2") == "X" && Cell(app, "A3") == "b" && Cell(app, "A4") == "X", "find_replace 替换 a→X");
        }

        private static void CheckTranspose(ToolRegistry registry, dynamic app, string sheet)
        {
            // 源仍在 A1:B3 已 seed 的数据 (a,3)(b,1)；转置写入 E1，不覆盖源
            var t = Need(registry, "transpose_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B2", ["sheet"] = sheet,
                ["target_sheet"] = sheet, ["target_address"] = "E1"
            });
            Assert("transpose-ok", res.Ok, "transpose_range 执行成功");
            if (!res.Ok) return;
            // 2x2 转置：A1:B2 = [[Col,B],[a,3]] → 行/列互换 → E1:F2 = [[Col,a],[B,3]]
            Assert("transpose-v", Cell(app, "E1") == "Col" && Cell(app, "F1") == "a" && Cell(app, "E2") == "B" && Cell(app, "F2") == "3", "转置结果 E1:F2 正确");
        }

        // 缺省写出位闭环：未传 target_address 时结果应写到源区域左上角（而非 A1）。
        private static void CheckTransposeDefault(ToolRegistry registry, dynamic app, string sheet)
        {
            var seed = new object[,]
            {
                { 1D, "x" },
                { 2D, "y" }
            };
            dynamic ws = app.ActiveSheet;
            ws.Range["A1"].Value2 = ""; // 清空 A1，用于识别旧缺陷默认落位
            var rng = ws.Range["B2:C3"];
            rng.Value2 = seed;
            Release((object)rng);
            Release((object)ws);

            var t = Need(registry, "transpose_range");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B2:C3", ["sheet"] = sheet, ["target_sheet"] = sheet
            });
            Assert("transposedefault-ok", res.Ok, "transpose_range 缺省写出位场景执行成功");
            if (!res.Ok) return;
            // 转置 B2:C3=[[1,x],[2,y]] → [[1,2],[x,y]]，缺省写回源左上角 B2：B2=1,C2=2,B3=x,C3=y；A1 保持空
            Assert("transposedefault-pos", Cell(app, "B2") == "1" && Cell(app, "C2") == "2" && Cell(app, "B3") == "x" && Cell(app, "C3") == "y", "转置缺省写回源区域左上角 B2");
            Assert("transposedefault-not-a1", Cell(app, "A1") == "", "缺省写出位不是 A1");
        }

        private static void CheckCondFormat(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "format_excel_conditional");
            var greater = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B2:B5", ["sheet"] = sheet, ["rule"] = "greater", ["value"] = 2
            });
            Assert("cond-greater", greater.Ok, "条件格式 greater 执行成功");

            var bar = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B2:B5", ["sheet"] = sheet, ["rule"] = "data_bar"
            });
            Assert("cond-databar", bar.Ok, "条件格式 data_bar 执行成功");

            var clear = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B2:B5", ["sheet"] = sheet, ["rule"] = "clear"
            });
            Assert("cond-clear", clear.Ok, "条件格式 clear 执行成功");

            var bad = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B2:B5", ["sheet"] = sheet, ["rule"] = "nonsense"
            });
            Assert("cond-invalid", !bad.Ok, "未知规则被拒绝");
        }

        // count 聚合按行计数（不限于可解析为数值的单元格）；越界列号显式 INVALID_COLUMN。
        private static void CheckPivot(ToolRegistry registry, dynamic app, string sheet)
        {
            var seed = new object[,]
            {
                { "cls", "amount" },
                { "A", 1D }, { "A", "x" }, { "B", 3D }, { "B", 5D }
            };
            dynamic ws = app.ActiveSheet;
            var rng = ws.Range["A1:B5"];
            rng.Value2 = seed;
            Release((object)rng);
            Release((object)ws);

            var t = Need(registry, "create_pivot_table");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["row_field"] = 1, ["value_field"] = 2, ["aggregation"] = "count",
                ["output_address"] = "D1"
            });
            Assert("pivot-count-ok", res.Ok, "create_pivot_table(count) 执行成功");
            if (!res.Ok) return;
            // A 组 2 行（1 与 x），count 应计 2（按行而非仅数值）；B 组 2 行
            // 输出 4 列：D=key，E=聚合值，F=聚合方式，G=取值。A/B 组各 2 行 → E2=E3=2
            var a = Cell(app, "E2");
            var b = Cell(app, "E3");
            Assert("pivot-count-rows", a == "2" && b == "2", "count 按分组行数计（A=2,B=2）");
        }

        private static void CheckPivotBound(ToolRegistry registry, dynamic app, string sheet)
        {
            var t = Need(registry, "create_pivot_table");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:B5", ["sheet"] = sheet,
                ["row_field"] = 1, ["value_field"] = 99, ["aggregation"] = "sum",
                ["output_address"] = "D1"
            });
            Assert("pivot-bound", !res.Ok && res.ErrorCode == "INVALID_COLUMN", "value_field 越界被显式拒绝(INVALID_COLUMN)");
        }

        // F8 闭环：以 JArray 形态传 actions（复现 JArray 非 IEnumerable<object> 的漏判），验证公式转换真正生效。
        private static void CheckTransformFormulas(ToolRegistry registry, dynamic app, string sheet)
        {
            dynamic ws = app.ActiveSheet;
            var rng = ws.Range["B1:B2"];
            rng.Cells[1, 1].Formula = "=A1*2";
            rng.Cells[2, 1].Value2 = 5D;
            Release((object)rng);
            Release((object)ws);

            var actions = new Newtonsoft.Json.Linq.JArray(
                new Newtonsoft.Json.Linq.JObject { ["type"] = "iferror", ["fallback"] = "0" });

            var t = Need(registry, "transform_excel_formulas");
            var res = t.Execute(new Dictionary<string, object>
            {
                ["address"] = "B1:B2", ["sheet"] = sheet, ["actions"] = actions
            });
            Assert("tf-ok", res.Ok, "transform_excel_formulas（JArray actions）执行成功");
            if (!res.Ok) return;
            var f1 = Convert.ToString(app.ActiveSheet.Range["B1"].Formula);
            Assert("tf-wrapped", f1.Contains("IFERROR"), "B1 公式被 IFERROR 包裹");
            var f2 = Convert.ToString(app.ActiveSheet.Range["B2"].Formula);
            Assert("tf-skip", f2 == "5", "B2 非公式单元格未被转换");
        }

        private static string Cell(dynamic app, string addr)
        {
            dynamic ws = app.ActiveSheet;
            try
            {
                var v = ws.Range[addr].Value2;
                return v == null ? string.Empty : Convert.ToString(v);
            }
            finally
            {
                Release((object)ws);
            }
        }

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