using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyJyyjCommon
{
    /// <summary>
    /// 「JYYJ快捷」常用 JYYJ 组（第二批）核心工具真机验证 harness。
    /// 独立 out-of-proc Excel + 未存盘工作簿，用 ExcelToolHost.CreateRegistry() 组装真实工具表，
    /// 直接 descriptor.Execute 跑通 8 个常用组工具的完整执行链（播种 → Execute → 回读单元格/结果域断言）。
    /// 覆盖本次对抗审查修复：fill_number_sequence 多格选区的 start 左上角起点；
    /// 行结构类（delete_blank_rows/insert_blank_rows/merge_duplicate）倒序删除与合并正确性。
    /// 结束按 PID 强杀自建实例，绝不触碰用户 Excel。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly HashSet<int> _excelBefore = new HashSet<int>();
        private static readonly HashSet<int> _spawned = new HashSet<int>();
        private static string _sheet = "Sheet1";

        private static int Main()
        {
            Console.WriteLine("== 常用组（第二批）8 工具真机验证 ==");
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                _excelBefore.Add(p.Id);
                p.Dispose();
            }

            dynamic app = null;
            ToolRegistry registry = null;
            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null)
                {
                    Console.WriteLine("[fatal] 未找到 Excel.Application COM（本机未安装 Excel？）");
                    return 2;
                }

                app = Activator.CreateInstance(type);
                app.Visible = false;
                app.DisplayAlerts = false;
                app.AutomationSecurity = 1;

                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!_excelBefore.Contains(p.Id)) _spawned.Add(p.Id);
                    p.Dispose();
                }

                app.Workbooks.Add();
                _sheet = Convert.ToString(app.ActiveWorkbook.Sheets.Item[1].Name);
                var host = new ExcelToolHost(app);
                registry = host.CreateRegistry();
                Console.WriteLine("[env] 工具注册表就绪，共 " + registry.Tools.Count + " 个工具。\n");

                VerifyInsertText(app, registry);
                VerifyDeleteText(app, registry);
                VerifyFillRandom(app, registry);
                VerifyFillSequenceMultiStart(app, registry);
                VerifyInsertBlankRows(app, registry);
                VerifyDeleteBlankRows(app, registry);
                VerifyFillBlankDown(app, registry);
                VerifyMergeDuplicate(app, registry);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[fatal] harness 异常：" + ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                var killed = new List<int>();
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (_spawned.Contains(p.Id))
                    {
                        try { p.Kill(); killed.Add(p.Id); } catch { }
                    }
                    p.Dispose();
                }

                Console.WriteLine("\n[cleanup] 已结束自建 EXCEL 进程 " + killed.Count + " 个");
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.WriteLine("\n== 结果：通过 " + _passed + "，失败 " + _failed + " ==");
            return _failed == 0 ? 0 : 1;
        }

        // ===================== 各工具验证 =====================

        private static void VerifyInsertText(dynamic app, ToolRegistry registry)
        {
            const string tool = "insert_text_into_cell";
            Seed(app, "$A$1", "abcdef");
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1",
                ["sheet"] = _sheet,
                ["insertion"] = "X",
                ["when"] = "char",
                ["where"] = "after",
                ["index"] = "2"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(CellText(app, "A1") == "abXcdef", tool + " 在字符位置后插入 X", "A1=" + CellText(app, "A1"));
        }

        private static void VerifyDeleteText(dynamic app, ToolRegistry registry)
        {
            const string tool = "delete_text_matches";
            Seed(app, "$A$1:$A$3", new object[,] { { "AB12CD34" }, { "x1y2" }, { "nomatch" } });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1:$A$3",
                ["sheet"] = _sheet,
                ["pattern"] = "\\d+"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(CellText(app, "A1") == "ABCD" && CellText(app, "A2") == "xy", tool + " 删除数字匹配", "A1=" + CellText(app, "A1") + ", A2=" + CellText(app, "A2"));
            Expect(CellText(app, "A3") == "nomatch", tool + " 无匹配保持原样", "A3=" + CellText(app, "A3"));
            Expect(IntField(result, "changed_cells") == 2, tool + " changed_cells=2", "changed=" + IntField(result, "changed_cells"));
        }

        private static void VerifyFillRandom(dynamic app, ToolRegistry registry)
        {
            const string tool = "fill_random_numbers";
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$B$1:$B$5",
                ["sheet"] = _sheet,
                ["min"] = "0",
                ["max"] = "100",
                ["decimals"] = "0"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            var inRange = true;
            for (var r = 1; r <= 5; r++)
            {
                var v = Convert.ToDouble(EqCell(app, "B" + r));
                if (v < 0 || v > 100) inRange = false;
            }
            Expect(inRange, tool + " 5 个随机数在 [0,100]", "B1=" + CellText(app, "B1") + ", B5=" + CellText(app, "B5"));
        }

        private static void VerifyFillSequenceMultiStart(dynamic app, ToolRegistry registry)
        {
            const string tool = "fill_number_sequence";
            // 修复回归：选区多格（B1:B3=5,10,15），未传 start，须以选区左上角（5）为起点向下填步长 1 的 4 个数。
            Seed(app, "$B$1:$B$3", new object[,] { { 5D }, { 10D }, { 15D } });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$B$1",
                ["sheet"] = _sheet,
                ["step"] = "1",
                ["count"] = "4",
                ["direction"] = "down"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqDouble(app, "B1") == 5D && EqDouble(app, "B2") == 6D && EqDouble(app, "B4") == 8D,
                tool + " 左上角 5 起向下 5,6,7,8", "B1..B4=" + CellText(app, "B1") + "," + CellText(app, "B2") + "," + CellText(app, "B3") + "," + CellText(app, "B4"));

            // 修复回归核心：选区为多格但显式 start 时必须用它，不被选区覆盖。
            Seed(app, "$B$1:$B$3", new object[,] { { 99D }, { 98D }, { 97D } });
            result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$B$1",
                ["sheet"] = _sheet,
                ["start"] = "100",
                ["step"] = "2",
                ["count"] = "3",
                ["direction"] = "down"
            });
            Expect(result.Ok, tool + " 显式 start 执行成功", Describe(result));
            Expect(EqDouble(app, "B1") == 100D && EqDouble(app, "B3") == 104D, tool + " 显式 start 100 起 100,102,104", "B1..B3=" + CellText(app, "B1") + "," + CellText(app, "B2") + "," + CellText(app, "B3"));
        }

        private static void VerifyInsertBlankRows(dynamic app, ToolRegistry registry)
        {
            const string tool = "insert_blank_rows_every_n";
            Seed(app, "$A$1:$A$4", new object[,] { { "r1" }, { "r2" }, { "r3" }, { "r4" } });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1:$A$4",
                ["sheet"] = _sheet,
                ["interval"] = "2",
                ["count"] = "1"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            // 每 2 行插 1 空行：r1,r2 | 空 | r3,r4
            Expect(CellText(app, "A1") == "r1" && CellText(app, "A2") == "r2", tool + " 前两行原样保留", "A1,A2=" + CellText(app, "A1") + "," + CellText(app, "A2"));
            Expect(CellText(app, "A3") == "" && CellText(app, "A4") == "r3" && CellText(app, "A5") == "r4", tool + " 第 2 行后插空行，原 r3/r4 顺延", "A3..A5=" + CellText(app, "A3") + "," + CellText(app, "A4") + "," + CellText(app, "A5"));
            // 更大间隔：interval=3 时只在该分组后插，末尾不足一组不额外补空
            Seed(app, "$C$1:$C$6", new object[,] { { 1D }, { 2D }, { 3D }, { 4D }, { 5D }, { 6D } });
            result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$C$1:$C$6",
                ["sheet"] = _sheet,
                ["interval"] = "3",
                ["count"] = "2"
            });
            Expect(result.Ok, tool + " interval=3 执行成功", Describe(result));
            // 分组 (1,2,3)|空|空|(4,5,6)
            Expect(CellText(app, "C3") == "3" && CellText(app, "C4") == "" && CellText(app, "C5") == "" && CellText(app, "C6") == "4",
                tool + " interval=3 在第 3 行后插 2 空行", "C3..C6=" + CellText(app, "C3") + "," + CellText(app, "C4") + "," + CellText(app, "C5") + "," + CellText(app, "C6"));
        }

        private static void VerifyDeleteBlankRows(dynamic app, ToolRegistry registry)
        {
            const string tool = "delete_blank_rows_range";
            // 首列空(A2/A4)删，整行空(A3)删；A1/A5 保留。
            Seed(app, "$A$1:$B$5", new object[,]
            {
                { "h", 1D }, { "", 2D }, { "", null }, { "", 4D }, { "keep", 5D }
            });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1:$B$5",
                ["sheet"] = _sheet,
                ["only_fully_blank"] = "false"
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(CellText(app, "A1") == "h" && CellText(app, "A2") == "keep", tool + " 首列为空行被删，保留列表头", "A1,A2=" + CellText(app, "A1") + "," + CellText(app, "A2"));
            Expect(IntField(result, "deleted_rows") == 3, tool + " deleted_rows=3", "deleted=" + IntField(result, "deleted_rows"));
        }

        private static void VerifyFillBlankDown(dynamic app, ToolRegistry registry)
        {
            const string tool = "fill_blank_cells_down";
            Seed(app, "$A$1:$A$4", new object[,] { { "x" }, { null }, { "y" }, { null } });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1:$A$4",
                ["sheet"] = _sheet
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(CellText(app, "A2") == "x" && CellText(app, "A4") == "y", tool + " 空单元格沿用上方值", "A2,A4=" + CellText(app, "A2") + "," + CellText(app, "A4"));
        }

        private static void VerifyMergeDuplicate(dynamic app, ToolRegistry registry)
        {
            const string tool = "merge_duplicate_cells";
            Seed(app, "$A$1:$A$5", new object[,] { { "甲" }, { "甲" }, { "乙" }, { "乙" }, { "乙" } });
            var result = Run(registry, tool, new Dictionary<string, object>
            {
                ["address"] = "$A$1:$A$5",
                ["sheet"] = _sheet
            });

            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(CellText(app, "A1") == "甲", tool + " 首组保留左上甲", "A1=" + CellText(app, "A1"));
            // 乙组 3 格合并，值保留在组顶部行（原第 3 行）。
            Expect(CellText(app, "A3") == "乙", tool + " 次组顶部乙", "A3=" + CellText(app, "A3"));
        }

        // ===================== 执行与断言助手 =====================

        private static ToolResult Run(ToolRegistry registry, string tool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }

            try
            {
                return descriptor.Execute(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EX-STACK][" + tool + "] " + ex + "\n");
                return new ToolResult { Ok = false, ErrorCode = "HARNESS_EXCEPTION", Data = ex.Message };
            }
        }

        private static void Seed(dynamic app, string address, object value)
        {
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = app.ActiveWorkbook.Sheets.Item[_sheet];
                range = sheet.Range[address];
                range.Value2 = value;
            }
            finally
            {
                Release(range);
                Release(sheet);
            }
        }

        private static void Release(object com)
        {
            if (com != null && System.Runtime.InteropServices.Marshal.IsComObject(com))
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
                catch { }
            }
        }

        private static string CellText(dynamic app, string address)
        {
            var v = EqCell(app, address);
            return v == null ? string.Empty : Convert.ToString(v);
        }

        private static object EqCell(dynamic app, string address)
        {
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = app.ActiveWorkbook.Sheets.Item[_sheet];
                range = sheet.Range[address];
                // 用 Value2 读原始值即可，无需设置 NumberFormat（避免对 COM Range 设置属性抛异常）
                var v = range.Value2;
                return v is DBNull ? null : v;
            }
            finally
            {
                Release(range);
                Release(sheet);
            }
        }

        private static double EqDouble(dynamic app, string address)
        {
            var v = EqCell(app, address);
            return v == null ? double.NaN : Convert.ToDouble(v);
        }

        private static int? IntField(ToolResult result, string key)
        {
            if (result == null || result.Data == null)
            {
                return null;
            }
            try
            {
                var jo = JToken.FromObject(result.Data) as JObject;
                var t = jo?[key];
                if (t == null)
                {
                    return null;
                }
                return Convert.ToInt32(t.ToObject<object>());
            }
            catch
            {
                return null;
            }
        }

        private static void Expect(bool condition, string name, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("[PASS] " + name + "　→ " + detail);
            }
            else
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name + "　→ " + detail);
            }
        }

        private static string Describe(ToolResult result)
        {
            if (result == null)
            {
                return "null";
            }
            return result.Ok ? "ok" : (result.ErrorCode ?? "failed") + " " + result.Data;
        }
    }
}