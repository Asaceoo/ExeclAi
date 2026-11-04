using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyParamOptions
{
    /// <summary>
    /// 「配置弹窗新增选项」真机验证（独立 out-of-proc Excel 实例，全新工作簿，不触碰任何用户文件）。
    /// 覆盖 v2.18.7.0 弹窗新增的真实选项，全部走真实注册表工具链（与弹窗驱动同一执行层）：
    ///   C1 转置跨表写回 target_sheet（Sheet1 → Sheet2）
    ///   C2 转置缺省写回原区域左上角（top_left）
    ///   C3 文本拆分 target_column=2（结果写到源列右侧相邻列）
    ///   C4 文本清洗 find/replace 查找替换
    ///   C5 透视汇总 target_address=G1 显式输出位置
    ///   C6 随机数填充 min=max=5 确定性校验
    /// 无状态 COM：每次读写都重新解析 工作簿/工作表（工具会释放共享 RCW，持旧代理会触发
    /// 「COM 对象与其基础 RCW 分开后就不能再使用」，见踩坑 §4.9/§7 无状态准则）。
    /// 渲染层（选项文本/Hint/默认值）由 tools/VerifyParamJson 闭环覆盖，本工具只验执行层真机闭环。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly HashSet<int> Spawned = new HashSet<int>();
        private static dynamic _app;

        private static int Main()
        {
            Console.WriteLine("== VerifyParamOptions：弹窗新增选项真机验证 ==");
            var before = new HashSet<int>();
            foreach (var p in Process.GetProcessesByName("EXCEL")) { before.Add(p.Id); p.Dispose(); }

            try
            {
                var t = Type.GetTypeFromProgID("Excel.Application");
                if (t == null) { Console.WriteLine("[fatal] 无 Excel.Application COM"); return 2; }
                _app = Activator.CreateInstance(t);
                _app.Visible = false;
                _app.DisplayAlerts = false;
                _app.AutomationSecurity = 1;

                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { Spawned.Add(p.Id); }
                    p.Dispose();
                }

                var wb = _app.Workbooks.Add();
                EnsureSheet(wb, "Sheet1");
                EnsureSheet(wb, "Sheet2");

                var host = new ExcelToolHost(_app);
                ToolRegistry registry = host.CreateRegistry();

                C1TransposeCrossSheet(registry);
                C2TransposeInPlace(registry);
                C3TextSplitTargetColumn(registry);
                C4TextCleanFindReplace(registry);
                C5PivotTargetAddress(registry);
                C6RandomNumbersRange(registry);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[fatal] harness 异常：" + ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                foreach (var id in Spawned)
                {
                    try { Process.GetProcessById(id).Kill(); } catch { }
                }
                Console.WriteLine("\n== 结果：通过 " + _passed + "，失败 " + _failed + " ==");
            }
            return _failed == 0 ? 0 : 1;
        }

        // ===================== C1 转置跨表写回（target_sheet） =====================
        private static void C1TransposeCrossSheet(ToolRegistry registry)
        {
            WriteTable("Sheet1", "A1:C2", new string[,] { { "a", "b", "c" }, { "d", "e", "f" } });
            Clear("Sheet2", "A1:H10");
            B(registry, "transpose_range", Range(
                ("address", "$A$1:$C$2"), ("sheet", "Sheet1"),
                ("target_sheet", "Sheet2"), ("target_address", "A1")),
                "C1 转置跨表写回 target_sheet=Sheet2");

            var m = ReadGrid("Sheet2", "A1:B3");
            Expect(m != null && Eq(m, 1, 1, "a") && Eq(m, 1, 2, "d")
                    && Eq(m, 2, 1, "b") && Eq(m, 2, 2, "e")
                    && Eq(m, 3, 1, "c") && Eq(m, 3, 2, "f"),
                "C1 Sheet2 收到转置结果（3 行 × 2 列：a d / b e / c f）",
                m == null ? "读取失败" : DumpGrid(m, 3, 2));
        }

        // ===================== C2 转置缺省写回源区左上角（top_left） =====================
        private static void C2TransposeInPlace(ToolRegistry registry)
        {
            WriteTable("Sheet1", "A1:C2", new string[,] { { "1", "2", "3" }, { "4", "5", "6" } });
            B(registry, "transpose_range", Range(
                ("address", "$A$1:$C$2"), ("sheet", "Sheet1")),
                "C2 转置缺省写回原区域左上角");

            var m = ReadGrid("Sheet1", "A1:B3");
            Expect(m != null && Eq(m, 1, 1, "1") && Eq(m, 1, 2, "4")
                    && Eq(m, 2, 1, "2") && Eq(m, 2, 2, "5")
                    && Eq(m, 3, 1, "3") && Eq(m, 3, 2, "6"),
                "C2 源区被原位覆盖为转置结果（1 4 / 2 5 / 3 6）",
                m == null ? "读取失败" : DumpGrid(m, 3, 2));
        }

        // ===================== C3 文本拆分 target_column=2（源列右侧相邻列） =====================
        private static void C3TextSplitTargetColumn(ToolRegistry registry)
        {
            WriteTable("Sheet1", "D1:D3", new string[,] { { "x,y" }, { "m,n" }, { "p,q" } });
            B(registry, "text_split", Range(
                ("address", "$D$1:$D$3"), ("sheet", "Sheet1"),
                ("column", "1"), ("delimiter", ","), ("target_column", "2"), ("keep_source", true)),
                "C3 文本拆分 target_column=2");

            var m = ReadGrid("Sheet1", "D1:F3");
            Expect(m != null && Eq(m, 1, 1, "x,y") && Eq(m, 1, 2, "x") && Eq(m, 1, 3, "y")
                    && Eq(m, 2, 1, "m,n") && Eq(m, 2, 2, "m") && Eq(m, 2, 3, "n")
                    && Eq(m, 3, 1, "p,q") && Eq(m, 3, 2, "p") && Eq(m, 3, 3, "q"),
                "C3 拆分结果从源列右侧相邻列起写（D 保留源值，E/F 为拆分片段）",
                m == null ? "读取失败" : DumpGrid(m, 3, 3));
        }

        // ===================== C4 文本清洗 find/replace =====================
        private static void C4TextCleanFindReplace(ToolRegistry registry)
        {
            WriteTable("Sheet1", "A1:A3", new string[,] { { "hello world" }, { "foo bar" }, { "x" } });
            B(registry, "text_clean", Range(
                ("address", "$A$1:$A$3"), ("sheet", "Sheet1"),
                ("trim", false), ("collapse_whitespace", false), ("remove_non_printable", false),
                ("case", "none"), ("find", "o"), ("replace", "0")),
                "C4 文本清洗 find/replace");

            var m = ReadGrid("Sheet1", "A1:A3");
            Expect(m != null && Eq(m, 1, 1, "hell0 w0rld") && Eq(m, 2, 1, "f00 bar") && Eq(m, 3, 1, "x"),
                "C4 查找到的 o 全部替换为 0",
                m == null ? "读取失败" : DumpGrid(m, 3, 1));
        }

        // ===================== C5 透视汇总 target_address=G1 =====================
        private static void C5PivotTargetAddress(ToolRegistry registry)
        {
            WriteTable("Sheet1", "A1:B5", new string[,] {
                { "类", "量" }, { "甲", "10" }, { "乙", "20" }, { "甲", "10" }, { "乙", "20" } });
            B(registry, "create_excel_pivot_summary", Range(
                ("source_address", "$A$1:$B$5"), ("sheet", "Sheet1"),
                ("label_column", "1"), ("value_column", "2"), ("target_address", "G1")),
                "C5 透视汇总 target_address=G1");

            var m = ReadGrid("Sheet1", "G1:H4");
            Expect(m != null && Eq(m, 1, 1, "Row Labels") && Eq(m, 1, 2, "Total")
                    && RowHas(m, 2, "甲", "20") && RowHas(m, 3, "乙", "40"),
                "C5 汇总结果写入 G1（甲=20，乙=40，带表头）",
                m == null ? "读取失败" : DumpGrid(m, 4, 2));
        }

        // ===================== C6 随机数填充 min=max=5 =====================
        private static void C6RandomNumbersRange(ToolRegistry registry)
        {
            Clear("Sheet1", "A1:A10");
            B(registry, "fill_random_numbers", Range(
                ("address", "$A$1:$A$10"), ("sheet", "Sheet1"),
                ("min", 5.0), ("max", 5.0), ("decimals", 0)),
                "C6 随机数填充 min=max=5");

            var m = ReadGrid("Sheet1", "A1:A10");
            var allFive = true;
            if (m != null)
            {
                for (var r = 1; r <= 10; r++)
                {
                    if (!Eq(m, r, 1, "5")) { allFive = false; break; }
                }
            }
            Expect(m != null && allFive, "C6 10 格全部 = 5（min=max 确定值）", m == null ? "读取失败" : DumpGrid(m, 10, 1));
        }

        // ===================== 工具链执行 =====================
        private static void B(ToolRegistry registry, string tool, Dictionary<string, object> args, string label)
        {
            try
            {
                if (!registry.TryGet(tool, out var d))
                {
                    Fail(label + "：工具未注册");
                    return;
                }
                ToolResult res;
                try
                {
                    res = d.Execute(args);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("      [RAW-THROW] " + label + " 执行抛异常：\n" + ex);
                    Fail(label + "：执行抛异常 " + ex.GetType().Name + " " + ex.Message);
                    return;
                }
                Expect(res.Ok, label, res.Ok ? (res.ErrorCode ?? "ok") : res.ErrorCode + " " + res.Data);
            }
            catch (Exception ex)
            {
                Console.WriteLine("      [OUTER-THROW] " + label + " 抛异常：\n" + ex);
                Fail(label + "：外层抛异常 " + ex.GetType().Name + " " + ex.Message);
            }
        }

        // ===================== Excel 助手（无状态：每次重新解析） =====================
        private static dynamic Wb()
        {
            return _app.Workbooks[1];
        }

        private static dynamic Ws(string name)
        {
            return Wb().Worksheets[name];
        }

        private static void EnsureSheet(dynamic wb, string name)
        {
            try { var _ = wb.Worksheets[name]; return; }
            catch
            {
                var last = wb.Worksheets.Count;
                var ws = wb.Worksheets.Add(Missing.Value, wb.Worksheets[last]);
                ws.Name = name;
            }
        }

        private static void WriteTable(string sheet, string range, string[,] cells)
        {
            dynamic ws = Ws(sheet);
            Clear(sheet, range);
            dynamic rng = ws.Range(range);
            dynamic target = rng.Resize[cells.GetLength(0), cells.GetLength(1)];
            var arr = new object[cells.GetLength(0), cells.GetLength(1)];
            for (var r = 0; r < cells.GetLength(0); r++)
            {
                for (var c = 0; c < cells.GetLength(1); c++)
                {
                    arr[r, c] = cells[r, c] == "" ? null : cells[r, c];
                }
            }
            target.Value2 = arr;
        }

        private static void Clear(string sheet, string range)
        {
            try { Ws(sheet).Range(range).ClearContents(); } catch { }
        }

        /// <summary>读回多格区域（Value2 返回 object[,]，1 基）。单格/整列返回标量/1 列数组，统一归整为 object[,]。</summary>
        private static object[,] ReadGrid(string sheet, string range)
        {
            try
            {
                dynamic ws = Ws(sheet);
                dynamic rng = ws.Range(range);
                var v = rng.Value2;
                if (v == null) { return null; }
                var rows = Convert.ToInt32(rng.Rows.Count);
                var cols = Convert.ToInt32(rng.Columns.Count);
                var result = new object[rows, cols];
                if (v is object[,] arr)
                {
                    for (var r = 0; r < rows; r++)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            result[r, c] = arr[r + 1, c + 1];
                        }
                    }
                }
                else
                {
                    result[0, 0] = v;
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine("      [READ] " + sheet + "!" + range + " 读取失败 " + ex.GetType().Name + " " + ex.Message);
                return null;
            }
        }

        private static bool Eq(object[,] m, int row, int col, string expected)
        {
            if (m == null || row < 1 || col < 1 || row > m.GetLength(0) || col > m.GetLength(1)) { return false; }
            var v = m[row - 1, col - 1];
            if (v == null || v == DBNull.Value) { return false; }
            return Convert.ToString(v, CultureInfo.InvariantCulture) == expected;
        }

        private static bool RowHas(object[,] m, int row, string label, string value)
        {
            return m != null && row >= 1 && row <= m.GetLength(0) && Eq(m, row, 1, label) && Eq(m, row, 2, value);
        }

        private static string DumpGrid(object[,] m, int rows, int cols)
        {
            var sb = new System.Text.StringBuilder();
            for (var r = 1; r <= rows; r++)
            {
                if (r > 1) { sb.Append(" | "); }
                for (var c = 1; c <= cols; c++)
                {
                    if (c > 1) { sb.Append(","); }
                    sb.Append(Convert.ToString(m[r - 1, c - 1] ?? "∅", CultureInfo.InvariantCulture));
                }
            }
            return sb.ToString();
        }

        private static Dictionary<string, object> Range(params (string k, object v)[] pairs)
        {
            var m = new Dictionary<string, object>();
            foreach (var (k, v) in pairs) { m[k] = v; }
            return m;
        }

        private static void Expect(bool c, string name, string detail)
        {
            if (c) { _passed++; Console.WriteLine("[PASS] " + name + " → " + detail); }
            else { _failed++; Console.WriteLine("[FAIL] " + name + " → " + detail); }
        }

        private static void Fail(string msg) { _failed++; Console.WriteLine("[FAIL] " + msg); }
    }
}
