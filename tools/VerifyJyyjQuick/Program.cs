using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyJyyjQuick
{
    /// <summary>
    /// 「JYYJ快捷」第一批（数据处理组 + 文本处理组）10 工具真机验证 harness。
    /// 独立 out-of-proc Excel + 未存盘工作簿，用 ExcelToolHost.CreateRegistry() 组装真实工具表，
    /// 用 RibbonCommandPlanner 组装的实参执行（与功能区按钮同一条执行链，仅绕过 WebView2 参数框 UI）。
    /// 每个工具：播种 → BuildFields/BuildArgs → Classify → Execute → 回读单元格/结果域断言 → BuildSummary 关键词。
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
            Console.WriteLine("== JYYJ快捷 第一批 10 工具真机验证 ==");
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

                VerifyClean(app, registry);
                VerifyDedup(app, registry);
                VerifySort(app, registry);
                VerifyTranspose(app, registry);
                VerifyFilter(app, registry);
                VerifyPivot(app, registry);
                VerifyTextSplit(app, registry);
                VerifyTextMerge(app, registry);
                VerifyTextClean(app, registry);
                VerifyFillFormula(app, registry);
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

                Console.WriteLine("[cleanup] 已结束自建 EXCEL 进程 " + killed.Count + " 个（PID：" + string.Join(",", killed) + "）");
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.WriteLine("\n== 结果：通过 " + _passed + "，失败 " + _failed + " ==");
            return _failed == 0 ? 0 : 1;
        }

        // ===================== 各工具验证 =====================

        private static void VerifyClean(dynamic app, ToolRegistry registry)
        {
            const string tool = "clean_excel_data";
            // 数据：第 2 行为可解析日期（触发 fix_dates），第 3 行全空（被 remove_blank_rows 删除），第 4 行日期列为空（被 fill_blanks 填「待补」）。
            Seed(app, "$A$1:$B$4", new object[,] { { "名称", "日期" }, { "甲", "2023/06/05" }, { "", "" }, { "丙", "" } });
            var sel = Sel("$A$1:$B$4", 4, 2, "名称", "日期");
            var values = new Dictionary<string, object>
            {
                ["remove_blank_rows"] = "true",
                ["fix_dates"] = "true",
                ["date_column"] = "2",
                ["fill_blanks"] = "true",
                ["fill_value"] = "待补"
            };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(IntField(result, "output_rows") == 3, tool + " 全空行被去除（output_rows=3）", "output_rows=" + IntField(result, "output_rows"));
            Expect(IntField(result, "removed_rows") == 1, tool + " 删除 1 个全空行", "removed_rows=" + IntField(result, "removed_rows"));
            Expect(CellText(app, "B2") != "2023/06/05", tool + " 日期列已格式修正（B2 不再为原文）", "B2=" + CellText(app, "B2"));
            Expect(EqCell(app, "B3") is string s && s == "待补", tool + " 空白单元格填「待补」", "B3=" + CellText(app, "B3"));
            Expect(SumContains(Summary(tool, result, sel), "清洗"), tool + " 摘要含「清洗」", Summary(tool, result, sel));
        }

        private static void VerifyDedup(dynamic app, ToolRegistry registry)
        {
            const string tool = "dedup_range";
            Seed(app, "$A$1:$B$4", new object[,] { { "姓名", "城市" }, { "甲", "北京" }, { "甲", "北京" }, { "乙", "上海" } });
            var sel = Sel("$A$1:$B$4", 4, 2, "姓名", "城市");
            var values = new Dictionary<string, object> { ["scope"] = "whole", ["keep"] = "first", ["has_header"] = "true" };

            var budget = RibbonCommandPlanner.CountRemovedRows(AppReadMatrix(app, "$A$1:$B$4"), 1, new int[0], false, false);
            Expect(budget == 1, tool + " 删行预算=1", "预算=" + budget);

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.DeleteRows);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(IntField(result, "deduped_rows_removed") == 1, tool + " 删除重复 1 行", "removed=" + IntField(result, "deduped_rows_removed"));
            Expect(EqCell(app, "A3") == "乙", tool + " 留存行上移到第 3 行", "A3=" + CellText(app, "A3"));
            Expect(SumContains(Summary(tool, result, sel), "删除重复 1"),
                tool + " 摘要含「删除重复 1」", Summary(tool, result, sel));
        }

        private static void VerifySort(dynamic app, ToolRegistry registry)
        {
            const string tool = "sort_range";
            Seed(app, "$A$1:$B$4", new object[,] { { "姓名", "分数" }, { "乙", 90D }, { "甲", 80D }, { "丙", 85D } });
            var sel = Sel("$A$1:$B$4", 4, 2, "姓名", "分数");
            var values = new Dictionary<string, object> { ["sort_column"] = "2", ["direction"] = "desc", ["has_header"] = "true" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.Overwrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqCell(app, "B2") == 90 && EqCell(app, "B3") == 85 && EqCell(app, "B4") == 80,
                tool + " 分数降序 90/85/80",
                $"B=[{CellText(app, "B2")},{CellText(app, "B3")},{CellText(app, "B4")}]");
            Expect(SumContains(Summary(tool, result, sel), "排序完成"), tool + " 摘要含「排序完成」", Summary(tool, result, sel));
        }

        private static void VerifyTranspose(dynamic app, ToolRegistry registry)
        {
            const string tool = "transpose_range";
            Seed(app, "$A$1:$B$3", new object[,] { { "a", "b" }, { 1D, 2D }, { 3D, 4D } });
            var sel = Sel("$A$1:$B$3", 3, 2, "h1", "h2");
            var values = new Dictionary<string, object> { ["target_mode"] = "top_left", ["target_address"] = "" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.Overwrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqCell(app, "A1") == "a" && EqCell(app, "A2") == "b", tool + " 转置第 1 行 [a,b]", $"A1={CellText(app, "A1")}, A2={CellText(app, "A2")}");
            Expect(EqCell(app, "C2") == 4, tool + " 转置后 C2=4", "C2=" + CellText(app, "C2"));
            Expect(SumContains(Summary(tool, result, sel), "转置"), tool + " 摘要含「转置」", Summary(tool, result, sel));
        }

        private static void VerifyFilter(dynamic app, ToolRegistry registry)
        {
            const string tool = "filter_excel_data";
            Seed(app, "$A$1:$B$4", new object[,] { { "姓名", "分数" }, { "乙", 90D }, { "甲", 80D }, { "丙", 85D } });
            var sel = Sel("$A$1:$B$4", 4, 2, "姓名", "分数");
            var values = new Dictionary<string, object> { ["column"] = "1", ["operator"] = "not_empty", ["value"] = "" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.Overwrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            // 需求文档 filter 无「是否含表头」选项：整区（含表头）按条件筛选，not_empty 下表头「姓名」亦非空 → matched=4。
            Expect(IntField(result, "matched_rows") == 4, tool + " 命中 4 行（含表头按需求契约参与筛选）", "matched=" + IntField(result, "matched_rows"));
            Expect(!string.IsNullOrWhiteSpace(StrField(result, "output_address")), tool + " 返回输出地址", StrField(result, "output_address") ?? "(null)");
            Expect(SumContains(Summary(tool, result, sel), "命中 4 行"), tool + " 摘要含「命中 4 行」", Summary(tool, result, sel));
        }

        private static void VerifyPivot(dynamic app, ToolRegistry registry)
        {
            const string tool = "create_excel_pivot_summary";
            Seed(app, "$A$1:$B$5", new object[,] { { "类", "金额" }, { "A", 10D }, { "B", 20D }, { "A", 5D }, { "B", 15D } });
            var sel = Sel("$A$1:$B$5", 5, 2, "类", "金额");
            var values = new Dictionary<string, object> { ["label_column"] = "1", ["value_column"] = "2" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(IntField(result, "row_count") == 2, tool + " 2 个分组", "row_count=" + IntField(result, "row_count"));
            var target = StrField(result, "target_address");
            Expect(!string.IsNullOrWhiteSpace(target), tool + " 返回输出地址", target ?? "(null)");
            var totals = ReadColumnValues(app, target);
            Expect(totals.Contains(15D) && totals.Contains(35D), tool + " 分组合计守恒 {15,35}", "Total=" + string.Join(",", totals));
            Expect(SumContains(Summary(tool, result, sel), "2 组"), tool + " 摘要含「2 组」", Summary(tool, result, sel));
        }

        private static void VerifyTextSplit(dynamic app, ToolRegistry registry)
        {
            const string tool = "text_split";
            Seed(app, "$A$1:$A$4", new object[,] { { "姓名" }, { "张三,李四" }, { "王五" }, { "赵六,钱七,孙八" } });
            var band = RibbonCommandPlanner.SplitTargetColumns(AppReadMatrix(app, "$A$1:$A$4"), 1, ",");
            Expect(band.Count == 3, tool + " 开放列带宽 3 列（写第 2..4 列）", "band=" + string.Join(",", band));
            var sel = Sel("$A$1:$A$4", 4, 1, "姓名");
            var values = new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = ",", ["keep_source"] = "true" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqCell(app, "A2") == "张三,李四", tool + " 保留源列", "A2=" + CellText(app, "A2"));
            Expect(EqCell(app, "D4") == "孙八", tool + " 最大 3 段写入第 4 列", "D4=" + CellText(app, "D4"));
            Expect(SumContains(Summary(tool, result, sel), "拆分完成"), tool + " 摘要含「拆分完成」", Summary(tool, result, sel));
        }

        private static void VerifyTextMerge(dynamic app, ToolRegistry registry)
        {
            const string tool = "text_merge";
            Seed(app, "$A$1:$C$4", new object[,] { { "姓", "名", "城市" }, { "张", "三", "北京" }, { "李", "四", "上海" }, { "王", "五", "广州" } });
            var sel = Sel("$A$1:$C$4", 4, 3, "姓", "名", "城市");
            var values = new Dictionary<string, object>
            {
                ["columns"] = new List<string> { "1", "2" },
                ["target_column"] = "4",
                ["delimiter"] = "",
                ["skip_empty"] = "true"
            };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqCell(app, "D2") == "张三" && EqCell(app, "D4") == "王五", tool + " 合并姓+名 写入右侧新列", "D2=" + CellText(app, "D2") + ", D4=" + CellText(app, "D4"));
            Expect(SumContains(Summary(tool, result, sel), "写入第 4 列"), tool + " 摘要含「写入第 4 列」", Summary(tool, result, sel));
        }

        private static void VerifyTextClean(dynamic app, ToolRegistry registry)
        {
            const string tool = "text_clean";
            Seed(app, "$A$1:$A$3", new object[,] { { "城市" }, { " 北京 " }, { "上海  浦东" } });
            var sel = Sel("$A$1:$A$3", 3, 1, "城市");
            var values = new Dictionary<string, object> { ["trim"] = "true", ["collapse"] = "true", ["remove_non_printable"] = "false", ["case"] = "none", ["remove_chars"] = "" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(EqCell(app, "A2") == "北京", tool + " 去首尾空格", "A2=[" + CellText(app, "A2") + "]");
            Expect(EqCell(app, "A3") == "上海 浦东", tool + " 合并连续空白", "A3=[" + CellText(app, "A3") + "]");
            Expect(SumContains(Summary(tool, result, sel), "文本清洗完成"), tool + " 摘要含「文本清洗完成」", Summary(tool, result, sel));
        }

        private static void VerifyFillFormula(dynamic app, ToolRegistry registry)
        {
            const string tool = "fill_formula";
            Seed(app, "$A$1:$A$3", new object[,] { { 10D }, { 20D }, { 30D } });
            var sel = Sel("$B$1:$B$3", 3, 1, "结果");
            var values = new Dictionary<string, object> { ["formula"] = "=A{row}+1", ["start_row"] = "" };

            var result = Run(app, registry, tool, sel, values);
            ExpectClassify(tool, CommandSafety.SafeWrite);
            Expect(result.Ok, tool + " 执行成功", Describe(result));
            Expect(IntField(result, "filled_cells") == 3, tool + " 填充 3 格", "filled=" + IntField(result, "filled_cells"));
            Expect(EqCell(app, "B1") == 11 && EqCell(app, "B3") == 31, tool + " 公式按行号替换并计算", "B1=" + CellText(app, "B1") + ", B3=" + CellText(app, "B3"));
            Expect(SumContains(Summary(tool, result, sel), "共 3 格"), tool + " 摘要含「共 3 格」", Summary(tool, result, sel));
        }

        // ===================== 执行与断言助手 =====================

        private static ToolResult Run(dynamic app, ToolRegistry registry, string tool, ParamSelection sel, Dictionary<string, object> values)
        {
            var fields = RibbonCommandPlanner.BuildFields(tool, sel);
            if (fields.Count == 0)
            {
                Fail(tool + "：BuildFields 未生成字段");
            }

            var build = RibbonCommandPlanner.BuildArgs(tool, sel, values);
            if (build is string err)
            {
                return new ToolResult { Ok = false, ErrorCode = "PLANNER_REJECT", Data = err };
            }

            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }

            try
            {
                return descriptor.Execute((Dictionary<string, object>)build);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[EX-STACK][" + tool + "] " + ex + "\n");
                return new ToolResult { Ok = false, ErrorCode = "HARNESS_EXCEPTION", Data = ex.Message };
            }
        }

        private static ParamSelection Sel(string address, int rows, int cols, params string[] headers)
        {
            var sel = new ParamSelection { SheetName = _sheet, Address = address, Rows = rows, Cols = cols };
            for (var c = 0; c < cols; c++)
            {
                sel.ColumnNames.Add(c < headers.Length ? headers[c] : "列" + (c + 1));
            }

            return sel;
        }

        private static void Seed(dynamic app, string address, object[,] matrix)
        {
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = app.ActiveWorkbook.Sheets.Item[_sheet];
                range = sheet.Range[address];
                range.Value2 = matrix;
            }
            finally
            {
                Release(range);
                Release(sheet);
            }
        }

        private static object[,] AppReadMatrix(dynamic app, string address)
        {
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = app.ActiveWorkbook.Sheets.Item[_sheet];
                range = sheet.Range[address];
                var v = range.Value2;
                return ExcelToolHostHelpersNormalize(v, Convert.ToInt32(range.Rows.Count), Convert.ToInt32(range.Columns.Count));
            }
            finally
            {
                Release(range);
                Release(sheet);
            }
        }

        private static object[,] ExcelToolHostHelpersNormalize(object v, int rows, int cols)
        {
            var result = new object[rows, cols];
            if (v is object[,] arr)
            {
                var lb0 = arr.GetLowerBound(0);
                var lb1 = arr.GetLowerBound(1);
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var val = (r < arr.GetLength(0) && c < arr.GetLength(1)) ? arr[r + lb0, c + lb1] : DBNull.Value;
                        result[r, c] = val is DBNull ? null : val;
                    }
                }
            }
            else
            {
                result[0, 0] = v;
            }

            return result;
        }

        private static object ReadCell(dynamic app, string address)
        {
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = app.ActiveWorkbook.Sheets.Item[_sheet];
                range = sheet.Range[address];
                return range.Value2;
            }
            finally
            {
                Release(range);
                Release(sheet);
            }
        }

        private static object EqCell(dynamic app, string address)
        {
            return ReadCell(app, address);
        }

        private static string CellText(dynamic app, string address)
        {
            var v = ReadCell(app, address);
            if (v == null || v is DBNull)
            {
                return "〈空〉";
            }

            return Convert.ToString(v);
        }

        /// <summary>读取目标区域完整写入块的 Total 列数值。address 形如 $F$6:$G$9（顶部为表头 Row Labels|Total）。</summary>
        private static List<double> ReadColumnValues(dynamic app, string address)
        {
            var result = new List<double>();
            var block = AppReadMatrix(app, address);
            var cols = block.GetLength(1);
            var lastCol = cols - 1;
            for (var r = 1; r < block.GetLength(0); r++)
            {
                if (block[r, lastCol] is double dv)
                {
                    result.Add(dv);
                }
            }

            return result;
        }

        private static string Summary(string tool, ToolResult result, ParamSelection sel)
        {
            try
            {
                return RibbonCommandPlanner.BuildSummary(tool, result, sel);
            }
            catch (Exception ex)
            {
                return "BuildSummary 异常：" + ex.Message;
            }
        }

        private static bool SumContains(string summary, string keyword)
        {
            return summary != null && summary.Contains(keyword);
        }

        private static int? IntField(ToolResult r, string key)
        {
            var jo = AsJObject(r?.Data);
            if (jo == null || jo[key] == null)
            {
                return null;
            }

            return long.TryParse(Convert.ToString(jo[key]), out var l) ? (int?)l : null;
        }

        private static string StrField(ToolResult r, string key)
        {
            var jo = AsJObject(r?.Data);
            if (jo == null || jo[key] == null)
            {
                return null;
            }

            return Convert.ToString(jo[key]);
        }

        private static JObject AsJObject(object data)
        {
            if (data == null)
            {
                return null;
            }

            if (data is JObject jo)
            {
                return jo;
            }

            try
            {
                return JToken.FromObject(data) as JObject;
            }
            catch
            {
                return null;
            }
        }

        private static void ExpectClassify(string tool, CommandSafety expected)
        {
            CommandSafety actual;
            try
            {
                actual = RibbonCommandPlanner.Classify(tool);
            }
            catch (Exception ex)
            {
                Fail(tool + "：Classify 异常 " + ex.Message);
                return;
            }

            Expect(actual == expected, tool + " 安全分级 = " + expected, "actual=" + actual);
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

        private static void Fail(string message)
        {
            _failed++;
            Console.WriteLine("[FAIL] " + message);
        }

        private static string Describe(ToolResult result)
        {
            if (result == null)
            {
                return "null";
            }

            return result.Ok ? "ok" : (result.ErrorCode ?? "failed") + " " + result.Data;
        }

        private static void Release(object com)
        {
            if (com != null && System.Runtime.InteropServices.Marshal.IsComObject(com))
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); }
                catch { }
            }
        }
    }
}