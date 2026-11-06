using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    /// <summary>
    /// dl 模式：v3.7.1.41 create_excel_chart 自动数据标签的真机闭环验证（可读性五件套②）。
    /// DL1 小数据柱形图缺省 → 自动开（HasDataLabels=True 且返回 data_labels=on）；
    /// DL2 data_labels=off 强制关闭；
    /// DL3 饼图缺省 → 百分比标签（ShowPercentage=True）；
    /// DL4 20 点折线缺省 → 超限不开（auto 保守）；
    /// DL5 散点图缺省 → 不开。
    /// 清理纪律沿用 Md1Probe：按 PID 强杀本 harness 实例。
    /// </summary>
    internal static class DataLabelProbe
    {
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();
        private static int pass, fail;

        public static int Run()
        {
            dynamic excel = null;
            try
            {
                var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); }
                    p.Dispose();
                }
                excel.Workbooks.Add();

                var host = new ExcelToolHost(excel);
                var registry = host.CreateRegistry();
                if (!registry.TryGet("create_excel_chart", out var chartTool))
                {
                    Console.WriteLine("[fatal] create_excel_chart 未注册");
                    return 2;
                }

                // 准备两张数据表：A1:B5 小数据（4 类目）；D1:E21 大数据（20 点）
                FillBlock(excel, "A1", new object[,] { { "类别", "销售额" }, { "华东", 320 }, { "华北", 280 }, { "华南", 410 }, { "西南", 190 } });
                var big = new object[21, 2];
                big[0, 0] = "日期"; big[0, 1] = "销量";
                for (var i = 1; i <= 20; i++) { big[i, 0] = "D" + i; big[i, 1] = 100 + i; }
                FillBlock(excel, "D1", big);

                Exec(registry, "delete_excel_worksheet", new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { ["name"] = "Sheet2" }, quiet: true);

                // DL1：小数据柱形图缺省 → 自动开
                var d1 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "A1:B5", ["chart_type"] = "ColumnClustered", ["anchor"] = "G1", ["title"] = "DL1 柱形（万元）· 不含退货"
                });
                var d1On = d1.Ok && ReadHasDataLabels(excel, 1) == true
                    && ReadDataLabelsField(d1) == "on";
                Report("DL1 小数据柱形缺省自动开标签", d1On, Describe(d1) + " HasDataLabels=" + ReadHasDataLabels(excel, 1));

                // DL2：data_labels=off 强制关闭
                var d2 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "A1:B5", ["chart_type"] = "BarClustered", ["anchor"] = "G16", ["data_labels"] = "off"
                });
                var d2Ok = d2.Ok && ReadHasDataLabels(excel, 2) == false;
                Report("DL2 data_labels=off 强制关闭", d2Ok, Describe(d2) + " HasDataLabels=" + ReadHasDataLabels(excel, 2));

                // DL3：饼图缺省 → 百分比标签
                var d3 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "A1:B5", ["chart_type"] = "Pie", ["anchor"] = "G31"
                });
                var d3On = d3.Ok && ReadHasDataLabels(excel, 3) == true && ReadShowPercentage(excel, 3) == true;
                Report("DL3 饼图缺省百分比标签", d3On, Describe(d3) + " HasDataLabels=" + ReadHasDataLabels(excel, 3) + " ShowPercentage=" + ReadShowPercentage(excel, 3));

                // DL4：20 点折线缺省 → 超限不开
                var d4 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "D1:E21", ["chart_type"] = "Line", ["anchor"] = "G46"
                });
                var d4Ok = d4.Ok && ReadHasDataLabels(excel, 4) == false;
                Report("DL4 20 点折线超限不开", d4Ok, Describe(d4) + " HasDataLabels=" + ReadHasDataLabels(excel, 4));

                // DL5：散点图缺省 → 不开
                var d5 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "A1:B5", ["chart_type"] = "Scatter", ["anchor"] = "G61"
                });
                var d5Ok = d5.Ok && ReadHasDataLabels(excel, 5) == false;
                Report("DL5 散点缺省不开", d5Ok, Describe(d5) + " HasDataLabels=" + ReadHasDataLabels(excel, 5));

                Console.WriteLine($"== dl 结果：{pass} 过 / {fail} 败 ==");
                return fail == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal] " + ex.Message + "\n" + ex.StackTrace);
                return 2;
            }
            finally
            {
                foreach (var pid in SpawnedPids)
                {
                    try { Process.GetProcessById(pid).Kill(); } catch { }
                }
            }
        }

        private static void FillBlock(dynamic excel, string anchor, object[,] values)
        {
            dynamic ws = excel.ActiveSheet;
            int rows = values.GetLength(0), cols = values.GetLength(1);
            dynamic start = ws.Range[anchor];
            dynamic block = ws.Range[start.Cells.Item[1, 1], start.Cells.Item[rows, cols]];
            block.Value2 = values;
        }

        /// <summary>按创建顺序取第 n 个图表（每个用例新建一个，顺序稳定）。</summary>
        private static dynamic ChartN(dynamic excel, int n)
        {
            dynamic ws = excel.ActiveSheet;
            dynamic charts = ws.ChartObjects();
            dynamic co = charts.Item(n);
            return co.Chart;
        }

        private static bool? ReadHasDataLabels(dynamic excel, int n)
        {
            dynamic chart = null;
            dynamic sc = null;
            dynamic s = null;
            try
            {
                chart = ChartN(excel, n);
                sc = chart.SeriesCollection();
                s = sc.Item(1);
                return Convert.ToBoolean(s.HasDataLabels);
            }
            catch { return null; }
            finally
            {
                Release(s); Release(sc); Release(chart);
            }
        }

        private static bool? ReadShowPercentage(dynamic excel, int n)
        {
            dynamic chart = null;
            dynamic sc = null;
            dynamic s = null;
            dynamic labels = null;
            try
            {
                chart = ChartN(excel, n);
                sc = chart.SeriesCollection();
                s = sc.Item(1);
                labels = s.DataLabels;
                return Convert.ToBoolean(labels.ShowPercentage);
            }
            catch { return null; }
            finally
            {
                Release(labels); Release(s); Release(sc); Release(chart);
            }
        }

        private static void Exec(ToolRegistry registry, string tool, Dictionary<string, object> args, bool quiet = false)
        {
            if (!registry.TryGet(tool, out var t)) { return; }
            var r = t.Execute(args);
            if (!quiet && !r.Ok) { Console.WriteLine($"[" + tool + "] " + r.ErrorCode + " " + Convert.ToString(r.Data)); }
        }

        private static string AsData(ToolResult r) => JsonConvertShim(r.Data);

        /// <summary>结构化读返回值里的 data_labels 字段（探针断言必须按字段解析，不能做带空格的文本匹配——踩坑手册断言口径家族）。</summary>
        private static string ReadDataLabelsField(ToolResult r)
        {
            try
            {
                var jo = Newtonsoft.Json.Linq.JObject.FromObject(r.Data);
                return Convert.ToString(jo["data_labels"]);
            }
            catch
            {
                return null;
            }
        }

        private static string JsonConvertShim(object data)
        {
            try { return Newtonsoft.Json.JsonConvert.SerializeObject(data) ?? ""; }
            catch { return Convert.ToString(data) ?? ""; }
        }

        private static string Describe(ToolResult r)
        {
            if (r.Ok)
            {
                var s = JsonConvertShim(r.Data);
                return s.Length <= 140 ? s : s.Substring(0, 140) + "…";
            }

            return r.ErrorCode + " " + Short(Convert.ToString(r.Data));
        }

        private static string Short(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            var flat = text.Replace("\r", " ").Replace("\n", " ");
            return flat.Length <= 200 ? flat : flat.Substring(0, 200) + "…";
        }

        private static void Report(string name, bool ok, string detail)
        {
            if (ok) { pass++; } else { fail++; }
            Console.WriteLine("[" + (ok ? "PASS" : "FAIL") + "] " + name + " | " + detail);
        }

        private static void Release(object value)
        {
            if (value == null) { return; }
            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
        }
    }
}
