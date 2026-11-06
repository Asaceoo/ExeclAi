using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyAdvancedTools
{
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            dynamic app = null;
            dynamic wb = null;
            string result = string.Empty;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                wb = app.Workbooks.Add();

                var sheet1 = wb.Worksheets.Item[1];
                sheet1.Name = "Sheet1";
                object[,] data = new object[,]
                {
                    { "City", "Sales", "Orders" },
                    { "Beijing", 100, "A" },
                    { "Shanghai", 200, "B" },
                    { "Shenzhen", 300, "C" },
                    { "Bei", 400, "D" }
                };
                var r = sheet1.Range["A1:C5"];
                r.Value2 = data;
                Release(r);

                var dims = wb.Worksheets.Add();
                dims.Name = "Dims";
                var dimRange = dims.Range["A1"];
                dimRange.Value2 = 1;
                Release(dimRange);

                var scratch = wb.Worksheets.Add();
                scratch.Name = "_AI_CALC";
                Release(scratch);

                var hidden = wb.Worksheets.Add();
                hidden.Name = "HiddenSheet";
                hidden.Visible = 2; // xlSheetHidden
                Release(hidden);

                var host = new ExcelToolHost((object)app);
                var registry = host.CreateRegistry();

                var sheet1Ref = wb.Worksheets.Item["Sheet1"];
                sheet1Ref.Activate();

                app.Range("A1:B2").Select();

                // 前置隔离测试：验证 copy_paste 是否受前序工具影响
                Check(registry, "copy_paste_excel_range", A("source_sheet", "Sheet1", "source_address", "A1:B2", "target_sheet", "Sheet1", "target_address", "G1"),
                    e => { var t = Str(e, "target_address"); return new R(e.Ok && t != null && t.Length > 0, "target=" + t); });

                Check(registry, "get_excel_sheet_list", new Dictionary<string, object>(),
                    e => Has(e, "sheet_count", val => Val(e, "sheet_count") >= 1, "sheet_count=" + Val(e, "sheet_count")));

                Check(registry, "preview_excel_sheet", A("sheet", "Sheet1"),
                    e => Has(e, "total_rows", 5, "total_rows=" + Val(e, "total_rows")));

                Check(registry, "get_excel_range_values_and_formulas", A("address", "A1:C3"),
                    e => Has(e, "cell_count", 9, "cell_count=" + Val(e, "cell_count")));

                Check(registry, "get_current_selection_values_and_formulas", new Dictionary<string, object>(),
                    e => Has(e, "cell_count", 4, "selection_cells=" + Val(e, "cell_count")));

                Check(registry, "select_excel_range", A("address", "C3"),
                    e => { var s = Str(e, "selected"); return new R(e.Ok && s == "$C$3", "selected=" + s); });

                // 拷贝(A1:B2→G1:H2)成功后会新增一个含 "Bei" 的单元格(G2="Beijing")，故期望为 3。
                Check(registry, "find_excel_cells", A("what", "Bei"),
                    e => Has(e, "count", 3, "matches=" + Val(e, "count")));

                Check(registry, "manage_excel_filter", A("action", "set", "address", "A1:C5", "column", 1, "operator", "contains", "value", "Bei"),
                    e => Has(e, "matched_rows", 2, "matched_rows=" + Val(e, "matched_rows")));
                Check(registry, "manage_excel_filter", A("action", "clear"),
                    e => new R(e.Ok, "cleared"));

                Check(registry, "manage_excel_named_range", A("action", "create", "name", "TestRange", "sheet", "Sheet1", "reference", "A1:B2"),
                    e => new R(e.Ok && Bool(e, "created"), "created"));
                Check(registry, "manage_excel_named_range", A("action", "query", "name", "TestRange"),
                    e => new R(e.Ok && Bool(e, "exists"), "exists=" + Bool(e, "exists")));
                Check(registry, "manage_excel_named_range", A("action", "delete", "name", "TestRange"),
                    e => new R(e.Ok && Bool(e, "deleted"), "deleted"));

                Check(registry, "write_excel_cells_batch", A("cells", new Dictionary<string, object>[]
                {
                    new Dictionary<string, object> { ["address"] = "E1", ["value"] = "x" },
                    new Dictionary<string, object> { ["address"] = "E2", ["value"] = "y" }
                }), e => Has(e, "written_cells", 2, "written=" + Val(e, "written_cells")));

                Check(registry, "append_calculation_scratchpad", A("input", "2", "formula", "2+3", "expected_output", "5"),
                    e => Has(e, "computed", 5, "computed=" + Str(e, "computed")));
                Check(registry, "clear_calculation_scratchpad", new Dictionary<string, object>(),
                    e => new R(e.Ok && Bool(e, "cleared"), "cleared"));

                Check(registry, "insert_excel_dimension", A("sheet", "Dims", "address", "B:B", "quantity", 1),
                    e => new R(e.Ok, "inserted"));
                Check(registry, "delete_excel_dimension", A("sheet", "Dims", "address", "B:B", "quantity", 1),
                    e => new R(e.Ok, "deleted"));

                Check(registry, "create_toc_sheet", A("sheet", "TOC"),
                    e => Has(e, "sheet_count", 3, "sheet_count=" + Val(e, "sheet_count")));

                // ==================== 第二期工具验证 ====================
                // 注意：工具内会对 ActiveWorkbook 调 FinalReleaseComObject，使本程序缓存的
                // wb/sheet1 RCW 失效；因此第二期所有直接 COM 操作都即时从 app 重新获取并及时释放。

                // 准备公式单元格（Sheet1 F1/F2）
                dynamic lw = null;
                dynamic ls = null;
                try
                {
                    lw = app.ActiveWorkbook;
                    ls = lw.Worksheets.Item["Sheet1"];
                    dynamic f1 = ls.Range["F1"];
                    f1.Formula = "=A2+B2";
                    Release(f1);
                    dynamic f2 = ls.Range["F2"];
                    f2.Formula = "=SUM(A2:A4)";
                    Release(f2);
                }
                finally
                {
                    Release((object)ls);
                    Release((object)lw);
                }

                // 可见复制源表 VisibleSrc（隐藏第 3 行）
                string srcSheetName = "VisibleSrc";
                var srcSheet = app.ActiveWorkbook.Worksheets.Add();
                try
                {
                    srcSheet.Name = srcSheetName;
                    var srcRange = srcSheet.Range["A1:B4"];
                    srcRange.Value2 = new object[,] { { "A", 1 }, { "B", 2 }, { "C", 3 }, { "D", 4 } };
                    Release(srcRange);
                    srcSheet.Rows.Item[3].Hidden = true;
                }
                finally
                {
                    Release((object)srcSheet);
                }

                var splitDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XAI_Split_" + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(splitDir);

                var mergePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XAI_Merge_" + Guid.NewGuid().ToString("N") + ".xlsx");
                dynamic mateWb = app.Workbooks.Add();
                try
                {
                    dynamic ms = mateWb.Worksheets.Item[1];
                    ms.Name = "Mate";
                    var mr = ms.Range["A1"];
                    mr.Value2 = "m";
                    Release(mr);
                    Release(ms);
                    mateWb.SaveAs(mergePath);
                    mateWb.Close(false);
                }
                finally
                {
                    Release((object)mateWb);
                }

                // 1) Markdown 双向转换
                Check(registry, "export_selection_to_markdown", A("sheet", "Sheet1", "address", "A1:B3", "include_headers", true),
                    e => { var md = Str(e, "markdown"); return new R(e.Ok && md != null && md.Contains("Beijing"), "has_data"); });
                Check(registry, "import_markdown_to_excel", A("sheet", "Sheet1", "address", "I1", "value_mode", "value", "markdown", "| P1 | P2 |\n| --- | --- |\n| x | 1 |"),
                    e => new R(e.Ok, "imported"));

                // 2) 单元格编辑增强
                Check(registry, "merge_or_split_cells", A("action", "merge", "sheet", "Sheet1", "address", "A10:B10"),
                    e => new R(e.Ok, "merged"));
                Check(registry, "merge_or_split_cells", A("action", "unmerge", "sheet", "Sheet1", "address", "A10:B10"),
                    e => new R(e.Ok, "unmerged"));
                Check(registry, "merge_or_split_cells", A("action", "merge_similar", "sheet", "VisibleSrc", "address", "A1:B2", "direction", "column"),
                    e => new R(e.Ok, "merge_similar"));
                Check(registry, "reset_cell_value_type", A("sheet", "Sheet1", "address", "E1", "target_type", "text"),
                    e => new R(e.Ok, "reset_type"));
                Check(registry, "transform_excel_formulas", A("sheet", "Sheet1", "address", "F1:F1", "actions", new Dictionary<string, object>[] { new Dictionary<string, object> { ["type"] = "iferror", ["fallback"] = "0" } }),
                    e => new R(e.Ok, "transform"));
                Check(registry, "convert_formula_references", A("sheet", "Sheet1", "address", "F1:F2", "mode", "to_absolute"),
                    e => Has(e, "changed_cells", 2, "changed=" + Val(e, "changed_cells")));

                // 3) 公式追踪 + 可见复制
                Check(registry, "trace_formula_references", A("sheet", "Sheet1", "address", "F1"),
                    e => { var f3 = Bool(e, "has_formula"); return new R(e.Ok && f3, "has_formula=" + f3); });
                Check(registry, "drill_excel_range", A("sheet", "Sheet1", "address", "F1"),
                    e => new R(e.Ok, "drilled"));
                Check(registry, "copy_visible_rows_to_sheet", A("source_sheet", "VisibleSrc", "source_address", "A1:B4", "target_sheet", "Sheet1", "target_address", "A12"),
                    e => new R(e.Ok, "copied_visible"));

                // 4) 打印/导出/簿级
                Check(registry, "apply_page_setup", A("sheet", "Sheet1", "orientation", "portrait", "paper_size", "a4", "center_horizontally", true),
                    e => new R(e.Ok, "pagesetup"));
                var pdfPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XAI_Export_" + Guid.NewGuid().ToString("N") + ".pdf");
                Check(registry, "export_sheets_to_pdf", A("file_path", pdfPath, "sheets", new object[] { "Sheet1" }),
                    e => new R(e.Ok || (e.ErrorCode != null && e.ErrorCode.Contains("PDF")), "ok_or_driver_err=" + e.ErrorCode));
                Check(registry, "merge_workbooks", A("source_workbooks", new object[] { mergePath }),
                    e => new R(e.Ok, "merged_sheets"));
                Check(registry, "split_workbook", A("output_folder", splitDir, "sheets", new object[] { "Sheet1" }),
                    e => new R(e.Ok, "split"));
                Check(registry, "list_external_links", new Dictionary<string, object>(),
                    e => new R(e.Ok, "link_count=" + Val(e, "link_count")));
                Check(registry, "replace_external_links", A("old_link", "C:\\fake.xlsx", "new_link", "C:\\fake2.xlsx"),
                    e => new R(e.Ok || e.ErrorCode != null, "ok_or_err=" + e.ErrorCode));

                // 5) 受控执行：默认关闭 → 门禁生效
                // 注意：ExcelToolHost 构造时会 Apply(ControlledExecutionStore.Load())，若持久化配置残留 enabled:true
                // 会覆盖 new ScriptOptions() 的默认 false，导致本段「默认关闭」断言随机器状态漂移。显式强制关断以隔离环境态。
                host.ScriptOptions.Enabled = false;
                Check(registry, "validate_vba_snippet", A("snippet", "Public Sub Foo()\nEnd Sub"),
                    e => new R(!e.Ok && e.ErrorCode == "DISABLED", "gated=" + e.ErrorCode));
                Check(registry, "run_vba_snippet", A("snippet", "Sub A()\nEnd Sub"),
                    e => new R(!e.Ok && e.ErrorCode == "DISABLED", "gated=" + e.ErrorCode));
                Check(registry, "run_python_code", A("code", "print(1)"),
                    e => new R(!e.Ok && e.ErrorCode == "DISABLED", "gated=" + e.ErrorCode));
                Check(registry, "run_cli_command", A("command", "cmd"),
                    e => new R(!e.Ok && e.ErrorCode == "DISABLED", "gated=" + e.ErrorCode));
                Check(registry, "run_exe_command", A("path", "C:\\Windows\\System32\\cmd.exe"),
                    e => new R(!e.Ok && e.ErrorCode == "DISABLED", "gated=" + e.ErrorCode));

                // 受控执行开启后验证真实可用性（独立 host，仅开启 CLI 白名单）
                var execHost = new ExcelToolHost((object)app);
                execHost.ScriptOptions.Enabled = true;
                execHost.ScriptOptions.CliAllowlist = new[] { "cmd" };
                var execRegistry = execHost.CreateRegistry();
                Check(execRegistry, "run_cli_command", A("command", "cmd", "args", new object[] { "/c", "echo", "hello-xai" }),
                    e => new R(e.Ok, "ran_cli"));

                // ==================== 第三期工具验证 ====================

                // A. 导入外部文件与批量处理
                var importDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XAI_Import_" + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(importDir);
                var csv1 = System.IO.Path.Combine(importDir, "a.csv");
                var csv2 = System.IO.Path.Combine(importDir, "b.csv");
                System.IO.File.WriteAllText(csv1, "Name,Value\nAlpha,10\nBeta,20");
                System.IO.File.WriteAllText(csv2, "Name,Value\nGamma,30");

                Check(registry, "preview_external_file", A("file_path", csv1, "sample_rows", 10),
                    e => Has(e, "column_count", v => (int)v >= 1, "column_count=" + Val(e, "column_count")));
                Check(registry, "import_delimited_file", A("file_path", csv1, "target_sheet", "Sheet1", "target_address", "K1"),
                    e => Has(e, "imported_rows", 3, "imported_rows=" + Val(e, "imported_rows")));
                Check(registry, "import_excel_file_sheet", A("file_path", mergePath),
                    e => { var m = Str(e, "mode"); return new R(e.Ok && m == "new_sheet", "mode=" + m); });
                Check(registry, "merge_delimited_files", A("file_paths", new object[] { csv1, csv2 }, "output_sheet", "Sheet1", "output_address", "K10"),
                    e => Has(e, "combined_rows", 4, "combined_rows=" + Val(e, "combined_rows")));

                // B. 数据分析与可视化
                Check(registry, "analyze_range_statistics", A("sheet", "Sheet1", "address", "A1:C5"),
                    e => Has(e, "columns", 3, "columns=" + Val(e, "columns")));
                Check(registry, "analyze_column_correlation", A("sheet", "Sheet1", "address", "A1:C5", "method", "pearson"),
                    e => Has(e, "numeric_column_count", v => (int)v >= 1, "num_cols=" + Val(e, "numeric_column_count")));
                Check(registry, "analyze_value_distribution", A("sheet", "Sheet1", "address", "A1:C5", "column", 1),
                    e => Has(e, "distinct_count", v => (int)v >= 1, "distinct=" + Val(e, "distinct_count")));
                Check(registry, "recommend_excel_chart", A("sheet", "Sheet1", "address", "A1:C5"),
                    e => { var t = Str(e, "recommended_chart_type"); return new R(e.Ok && t != null && t != "null", "type=" + t); });

                // 准备一张图表用于导出图片（直接 COM 即时获取并释放）
                var chartExportDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "XAI_Chart_" + Guid.NewGuid().ToString("N"));
                System.IO.Directory.CreateDirectory(chartExportDir);
                var chartPng = System.IO.Path.Combine(chartExportDir, "chart.png");
                dynamic chartSheetRef = null;
                dynamic chartWbRef = null;
                try
                {
                    chartWbRef = app.ActiveWorkbook;
                    chartSheetRef = chartWbRef.Worksheets.Item["Sheet1"];
                    dynamic chartObj = chartSheetRef.ChartObjects().Add(100, 60, 420, 260);
                    chartObj.Name = "XaiChart";
                    chartObj.Chart.SetSourceData(chartSheetRef.Range("A1:C5"));
                    Release(chartObj);
                }
                finally
                {
                    Release((object)chartSheetRef);
                    Release((object)chartWbRef);
                }

                Check(registry, "create_excel_sparkline", A("sheet", "Sheet1", "data_range", "B2:E2", "dest_range", "E10", "type", "line"),
                    e => new R(e.Ok && Bool(e, "created"), "created=" + Bool(e, "created")));
                Check(registry, "export_excel_chart_image", A("sheet", "Sheet1", "chart_name", "XaiChart", "file_path", chartPng),
                    e => new R(e.Ok && Bool(e, "exported") && System.IO.File.Exists(chartPng), "exported=" + Bool(e, "exported") + (System.IO.File.Exists(chartPng) ? "&&file_exists" : "&&no_file")));

                Console.WriteLine();
                Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
                if (Failures.Count > 0)
                {
                    foreach (var f in Failures)
                    {
                        Console.WriteLine("FAILED: " + f);
                    }

                    result = "FAIL";
                }
                else
                {
                    Console.WriteLine("ALL TOOLS VERIFIED OK");
                    result = "PASS";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("HARNESS ERROR: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
                result = "ERROR";
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Close(false); // 不保存
                        Release(wb);
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null)
                    {
                        app.Quit();
                        Release(app);
                    }
                }
                catch
                {
                }
            }

            return result == "PASS" ? 0 : 1;
        }

        private static R Has(ToolResult e, string key, int expected, string detail)
        {
            return new R(e.Ok && Val(e, key) == expected, detail);
        }

        private static R Has(ToolResult e, string key, Func<object, bool> predicate, string detail)
        {
            return new R(e.Ok && predicate(Val(e, key)), detail);
        }

        private static int Val(ToolResult e, string key)
        {
            if (e.Data == null || e.Data is string)
            {
                return 0;
            }

            var token = (JToken)JObject.FromObject(e.Data)[key];
            return token == null || token.Type == JTokenType.Null ? 0 : token.ToObject<int>();
        }

        private static string Str(ToolResult e, string key)
        {
            if (e.Data == null || e.Data is string)
            {
                return null;
            }

            return (string)JObject.FromObject(e.Data)[key];
        }

        private static bool Bool(ToolResult e, string key)
        {
            if (e.Data == null || e.Data is string)
            {
                return false;
            }

            var token = (JToken)JObject.FromObject(e.Data)[key];
            return token != null && token.Type != JTokenType.Null && (bool)token;
        }

        private static Dictionary<string, object> A(params object[] kvs)
        {
            var map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < kvs.Length; i += 2)
            {
                map[Convert.ToString(kvs[i])] = kvs[i + 1];
            }

            return map;
        }

        private static void Check(ToolRegistry registry, string name, Dictionary<string, object> args, Func<ToolResult, R> assert)
        {
            try
            {
                if (!registry.TryGet(name, out var tool))
                {
                    Fail(name, "NOT REGISTERED");
                    return;
                }

                var res = tool.Execute(args);
                var r = assert(res);
                if (r.Ok)
                {
                    _passed++;
                    Console.WriteLine("PASS  " + name + "  [" + r.Detail + "]");
                }
                else
                {
                    Fail(name, r.Detail + " (error=" + res.ErrorCode + ", data=" + (res.Data == null ? "null" : res.Data.ToString()) + ")");
                }
            }
            catch (Exception ex)
            {
                Fail(name, "EXCEPTION: " + ex.Message);
            }
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

        private struct R
        {
            public bool Ok;
            public string Detail;

            public R(bool ok, string detail)
            {
                Ok = ok;
                Detail = detail;
            }
        }
    }
}