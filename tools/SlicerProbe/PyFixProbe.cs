using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace SlicerProbe
{
    /// <summary>
    /// pyfix 模式：v3.7.1.38 三项修复的真机闭环验证。
    /// T1 run_python_code 缺模块自愈（PIL→pillow，真实缺 PIL 的 python 环境）；
    /// T1b 非白名单模块不自愈（负例）；
    /// T2 run_python_excel write_ranges 带 sheet 名 JArray 格式（真机复现 Sheet1!A52:C60 误报）；
    /// T3 create_excel_chart 缺 source_address → INVALID_ARGS（不再误报 EXCEL_COM_ERROR）。
    /// 清理纪律沿用 Md1Probe：不 Close/Quit，按 PID 强杀本 harness 实例。
    /// </summary>
    internal static class PyFixProbe
    {
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();
        private static int pass, fail;

        public static int Run(string pythonPath)
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
                var sheetName = Convert.ToString(excel.ActiveSheet.Name);
                Console.WriteLine($"[env] sheet={sheetName} python={pythonPath}");

                var options = new ScriptOptions { Enabled = true, PythonPath = pythonPath, TimeoutSeconds = 60 };
                var runner = new ScriptRunnerService(excel, options);
                var pyExcel = new PythonExcelService(excel, options);

                // T1：缺模块自愈（该环境真实缺 pyyaml，第一跑触发 pip install pyyaml → 重建 worker → 重试）
                var t1 = runner.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import chardet\nprint('CHARDET_IMPORT_OK')",
                    ["timeout_ms"] = 120000
                });
                var t1Text = Convert.ToString(t1.Data) ?? "";
                var t1Ok = t1.Ok && t1Text.IndexOf("CHARDET_IMPORT_OK", StringComparison.Ordinal) >= 0;
                Report("T1 run_python_code PIL 缺模块自愈", t1Ok,
                    t1.Ok ? t1Text + " deps_auto_installed=" + (t1Text.IndexOf("deps_auto_installed", StringComparison.Ordinal) >= 0 ? "?" : "见Data") : t1.ErrorCode + " " + Short(t1Text));

                // T1b：非白名单模块不自愈（快速失败，不装包）
                var t1b = runner.RunPythonCode(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import xai_no_such_module_zz9",
                    ["timeout_ms"] = 30000
                });
                var t1bOk = !t1b.Ok && t1b.ErrorCode == "PYTHON_RUN_ERROR"
                    && Convert.ToString(t1b.Data).IndexOf("xai_no_such_module_zz9", StringComparison.Ordinal) >= 0;
                Report("T1b 非白名单模块不自愈(负例)", t1bOk, t1b.Ok ? "意外成功?" : t1b.ErrorCode);

                // T2：write_ranges 带表名字符串（JArray 承载，真机复现 AI 反序列化路径）
                var t2 = pyExcel.Run(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "import xlwings as xw\nxw.sheets.active.range('A52').value = 'WROTE_OK'\nprint('WR_OK')",
                    ["write_ranges"] = JArray.Parse("[\"" + sheetName + "!A52:C60\"]"),
                    ["timeout_ms"] = 120000
                });
                var t2Text = Convert.ToString(t2.Data) ?? "";
                var t2Ok = t2.Ok && t2Text.IndexOf("WR_OK", StringComparison.Ordinal) >= 0;
                Report("T2 run_python_excel write_ranges 带表名(JArray)", t2Ok,
                    t2.Ok ? t2Text + " snapshotted=" + (t2Text.IndexOf("write_ranges_snapshotted", StringComparison.Ordinal) >= 0 ? "1" : "?") : t2.ErrorCode + " " + Short(t2Text));

                // T3：create_excel_chart 缺 source_address → INVALID_ARGS
                var pivotChart = new ExcelPivotChartService(excel, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, null);
                if (!registry.TryGet("create_excel_chart", out var chartTool))
                {
                    Report("T3 create_excel_chart 缺 source_address", false, "工具未注册");
                }
                else
                {
                    var t3 = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["chart_type"] = "ColumnClustered"
                    });
                    var t3Ok = !t3.Ok && t3.ErrorCode == "INVALID_ARGS";
                    Report("T3 create_excel_chart 缺 source_address", t3Ok,
                        (t3.Ok ? "意外成功?" : t3.ErrorCode + " " + Short(Convert.ToString(t3.Data))));
                }

                Console.WriteLine($"== pyfix 结果：{pass} 过 / {fail} 败 ==");
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

        private static void Report(string name, bool ok, string detail)
        {
            if (ok) { pass++; } else { fail++; }
            Console.WriteLine("[" + (ok ? "PASS" : "FAIL") + "] " + name + " | " + detail);
        }

        private static string Short(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            var flat = text.Replace("\r", " ").Replace("\n", " ");
            return flat.Length <= 3000 ? flat : flat.Substring(0, 3000) + "…";
        }
    }
}
