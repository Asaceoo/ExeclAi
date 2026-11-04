using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyPhaseC
{
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();
        private static string lastChartName;

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

                // 准备一张可供分析的分类+数值数据表
                dynamic setupSheet = wb.Worksheets.Item[1];
                setupSheet.Name = "p4c";
                object[,] data = new object[,]
                {
                    { "类别", "数量", "金额" },
                    { "甲", 10D, 100D },
                    { "乙", 20D, 200D },
                    { "丙", 30D, 300D },
                    { "丁", 40D, 400D }
                };
                var setupRange = setupSheet.Range["A1:C5"];
                setupRange.Value2 = data;
                Release((object)setupRange);
                Release((object)setupSheet);

                var host = new ExcelToolHost((object)app);
                var registry = host.CreateRegistry();

                // ==================== C1: 一键插入推荐图表 ====================
                Check(registry, "insert_recommended_chart", A("sheet", "p4c", "address", "A1:C5"),
                    e =>
                    {
                        var name = Str(e, "chart_name");
                        var type = Str(e, "chart_type");
                        var title = Str(e, "title");
                        return new R(e.Ok && !string.IsNullOrEmpty(name) && type == "ColumnClustered" && !string.IsNullOrEmpty(title),
                            "chart=" + name + " type=" + type + " title=" + title);
                    });

                // 校验图表对象确实插入到工作表
                dynamic verifySheet = null;
                dynamic verifyWb = null;
                try
                {
                    verifyWb = app.ActiveWorkbook;
                    verifySheet = verifyWb.Worksheets.Item["p4c"];
                    var chartCount = verifySheet.ChartObjects().Count;
                    var chartName = lastChartName;
                    var found = false;
                    for (var i = 1; i <= Convert.ToInt32(chartCount); i++)
                    {
                        var co = verifySheet.ChartObjects(i);
                        if (Convert.ToString(co.Name) == chartName)
                        {
                            found = true;
                        }

                        Release((object)co);
                    }

                    if (found)
                    {
                        _passed++;
                        Console.WriteLine("PASS  chart-object-present  [name=" + chartName + "]");
                    }
                    else
                    {
                        Fail("chart-object-present", "not found: " + chartName);
                    }
                }
                finally
                {
                    Release((object)verifySheet);
                    Release((object)verifyWb);
                }

                // ==================== C2: 模板化文字报告 ====================
                // 模式1：工作表输出
                Check(registry, "generate_analysis_report", A("sheet", "p4c", "address", "A1:C5", "output", "worksheet", "sheet_name", "分析报告"),
                    e =>
                    {
                        var wsName = Str(e, "worksheet_name");
                        return new R(e.Ok && !string.IsNullOrEmpty(wsName) && Val(e, "word_count") > 0, "sheet=" + wsName + " words=" + Val(e, "word_count"));
                    });

                // 报告各章节应齐全
                Check(registry, "generate_analysis_report", A("sheet", "p4c", "address", "A1:C5", "output", "worksheet"),
                    e =>
                    {
                        if (!e.Ok || e.Data == null) { return new R(false, "no_data"); }
                        var tokens = JObject.FromObject(e.Data)["sections"] as JArray;
                        if (tokens == null) { return new R(false, "no_sections"); }
                        foreach (var expected in new[] { "概览", "趋势", "异常", "重点列摘要", "相关性", "建议" })
                        {
                            var has = false;
                            foreach (var token in tokens)
                            {
                                if (string.Equals(Convert.ToString(token["title"]), expected))
                                {
                                    has = true;
                                    break;
                                }
                            }

                            if (!has) { return new R(false, "missing=" + expected); }
                        }

                        return new R(true, "sections=" + tokens.Count);
                    });

                // 模式2：Markdown 文件
                var mdPath = Path.Combine(Path.GetTempPath(), "XAI_Report_" + Guid.NewGuid().ToString("N") + ".md");
                Check(registry, "generate_analysis_report", A("sheet", "p4c", "address", "A1:C5", "output", "markdown", "output_path", mdPath),
                    e =>
                    {
                        var ok = e.Ok && File.Exists(mdPath)
                            && File.ReadAllText(mdPath).Contains("# XAI 分析报告")
                            && File.ReadAllText(mdPath).Contains("## 概览");
                        return new R(ok, "file=" + Str(e, "file_path") + " exists=" + File.Exists(mdPath));
                    });

                // 模式3：HTML 文件
                var htmlPath = Path.Combine(Path.GetTempPath(), "XAI_Report_" + Guid.NewGuid().ToString("N") + ".html");
                Check(registry, "generate_analysis_report", A("sheet", "p4c", "address", "A1:C5", "output", "html", "output_path", htmlPath),
                    e =>
                    {
                        var ok = e.Ok && File.Exists(htmlPath) && File.ReadAllText(htmlPath).Contains("<h1>XAI 分析报告</h1>");
                        return new R(ok, "file=" + Str(e, "file_path") + " exists=" + File.Exists(htmlPath));
                    });

                Console.WriteLine();
                Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
                if (Failures.Count > 0)
                {
                    foreach (var failure in Failures)
                    {
                        Console.WriteLine("FAILED: " + failure);
                    }

                    result = "FAIL";
                }
                else
                {
                    Console.WriteLine("ALL P4-C TOOLS VERIFIED OK");
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
                        wb.Close(false);
                        Release((object)wb);
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
                        Release((object)app);
                    }
                }
                catch
                {
                }
            }

            return result == "PASS" ? 0 : 1;
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
                if (res.Ok && name == "insert_recommended_chart")
                {
                    lastChartName = Str(res, "chart_name");
                }

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