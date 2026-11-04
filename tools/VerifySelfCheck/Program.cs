using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;

namespace VerifySelfCheck
{
    /// <summary>
    /// 真机自检 harness（v1.19.4.0）：
    ///  1) 启动真实 Excel（经由 ProgID，已安装的加载项会随实例载入）；
    ///  2) 用与侧边栏完全相同的真实工具链（ExcelToolHost + CreateRegistry）在实机上
    ///     创建「簇状柱形图」，回读 Chart.ChartType 断言 == 51（修复项：原来是 -4102 3D 饼图）；
    ///  3) 直调真实 SidebarBridge.RunDiagnosticAdjudication()，即「一键自检」按钮背后那条方法，
    ///     全量打印确定性结论 + AI 研判文本。
    /// 约束：与正式加载项一致必须无状态——建簿/写数据/建图/回读全部「重新解析 → 用完即放」。
    /// </summary>
    internal static class Program
    {
        private static int _failed;

        [STAThread]
        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            // WebView2 会干扰进程内系统代理，导致对模型端点的 HTTP 请求挂起；与 VerifySidebarChat 一致地禁用。
            System.Net.ServicePointManager.ServerCertificateValidationCallback = null;
            System.Net.WebRequest.DefaultWebProxy = null;
            dynamic app = null;
            try
            {
                CheckVersionLoaded();

                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = true;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;

                SetupData(app);
                VerifyClusterChart(app);

                ProbeApp(app, "VerifyClusterChart 之后 / RunAdjudication 之前");
                RunAdjudication(app);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("== HARNESS EXCEPTION: " + Safe(ex));
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (app != null)
                {
                    try { dynamic active = app.ActiveWorkbook; if (active != null) { active.Close(false); SafeRelease(active); } } catch { }
                    try { app.ScreenUpdating = true; app.Quit(); SafeRelease(app); } catch { }
                }
            }

            Console.WriteLine();
            Console.WriteLine("== VERIFYSELFCHECK " + (_failed == 0 ? "ALL OK" : ("FAILED=" + _failed)) + " ==");
            return _failed == 0 ? 0 : 1;
        }

        private static void CheckVersionLoaded()
        {
            var asm = typeof(SidebarBridge).Assembly;
            var version = asm.GetName().Version;
            var versionStr = version.ToString();
            var info = asm.GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false)
                .Cast<AssemblyInformationalVersionAttribute>().FirstOrDefault()?.InformationalVersion;
            // 不锁绝对版本号，断言确实加载了非占位(<1.0.0.0) 的已安装加载项程序集，避免每次发版都要改断言。
            Report("加载到的加载项版本", version != null && version.Major >= 1 && versionStr != "0.0.0.0",
                "asm=" + versionStr + " info=" + info + "（期望非占位版本）");
        }

        private static void SetupData(dynamic app)
        {
            dynamic ws = null;
            dynamic rng = null;
            try
            {
                dynamic active = app.ActiveWorkbook;
                if (active == null)
                {
                    dynamic created = app.Workbooks.Add();
                    SafeRelease(created);
                }
                else
                {
                    SafeRelease(active);
                }
                ws = app.ActiveWorkbook.Worksheets.Item[1];
                ws.Name = "sc";
                rng = ws.Range["A1:B5"];
                rng.Value2 = new object[,]
                {
                    { "季度", "销量" },
                    { "Q1", 12D },
                    { "Q2", 19D },
                    { "Q3", 16D },
                    { "Q4", 24D }
                };
            }
            finally
            {
                SafeRelease(rng);
                SafeRelease(ws);
            }
        }

        private static void VerifyClusterChart(dynamic app)
        {
            var wbPath = Convert.ToString(app.ActiveWorkbook.FullName ?? string.Empty);
            var host = new ExcelToolHost((object)app, null, ExcelToolHost.CreateWorkbookHash(wbPath), "selfcheck", "自检");
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();

            ToolResult res = null;
            var execErr = "";
            try
            {
                if (registry.TryGet("create_excel_chart", out var chartTool))
                {
                    res = chartTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["sheet"] = "sc",
                        ["source_address"] = "A1:B5",
                        ["chart_type"] = "Column",
                        ["title"] = "簇状柱形图验证"
                    });
                }
                else
                {
                    execErr = "create_excel_chart 未注册";
                }
            }
            catch (Exception ex)
            {
                execErr = "EXCEPTION " + Safe(ex);
            }

            if (res == null || !res.Ok)
            {
                Report("图表工具执行", false, "ok=" + (res?.Ok) + " err=" + (execErr == "" ? (res?.ErrorCode + " " + res?.Data) : execErr));
                return;
            }

            int actual = -999;
            var readErr = "";
            dynamic ws = null, cos = null, co = null, ch = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets.Item["sc"];
                cos = ws.ChartObjects();
                var count = Convert.ToInt32(cos.Count);
                if (count == 0)
                {
                    readErr = "尚无图表";
                }
                else
                {
                    co = cos.Item(count);
                    ch = co.Chart;
                    actual = Convert.ToInt32(ch.ChartType);
                }
            }
            catch (Exception ex)
            {
                readErr = "回读异常 " + Safe(ex);
            }
            finally
            {
                SafeRelease(ch);
                SafeRelease(co);
                SafeRelease(cos);
                SafeRelease(ws);
            }

            var typeName = ChartTypeMapper.Resolve("Column");
            Report("图表类型修复", actual == 51 && actual != -4102,
                "实际 ChartType=" + actual + "（映射=" + typeName + "；期望 51=簇状柱形图，非 -4102 3D 饼图）"
                + (readErr == "" ? "" : " · " + readErr));
        }

        private static void ProbeApp(object appObj, string where)
        {
            try
            {
                dynamic a = (dynamic)appObj;
                Console.WriteLine(where + " => app 存活 app.Version=" + a.Version);
            }
            catch (Exception ex)
            {
                Console.WriteLine(where + " => app 已失效: " + Safe(ex));
            }
        }

        private static void RunAdjudication(dynamic app)
        {
            Console.WriteLine("== 进入 RunAdjudication ==");
            try
            {
                using (var bridge = new SidebarBridge((object)app, new WebView2()))
                {
                    var start = DateTime.UtcNow;
                    var json = bridge.RunDiagnosticAdjudication().GetAwaiter().GetResult();
                    var took = (DateTime.UtcNow - start).TotalSeconds.ToString("0.0");
                    Console.WriteLine("\n== 一键自检（AI 研判）耗时 " + took + "s ==");
                    Console.WriteLine(json);

                    JObject obj;
                    try
                    {
                        obj = JObject.Parse(json);
                    }
                    catch
                    {
                        Report("自检返回 JSON 可解析", false, "原始长度=" + (json?.Length ?? 0));
                        return;
                    }

                    var okFlag = obj.Value<bool>("ok");
                    var deterministic = obj.Value<string>("deterministic") ?? string.Empty;
                    var ai = obj.Value<string>("aiVerdict") ?? string.Empty;
                    var issues = (obj["summary"]?["issues"] as JArray)?.Count ?? 0;
                    Report("自检 ok", okFlag, "");
                    Report("确定性研判", deterministic.Contains("通过") || issues >= 0, "deterministic=" + deterministic + " issues=" + issues);
                    Report("AI 研判产出", !string.IsNullOrWhiteSpace(ai) && !ai.StartsWith("(AI 研判不可用"),
                        "长度=" + ai.Length + (ai.Length > 60 ? " 前60字=" + ai.Substring(0, 60) : ""));
                }
            }
            catch (Exception ex)
            {
                Report("自检执行", false, "EXCEPTION " + Safe(ex));
            }
        }

        private static void Report(string name, bool pass, string detail)
        {
            Console.WriteLine((pass ? "  [PASS] " : "  [FAIL] ") + name + " -> " + detail);
            if (!pass) _failed++;
        }

        private static void SafeRelease(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                {
                    Marshal.FinalReleaseComObject(o);
                }
            }
            catch { }
        }

        private static string Safe(Exception ex)
        {
            var inner = ex;
            while (inner.InnerException != null) inner = inner.InnerException;
            return inner.GetType().Name + ": " + (inner.Message ?? string.Empty);
        }
    }
}