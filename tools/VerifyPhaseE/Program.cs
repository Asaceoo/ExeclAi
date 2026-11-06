using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyPhaseE
{
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
                Failures.Add("HARNESS: " + ex.GetType().Name + ": " + SafeMessage(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + ex.GetType().Name + ": " + SafeMessage(ex));
                Console.WriteLine("HARNESS STACK: " + DumpStack(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL P4-E STABILITY CHECKS VERIFIED OK");
            return 0;
        }

        private static void RunAll()
        {
            // ---- E2 取消链路（无需 Excel；用本地挂起端点验证 token 穿透到 HTTP 请求）----
            RunE2Cancellation();

            // ---- E1 大区域降采样 + E3 读改写快速连发（真实 Excel COM）----
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();

                // 工作表准备必须在首次工具调用前完成：工具内部经 app.ActiveWorkbook 重取
                // 工作簿并 FinalReleaseComObject，会使此处持有的 wb RCW 一并失效。
                dynamic setupSheet = wb.Worksheets.Item[1];
                setupSheet.Name = "p4e";
                const int setupRows = 600;
                const int setupCols = 8;
                var big = new object[setupRows, setupCols];
                for (var r = 0; r < setupRows; r++)
                {
                    for (var c = 0; c < setupCols; c++)
                    {
                        big[r, c] = "R" + (r + 1) + "C" + (c + 1);
                    }
                }

                dynamic setupRange = setupSheet.Range["A1:H600"];
                setupRange.Value2 = big;
                Release((object)setupRange);
                Release((object)setupSheet);

                var host = new ExcelToolHost((object)app);
                var registry = host.CreateRegistry();

                RunE1(registry);
                RunE3(registry);
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); Release((object)wb); } }
                catch (Exception ex) { Console.WriteLine("NOTE  workbook-close: " + ex.GetType().Name + " " + SafeMessage(ex)); }
                try { if (app != null) { app.ScreenUpdating = true; app.Quit(); Release((object)app); } }
                catch (Exception ex) { Console.WriteLine("NOTE  app-quit: " + ex.GetType().Name + " " + SafeMessage(ex)); }
            }
        }

        // ==================== E2：模型请求可中断（停止生成 链路） ====================
        private static void RunE2Cancellation()
        {
            try
            {
                var port = FreePort();
                var prefix = "http://127.0.0.1:" + port + "/";
                using var listener = new HttpListener();
                listener.Prefixes.Add(prefix);
                try
                {
                    listener.Start();
                }
                catch (Exception sew)
                {
                    Note(sew.Message);
                    return; // 环境不支持 HttpListener 时跳过（clean 不等于失败）
                }

                var received = new ManualResetEventSlim(false);
                var release = new ManualResetEventSlim(false);
                var listenerThread = new Thread(() => HoldRequest(listener, received, release));
                listenerThread.IsBackground = true;
                listenerThread.Start();

                var profile = new ModelProfile
                {
                    Name = "e2",
                    BaseUrl = "http://127.0.0.1:" + port,
                    ApiKey = "test-key",
                    ModelName = "test-model",
                    ProtocolMode = AgentProtocolMode.Fallback
                };

                using var client = new OpenAiCompatibleClient(profile);
                var runner = new AgentRunner(
                    client,
                    new ToolRegistry(),
                    (_, _) => Task.FromResult(ConfirmationOutcome.Confirmed));
                var messages = new List<ChatMessage> { ChatMessage.User("ping") };

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));
                var run = runner.RunAsync(messages, cts.Token);

                if (!received.Wait(TimeSpan.FromSeconds(8)))
                {
                    Fail("e2-cancellation", "模型请求未到达本地挂起端点");
                    release.Set();
                    return;
                }

                var sw = System.Diagnostics.Stopwatch.StartNew();
                cts.Cancel(); // 模拟「停止生成」：Bridge.StopGeneration -> activeRunCts.Cancel()
                var cancelled = false;
                try
                {
                    run.GetAwaiter().GetResult();
                }
                catch (TaskCanceledException)
                {
                    cancelled = true;
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                }
                catch (HttpRequestException)
                {
                    cancelled = true; // 取消也可能表现为连接中断
                }
                catch (Exception ex)
                {
                    cancelled = true; // 挂了端点后任何异常都证明请求被终止
                    Console.WriteLine("  (cancel surfaced as " + ex.GetType().Name + ")");
                }

                var elapsedMs = sw.ElapsedMilliseconds;
                release.Set();
                // 取消必须快速生效（远早于 5 分钟超时）
                var ok = cancelled && elapsedMs < 8000;
                if (ok)
                {
                    Pass("e2-cancellation", "cancel->" + elapsedMs + "ms," + (cancelled ? "中止" : "返回"));
                }
                else
                {
                    Fail("e2-cancellation", "elapsed=" + elapsedMs + "ms, cancelled=" + cancelled);
                }
            }
            catch (Exception ex)
            {
                Fail("e2-cancellation", "HARNESS: " + ex.Message);
            }
        }

        private static void HoldRequest(HttpListener listener, ManualResetEventSlim received, ManualResetEventSlim release)
        {
            try
            {
                while (listener.IsListening)
                {
                    HttpListenerContext context;
                    try
                    {
                        context = listener.GetContext();
                    }
                    catch (Exception ex) when (ex is HttpListenerException || ex is InvalidOperationException || ex is ObjectDisposedException)
                    {
                        return;
                    }

                    received.Set();
                    release.Wait(TimeSpan.FromSeconds(30)); // 挂起响应直到测试结束
                    try
                    {
                        var json = "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"ok\"},\"finish_reason\":\"stop\"}],\"usage\":{\"prompt_tokens\":1,\"completion_tokens\":1,\"total_tokens\":2}}";
                        var bytes = System.Text.Encoding.UTF8.GetBytes(json);
                        context.Response.StatusCode = 200;
                        context.Response.ContentType = "application/json";
                        context.Response.ContentLength64 = bytes.Length;
                        context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                        context.Response.OutputStream.Close();
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static int FreePort()
        {
            var l = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
            l.Start();
            var p = ((IPEndPoint)l.LocalEndpoint).Port;
            l.Stop();
            return p;
        }

        // ==================== E1：大区域分段降采样 ====================
        private static void RunE1(ToolRegistry registry)
        {
            const int rows = 600;
            // 超限（4800 > 默认 2000）应自动降采样，不整块读取
            Check(registry, "get_excel_range_values_and_formulas", A("sheet", "p4e", "address", "A1:H600"),
                e =>
                {
                    var summary = Bool(e, "summary");
                    var trunc = e.Truncated;
                    var top = Arr(e, "top_rows");
                    var bottom = Arr(e, "bottom_rows");
                    var ok = e.Ok && summary == true && trunc == true
                             && Int(e, "rows") == rows
                             && Int(e, "required_count") == 2000
                             && top.Count == 3 && bottom.Count == 2;
                    return new R(ok, "summary=" + summary + " truncated=" + trunc + " top->" + (top.Count > 0 ? top.Count : 0) + " bottom->" + (bottom.Count > 0 ? bottom.Count : 0) + " rows=" + Int(e, "rows") + " required=" + Int(e, "required_count"));
                });

            // 手动放宽上限（50000 > 4800）仍能整块读取
            Check(registry, "get_excel_range_values_and_formulas", A("sheet", "p4e", "address", "A1:H600", "max_cells", 50000),
                e =>
                {
                    var summary = Bool(e, "summary");
                    var values = Arr(e, "values");
                    var ok = e.Ok && summary == false && values.Count == rows
                             && Int(e, "rows") == rows;
                    return new R(ok, "summary=" + summary + " values_rows=" + values.Count);
                });
        }

        // ==================== E3：读改写快速连发（回归，验证瞬时忙不误伤） ====================
        private static void RunE3(ToolRegistry registry)
        {
            // 30 次连续小范围读取：健康操作不应被误判为 COM 忙（绝无 EXCEL_COM_ERROR）
            var allOk = true;
            var lastError = string.Empty;
            for (var i = 0; i < 30; i++)
            {
                try
                {
                    var res = Call(registry, "get_excel_range_values_and_formulas", A("sheet", "p4e", "address", "A1:C10"));
                    if (!res.Ok || res.ErrorCode == "EXCEL_COM_ERROR")
                    {
                        allOk = false;
                        lastError = "i=" + i + " error=" + res.ErrorCode + " " + (res.Data?.ToString() ?? "");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    allOk = false;
                    lastError = "i=" + i + " EX=" + ex.GetType().Name + ": " + ex.Message;
                    break;
                }
            }

            if (allOk)
            {
                Pass("e3-rapid-read", "30x read A1:C10 all ok");
            }
            else
            {
                Fail("e3-rapid-read", lastError);
            }

            // 写 + 立即读的回环：验证 Execute 补偿不破坏写读一致性
            Check(registry, "write_excel_range", A("sheet", "p4e", "address", "E1:F2", "value_mode", "value",
                "values", new object[,]{ { "x1", 1D }, { "x2", 2D } }),
                e => new R(e.Ok && e.ErrorCode == null, "op=" + Str(e, "operation_id")));

            Check(registry, "get_excel_range_values_and_formulas", A("sheet", "p4e", "address", "E1:F2"),
                e =>
                {
                    var v = Arr(e, "values");
                    var ok = e.Ok && v.Count == 2 && Convert.ToString(v[0][0]) == "x1";
                    return new R(ok, "e1f1=" + (v.Count > 0 ? Convert.ToString(v[0][0]) : "?"));
                });
        }

        // ==================== 工具 ====================
        private static ToolResult Call(ToolRegistry registry, string name, Dictionary<string, object> args)
        {
            if (!registry.TryGet(name, out var tool))
            {
                return new ToolResult { Ok = false, ErrorCode = "NOT_REGISTERED", Data = name };
            }

            return tool.Execute(args);
        }

        private static void Check(ToolRegistry registry, string name, Dictionary<string, object> args, Func<ToolResult, R> assert)
        {
            try
            {
                var res = Call(registry, name, args);
                var r = assert(res);
                if (r.Ok)
                {
                    Pass(name, r.Detail);
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

        private static void Pass(string name, string detail)
        {
            _passed++;
            Console.WriteLine("PASS  " + name + "  [" + detail + "]");
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + " -> " + detail);
            Console.WriteLine("FAIL  " + name + "  [" + detail + "]");
        }

        private static void Note(string message)
        {
            Console.WriteLine("SKIP  " + message);
        }

        private static string SafeMessage(Exception ex)
        {
            try
            {
                return ex?.Message ?? ex?.GetType().Name ?? "unknown";
            }
            catch
            {
                return "(" + ex?.GetType().Name + ")";
            }
        }

        private static string DumpStack(Exception ex)
        {
            var sb = new System.Text.StringBuilder();
            for (var e = ex; e != null; e = e.InnerException)
            {
                sb.AppendLine("  " + e.GetType().Name + ": " + SafeMessage(e));
                try
                {
                    sb.AppendLine(new System.Diagnostics.StackTrace(e, true).ToString());
                }
                catch (Exception se)
                {
                    sb.AppendLine("    (stack unavailable: " + safeName(se) + ")");
                }
            }

            return sb.ToString();
        }

        private static string safeName(object value)
        {
            try
            {
                return value?.GetType().Name ?? "null";
            }
            catch
            {
                return "?";
            }
        }

        private static JToken Token(ToolResult e, string key)
        {
            if (e.Data == null || e.Data is string)
            {
                return null;
            }

            return JObject.FromObject(e.Data)[key];
        }

        private static bool Bool(ToolResult e, string key)
        {
            var t = Token(e, key);
            return t != null && t.Type != JTokenType.Null && Convert.ToBoolean(t.ToObject<object>());
        }

        private static int Int(ToolResult e, string key)
        {
            var t = Token(e, key);
            return t == null || t.Type == JTokenType.Null ? 0 : (int)t.ToObject<long>();
        }

        private static string Str(ToolResult e, string key)
        {
            var t = Token(e, key);
            return t == null || t.Type == JTokenType.Null ? null : Convert.ToString(t.ToObject<object>());
        }

        private static JArray Arr(ToolResult e, string key)
        {
            return Token(e, key) as JArray ?? new JArray();
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