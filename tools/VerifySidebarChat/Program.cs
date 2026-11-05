using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace VerifySidebarChat
{
    internal static class Program
    {
        private const string Origin = "https://appassets.local";
        private static readonly string LogPath =
            AppDomain.CurrentDomain.BaseDirectory + "VerifySidebarChat.log";
        private static readonly string UserDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelAiAssistant", "WebView2", "VerifySidebarChat");
        private static dynamic excel;
        private static int exitCode;
        private static WebView2 webView;
        private static SidebarBridge bridge;
        private static int progress = -1;

        [STAThread]
        private static int Main()
        {
            Log("start");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Log("DefaultWebProxy=" + (System.Net.WebRequest.DefaultWebProxy?.GetProxy(new Uri("https://token.sensenova.cn"))?.ToString() ?? "null"));
            System.Net.ServicePointManager.ServerCertificateValidationCallback = null;
            System.Net.WebRequest.DefaultWebProxy = null; // 假设 WebView2 干扰了系统代理，先禁用验证
            Log("DefaultWebProxy disabled");

            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            excel.Visible = true;
            var workbook = excel.Workbooks.Add();
            workbook.Sheets[1].Cells[1, 1].Value = "XAI-LIVE";

            var window = new Form { Width = 480, Height = 780, Text = "JYYJ助手 - 实机对话验证" };
            webView = PrepareHost(window);

            var timer = new System.Windows.Forms.Timer { Interval = 500 };
            timer.Tick += (s, e) => Pump(window, timer);
            window.Load += async (s, e) =>
            {
                timer.Start();
                try
                {
                    var env = await CoreWebView2Environment.CreateAsync(userDataFolder: UserDataFolder);
                    await webView.EnsureCoreWebView2Async(env);
                    Log("EnsureCoreWebView2Async completed");
                }
                catch (Exception ex) { Log("EnsureCoreWebView2Async failure: " + ex); }
            };
            window.FormClosing += (s, e) => TryQuitExcel();

            Application.Run(window);
            return exitCode;
        }

        private static WebView2 PrepareHost(Form window)
        {
            webView = new WebView2 { Dock = DockStyle.Fill };
            webView.CoreWebView2InitializationCompleted += (s, e) =>
                Log("webview init success=" + e.IsSuccess);
            webView.WebMessageReceived += OnWebMessage;
            webView.CoreWebView2InitializationCompleted += OnInitialized;
            var host = new UserControl { Dock = DockStyle.Fill };
            host.Controls.Add(webView);
            window.Controls.Add(host);
            Log("host added, creating environment folder=" + UserDataFolder);
            Directory.CreateDirectory(UserDataFolder);
            return webView;
        }

        private static void OnInitialized(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                Log("OnInitialized success=" + e.IsSuccess);
                if (!e.IsSuccess) return;
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.AddWebResourceRequestedFilter(Origin + "/*", CoreWebView2WebResourceContext.All);
                webView.CoreWebView2.WebResourceRequested += (s2, req) =>
                {
                    if (string.Equals(req.Request.Uri, Origin + "/sidebar", StringComparison.OrdinalIgnoreCase))
                    {
                        Stream stream = null;
                        try
                        {
                            stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
                                "ExcelAiAssistant.Addin.Sidebar.index.html");
                            if (stream == null)
                            {
                                req.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(
                                    null, 404, "Missing resource", "Content-Type: text/html");
                                return;
                            }
                            var ms = new MemoryStream();
                            stream.CopyTo(ms);
                            ms.Position = 0;
                            req.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(
                                ms, 200, "OK", "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
                        }
                        finally { stream?.Dispose(); }
                    }
                };
                bridge = new SidebarBridge(excel, webView);
                webView.CoreWebView2.AddHostObjectToScript("excelAiBridge", bridge);
                webView.Source = new Uri(Origin + "/sidebar");
                Log("source assigned, bridge ready, progress=" + progress);
                progress = 0;
            }
            catch (Exception ex)
            {
                Log("OnInitialized exception: " + ex);
            }
        }

        private static void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            Log("web-message: " + e.TryGetWebMessageAsString());
        }

        private static void Pump(Form window, System.Windows.Forms.Timer timer)
        {
            try
            {
                if (progress != 0) return;
                timer.Stop();
                RunConversation(window);
            }
            catch (Exception ex)
            {
                Log("Pump exception: " + ex);
                window.BeginInvoke((Action)(() => window.Close()));
            }
        }

        private static async void RunConversation(Form window)
        {
            try
            {
                Log("conversation started");
                await RawProbeAsync();

                var chat = await Send(
                    "请用一句话自我介绍，并说明你是嵌入 Excel 的助手，可以用工具操作工作表。");
                FinishChat(chat);

                var read = await Send(
                    "请读取当前工作表 A1 单元格的内容并原样告诉我。");
                FinishRead(read);

                Log("done, scheduling close");
                window.BeginInvoke((Action)(() => window.Close()));
            }
            catch (Exception ex)
            {
                Log("RunConversation exception: " + ex);
                exitCode = 5;
                window.BeginInvoke((Action)(() => window.Close()));
            }
        }

        // 与 OpenAICompatibleClient 直接联调的原始探测，输出完整异常以定位“发送请求时出错”根因。
        private static async System.Threading.Tasks.Task RawProbeAsync()
        {
            try
            {
                var cfgPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAiAssistant", "models.json");
                var store = new ConfigurationStore(
                    cfgPath,
                    apiKey => System.Security.Cryptography.ProtectedData.Protect(
                        System.Text.Encoding.UTF8.GetBytes(apiKey),
                        System.Text.Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                        System.Security.Cryptography.DataProtectionScope.CurrentUser),
                    cipher => System.Text.Encoding.UTF8.GetString(
                        System.Security.Cryptography.ProtectedData.Unprotect(
                            cipher,
                            System.Text.Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                            System.Security.Cryptography.DataProtectionScope.CurrentUser)));
                var profile = store.Load().FirstOrDefault();
                if (profile == null) { Log("PROBE: no profile"); return; }
                Log("PROBE profile=" + profile.ModelName
                    + " key=" + profile.ApiKey.Substring(0, Math.Min(7, profile.ApiKey.Length)) + "***"
                    + " base=" + profile.BaseUrl + " mode=" + profile.ProtocolMode);

                using (var client = new OpenAiCompatibleClient(profile))
                {
                    var msgs = new[] { ChatMessage.System("你是嵌入 Excel 的操作助手。"), ChatMessage.User("hi") };
                    try
                    {
                        var r = await client.CompleteAsync(msgs, CancellationToken.None).ConfigureAwait(false);
                        Log("PROBE-A(no-tools) OK content=" + (r.Content ?? "").Substring(0, Math.Min(40, (r.Content ?? "").Length)).Replace("\n", " "));
                    }
                    catch (Exception ex) { Log("PROBE-A(no-tools) FAIL: " + ex.GetType().Name + ": " + ex.Message); }

                    var tool = new ToolDescriptor(
                        "read_cells",
                        "读取当前选定单元格范围的内容",
                        RiskLevel.ReadOnly,
                        _ => new ToolResult { Ok = true, Data = "ok" },
                        new Dictionary<string, object> { ["range"] = new Dictionary<string, object> { ["type"] = "string" } });
                    try
                    {
                        var r = await client.CompleteAsync(msgs, new[] { tool }, CancellationToken.None).ConfigureAwait(false);
                        Log("PROBE-B(with-tools) OK calls=" + r.ToolCalls.Count);
                    }
                    catch (Exception ex) { Log("PROBE-B(with-tools) FAIL: " + ex.GetType().Name + ": " + ex.Message); }
                }
            }
            catch (Exception ex) { Log("PROBE exception: " + ex); }
        }

        // 直接调用真实 SidebarBridge 的 SendPrompt —— 与侧边栏 JS 调用的方法/路径完全一致。
        private static async System.Threading.Tasks.Task<string> Send(string prompt)
        {
            var start = DateTime.UtcNow;
            var json = await bridge.SendPrompt(prompt);
            Log("SendPrompt(" + prompt.Substring(0, System.Math.Min(prompt.Length, 12)) + "...) took "
                + (DateTime.UtcNow - start).TotalSeconds.ToString("0.0") + "s -> " + json.Length + " chars");
            return json;
        }

        private static void FinishChat(string raw)
        {
            Log("CHAT-RESULT raw=" + raw);
            try
            {
                var json = Unwrap(raw);
                var obj = JObject.Parse(json);
                bool ok = obj.Value<bool>("ok");
                var answer = obj.Value<string>("answer") ?? "";
                Log("CHAT ok=" + ok + " answer=" + answer);
                if (ok && !string.IsNullOrWhiteSpace(answer)) Log("CHAT PASS");
                else { exitCode = 3; Log("CHAT FAIL"); }
            }
            catch (Exception ex) { exitCode = 3; Log("CHAT parse fail: " + ex); }
        }

        private static void FinishRead(string raw)
        {
            Log("READ-RESULT raw=" + raw);
            try
            {
                var json = Unwrap(raw);
                var obj = JObject.Parse(json);
                bool ok = obj.Value<bool>("ok");
                var answer = obj.Value<string>("answer") ?? "";
                int toolCount = (obj["tool_events"] as JArray)?.Count ?? 0;
                Log("READ ok=" + ok + " tools=" + toolCount + " answer=" + answer);
                if (ok && toolCount > 0) Log("READ PASS (tool used)");
                else { exitCode = 4; Log("READ FAIL"); }
            }
            catch (Exception ex) { exitCode = 4; Log("READ parse fail: " + ex); }
        }

        private static string Unwrap(string raw)
        {
            var s = (raw ?? "").Trim();
            // ExecuteScriptAsync 会 JSON 编码返回值；若返回的是带引号字符串则再解码。
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
            {
                try { return JsonConvert.DeserializeObject<string>(s) ?? string.Empty; }
                catch { return s; }
            }
            // .NET Task<string> 解析后可能仍是字符串，再剥一层。
            try
            {
                var tok = JToken.Parse(s);
                if (tok.Type == JTokenType.String)
                    return Unwrap(tok.Value<string>());
            }
            catch { }
            return s;
        }

        private static void TryQuitExcel()
        {
            try { bridge?.Dispose(); } catch { }
            try { excel?.Quit(); } catch { }
        }

        private static void Log(string msg)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine); }
            catch { }
        }
    }
}