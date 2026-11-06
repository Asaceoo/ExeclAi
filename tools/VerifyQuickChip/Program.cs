using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAiAssistant.Addin;
using Newtonsoft.Json.Linq;

namespace VerifyQuickChip
{
    /// <summary>
    /// 实机验证「快捷/技能 chip 一键触达」修复：
    /// 点击一个 quick chip → send() 需被同步触发（发送按钮立即隐藏、prompt 被清空消费），
    /// 且真实 SidebarBridge.SendPrompt 被调用并收到该 chip 的 prompt。
    /// 判据（修复前：仅填入 prompt 不发送 → after 为空、prompt 非空；修复后：after='none'、prompt 为空）。
    /// </summary>
    internal static class Program
    {
        private const string Origin = "https://appassets.local";
        private static readonly string MyLog =
            AppDomain.CurrentDomain.BaseDirectory + "VerifyQuickChip.log";
        private static readonly string AppLog =
            Path.Combine(Path.GetTempPath(), "ExcelAiAssistant.log");
        private static readonly string UserDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelAiAssistant", "WebView2", "VerifyQuickChip");
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
            System.Net.ServicePointManager.ServerCertificateValidationCallback = null;
            System.Net.WebRequest.DefaultWebProxy = null;
            Log("addin source version=" + typeof(SidebarBridge).Assembly.GetName().Version);

            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            excel.Visible = true;
            var workbook = excel.Workbooks.Add();
            workbook.Sheets[1].Cells[1, 1].Value = "XAI-CHIP";

            var window = new Form { Width = 480, Height = 780, Text = "JYYJ助手 - 快捷chip验证" };
            webView = new WebView2 { Dock = DockStyle.Fill };
            webView.CoreWebView2InitializationCompleted += (s, e) => Log("webview init success=" + e.IsSuccess);
            webView.CoreWebView2InitializationCompleted += OnInitialized;
            var host = new UserControl { Dock = DockStyle.Fill };
            host.Controls.Add(webView);
            window.Controls.Add(host);
            Directory.CreateDirectory(UserDataFolder);

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
            window.FormClosing += (s, e) =>
            {
                try { bridge?.Dispose(); } catch { }
                try { excel?.Quit(); } catch { }
            };

            Application.Run(window);
            Log("exitCode=" + exitCode + " PASS=" + (exitCode == 0));
            return exitCode;
        }

        private static void OnInitialized(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                if (!e.IsSuccess) { Log("OnInitialized failure"); return; }
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.AddWebResourceRequestedFilter(Origin + "/*", CoreWebView2WebResourceContext.All);
                webView.CoreWebView2.WebResourceRequested += (s2, req) =>
                {
                    if (string.Equals(req.Request.Uri, Origin + "/sidebar", StringComparison.OrdinalIgnoreCase))
                    {
                        Stream stream = null;
                        try
                        {
                            stream = typeof(ExcelAiAssistant.Addin.SidebarBridge).Assembly
                                .GetManifestResourceStream("ExcelAiAssistant.Addin.Sidebar.index.html");
                            if (stream == null) { req.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(null, 404, "Missing", "Content-Type: text/html"); return; }
                            var ms = new MemoryStream();
                            stream.CopyTo(ms); ms.Position = 0;
                            req.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(
                                ms, 200, "OK", "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
                        }
                        finally { stream?.Dispose(); }
                    }
                };
                bridge = new SidebarBridge(excel, webView);
                webView.CoreWebView2.AddHostObjectToScript("excelAiBridge", bridge);
                webView.Source = new Uri(Origin + "/sidebar");
                Log("source assigned, bridge ready");
                progress = 0;
            }
            catch (Exception ex) { Log("OnInitialized exception: " + ex); }
        }

        private static void Pump(Form window, System.Windows.Forms.Timer timer)
        {
            if (progress != 0) return;
            timer.Stop();
            RunProbe(window);
        }

        private static async void RunProbe(Form window)
        {
            try
            {
                Log("probe started");
                // 前端异步渲染（skills 加载后重建 DOM），探测时机有抖动：从固定延时改为「C# 同步轮询就绪」。
                var status = await PollReady(window);
                Log("readyStatus=" + status);
                if (status == "timeout")
                {
                    Log("PROBE raw=null (ready poll timeout)");
                    exitCode = 8;
                    window.BeginInvoke((Action)(() => window.Close()));
                    return;
                }

                var raw = await webView.CoreWebView2.ExecuteScriptAsync(
                    "(function(){" +
                    "  var chip=document.querySelector('.quick-chip:not(.add)');" +
                    "  var sendBtn=document.getElementById('sendButton');" +
                    "  var promptEl=document.getElementById('prompt');" +
                    "  var out={ok:false};" +
                    "  try{" +
                    "    out={ok:true, name:(chip.textContent||'').replace(/\\s+/g,' ').trim(), title:chip.title||''," +
                    "         beforeBefore:sendBtn.style.display, afterSend:'(no)', promptAfter:'(n-a)'};" +
                    "    chip.click();" +
                    "    out.afterSend=sendBtn.style.display; out.promptAfter=JSON.stringify(promptEl.value);" +
                    "  }catch(e){ out.ok=false; out.clickError=String(e); }" +
                    "  return out;" +
                    "})()");
                Log("PROBE raw=" + raw);
                var obj = JObject.Parse(raw);
                bool ok = obj.Value<bool?>("ok") ?? true;
                if (obj.Value<string>("clickError") != null) { Log("clickError=" + obj.Value<string>("clickError")); }
                if (!ok) { Log("PROBE click failed (err " + obj.Value<string>("clickError") + ")"); exitCode = 9; window.BeginInvoke((Action)(() => window.Close())); return; }
                string name = obj.Value<string>("name");
                string title = obj.Value<string>("title");
                string afterSend = obj.Value<string>("afterSend");
                string promptAfter = (obj.Value<string>("promptAfter") ?? "").Trim();
                Log("chip.name=" + (name ?? "(none)"));
                Log("chip.title=" + (title ?? "(none)"));
                Log("afterSend=" + (afterSend ?? "(none)") + " promptAfter=" + promptAfter);

                // 判据1：新代码标题包含「点击立即执行」→ 确认部署的是修复版
                bool titleOk = title.Contains("点击立即执行");
                // 判据2：send() 被同步触发 → 发送按钮 style.display='none'
                bool sendRan = afterSend == "none";
                // 判据3：prompt 被 send() 消费清空
                bool promptConsumed = promptAfter == "\"\"" || promptAfter == "";
                Log("titleOk=" + titleOk + " sendRan=" + sendRan + " promptConsumed=" + promptConsumed);

                // 判据4：真实 SidebarBridge.SendPrompt 被调用并收到该 prompt（读应用日志末尾）
                await System.Threading.Tasks.Task.Delay(1200);
                string needle = "SendPrompt entered preview=";
                var tail = ReadTail(AppLog, 200);
                bool bridgeCalled = tail.Contains(needle);
                bool previewMatch = false;
                if (bridgeCalled)
                {
                    string prefix = (title.Contains("（点击立即执行）")
                        ? title.Substring(0, title.IndexOf("（点击立即执行）"))
                        : title).Trim();
                    if (!string.IsNullOrEmpty(prefix))
                    {
                        string shortPrefix = prefix.Length > 20 ? prefix.Substring(0, 20) : prefix;
                        previewMatch = tail.Contains(needle + shortPrefix);
                    }
                }
                Log("bridgeCalled=" + bridgeCalled + " previewMatch=" + previewMatch);
                Log("---- latest app log (last 600 chars) ----");
                Log(tail.Length > 600 ? tail.Substring(tail.Length - 600) : tail);

                bool pass = titleOk && sendRan && promptConsumed && bridgeCalled && previewMatch;
                if (!pass) exitCode = 6;
                Log(pass ? "RESULT PASS (chip 一键触达闭环通)" : "RESULT FAIL");
                window.BeginInvoke((Action)(() => window.Close()));
            }
            catch (Exception ex)
            {
                exitCode = 7;
                Log("RunProbe exception: " + ex);
                window.BeginInvoke((Action)(() => window.Close()));
            }
        }

        /// <summary>
        /// C# 侧同步轮询：每隔 250ms 查一次前端 DOM，直到 quick chip 与输入元素就绪，
        /// 或超过 ~15s 判定超时。返回 "ready" 或 "timeout"。
        /// </summary>
        private static async System.Threading.Tasks.Task<string> PollReady(Form window)
        {
            const int tries = 60;
            for (int i = 1; i <= tries; i++)
            {
                var raw = await webView.CoreWebView2.ExecuteScriptAsync(
                    "((function(){" +
                    "  var chip=document.querySelector('.quick-chip:not(.add)');" +
                    "  var sendBtn=document.getElementById('sendButton');" +
                    "  var promptEl=document.getElementById('prompt');" +
                    "  return {chip:!!chip, send:!!sendBtn, prompt:!!promptEl};" +
                    "})())");
                var obj = JObject.Parse(raw);
                bool chip = obj.Value<bool?>("chip") ?? false;
                bool send = obj.Value<bool?>("send") ?? false;
                bool prompt = obj.Value<bool?>("prompt") ?? false;
                if (i==1 || i==tries) Log("  poll#" + i + " chip=" + chip + " send=" + send + " prompt=" + prompt);
                if (chip && send && prompt) return "ready";
                await System.Threading.Tasks.Task.Delay(250);
            }
            return "timeout";
        }

        private static string ReadTail(string path, int maxLines)
        {
            try
            {
                if (!File.Exists(path)) return "(log missing)";
                var lines = File.ReadAllLines(path);
                int skip = Math.Max(0, lines.Length - maxLines);
                return string.Join("\n", lines.Skip(skip));
            }
            catch (Exception ex) { return "(log read error: " + ex.Message + ")"; }
        }

        private static void Log(string msg)
        {
            try { File.AppendAllText(MyLog, DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine); }
            catch { }
            Console.WriteLine(msg);
        }
    }
}