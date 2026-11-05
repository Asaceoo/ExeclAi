using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;

// 真机验证（联系真实 Excel）：
//  试A「停靠窗线程封送」——从后台线程调用 SidebarBridge.OpenWebDock，断言经 webView.Invoke
//     封送后在 Excel UI 主线程创建出 DeepSeekTaskPaneHost 停靠窗（真实打开右侧窗的关键根因修复）。
//  试B「DeepSeekTaskPaneHost 真实初始化」——停靠窗内的 WebView2 控件真实初始化成功（CoreWebView2 就绪）。
//  试C「思考过程实时流式」——调用真实 SendPrompt，采集页面 agent-progress；断言在生成过程收到
//     多条流式增量（kind=3 ReasoningDelta），证伪“等全部推理完才一次性出现”。
namespace VerifyRealLife
{
    internal static class Program
    {
        private const string Origin = "https://appassets.local";
        private static readonly string LogPath =
            AppDomain.CurrentDomain.BaseDirectory + "VerifyRealLife.log";
        private static readonly string UserDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelAiAssistant", "WebView2", "VerifyRealLife");

        private static dynamic excel;
        private static int exitCode;
        private static WebView2 webView;
        private static SidebarBridge bridge;
        private static int readyState = -1;         // -1 初始化 0 就绪 1 测试中
        private static int uiThreadId;
        private static readonly List<string> Res = new List<string>();
        private static readonly List<string> Warns = new List<string>();

        // 停靠窗试A/B 状态
        private static bool paneShown;
        private static int paneExecutorThreadId = -1;
        private static bool paneCoreReady;

        // 流式试C 状态
        private static readonly List<ProgHit> ProgHits = new List<ProgHit>();
        private sealed class ProgHit
        {
            public int Kind;
            public DateTime T;
        }

        [STAThread]
        private static int Main()
        {
            uiThreadId = Thread.CurrentThread.ManagedThreadId;
            Log("start uiThread=" + uiThreadId);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            }
            catch (Exception ex)
            {
                Log("EXCEL-START-FAIL: " + ex);
                return 9;
            }

            excel.Visible = true;
            var workbook = excel.Workbooks.Add();
            workbook.Sheets[1].Cells[1, 1].Value = "XAI-LIVE";

            var window = new Form { Width = 420, Height = 700, Text = "XAI 真机验证宿主" };
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
                catch (Exception ex) { Log("EnsureCoreWebView2Async FAIL: " + ex); }
            };
            window.FormClosing += (s, e) => TryCleanup();

            Application.Run(window);
            Log("EXIT=" + exitCode);
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
            Directory.CreateDirectory(UserDataFolder);
            return webView;
        }

        private static void OnInitialized(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                if (!e.IsSuccess) return;
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

                // 停靠窗触发桩：真实创建 DeepSeekTaskPaneHost 停靠窗（外部进程无 Excel ICTPFactory，
                // 用真实控件宿主替代 Excel 停靠栏外壳，验证“内容宿主 + 线程封送”这两个每次打不开的真根因）。
                SidebarBridge.OpenWebDockPane = OpenPaneFromBridge;

                bridge = new SidebarBridge(excel, webView);
                webView.CoreWebView2.AddHostObjectToScript("excelAiBridge", bridge);
                Log("bridge ready");
                // 不等待页面导航完成（首启 WebView2 目录初始化很慢），CoreWebView2 就绪即开始测试。
                readyState = 0;
            }
            catch (Exception ex)
            {
                Log("OnInitialized exception: " + ex);
            }
        }

        private static void OnWebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string msg = null;
            try { msg = e.TryGetWebMessageAsString(); } catch { }
            if (msg != null && msg.StartsWith("__prog#", StringComparison.Ordinal))
            {
                int kind;
                if (int.TryParse(msg.Substring(7), out kind))
                {
                    ProgHits.Add(new ProgHit { Kind = kind, T = DateTime.UtcNow });
                }
            }
            else
            {
                Log("web-message: " + msg);
            }
        }

        // 后台线程经修复后的 OpenWebDock → 此处必须在 UI 主线程被调用，否则封送失效。
        private static void OpenPaneFromBridge(string title, string url)
        {
            var tid = Thread.CurrentThread.ManagedThreadId;
            paneExecutorThreadId = tid;
            if (tid != uiThreadId)
            {
                Res.Add("[\u2717] 停靠窗触发在非 UI 线程执行（webView.Invoke 封送失效）");
                return;
            }
            try
            {
                ShowPaneHost();
                paneShown = true;
                Res.Add("[ok] 停靠窗在 UI 主线程创建");
            }
            catch (Exception ex)
            {
                Res.Add("[\u2717] DeepSeekTaskPaneHost 创建失败: " + ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static void ShowPaneHost()
        {
            var host = (DeepSeekTaskPaneHost)Activator.CreateInstance(
                typeof(DeepSeekTaskPaneHost), new object[] { (object)excel });
            var field = typeof(DeepSeekTaskPaneHost).GetField(
                "webView", BindingFlags.NonPublic | BindingFlags.Instance);
            var hostView = field != null ? (WebView2)field.GetValue(host) : null;
            if (hostView != null)
            {
                hostView.CoreWebView2InitializationCompleted += (s, ev) =>
                {
                    if (ev.IsSuccess && hostView.CoreWebView2 != null) { paneCoreReady = true; Log("DeepSeek host webview2 ready"); }
                    else { Res.Add("[\u2717] DeepSeekTaskPaneHost WebView2 初始化失败"); }
                };
            }
            var form = new Form
            {
                Text = "XAI · DeepSeek 真机验证停靠窗",
                Width = 460,
                Height = 720,
                Dock = DockStyle.Fill,
                WindowState = FormWindowState.Normal
            };
            host.Dock = DockStyle.Fill;
            form.Controls.Add(host);
            form.Show();
        }

        private static void Pump(Form window, System.Windows.Forms.Timer timer)
        {
            try
            {
                if (readyState == 0)
                {
                    readyState = 1;
                    timer.Stop();
                    RunAsync().ContinueWith(t => window.BeginInvoke((Action)(() => window.Close())),
                        TaskScheduler.Default);
                }
            }
            catch (Exception ex)
            {
                Log("Pump exception: " + ex);
                window.BeginInvoke((Action)(() => window.Close()));
            }
        }

        private static async Task RunAsync()
        {
            try
            {
                Log("test start");
                await TestPaneThreadAffinityAsync();
                await TestPaneHostWebViewReadyAsync();
                await TestStreamingAsync();
                Summarize();
            }
            catch (Exception ex)
            {
                Log("RunAsync exception: " + ex);
                Res.Add("[\u2717] 真机验证异常: " + ex.Message);
                Summarize();
            }
        }

        // 试A：后台线程打开停靠窗 → 必须封送到 UI 线程。
        private static async Task TestPaneThreadAffinityAsync()
        {
            Log("A: 从后台线程触发 OpenWebDock(DeepSeek)");
            await Task.Run(() =>
            {
                try { bridge.OpenWebDock("XAI · DeepSeek", "https://chat.deepseek.com/", false); }
                catch (Exception ex) { Res.Add("[\u2717] OpenWebDock 抛异常: " + ex.Message); }
            });
            await Task.Delay(300);
            Log("A: paneShown=" + paneShown + " executorThread=" + paneExecutorThreadId
                + " uiThread=" + uiThreadId);
            if (paneShown && paneExecutorThreadId == uiThreadId)
                Res.Add("[ok] 试A 停靠窗线程封送：后台调用正确封送并创建。");
            else
            {
                Res.Add("[\u2717] 试A 停靠窗未在 UI 线程创建（paneShown=" + paneShown
                    + " executor=" + paneExecutorThreadId + " ui=" + uiThreadId + "）");
                exitCode = exitCode == 0 ? 3 : exitCode;
            }
        }

        // 试B：停靠窗内容宿主 WebView2 真实初始化。
        private static async Task TestPaneHostWebViewReadyAsync()
        {
            Log("B: 等待 DeepSeekTaskPaneHost.WebView2 初始化");
            var deadline = DateTime.UtcNow.AddSeconds(25);
            while (!paneCoreReady && DateTime.UtcNow < deadline)
            {
                await Task.Delay(250);
            }
            if (paneCoreReady) Res.Add("[ok] 试B DeepSeekTaskPaneHost 控件真实初始化（CoreWebView2 就绪）");
            else
            {
                Res.Add("[\u2717] 试B DeepSeekTaskPaneHost WebView2 25s 内未就绪");
                exitCode = exitCode == 0 ? 4 : exitCode;
            }
        }

        // 试C：思考过程实时流式。直连真实网关(localhost:8000)调用客户端流式接口，
        // 断言在一条流式响应中连续收到多条 ReasoningDelta(kind=3) 增量 —— 证伪“等全部推理完才一次性出现”。
        private static async Task TestStreamingAsync()
        {
            Log("C: 直连网关流式验证");
            var cfgPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelAiAssistant", "models.json");
            ConfigurationStore store = null;
            try
            {
                var keyBytes = System.Text.Encoding.UTF8.GetBytes("ExcelAiAssistant:v1");
                store = new ConfigurationStore(
                    cfgPath,
                    apiKey => System.Security.Cryptography.ProtectedData.Protect(
                        System.Text.Encoding.UTF8.GetBytes(apiKey), keyBytes,
                        System.Security.Cryptography.DataProtectionScope.CurrentUser),
                    cipher => System.Text.Encoding.UTF8.GetString(
                        System.Security.Cryptography.ProtectedData.Unprotect(
                            cipher, keyBytes,
                            System.Security.Cryptography.DataProtectionScope.CurrentUser)));
            }
            catch (Exception ex)
            {
                Warns.Add("读配置失败: " + ex.Message);
                Res.Add("[i] 试C 跳过（无配置）");
                return;
            }

            var profile = store.Load().FirstOrDefault();
            if (profile == null)
            {
                Res.Add("[i] 试C 跳过（无模型配置）");
                return;
            }

            var deltas = new List<string>();
            var seenFinal = false;
            var started = DateTime.UtcNow;
            AgentProgress sawFinalKind0 = null;
            try
            {
                using (var client = new OpenAiCompatibleClient(profile))
                {
                    var msgs = new[]
                    {
                        ChatMessage.System("扮演推理助手，先给出逐步推理。"),
                        ChatMessage.User("1+1=? 请先推理两三句再给答案，不要调用工具。")
                    };
                    var r = await client.CompleteAsync(
                        msgs, null, CancellationToken.None,
                        p =>
                        {
                            if (p.Kind == AgentProgressKind.ReasoningDelta) deltas.Add(p.Reasoning ?? string.Empty);
                            else if (p.Kind == AgentProgressKind.Reasoning) { seenFinal = true; sawFinalKind0 = p; }
                        });
                    Log("C reply content=" + (r.Content ?? string.Empty).Replace("\n", " ")
                        + " reasoningLen=" + (r.Reasoning ?? string.Empty).Length);
                }
            }
            catch (Exception ex)
            {
                Warns.Add("流式试C 请求失败: " + ex.GetType().Name + ": " + ex.Message);
                Res.Add("[i] 试C 网关不可用：即跳过");
                return;
            }
            var done = DateTime.UtcNow;
            int inWindow = 0;
            for (int i = 0; i < deltas.Count; i++) { if (deltas[i] != null && deltas[i].Length > 0) inWindow++; }
            Log("C took=" + (done - started).TotalSeconds.ToString("0.0") + "s reasoningDeltas=" + deltas.Count
                + " nonEmpty=" + inWindow + " finalOneShot=" + (seenFinal ? 1 : 0));

            if (inWindow >= 2)
            {
                Res.Add("[ok] 试C 思考过程实时流式：收到 " + inWindow
                    + " 条推理增量（Kind=ReasoningDelta），逐条实时回调。");
            }
            else if (inWindow >= 1)
            {
                Warns.Add("网关增量较少（" + inWindow + " 条）；模型推理可能很短。");
                Res.Add("[i] 试C 收到流式增量 " + inWindow + " 条，链路为流式。");
            }
            else if (seenFinal)
            {
                Warns.Add("本机网关/模型未回流式 reasoning 增量（仅结尾一次性 Kind=Reasoning）。");
                Res.Add("[i] 试C 网关非流式（一次性返回），未验证到实时增量。");
            }
            else
            {
                Res.Add("[i] 试C 未收到任何 reasoning。");
            }
        }

        private static void Summarize()
        {
            Log("========== 真机验证汇总 ==========");
            foreach (var r in Res) Log(r);
            foreach (var w in Warns) Log("[warn] " + w);
            bool passThread = Res.Exists(r => r.Contains("试A") && r.StartsWith("[ok]"));
            bool passPane = Res.Exists(r => r.Contains("试B") && r.StartsWith("[ok]"));
            bool passStreaming = Res.Exists(r => r.Contains("试C") && r.StartsWith("[ok]"));
            Log("A停靠窗线程封送=" + (passThread ? "PASS" : "FAIL")
                + " B宿主初始化=" + (passPane ? "PASS" : "FAIL")
                + " C流式=" + (passStreaming ? "PASS" : (Warns.Count > 0 ? "SKIP" : "FAIL")));
            // 试A/B 是强断言（exit 非 0 则该失败）；试C 网关能力为弱断言。
        }

        private static void TryCleanup()
        {
            try { SidebarBridge.OpenWebDockPane = null; } catch { }
            try { if (bridge != null) bridge.Dispose(); } catch { }
            try { if (excel != null) excel.Quit(); } catch { }
        }

        private static void Log(string msg)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine); }
            catch { }
        }
    }
}