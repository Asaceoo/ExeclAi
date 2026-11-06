using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// DS 右侧停靠侧边栏：与 JYYJ助手一样作为 CustomTaskPane 停靠在 Excel 右侧，
    /// 以顶层导航（非 iframe）承载 chat.deepseek.com，可正常登录、提问、缩放。
    ///
    /// 初始化全程无阻塞（镜像 ExcelAiTaskPaneHost 已被证实的 async void OnLoaded + await 路径），
    /// 绝不用 .GetResult() 同步阻塞 Excel 主线程，避免整体 Excel 卡死。
    /// DeepSeek 拒绝 iframe 内嵌（X-Frame-Options/CSP frame-ancestors），顶层导航不受影响。
    /// </summary>
    [ComVisible(true)]
    [Guid("8F7A6D91-4C52-45BE-9F55-7D8C2A31E003")]
    [ProgId("ExcelAiAssistant.DeepSeekPaneHost")]
    public sealed class DeepSeekTaskPaneHost : UserControl
    {
        internal static dynamic CurrentExcelApplication;
        internal static Action HidePane;
        internal static Action<string> OpenSystemBrowser;

        // 泛化后的通用网页停靠宿主：ExcelAiAddIn 在 CreateCTP 前设定 `CurrentUrl` 后，
        // 本实例即承载该站点；之后可经 NavigateTo 切到任意地址（如自定义 URL）。
        internal static string CurrentUrl = "https://chat.deepseek.com/";

        /// <summary>最近成功加入 CTP 的宿主实例。用于已存在窗格时切换站点——
        /// CTP 的 .Control 在嵌入式互操作类型下既不能强类型也无 dynamic 可达，而实例一定
        /// 在 CTP 内构造，故用静态引用导航最可靠。</summary>
        internal static DeepSeekTaskPaneHost LastInstance;

        private readonly dynamic excelApplication;
        private readonly WebView2 webView;
        private string _initialUrl;
        private Label lblZoom;

        public DeepSeekTaskPaneHost()
            : this((object)CurrentExcelApplication)
        {
        }

        public DeepSeekTaskPaneHost(dynamic excelApplication)
        {
            this.excelApplication = excelApplication ?? throw new ArgumentNullException(nameof(excelApplication));
            _initialUrl = NormalizeUrl(CurrentUrl ?? "https://chat.deepseek.com/");
            LastInstance = this;
            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(245, 247, 250);

            webView = new WebView2 { Dock = DockStyle.Fill };
            webView.CoreWebView2InitializationCompleted += OnInitialized;

            Controls.Add(webView);
            Controls.Add(BuildToolbar());
            Load += OnLoaded;
        }

        /// <summary>把当前停靠窗格导航到指定网址（任意站点 / 自定义 URL 通用）。</summary>
        public void NavigateTo(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return;
            }

            url = NormalizeUrl(url.Trim());
            if (webView?.CoreWebView2 != null)
                {
                try { webView.CoreWebView2.Navigate(url); }
                catch (Exception ex) { DiagnosticsLog.Write("WebDock navigate failed: " + ex); }
            }

            _targetUrl = url;
        }

        private string _targetUrl;

        private static string NormalizeUrl(string url)
        {
            if (url.IndexOf("://", StringComparison.Ordinal) < 0)
            {
                return "https://" + url;
            }

            return url;
        }

        /// <summary>顶部浏览器工具栏：浅色贴近侧边栏风格，避免原生 ToolStrip 字体/箭头渲染异常。</summary>
        private Control BuildToolbar()
        {
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(6, 4, 6, 4),
                WrapContents = false,
                BackColor = Color.FromArgb(240, 243, 248)
            };

            bar.Controls.Add(ToolButton("− 缩小", () => SetZoom(CurrentZoom() - 0.15)));

            lblZoom = new Label
            {
                Text = "100%",
                AutoSize = true,
                ForeColor = Color.FromArgb(31, 105, 224),
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold),
                Margin = new Padding(6, 8, 6, 0),
                TextAlign = ContentAlignment.MiddleLeft
            };
            bar.Controls.Add(lblZoom);

            bar.Controls.Add(ToolButton("＋ 放大", () => SetZoom(CurrentZoom() + 0.15)));
            bar.Controls.Add(ToolButton("重置", () => SetZoom(1.0)));
            bar.Controls.Add(ToolButton("刷新", () => Reload()));
            bar.Controls.Add(ToolButton("系统浏览器 ↗", () => OpenBrowser()));
            bar.Controls.Add(ToolButton("关闭 ✕", () => { try { HidePane?.Invoke(); } catch { } }));

            return bar;
        }

        private static Button ToolButton(string text, Action onClick)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(31, 105, 224),
                Font = new Font("Microsoft YaHei UI", 9f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Margin = new Padding(3),
                FlatAppearance = { BorderColor = Color.FromArgb(219, 224, 235) },
                UseVisualStyleBackColor = false
            };
            if (onClick != null)
            {
                button.Click += (_, __) => onClick();
            }

            return button;
        }

        private async void OnLoaded(object sender, EventArgs eventArgs)
        {
            try
            {
                var folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "XaiAssistant", "ds-browser");
                Directory.CreateDirectory(folder);
                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: folder);
                await webView.EnsureCoreWebView2Async(environment);
            }
            catch (Exception ex)
            {
                DiagnosticsLog.Write("DeepSeek pane init failed: " + ex);
                // data folder 可能被占导致白屏：回落用独立临时 userDataFolder 再试一次。
                try
                {
                    var alt = Path.Combine(Path.GetTempPath(), "XaiDock", Guid.NewGuid().ToString("N"));
                    var env2 = await CoreWebView2Environment.CreateAsync(userDataFolder: alt);
                    await webView.EnsureCoreWebView2Async(env2);
                }
                catch (Exception ex2) { DiagnosticsLog.Write("DeepSeek pane init retry failed: " + ex2); }
            }
        }

        private void OnInitialized(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess || webView.CoreWebView2 == null)
            {
                DiagnosticsLog.Write("DeepSeek pane init failed: " + e.InitializationException);
                return;
            }

            // 以接近普通浏览器的 UA 顶层加载，站点才接受并给出登录/提问界面。
            webView.CoreWebView2.Settings.UserAgent =
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
            try { webView.CoreWebView2.Navigate(ResolveTargetUrl()); }
            catch (Exception ex) { DiagnosticsLog.Write("WebDock pane navigate failed: " + ex); }
        }

        private double CurrentZoom()
        {
            try { return webView.ZoomFactor; }
            catch { return 1.0; }
        }

        private void SetZoom(double zoom)
        {
            if (webView?.CoreWebView2 == null)
            {
                return;
            }

            zoom = Math.Round(Math.Min(3, Math.Max(0.5, zoom)) * 100) / 100;
            try { webView.ZoomFactor = zoom; }
            catch { return; }
            if (lblZoom != null)
            {
                lblZoom.Text = Math.Round(zoom * 100) + "%";
            }
        }

        private void Reload()
        {
            try { webView?.CoreWebView2?.Reload(); }
            catch { }
        }

        // 当前承载地址（供「系统浏览器打开」用）；若中途 NavigateTo 换过则取最新值。
        private string TargetWebUrl => ResolveTargetUrl();

        // _targetUrl 优先（可能是初始化前经 NavigateTo 设定的新地址），其次初始地址。
        private string ResolveTargetUrl() =>
            !string.IsNullOrWhiteSpace(_targetUrl) ? _targetUrl : _initialUrl;

        private void OpenBrowser()
        {
            try
            {
                var url = TargetWebUrl;
                if (OpenSystemBrowser != null)
                {
                    OpenSystemBrowser(url);
                }
                else
                {
                    Process.Start(url);
                }
            }
            catch { }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (webView != null)
                {
                    webView.CoreWebView2InitializationCompleted -= OnInitialized;
                }

                if (LastInstance == this)
                {
                    LastInstance = null; // 窗格销毁后清悬垂引用，避免后续切站导航到已释放实例
                }
            }

            base.Dispose(disposing);
        }
    }
}