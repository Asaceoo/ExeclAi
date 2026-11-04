using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace ExcelAiAssistant.Addin
{
    [ComVisible(true)]
    [Guid("8F7A6D91-4C52-45BE-9F55-7D8C2A31E002")]
    [ProgId("ExcelAiAssistant.TaskPaneHost")]
    public sealed class ExcelAiTaskPaneHost : UserControl
    {
        private const string AssetOrigin = "https://appassets.local";
        internal static dynamic CurrentExcelApplication;
        private readonly dynamic excelApplication;
        private readonly WebView2 webView;
        private SidebarBridge sidebarBridge;

        public ExcelAiTaskPaneHost()
            : this((object)CurrentExcelApplication)
        {
        }

        public ExcelAiTaskPaneHost(dynamic excelApplication)
        {
            this.excelApplication = excelApplication ?? throw new ArgumentNullException(nameof(excelApplication));
            DiagnosticsLog.Write("Task pane host constructor entered");
            Text = "JYYJ助手";
            Dock = DockStyle.Fill;
            webView = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(webView);
            webView.CoreWebView2InitializationCompleted += OnInitializationCompleted;
            Load += OnLoaded;
            DiagnosticsLog.Write("Task pane host constructor completed");
        }

        private async void OnLoaded(object sender, EventArgs eventArgs)
        {
            try
            {
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAiAssistant",
                    "WebView2",
                    IntPtr.Size == 4 ? "x86" : "x64");
                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
                await webView.EnsureCoreWebView2Async(environment);
            }
            catch (Exception exception) when (
                exception is WebView2RuntimeNotFoundException ||
                exception is FileNotFoundException ||
                exception is InvalidOperationException ||
                exception is TimeoutException)
            {
                MessageBox.Show(
                    this,
                    $"无法初始化 WebView2：{exception.Message}",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OnInitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs eventArgs)
        {
            DiagnosticsLog.Write($"WebView2 initialization completed success={eventArgs.IsSuccess}");
            if (!eventArgs.IsSuccess)
            {
                MessageBox.Show(
                    this,
                    $"WebView2 初始化失败：{eventArgs.InitializationException?.Message ?? "未知错误"}",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                DiagnosticsLog.Write($"WebView2 initialization exception={eventArgs.InitializationException}");
                return;
            }

            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            webView.CoreWebView2.Settings.IsZoomControlEnabled = false;
            webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            webView.CoreWebView2.AddWebResourceRequestedFilter(
                $"{AssetOrigin}/*",
                CoreWebView2WebResourceContext.All);
            webView.CoreWebView2.WebResourceRequested += OnWebResourceRequested;
            sidebarBridge = new SidebarBridge(excelApplication, webView);
            webView.CoreWebView2.AddHostObjectToScript("excelAiBridge", sidebarBridge);
            webView.Source = new Uri($"{AssetOrigin}/sidebar");
            DiagnosticsLog.Write($"WebView2 source assigned uri={webView.Source}");
        }

        private void OnWebResourceRequested(object sender, CoreWebView2WebResourceRequestedEventArgs eventArgs)
        {
            if (!string.Equals(eventArgs.Request.Uri, $"{AssetOrigin}/sidebar", StringComparison.OrdinalIgnoreCase))
            {
                DiagnosticsLog.Write($"Web resource ignored uri={eventArgs.Request.Uri}");
                return;
            }

            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(
                "ExcelAiAssistant.Addin.Sidebar.index.html");
            DiagnosticsLog.Write($"Sidebar resource requested uri={eventArgs.Request.Uri} found={stream != null}");
            var memoryStream = new MemoryStream();
            stream?.CopyTo(memoryStream);
            memoryStream.Position = 0;
            eventArgs.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(
                memoryStream,
                200,
                "OK",
                "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
            DiagnosticsLog.Write("Sidebar resource response created");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && webView != null)
            {
                webView.CoreWebView2InitializationCompleted -= OnInitializationCompleted;
                if (webView.CoreWebView2 != null)
                {
                    webView.CoreWebView2.WebResourceRequested -= OnWebResourceRequested;
                    if (sidebarBridge != null)
                    {
                        webView.CoreWebView2.RemoveHostObjectFromScript("excelAiBridge");
                        sidebarBridge.Dispose();
                        sidebarBridge = null;
                    }
                }

                webView.Dispose();
            }

            base.Dispose(disposing);
            if (CurrentExcelApplication == excelApplication)
            {
                CurrentExcelApplication = null;
            }
        }
    }
}


