using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 「JYYJ快捷」参数框：独立静态 HTML 弹窗（内嵌 ParameterForm.html，离线可用），不复用侧边栏会话。
    /// 下拉/多选由 C# 注入当前选区列名；确认/取消为原生按钮，取消零执行零变更。
    /// 展示当前选区内参数；确认时经 WebView2 执行 collect() 采集值。
    /// </summary>
    public sealed class RibbonParameterDialog : Form
    {
        private const string ParamJsonToken = "__PARAM_JSON__";
        private readonly string title;
        private readonly string address;
        private readonly IReadOnlyList<ParamField> fields;
        private WebView2 webView;
        private Label overlayHintLabel;
        private Label dangerHintLabel;
        private Button confirmButton;

        /// <summary>确认时采集到的字段值（name → string 或 List&lt;string&gt;）。</summary>
        public Dictionary<string, object> Values { get; private set; }

        /// <summary>置顶 + 可自由缩放 + 模态显示在 UI 线程（Ribbon onAction 本就在 UI 线程）。</summary>
        public RibbonParameterDialog(string title, string address, IReadOnlyList<ParamField> fields, string overlayHint, string dangerHint)
        {
            this.title = title ?? "参数设置";
            this.address = address ?? string.Empty;
            this.fields = fields ?? new List<ParamField>();
            Text = this.title;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            // 可自由缩放：Sizable 边框 + 允许最大化/最小化；拖拽边缘或点最大化即可放大，避免字段多时内容截断。
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            MinimumSize = new Size(440, 420);
            ShowInTaskbar = true;
            ClientSize = new Size(600, 640);
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.White;
            Load += OnLoaded;
            BuildLayout(overlayHint, dangerHint);
        }

        private void BuildLayout(string overlayHint, string dangerHint)
        {
            webView = new WebView2 { Dock = DockStyle.Fill, Location = new Point(0, 0) };
            Controls.Add(webView);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 78, BackColor = Color.White };
            Controls.Add(bottom);

            if (!string.IsNullOrWhiteSpace(overlayHint))
            {
                overlayHintLabel = new Label
                {
                    Text = overlayHint,
                    AutoSize = true,
                    Location = new Point(14, 6),
                    Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(196, 152, 32) // 黄字
                };
                bottom.Controls.Add(overlayHintLabel);
            }

            if (!string.IsNullOrWhiteSpace(dangerHint))
            {
                dangerHintLabel = new Label
                {
                    Text = dangerHint,
                    AutoSize = true,
                    Location = new Point(14, overlayHintLabel == null ? 6 : 30),
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(196, 43, 28) // 红字二次警示
                };
                bottom.Controls.Add(dangerHintLabel);
            }

            var cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Size = new Size(96, 32),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(ClientSize.Width - 216, 38)
            };
            cancelButton.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            bottom.Controls.Add(cancelButton);

            confirmButton = new Button
            {
                Text = "确认执行",
                Size = new Size(108, 32),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(ClientSize.Width - 112, 38),
                BackColor = Color.FromArgb(43, 87, 154),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            confirmButton.Click += OnConfirm;
            bottom.Controls.Add(confirmButton);
            AcceptButton = confirmButton;
            CancelButton = cancelButton;
        }

        private async void OnLoaded(object sender, EventArgs eventArgs)
        {
            try
            {
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAiAssistant",
                    "WebView2",
                    "ribbon",
                    IntPtr.Size == 4 ? "x86" : "x64");
                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
                await webView.EnsureCoreWebView2Async(environment);
                if (!IsDisposed && webView.CoreWebView2 != null)
                {
                    webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    webView.CoreWebView2.Settings.IsZoomControlEnabled = false;
                    webView.CoreWebView2.Settings.IsStatusBarEnabled = false;
                    webView.CoreWebView2.NavigateToString(BuildHtml());
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"RibbonParameterDialog init failed: {exception.GetType().Name}: {exception.Message}");
                MessageBox.Show(this, $"参数框初始化失败：{exception.Message}\n\n请重试；不影响 Excel 使用。", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        private string BuildHtml()
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelAiAssistant.Addin.Ribbon.ParameterForm.html");
            if (stream == null)
            {
                return "<html><body>参数模板缺失。</body></html>";
            }

            using var reader = new StreamReader(stream);
            var template = reader.ReadToEnd();
            var payload = new
            {
                title,
                address,
                fields
            };
            // 必须 camelCase：前端 on() 按 f.type/f.label/f.name/f.options 读取，而 ParamField/ParamOption 是
            // PascalCase 属性。若用默认序列化，fields 属性名全大写 → JS 读到 undefined → 四个类型分支全失配 →
            // 参数框 body 空白，而 Collect() 空转用默认值兜底，掩盖了"用户看不到参数"这个缺陷。
            var jsonSettings = new Newtonsoft.Json.JsonSerializerSettings
            {
                ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
            };
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(payload, Newtonsoft.Json.Formatting.None, jsonSettings);
            return template.Replace(ParamJsonToken, json);
        }

        /// <summary>采集并校验参数；确认后须返回正常字典，否则提示不关闭。</summary>
        private async void OnConfirm(object sender, EventArgs eventArgs)
        {
            try
            {
                if (webView == null || webView.CoreWebView2 == null)
                {
                    MessageBox.Show(this, "参数框尚未就绪，请稍后再点“确认执行”。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 不能 Stringify：WebView2.ExecuteScriptAsync 对返回「字符串」的结果会再套一层引号
                // （返回 JSON 编码的字符串字面量），导致 JObject.Parse 拿到字符串而非对象而报
                // 「Current JsonReader item is not an object: String」。collect() 返回对象，直接让
                // ExecuteScriptAsync 按对象序列化，得到 {ok,values} 顶层对象。
                var raw = await webView.CoreWebView2.ExecuteScriptAsync("collect()");
                var jo = JObject.Parse(raw);
                // ok 缺失/null/非布尔（含字符串 "false"）统一视为未通过，避免静默读走脏数据
                if (!(jo.Value<bool?>("ok") ?? false))
                {
                    return;
                }

                var values = new Dictionary<string, object>();
                var obj = jo["values"] as JObject;
                if (obj != null)
                {
                    foreach (var kv in obj)
                    {
                        values[kv.Key] = kv.Value is JArray arr
                            ? new List<string>(arr.ToObject<string[]>())
                            : (object)kv.Value.ToString();
                    }
                }

                Values = values;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"RibbonParameterDialog collect failed: {exception}");
                MessageBox.Show(this, "读取参数失败：" + exception.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing && webView != null)
                {
                    webView.Dispose();
                    webView = null;
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }
    }
}