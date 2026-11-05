using System;
using System.Drawing;
using System.Media;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Agent;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 原生置顶模态确认框：替代「仅侧边栏内 overlay」的确认方式。置顶 + 任务栏闪烁 + 系统提示音，
    /// 保证用户一定能看到；除非用户明确点「允许执行 / 取消」，否则不静默超时（不再 120s 干等误判成
    /// 执行慢/用户取消）。在独立 STA 线程上以模态方式运行，不依赖调用方 UI 线程上下文。
    /// </summary>
    public static class NativeConfirmDialog
    {
        public static ConfirmationOutcome ShowDialog(
            string title,
            string tool,
            string target,
            string reason,
            string riskText,
            string code,
            string language)
        {
            ConfirmationOutcome result = ConfirmationOutcome.TimedOut;
            Exception caught = null;

            var thread = new Thread(() =>
            {
                try
                {
                    using var form = new ConfirmForm(title, tool, target, reason, riskText, code, language);
                    form.ShowDialog();
                    result = form.Outcome;
                }
                catch (Exception ex)
                {
                    caught = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
            thread.Join();

            if (caught != null)
            {
                throw new InvalidOperationException("原生确认框打开失败：" + caught.Message, caught);
            }

            return result;
        }

        private sealed class ConfirmForm : Form
        {
            public ConfirmationOutcome Outcome { get; private set; } = ConfirmationOutcome.TimedOut;

            private readonly string tool;
            private readonly string target;
            private readonly string reason;
            private readonly string code;
            private readonly string language;

            public ConfirmForm(string title, string tool, string target, string reason, string riskText, string code, string language)
            {
                this.tool = tool;
                this.target = target;
                this.reason = reason;
                this.code = code;
                this.language = language;

                Text = title ?? "确认执行";
                StartPosition = FormStartPosition.CenterScreen;
                TopMost = true;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = true;
                ClientSize = new Size(560, codeVisible ? 430 : 250);
                AutoScroll = true;
                Font = new Font("Microsoft YaHei UI", 9F);
                BackColor = Color.White;

                BuildLayout(riskText);
                Shown += (_, __) => FlashWindowEx(Handle);
                SystemSounds.Exclamation.Play();
            }

            private bool codeVisible => !string.IsNullOrWhiteSpace(code);

            private void BuildLayout(string riskText)
            {
                int y = 14;
                var label = new Label
                {
                    Text = $"[{riskText}] {tool}",
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                    AutoSize = true,
                    Location = new Point(16, y),
                    ForeColor = Color.FromArgb(196, 43, 28)
                };
                Controls.Add(label);
                y += 30;

                AddField("操作对象", target, ref y);
                AddField("原因", reason, ref y);

                var tip = new Label
                {
                    Text = "执行前已创建快照，可在历史中回滚。请确认后继续；本框不会自动超时。",
                    AutoSize = true,
                    Location = new Point(16, y),
                    ForeColor = Color.FromArgb(120, 120, 120)
                };
                Controls.Add(tip);
                y += 26;

                if (codeVisible)
                {
                    var head = new Label
                    {
                        Text = (language ?? "代码") + " 预览：",
                        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                        AutoSize = true,
                        Location = new Point(16, y)
                    };
                    Controls.Add(head);
                    y += 24;

                    var box = new TextBox
                    {
                        ReadOnly = true,
                        Multiline = true,
                        ScrollBars = ScrollBars.Both,
                        WordWrap = false,
                        Font = new Font("Consolas", 9F),
                        Location = new Point(16, y),
                        Size = new Size(ClientSize.Width - 32, 150),
                        Text = code,
                        BackColor = Color.FromArgb(247, 247, 247)
                    };
                    Controls.Add(box);
                    y += 160;
                }

                var reject = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Size = new Size(100, 34),
                    Location = new Point(ClientSize.Width - 228, y + 6)
                };
                reject.Click += (_, __) => Outcome = ConfirmationOutcome.Declined;
                Controls.Add(reject);

                var approve = new Button
                {
                    Text = "允许执行",
                    DialogResult = DialogResult.OK,
                    Size = new Size(110, 34),
                    Location = new Point(ClientSize.Width - 112, y + 6),
                    BackColor = Color.FromArgb(16, 113, 234),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                approve.Click += (_, __) => Outcome = ConfirmationOutcome.Confirmed;
                Controls.Add(approve);

                AcceptButton = approve;
                CancelButton = reject;
            }

            private void AddField(string name, string value, ref int y)
            {
                var caption = new Label
                {
                    Text = name + "：",
                    AutoSize = true,
                    Location = new Point(16, y),
                    ForeColor = Color.FromArgb(90, 90, 90)
                };
                Controls.Add(caption);
                var text = string.IsNullOrWhiteSpace(value) ? "—" : value;
                var valueLabel = new Label
                {
                    Text = text,
                    AutoSize = true,
                    Location = new Point(90, y),
                    MaximumSize = new Size(ClientSize.Width - 110, 0)
                };
                Controls.Add(valueLabel);
                y += 24;
            }
        }

        [DllImport("user32.dll")]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        private static void FlashWindowEx(IntPtr hWnd)
        {
            var info = new FLASHWINFO
            {
                cbSize = (uint)Marshal.SizeOf(typeof(FLASHWINFO)),
                hwnd = hWnd,
                dwFlags = 3, // FLASHW_ALL
                uCount = 5,
                dwTimeout = 0
            };
            FlashWindowEx(ref info);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }
    }
}