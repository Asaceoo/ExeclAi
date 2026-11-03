using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Skills;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// SidebarBridge 分域 partial（阶段 A 机械拆分，见 docs/AI交接文档.md）。
    /// 公共 API 即 WebView2 前端契约（hostObjects.excelAiBridge.*），签名不得改动。
    /// </summary>
    public sealed partial class SidebarBridge : IDisposable
    {
        private string TryHandleWorkflowDirective(string prompt)
        {
            var match = Regex.Match(prompt, @"^\s*(?:@workflow|@流程|#流程)\s*[:：]\s*([^ \t\r\n]+)\s*(.*)$");
            if (!match.Success)
            {
                return null;
            }

            var name = match.Groups[1].Value.Trim();
            var overrideText = match.Groups[2].Value.Trim();
            var overrides = workflowManager.ParseOverrides(overrideText);
            var run = workflowManager.ExecuteWorkflowInternal(name, overrides, 0, null);
            var answer = workflowManager.BuildWorkflowAnswer(run);
            return JsonConvert.SerializeObject(new
            {
                ok = run.Completed,
                answer,
                used_skill = (string)null,
                turn_count = conversation.Count / 2 + 1,
                run = workflowManager.BuildRunPayload(run),
                session = GetActiveSessionObject(),
                sessions = ListSessionObjects()
            });
        }

        /// <summary>流程清单（内置 + 用户自定义）。</summary>
        public Task<string> ListWorkflows()
        {
            return workflowManager.ListWorkflows();
        }

        /// <summary>预览当前会话可提取为流程的工具步骤（不保存）。</summary>
        public Task<string> PreviewWorkflowSteps()
        {
            return workflowManager.PreviewWorkflowSteps();
        }

        /// <summary>另存当前会话为流程模板。</summary>
        public Task<string> SaveWorkflow(string name, string description, bool deScope = true)
        {
            return workflowManager.SaveWorkflow(name, description, deScope);
        }

        /// <summary>脱域流程模板（硬编码作用域改写为选区引用）。</summary>
        public Task<string> DeScopeWorkflow(string name)
        {
            return workflowManager.DeScopeWorkflow(name);
        }

        /// <summary>删除流程模板。</summary>
        public Task<string> DeleteWorkflow(string name)
        {
            return workflowManager.DeleteWorkflow(name);
        }

        /// <summary>运行流程模板。</summary>
        public Task<string> RunWorkflow(string name, string overridesJson)
        {
            return workflowManager.RunWorkflow(name, overridesJson);
        }

        /// <summary>单步重试：仅重跑失败步骤及其后步骤。</summary>
        public Task<string> RetryWorkflowStep(string name, int index)
        {
            return workflowManager.RetryWorkflowStep(name, index);
        }

        /// <summary>工作簿绑定可视化（C2）：返回当前活动工作簿与外部只读句柄清单，供侧边栏切换上下文锚点。</summary>
        public Task<string> ListConnectedBooks()
        {
            return ExecuteAsync(() =>
            {
                dynamic active = null;
                string activeHash = null;
                string activeName = null;
                string activePath = null;
                try
                {
                    active = excelApplication.ActiveWorkbook;
                    if (active != null)
                    {
                        activeName = Convert.ToString(active.Name);
                        activePath = Convert.ToString(active.FullName);
                        activeHash = ExcelToolHost.CreateWorkbookHash(activePath);
                    }
                }
                catch { }

                var external = new List<object>();
                try
                {
                    foreach (var info in externalBooks.List())
                    {
                        external.Add(new
                        {
                            handle = info.Token,
                            name = info.Name,
                            path = info.Path,
                            sheets = info.Sheets
                        });
                    }
                }
                catch { }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    active = new { name = activeName, path = activePath, hash = activeHash },
                    external = external.ToArray()
                }));
            });
        }

        /// <summary>停止当前生成。</summary>
        public Task<string> StopGeneration()
        {
            activeRunCts?.Cancel();
            return Task.FromResult("{\"ok\":true}");
        }

    }
}
