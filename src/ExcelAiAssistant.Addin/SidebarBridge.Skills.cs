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
        public Task<string> SinkVbaTool(string name, string description, string code)
        {
            return ExecuteAsync(() =>
            {
                var outcome = SinkVbaCore(name, description, code);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = outcome.Ok,
                    name = outcome.Name,
                    renamed = outcome.Ok && outcome.RenamedFrom != null
                        ? new { from = outcome.RenamedFrom, to = outcome.RenamedTo }
                        : null,
                    reused_scope = outcome.ReusedScope,
                    error = outcome.Error,
                    message = outcome.Message
                }));
            });
        }

        private sealed class SinkVbaOutcome
        {
            public bool Ok;
            public string Name;
            public string RenamedFrom;
            public string RenamedTo;
            public bool ReusedScope;
            public string Error;
            public string Message;
        }

        /// <summary>沉淀核心：把最近的 VBA 代码存为自定义工具，供手动按钮与聊天自动沉淀共用。
        /// 永不抛异常——任何失败都折叠为 SinkVbaOutcome 返回，保证调用方（自动沉淀/手动按钮）都不会被拖垮。</summary>
        private SinkVbaOutcome SinkVbaCore(string name, string description, string code)
        {
            try
            {
                var codeText = string.IsNullOrWhiteSpace(code) ? (_lastVbaSnippet ?? "") : code;
                if (string.IsNullOrWhiteSpace(codeText))
                {
                    return new SinkVbaOutcome { Ok = false, Error = "没有可沉淀的 VBA 代码，请先成功执行一项操作。" };
                }

                // 自动把硬编码作用域改写为「选择要处理的数据区域」的可复用模板（同预置模板形态）。
                codeText = ExcelAiAssistant.Core.Tools.VbaReusabilizer.MakeReusable(codeText, out var reusedVar);

                var store = new CustomToolStore(CustomToolStore.DefaultRoot());
                var baseName = string.IsNullOrWhiteSpace(name) ? AutoToolName(codeText) : name.Trim();
                var nameText = store.FindAvailableName(baseName); // 名字去重，绝不静默覆盖
                var descText = string.IsNullOrWhiteSpace(description)
                    ? BuildToolDescription(codeText, nameText)
                    : description.Trim();
                var result = store.Save(nameText, descText, codeText);
                if (result.Ok)
                {
                    _lastVbaSnippet = null; // 已沉淀后清空，避免重复沉淀旧代码；失败则保留供重试
                }

                return new SinkVbaOutcome
                {
                    Ok = result.Ok,
                    Name = result.Ok ? nameText : null,
                    RenamedFrom = result.Ok && !string.Equals(baseName, nameText, StringComparison.OrdinalIgnoreCase) ? baseName : null,
                    RenamedTo = result.Ok && !string.Equals(baseName, nameText, StringComparison.OrdinalIgnoreCase) ? nameText : null,
                    ReusedScope = result.Ok && reusedVar,
                    Error = result.Ok ? null : (result.Data?.ToString() ?? result.ErrorCode),
                    Message = result.Ok ? (result.Data + (reusedVar ? "（已把硬编码作用域自动改写为“选择要处理的数据区域”，可直接复用）" : "")) : null
                };
            }
            catch (Exception ex)
            {
DiagnosticsLog.Write("[LogCover] SidebarBridge.Skills.cs:104  {ex.GetType().Name}: {ex.Message}");
                // 沉淀是「锦上添花」：目录不可建 / 磁盘不可写 / 名称异常等一律折叠为失败结果，
                // 绝不把异常抛给调用方（否则自动沉淀会拖垮整轮 AI 回复）。
                return new SinkVbaOutcome { Ok = false, Error = "沉淀失败：" + ex.Message };
            }
        }

        /// <summary>导入技能包（zip）。</summary>
        public Task<string> ImportSkillPack(string path)
        {
            return skillPackInstaller.ImportSkillPack(path);
        }

        /// <summary>从文本安装技能（SKILL.md 原文）。</summary>
        public Task<string> InstallSkillFromText(string markdown)
        {
            return skillPackInstaller.InstallSkillFromText(markdown);
        }

        /// <summary>从 URL 安装技能（下载 zip/md 并安装）。</summary>
        public Task<string> InstallSkillFromUrl(string url)
        {
            return skillPackInstaller.InstallSkillFromUrl(url);
        }

        /// <summary>从仓库坐标安装技能（owner/repo、owner/repo/技能名、GitHub 链接、skills.sh 技能页链接；AI Skill P1）。</summary>
        public Task<string> InstallSkillFromRepo(string input)
        {
            return skillPackInstaller.InstallSkillFromRepo(input);
        }

        /// <summary>删除用户技能（内置技能拒绝）。</summary>
        public Task<string> DeleteSkill(string name)
        {
            return skillPackInstaller.DeleteSkill(name);
        }

        public Task<string> GetSendScope()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(ComputeSendScope())));
        }

        public Task<string> GetMemory()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                preferences = memoryStore.Preferences,
                notes = memoryStore.Notes
            })));
        }

        public Task<string> SetMemory(string preferences, string notes)
        {
            return ExecuteAsync(() =>
            {
                memoryStore.Preferences = preferences ?? string.Empty;
                memoryStore.Notes = notes ?? string.Empty;
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true }));
            });
        }

    }
}
