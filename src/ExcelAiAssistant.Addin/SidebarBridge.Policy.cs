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
        public Task<string> GetControlledExecution()
        {
            return ExecuteAsync(() =>
            {
                var config = new ExcelToolHost(excelApplication).GetControlledExecutionConfig();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    file_path = new ControlledExecutionStore().FilePath,
                    config = new
                    {
                        enabled = config.Enabled,
                        cli_allowlist = config.CliAllowlist,
                        cli_denylist = config.CliDenylist,
                        exe_allowlist = config.ExeAllowlist,
                        exe_denylist = config.ExeDenylist,
                        script_allowlist = config.ScriptAllowlist,
                        script_denylist = config.ScriptDenylist,
                        timeout_seconds = config.TimeoutSeconds,
                        output_limit_kb = config.OutputLimitKb,
                        python_path = config.PythonPath
                    }
                }));
            });
        }

        public Task<string> SetControlledExecution(string configJson)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(configJson))
                {
                    throw new ArgumentException("受控执行配置不能为空。", nameof(configJson));
                }

                var parsed = JsonConvert.DeserializeObject<ControlledExecutionConfig>(configJson)
                             ?? throw new ArgumentException("受控执行配置解析失败。", nameof(configJson));
                new ExcelToolHost(excelApplication).SaveControlledExecution(parsed);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    file_path = new ControlledExecutionStore().FilePath
                }));
            });
        }

        /// <summary>导出受控执行策略到指定 json 路径（C4），返回导出路径。</summary>
        public Task<string> ExportControlledPolicy(string exportPath)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(exportPath))
                {
                    var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    exportPath = Path.Combine(baseDir, "XaiAssistant",
                        "controlled-execution-export-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".json");
                }

                var config = new ExcelToolHost(excelApplication).GetControlledExecutionConfig();
                var dir = System.IO.Path.GetDirectoryName(exportPath);
                if (!string.IsNullOrEmpty(dir))
                {
                    System.IO.Directory.CreateDirectory(dir);
                }

                System.IO.File.WriteAllText(
                    exportPath,
                    JsonConvert.SerializeObject(config, Formatting.Indented),
                    Encoding.UTF8);
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, path = exportPath }));
            });
        }

        /// <summary>从指定 json 路径导入受控执行策略并立即生效（C4）。</summary>
        public Task<string> ImportControlledPolicy(string importPath)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(importPath) || !System.IO.File.Exists(importPath))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "导入文件不存在：" + importPath }));
                }

                try
                {
                    var json = System.IO.File.ReadAllText(importPath, Encoding.UTF8);
                    var parsed = JsonConvert.DeserializeObject<ControlledExecutionConfig>(json)
                                 ?? throw new ArgumentException("受控执行配置解析失败。", nameof(importPath));
                    var store = new ControlledExecutionStore();
                    store.Save(parsed);
                    new ExcelToolHost(excelApplication).SaveControlledExecution(parsed);
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        file_path = store.FilePath,
                        enabled = parsed.Enabled,
                        cli_allow = parsed.CliAllowlist.Length,
                        script_allow = parsed.ScriptAllowlist.Length,
                        exe_allow = parsed.ExeAllowlist.Length
                    }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] SidebarBridge.Policy.cs:137  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "导入失败：" + ex.Message }));
                }
            });
        }

        /// <summary>设置当前 C2 上下文锚点句柄；传空串则清除。校验句柄必须处于打开状态。</summary>
        /// <summary>读取当前写入输出偏好：auto / formula / value。</summary>
        public Task<string> GetWritePreference()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                write_preference = CurrentWriteMode()
            })));
        }

        public Task<string> SetWritePreference(string mode)
        {
            return ExecuteAsync(() =>
            {
                var value = (mode ?? "auto").Trim().ToLowerInvariant();
                if (value != "auto" && value != "formula" && value != "value")
                {
                    throw new ArgumentException("写入偏好只能是 auto / formula / value。");
                }

                new WritePreferenceStore().Save(new WritePreference { ValueMode = value });
                cachedWriteMode = value;
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, write_preference = value }));
            });
        }

        /// <summary>读取当前"单轮工具调用上限"。</summary>
        public Task<string> GetMaxToolCalls()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                max_tool_calls_per_turn = cachedMaxToolCalls
            })));
        }

        /// <summary>设置"单轮工具调用上限"（1–10000），立即持久化并生效。</summary>
        public Task<string> SetMaxToolCalls(int maxToolCalls)
        {
            return ExecuteAsync(() =>
            {
                var clamped = Math.Max(1, Math.Min(10000, maxToolCalls));
                var settings = agentSettingsStore.Load();
                settings.MaxToolCallsPerTurn = clamped;
                agentSettingsStore.Save(settings);
                cachedMaxToolCalls = clamped;
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    max_tool_calls_per_turn = clamped
                }));
            });
        }

        /// <summary>读取联网搜索总开关（v3.8.1.83）。</summary>
        public Task<string> GetWebSearchEnabled()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                web_search_enabled = cachedWebSearchEnabled
            })));
        }

        /// <summary>设置联网搜索总开关，立即持久化并生效。</summary>
        public Task<string> SetWebSearchEnabled(bool enabled)
        {
            return ExecuteAsync(() =>
            {
                var settings = agentSettingsStore.Load();
                settings.WebSearchEnabled = enabled;
                agentSettingsStore.Save(settings);
                cachedWebSearchEnabled = enabled;
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    web_search_enabled = enabled
                }));
            });
        }

        /// <summary>
        /// 手动把当前对话压缩成摘要然后继续：用模型总结现有对话，清空原始消息，
        /// 把摘要作为后续对话的系统上下文保留（可跨重启持久化）。
        /// </summary>
        public Task<string> CompressContext()
        {
            return ExecuteAsync(async () =>
            {
                var profile = RequireProfile();
                if (conversation.Count == 0 && string.IsNullOrWhiteSpace(activeSummary))
                {
                    throw new InvalidOperationException("当前没有可压缩的对话内容。");
                }

                using var client = new OpenAiCompatibleClient(profile);
                using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(120));
                var input = new List<ChatMessage>
                {
                    ChatMessage.System(
                        "你是 Excel JYYJ助手的上下文压缩器。请把下面的对话压缩成一段不超过 300 字的中文摘要，" +
                        "保留：用户意图、涉及的工作表与区域、已完成的关键修改、返回的关键结果、遗留的待办、确定的偏好。" +
                        "不要遗漏重要事实，不要编造信息。只输出摘要本身。")
                };
                if (!string.IsNullOrWhiteSpace(activeSummary))
                {
                    input.Add(ChatMessage.System("已有摘要（请并入新摘要，避免重复）：\n" + activeSummary));
                }

                input.AddRange(conversation);

                var result = await client.CompleteAsync(input, timeoutSource.Token).ConfigureAwait(false);
                var summary = (result?.Content ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(summary))
                {
                    throw new InvalidOperationException("模型没有返回有效摘要，未做压缩。");
                }

                activeSummary = "此前对话摘要：\n" + summary;
                conversation.Clear();

                if (string.IsNullOrWhiteSpace(activeSessionId))
                {
                    activeSessionId = sessionStore.Create(ResolveSessionKey());
                }

                var key = ResolveSessionKey();
                sessionStore.Save(key, activeSessionId, conversation);
                sessionStore.SaveSummary(key, activeSessionId, activeSummary);
                memoryStore.AddRecentSummary(key, summary);

                return JsonConvert.SerializeObject(new
                {
                    ok = true,
                    summary,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                });
            });
        }

        /// <summary>
        /// 公式助手：根据数据选区（可多选）与目标选区、需求，让模型生成一条可复制粘贴的公式及其通俗说明（纯文本生成，不改动工作表）。
        /// dataAreas / targetAreas 用换行或逗号分隔多个区域地址。
        /// </summary>
    }
}
