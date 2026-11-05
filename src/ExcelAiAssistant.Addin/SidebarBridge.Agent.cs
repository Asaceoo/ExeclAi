using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using ExcelAiAssistant.Core.Search;
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
        public Task<string> SendPrompt(string prompt)
        {
            var preview = prompt == null ? "<null>" : (prompt.Length > 60 ? prompt.Substring(0, 60) + "…" : prompt);
            DiagnosticsLog.Write("SendPrompt entered preview=" + preview);
            return ExecuteAsync(async () =>
            {
                if (string.IsNullOrWhiteSpace(prompt))
                {
                    throw new ArgumentException("提示词不能为空。", nameof(prompt));
                }

                var directive = TryHandleWorkflowDirective(prompt);
                if (directive != null)
                {
                    return directive;
                }

                var profile = RequireProfile();
                var thisTurnLabel = CurrentTurnLabel();
                var registry = CreateRegistry(thisTurnLabel);
                var difficulty = ApplySkillDirective(registry, prompt, out var skillBody, out var skillSource);

                // 联网搜索（v3.8.1.83，主 A 兜 B）：按设置开关与 baseUrl 探测提供方；
                // ProviderNative 时请求层注入服务端搜索参数，FallbackBing 时由 web_search 工具兜底。
                var searchDetection = SearchProviderDetector.Detect(
                    profile.BaseUrl, agentSettingsStore.Load().WebSearchEnabled);

                var runner = new AgentRunner(
                    new OpenAiCompatibleClient(profile)
                    {
                        SearchDetection = searchDetection
                    },
                    registry,
                    ConfirmTool,
                    new AgentRunnerOptions
                    {
                        MaxToolCallsPerTurn = CurrentMaxToolCalls(),
                        TurnTimeoutSeconds = 3600,
                        ConfirmNormalWrites = true,
                        Trace = (tool, ok, status, detail) =>
                            DiagnosticsLog.Write(
                                string.IsNullOrWhiteSpace(detail)
                                    ? $"工具[{tool}] {status}"
                                    : $"工具[{tool}] {status} {detail}")
                    });

                // 先把用户输入并入会话并立即落盘：即使本轮执行失败/中断，会话内容也绝不丢失。
                conversation.Add(ChatMessage.User(CreateUserPrompt(prompt)));
                TrimConversation();

                var messages = new List<ChatMessage>
                {
                    ChatMessage.System(CreateSystemPrompt(registry, profile, skillBody))
                };
                if (!string.IsNullOrWhiteSpace(activeSummary))
                {
                    messages.Add(ChatMessage.System(activeSummary));
                }
                messages.AddRange(conversation);

                using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(3600));
                activeRunCts?.Cancel();
                activeRunCts?.Dispose();
                activeRunCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutSource.Token);
                AgentRunResult result;
                try
                {
                    result = await runner.RunAsync(messages, activeRunCts.Token,
                        onProgress: p => Publish("agent-progress", ProgressPayload(p)));
                }
                catch (OperationCanceledException)
                {
                    PersistSession(out _);
                    Publish("session-updated", GetActiveSessionObject());
                    return JsonConvert.SerializeObject(new { ok = false, error = "已停止生成。", stopped = true });
                }
                catch (Exception ex)
                {
                    // 工具/模型异常：保留已入会话语料并落盘，避免「失败后对话内容丢失」。
                    DiagnosticsLog.Error("SendPrompt RunAsync 异常: " + ex);
                    PersistSession(out _);
                    Publish("session-updated", GetActiveSessionObject());
                    Publish("tool-error", "本轮执行失败：" + ex.Message);
                    return JsonConvert.SerializeObject(new { ok = false, error = "执行失败：" + ex.Message, lost = false });
                }
                foreach (var toolEvent in result.ToolEvents.Where(item => item.Result?.Ok == false))
                {
                    if (toolEvent.Args != null)
                    {
                        retryableToolArgs[toolEvent.Tool] = toolEvent.Args
                            .ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
                    }

                    Publish("tool-error", $"{toolEvent.Tool}: {FriendlyError.For(toolEvent.Result.ErrorCode, toolEvent.Result.ErrorCode)}");
                }

                lastToolRecords.Clear();
                lastToolRecords.AddRange(result.ToolEvents
                    .Where(item => item.Result?.Ok == true && !string.IsNullOrWhiteSpace(item.Tool)));

                // 自检可观测性：本轮元监控摘要落库（趋势数据）并实时推给前端，让用户看得见 AI 自检了什么。
                if (!string.IsNullOrWhiteSpace(result.AutoVerifySummary))
                {
                    try
                    {
                        SelfCheckHistory?.Append(new SelfCheckRecord
                        {
                            TimestampLocal = DateTime.Now,
                            Kind = "turn",
                            Healthy = result.AutoVerifyExecuted > 0,
                            Ops = result.AutoVerifyOps,
                            Executed = result.AutoVerifyExecuted,
                            Skipped = result.AutoVerifySkipped,
                            Summary = result.AutoVerifySummary
                        });
                        Publish("selfcheck-summary", new
                        {
                            summary = result.AutoVerifySummary,
                            ops = result.AutoVerifyOps,
                            executed = result.AutoVerifyExecuted,
                            skipped = result.AutoVerifySkipped,
                            healthy = result.AutoVerifyExecuted > 0
                        });
                    }
                    catch (Exception exStore)
                    {
                        DiagnosticsLog.Error("自检历史落库失败：" + exStore.Message);
                    }
                }

                foreach (var ev in result.ToolEvents)
                {
                    if (ev.Result?.Ok == true && ev.Tool == "run_vba_snippet"
                        && ev.Args != null && ev.Args.TryGetValue("snippet", out var code)
                        && !string.IsNullOrWhiteSpace(Convert.ToString(code)))
                    {
                        _lastVbaSnippet = Convert.ToString(code);
                    }
                }

                // 把本轮每个工具（含写入数据 / VBA 代码 / 图表）的成败与关键信息记入运行日志，
                // 供设置页「运行日志」核对 AI 生成的数据 / VBA 代码 / 图表是否正确。
                foreach (var ev in result.ToolEvents)
                {
                    if (string.IsNullOrWhiteSpace(ev.Tool))
                    {
                        continue;
                    }

                    try
                    {
                        var ok = ev.Result?.Ok == true;
                        var detail = string.Empty;
                        if (ok)
                        {
                            detail = RenderedToolSummary(ev) ?? string.Empty;
                        }
                        else
                        {
                            detail = ev.Result?.ErrorCode ?? string.Empty;
                            // 失败时把真实错误正文并入同一行日志（如 VBA_RUN_ERROR VBA 运行时错误 1004: ...），
                            // 仅记 ErrorCode 让设置页日志无法看出真实失败原因。
                            if (ev.Result?.Data is string errText && !string.IsNullOrWhiteSpace(errText) && !detail.Contains(errText))
                            {
                                detail = detail + (string.IsNullOrEmpty(detail) ? "" : " ") + CompactLogText(errText, 300);
                            }
                        }

                        if ((ev.Tool == "run_vba_snippet" || ev.Tool == "run_python_code") && ev.Args != null
                            && ev.Args.TryGetValue(ev.Tool == "run_vba_snippet" ? "snippet" : "code", out var codeToken)
                            && !string.IsNullOrWhiteSpace(Convert.ToString(codeToken)))
                        {
                            var codeText = Convert.ToString(codeToken);
                            var head = SummarizeSnippet(codeText);
                            detail = string.IsNullOrWhiteSpace(detail)
                                ? head
                                : detail + " 代码=" + head;
                            // 完整代码落日志，保证“数据被清空/结果不符”类问题可事后从日志还原真实注入内容定位根因。
                            DiagnosticsLog.Write(LogLevel.Info,
                                $"[{ev.Tool.ToUpperInvariant()}-SNIPPET] {head}\n" + TruncateCode(codeText));
                        }
                        else if (ev.Tool == "create_excel_chart" && ev.Args != null
                                 && ev.Args.TryGetValue("name", out var chartName)
                                 && !string.IsNullOrWhiteSpace(Convert.ToString(chartName)))
                        {
                            detail = "图表=" + Convert.ToString(chartName) + (string.IsNullOrWhiteSpace(detail) ? "" : " " + detail);
                        }

                        DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error,
                            $"tool:{ev.Tool} ok={ok} {detail}");
                    }
                    catch (Exception)
                    {
DiagnosticsLog.Write("[LogCover] SidebarBridge.Agent.cs:228  silent-catch covered (no exception variable)");
                        // 日志绝不阻断主流程
                    }
                }

                var answer = result.Content;
                if (result.ReachedToolLimit && string.IsNullOrWhiteSpace(answer))
                {
                    answer = "已达单轮工具调用上限，请继续或缩小任务范围。";
                }

                ChatMessage answerMessage = null;
                if (!string.IsNullOrWhiteSpace(answer))
                {
                    answerMessage = ChatMessage.Assistant(answer);
                    answerMessage.TurnLabel = thisTurnLabel;
                    answerMessage.Reasoning = result.Reasoning;
                    answerMessage.Steps = result.Steps;
                    conversation.Add(answerMessage);
                }

                TrimConversation();

                var toolViews = result.ToolEvents
                    .Select(ToolEventViewBuilder.Build)
                    .Where(item => item != null)
                    .Cast<ToolEventView>()
                    .ToArray();

                if (toolViews.Length > 0)
                {
                    if (string.IsNullOrWhiteSpace(answer))
                    {
                        // 仅有工具调用、无正文回复：新增一条空助手消息承载工具卡片
                        var eventOnlyMsg = ChatMessage.Assistant(string.Empty);
                        eventOnlyMsg.TurnLabel = thisTurnLabel;
                        eventOnlyMsg.Reasoning = result.Reasoning;
                        eventOnlyMsg.Steps = result.Steps;
                        eventOnlyMsg.ToolEvents = toolViews;
                        conversation.Add(eventOnlyMsg);
                    }
                    else
                    {
                        // 把工具卡片挂到本轮已添加的正文助手消息上，避免产生重复 assistant
                        var lastAssistant = conversation.LastOrDefault(message => message.Role == "assistant");
                        if (lastAssistant != null)
                        {
                            lastAssistant.ToolEvents = toolViews;
                        }
                    }
                }

                // 自动沉淀兜底：本轮成功执行了 run_vba_snippet 但未调用 save_vba_tool 时，系统自动把代码
                // 沉淀为可复用工具——不依赖模型是否记得调用（模型常跳过，导致聊天里执行的 VBA 进不了模板库）。
                // 沉淀是「锦上添花」：整段（含标志计算与徽标赋值）统一 try/catch，任何异常都只记日志，绝不拖垮主回复。
                try
                {
                    var vbaSinkedThisTurn = result.ToolEvents.Any(ev =>
                        ev.Result?.Ok == true && string.Equals(ev.Tool, "save_vba_tool", StringComparison.OrdinalIgnoreCase));
                    if (ExcelAiAssistant.Core.Tools.VbaAutoSinkPolicy.ShouldAutoSink(result.ToolEvents, _lastVbaSnippet))
                    {
                        var sink = SinkVbaCore(null, null, null);
                        if (sink.Ok)
                        {
                            vbaSinkedThisTurn = true;
                            DiagnosticsLog.Write(LogLevel.Info, $"auto-sink-vba ok name={sink.Name} reused_scope={sink.ReusedScope}");
                            conversation.Add(ChatMessage.System("本轮成功执行的 VBA 代码已自动沉淀到 VBA 代码库（模板“" + sink.Name + "”），可在「VBA」面板的模板库中查看与复用。"));
                        }
                        else
                        {
                            DiagnosticsLog.Write(LogLevel.Warn, "auto-sink-vba failed: " + sink.Error);
                        }
                    }

                    // 无论 AI 主动调 save_vba_tool 成功，还是系统自动沉淀成功，本轮 VBA 都已进模板库，
                    // 前端据此显示「已自动沉淀」徽标，不再显示会生成重复工具的「沉淀为工具」按钮。
                    if (vbaSinkedThisTurn)
                    {
                        var sinkedAssistant = conversation.LastOrDefault(message => message.Role == "assistant");
                        if (sinkedAssistant != null)
                        {
                            sinkedAssistant.AutoSinkedVba = true;
                        }
                    }
                }
                catch (Exception sinkException)
                {
                    DiagnosticsLog.Write(LogLevel.Error, "auto-sink-vba exception: " + sinkException.Message);
                }

                int assignedTurn; 
                PersistSession(out assignedTurn);
                TryAutoTitleSession(prompt);
                _ = difficulty;
                return JsonConvert.SerializeObject(new
                {
                    ok = true,
                    answer = answer ?? string.Empty,
                    used_skill = string.IsNullOrWhiteSpace(skillBody) ? null : skillBody.Trim(),
                    used_skill_source = skillSource,
                    turn_count = assignedTurn,
                    context_restored = false,
                    token_estimate = result.EstimatedTokens,
                    usage = new
                    {
                        prompt_tokens = result.PromptTokens,
                        completion_tokens = result.CompletionTokens,
                        total_tokens = result.TotalTokens
                    },
                    auto_degraded = result.AutoDegraded,
                    tool_events = ProjectToolEvents(toolViews),
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                });
            });
        }

        public Task<string> RetryTool(string tool)
        {
            return ExecuteAsync(async () =>
            {
                if (string.IsNullOrWhiteSpace(tool))
                {
                    throw new ArgumentException("工具名不能为空。", nameof(tool));
                }

                if (!retryableToolArgs.TryGetValue(tool, out var args) || args == null)
                {
                    throw new InvalidOperationException($"工具「{tool}」没有可重试的参数。");
                }

                var registry = CreateRegistry(CurrentTurnLabel());
                if (!registry.TryGet(tool, out var descriptor))
                {
                    throw new InvalidOperationException($"未知工具：{tool}");
                }

                var envelope = new ToolEnvelope
                {
                    Tool = tool,
                    Args = args,
                    CallId = "retry-" + Guid.NewGuid().ToString("N"),
                    Reason = "用户点击了重试"
                };
                var outcome = await ConfirmTool(envelope, descriptor.RiskLevel);
                if (outcome != ConfirmationOutcome.Confirmed)
                {
                    return JsonConvert.SerializeObject(new { ok = false, error = "确认已取消或超时，未重试。", cancelled = true });
                }

                var result = descriptor.Execute(args);
                if (result.Ok)
                {
                    retryableToolArgs.Remove(tool);
                }

                return JsonConvert.SerializeObject(new
                {
                    ok = result.Ok,
                    tool,
                    error_code = result.Ok ? null : result.ErrorCode,
                    error = result.Ok ? null : (result.Data?.ToString() ?? result.ErrorCode),
                    summary = result.Ok ? SummarizeToolResult(result) : null
                });
            });
        }

        // ===================== 流程编排（P5-B） =====================

        /// <summary>解析聊天中的 @workflow:&lt;名称&gt; 指令；命中则执行并返回序列化响应，否则返回 null。</summary>
    }
}
