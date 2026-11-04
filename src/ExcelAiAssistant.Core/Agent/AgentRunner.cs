using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Search;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 工具确认结果三态：Confirmed=用户点「允许执行」；
    /// Declined=用户点「取消」；TimedOut=等待超时系统中止（不是用户主动取消）。
    /// 刻意与 bool 区分，避免「超时」被误报成「用户取消了本次操作」。
    /// </summary>
    public enum ConfirmationOutcome
    {
        Confirmed = 0,
        Declined = 1,
        TimedOut = 2
    }

    public sealed class AgentRunResult
    {
        public string Content { get; set; }

        /// <summary>本轮模型的思维链推理文字（最后一次非空者胜出）。</summary>
        public string Reasoning { get; set; }

        /// <summary>回合结束时的自动自检元监控摘要（null = 本回合无可自检写操作）。</summary>
        public string AutoVerifySummary { get; set; }

        /// <summary>本回合可自检的写操作数（AutoVerifySummary 非 null 时有效）。</summary>
        public int AutoVerifyOps { get; set; }

        /// <summary>本回合实际执行的自检计划数。</summary>
        public int AutoVerifyExecuted { get; set; }

        /// <summary>本回合跳过的自检计划数。</summary>
        public int AutoVerifySkipped { get; set; }

        public List<ToolEventRecord> ToolEvents { get; } =
            new List<ToolEventRecord>();

        /// <summary>按真实执行顺序组织的步骤序列（思考片段 / 工具调用），前端据此交替渲染。</summary>
        public List<AgentStepView> Steps { get; } =
            new List<AgentStepView>();

        public bool ReachedToolLimit { get; set; }

        public bool AutoDegraded { get; set; }

        public int PromptTokens { get; set; }

        public int CompletionTokens { get; set; }

        public int TotalTokens { get; set; }

        public int EstimatedTokens { get; set; }
    }

    public sealed class ToolEventRecord
    {
        public string CallId { get; set; }

        public string Tool { get; set; }

        public IReadOnlyDictionary<string, object> Args { get; set; }

        public ToolResult Result { get; set; }

        public bool Confirmed { get; set; }
    }

    /// <summary>运行过程中的实时进度事件类别，供 SidebarBridge 推送给前端实时渲染「思考过程」。</summary>
    public enum AgentProgressKind
    {
        Reasoning = 0,
        ToolStart = 1,
        ToolResult = 2,

        /// <summary>模型推理的逐字增量（流式实时）。Reasoning 为本条新增的增量片段。</summary>
        ReasoningDelta = 3
    }

    /// <summary>每一跳（模型推理 / 工具调用）的实时进度。</summary>
    public sealed class AgentProgress
    {
        public AgentProgressKind Kind { get; set; }

        public string Tool { get; set; }

        public bool Ok { get; set; }

        public bool Confirmed { get; set; }

        /// <summary>工具结果的简短摘要（成功时通常来自结果 Data）。</summary>
        public string Summary { get; set; }

        /// <summary>工具结果的完整 JSON（供前端点击「详情」展开查看原始数据）。</summary>
        public string Detail { get; set; }

        public string Reasoning { get; set; }

        /// <summary>本回合内工具步骤序号（从 0 起），用于前端把 ToolStart 与 ToolResult 配对。</summary>
        public int ToolIndex { get; set; } = -1;

        /// <summary>本回合已产生的工具步骤总数。</summary>
        public int ToolCount { get; set; }

        /// <summary>
        /// 进度来源命名空间：null/缺省 = agent 回合；"workflow" = run_workflow 内层步骤。
        /// 前端据此区分 ToolStart/ToolResult 的配对键，避免内层流程步骤与外围工具序号互相覆盖。
        /// </summary>
        public string Namespace { get; set; }
    }

    /// <summary>工具执行轨迹回调：由宿主桥接写入运行日志，用于监控「AI 生成了什么数据/哪个工具成功或失败」。</summary>
    public delegate void ToolTracerCallback(string tool, bool ok, string status, string detail);

    public sealed class AgentRunnerOptions
    {
        public int MaxToolCallsPerTurn { get; set; } = 10000;

        public int TurnTimeoutSeconds { get; set; } = 3600;

        public bool ConfirmNormalWrites { get; set; } = true;

        /// <summary>
        /// 执行层自动自检：主写操作成功后由服务端自动调用配套 verify_* 工具回读核对，
        /// 结果作为附加观察消息交给模型与用户。默认开启；模型仍可显式调用 verify_* 做更细断言。
        /// </summary>
        public bool AutoVerify { get; set; } = true;

        /// <summary>可选：工具生命周期回调（开始/成功/失败），用于把工具执行轨迹落到结构化日志。</summary>
        public ToolTracerCallback Trace { get; set; }
    }

    public sealed class AgentRunner
    {
        private readonly IChatCompletionClient client;
        private readonly ToolRegistry toolRegistry;
        private readonly Func<ToolEnvelope, RiskLevel, Task<ConfirmationOutcome>> confirmationHandler;
        private readonly AgentRunnerOptions options;

        public AgentRunner(
            IChatCompletionClient client,
            ToolRegistry toolRegistry,
            Func<ToolEnvelope, RiskLevel, Task<ConfirmationOutcome>> confirmationHandler,
            AgentRunnerOptions options = null)
        {
            this.client = client ?? throw new ArgumentNullException(nameof(client));
            this.toolRegistry = toolRegistry ?? throw new ArgumentNullException(nameof(toolRegistry));
            this.confirmationHandler = confirmationHandler ??
                throw new ArgumentNullException(nameof(confirmationHandler));
            this.options = options ?? new AgentRunnerOptions();
        }

        public async Task<AgentRunResult> RunAsync(
            IReadOnlyList<ChatMessage> messages,
            CancellationToken cancellationToken,
            Action<AgentProgress> onProgress = null)
        {
            if (messages == null)
            {
                throw new ArgumentNullException(nameof(messages));
            }

            var workingMessages = messages.ToList();
            var result = new AgentRunResult();
            var autoVerifyMonitor = new AutoVerifyMonitor();
            using var timeoutSource = new CancellationTokenSource(
                TimeSpan.FromSeconds(options.TurnTimeoutSeconds));
            using var linkedSource = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken,
                timeoutSource.Token);
            var toolDescriptors = toolRegistry.Tools.ToList();
            var executions = 0;
            var toolIndex = 0;
            var turnWatch = System.Diagnostics.Stopwatch.StartNew();
            AppLog.Info("agent", $"回合开始 工具上限={options.MaxToolCallsPerTurn} 空闲超时={options.TurnTimeoutSeconds}s");

            try
            {
            while (true)
            {
                // 空闲超时（每轮重置）：单次模型调用+其工具执行的卡死保护；
                // 不再作为整轮总时长上限——仪表盘等长流程曾被 300s 一刀切中断（用户实证）
                timeoutSource.CancelAfter(TimeSpan.FromSeconds(options.TurnTimeoutSeconds));

                if (executions >= options.MaxToolCallsPerTurn)
                {
                    result.ReachedToolLimit = true;
                    break;
                }

                var completion = await client.CompleteAsync(
                    workingMessages,
                    toolDescriptors,
                    linkedSource.Token,
                    onProgress);

                if (completion.FallbackUsed)
                {
                    result.AutoDegraded = true;
                    toolDescriptors = new List<ToolDescriptor>();
                }

                result.PromptTokens += completion.PromptTokens;
                result.CompletionTokens += completion.CompletionTokens;
                result.TotalTokens += completion.TotalTokens;
                result.EstimatedTokens = TokenEstimator.Estimate(workingMessages);

                if (onProgress != null && !string.IsNullOrWhiteSpace(completion.Reasoning))
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.Reasoning,
                        Reasoning = completion.Reasoning
                    });
                }

                // 按真实顺序记录思考片段：紧随其后的（若有）工具调用会继续追加到 Steps。
                if (!string.IsNullOrWhiteSpace(completion.Reasoning))
                {
                    result.Steps.Add(new AgentStepView
                    {
                        Type = "reasoning",
                        Text = completion.Reasoning
                    });
                }

                if (completion.ToolCalls.Count > 0)
                {
                    var nativeCalls = completion.ToolCalls.ToList();
                    workingMessages.Add(ChatMessage.AssistantWithToolCalls(nativeCalls));
                    var executedAny = false;
                    var autoVerifies = new List<ToolEventRecord>();
                    foreach (var call in nativeCalls)
                    {
                        if (executions >= options.MaxToolCallsPerTurn)
                        {
                            result.ReachedToolLimit = true;
                            break;
                        }

                        var envelope = new ToolEnvelope
                        {
                            Tool = call.Name,
                            Args = call.Arguments ?? new Dictionary<string, object>(),
                            CallId = call.Id ?? call.Name
                        };
                        var beforeEvs = result.ToolEvents.Count;
                        var executed = await ExecuteToolAsync(envelope, toolRegistry, workingMessages, result, true,
                            onProgress, toolIndex, autoVerifies, autoVerifyMonitor);
                        // ExecuteToolAsync 追加主工具事件，自动自检成功时还会追加一条 verify 事件记录；
                        // 逐条记录到 Steps，保证 Steps 中的顺序就是真实执行顺序。
                        for (var eventIndex = beforeEvs; eventIndex < result.ToolEvents.Count; eventIndex++)
                        {
                            var view = ToolEventView.FromRecord(result.ToolEvents[eventIndex]);
                            result.Steps.Add(new AgentStepView
                            {
                                Type = "tool",
                                ToolEvent = view
                            });
                        }
                        toolIndex++;
                        executions++;
                        if (executed)
                        {
                            executedAny = true;
                        }
                    }

                    FlushAutoVerify(workingMessages, autoVerifies);

                    if (executions >= options.MaxToolCallsPerTurn)
                    {
                        result.ReachedToolLimit = true;
                    }

                    if (!executedAny && result.ReachedToolLimit)
                    {
                        break;
                    }

                    continue;
                }

                if (!AgentProtocolParser.TryParse(completion.Content, out var action))
                {
                    action = new ParsedAgentAction
                    {
                        Action = "final",
                        Content = completion.Content
                    };
                }

                if (!string.IsNullOrWhiteSpace(completion.Reasoning)) { result.Reasoning = completion.Reasoning; }

                if (action.Action == "final")
                {
                    result.Content = action.Content;
                    EmitAutoVerifySummary(result, autoVerifyMonitor);
                    return result;
                }

                var envelope2 = action.ToolCall ??
                                throw new FormatException("tool action 缺少 tool call。");
                var singleAutoVerifies = new List<ToolEventRecord>();
                await ExecuteToolAsync(envelope2, toolRegistry, workingMessages, result, false,
                    onProgress, toolIndex, singleAutoVerifies, autoVerifyMonitor);
                FlushAutoVerify(workingMessages, singleAutoVerifies);
                toolIndex++;
                executions++;
            }

            if (result.ReachedToolLimit && string.IsNullOrEmpty(result.Content))
            {
                // 触顶不再静默戛然而止：向用户明确说明并指路续做
                result.Content = "本轮已连续调用 " + executions + " 次工具，达到单轮上限，已暂停以保护会话。"
                    + "已完成的部分全部保留在工作簿中——请发送「继续」接续任务。";
                AppLog.Info("agent", $"回合触顶中断 工具={executions}");
            }

            // 回合正常结束摘要（日志全覆盖批次）：与触顶/超时/异常中断日志对应，正常路径也留痕，
            // 工具数+耗时+是否触顶，配合工具级 trace 可完整还原一轮执行轨迹。
            AppLog.Info("agent", $"回合正常结束 工具={executions} 耗时={turnWatch.ElapsedMilliseconds}ms 触顶={result.ReachedToolLimit}");
            EmitAutoVerifySummary(result, autoVerifyMonitor);
            return result;
            }
            catch (OperationCanceledException) when (timeoutSource.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
            {
                // 空闲超时中断：绝不静默——把已完成工作与续做指引交给用户
                AppLog.Error("agent", "回合空闲超时中断");
                result.Content = InterruptedContent(result, "处理超时（单步空闲超过 " + options.TurnTimeoutSeconds + " 秒）");
                EmitAutoVerifySummary(result, autoVerifyMonitor);
                return result;
            }
            catch (Exception ex) when (ShouldContain(ex))
            {
                // 模型/网络/解析异常兜底：转成可见错误消息，回复不再无声消失
                AppLog.Error("agent", "回合异常中断: " + ex.GetType().Name + " " + ex.Message);
                result.Content = InterruptedContent(result, "处理异常（" + ex.GetType().Name + "）");
                EmitAutoVerifySummary(result, autoVerifyMonitor);
                return result;
            }
        }

        /// <summary>回合被系统中断时的用户可见文案：说明已完成工作保留、指路续做。</summary>
        private static string InterruptedContent(AgentRunResult result, string reason)
        {
            int toolSteps = 0;
            foreach (var step in result.Steps)
            {
                if (step.Type == "tool") { toolSteps++; }
            }
            return "⚠️ 本轮因" + reason + "被系统中断（并非您取消）。已完成 " + toolSteps +
                " 步工具调用，结果均保留在工作簿中；自检与回滚记录见「运行日志」。请发送「继续」接续任务。";
        }

        /// <summary>回合结束时输出自动自检元监控摘要（进结果对象供会话层展示）；链路疑似失效时记为告警级轨迹，进入运行日志供 AI 研判识别。</summary>
        private void EmitAutoVerifySummary(AgentRunResult result, AutoVerifyMonitor monitor)
        {
            result.AutoVerifySummary = monitor?.BuildSummary();
            if (string.IsNullOrEmpty(result.AutoVerifySummary))
            {
                return;
            }

            result.AutoVerifyOps = monitor.SelfCheckableOps;
            result.AutoVerifyExecuted = monitor.ExecutedPlans;
            result.AutoVerifySkipped = monitor.SkippedPlans;
            var healthy = monitor.ExecutedPlans > 0;
            // 回合级自检覆盖统计落日志：可自检操作数/已执行/跳过，便于离线排查自检链路失效。
            AppLog.Info("verify", $"回合自检汇总 {result.AutoVerifySummary} 可自检={result.AutoVerifyOps} 已执行={result.AutoVerifyExecuted} 跳过={result.AutoVerifySkipped}");
            options.Trace?.Invoke(
                "auto_verify_monitor",
                healthy,
                healthy ? "ok" : "SELF_CHECK_LINK_SUSPECT",
                result.AutoVerifySummary);
        }

        /// <summary>把自检失败结果提炼成一行可读明细：判定 issues 优先，其次错误码+数据摘要。</summary>
        private static string DescribeVerifyFailure(ToolResult verifyResult)
        {
            if (verifyResult == null)
            {
                return "<无结果>";
            }

            try
            {
                if (JToken.FromObject(verifyResult.Data) is JObject data)
                {
                    var issues = data["issues"] as JArray;
                    if (issues != null && issues.Count > 0)
                    {
                        return string.Join("；", issues.Select(t => t.ToString()));
                    }

                    var verdict = data.Value<string>("verdict");
                    if (!string.IsNullOrEmpty(verdict))
                    {
                        return "verdict=" + verdict;
                    }
                }
            }
            catch (Newtonsoft.Json.JsonException)
            {
                AppLog.Warn("AgentRunner", "[LogCover] AgentRunner.cs:413  silent-catch covered (no exception variable)");
                // Data 非 JSON 对象（异常文案字符串），走下方兜底
            }

            var text = verifyResult.Data as string;
            if (!string.IsNullOrEmpty(text))
            {
                return text.Length > 300 ? text.Substring(0, 300) + "…" : text;
            }

            return verifyResult.ErrorCode ?? "<空>";
        }

        private async Task<bool> ExecuteToolAsync(
            ToolEnvelope envelope,
            ToolRegistry registry,
            List<ChatMessage> messages,
            AgentRunResult result,
            bool nativeAppend,
            Action<AgentProgress> onProgress = null,
            int toolIndex = -1,
            List<ToolEventRecord> autoVerifyCollector = null,
            AutoVerifyMonitor autoVerifyMonitor = null)
        {
            if (onProgress != null)
            {
                onProgress(new AgentProgress
                {
                    Kind = AgentProgressKind.ToolStart,
                    Tool = envelope.Tool,
                    ToolIndex = toolIndex,
                    ToolCount = toolIndex + 1
                });
            }

            // 提供方内置搜索函数（如 Kimi builtin_function $web_search）不在注册表：
            // 拦截后走本地 bing 抓取兜底，保证模型拿到真实搜索结果而不是 UNKNOWN_TOOL。
            if (envelope.Tool != null && envelope.Tool.StartsWith("$", StringComparison.Ordinal))
            {
                var builtinArgs = new Dictionary<string, object>(envelope.Args ?? new Dictionary<string, object>());
                if (!builtinArgs.ContainsKey("query") && builtinArgs.TryGetValue("search_query", out var searchQuery) && searchQuery != null)
                {
                    builtinArgs["query"] = searchQuery;
                }

                var builtinResult = Search.WebSearchService.Search(builtinArgs);
                var builtinSummary = builtinResult.Ok
                    ? "内置搜索已执行（bing 兜底），返回 " + Search.WebSearchService.CountOf(builtinResult) + " 条结果"
                    : DescribeVerifyFailure(builtinResult);
                options.Trace?.Invoke(envelope.Tool, builtinResult.Ok, builtinResult.Ok ? "ok" : (builtinResult.ErrorCode ?? "failed"), "内置函数 bing 兜底");
                AppendResult(messages, envelope, builtinResult, nativeAppend);
                if (onProgress != null)
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.ToolResult,
                        Tool = envelope.Tool,
                        Ok = builtinResult.Ok,
                        Confirmed = true,
                        Summary = builtinSummary,
                        ToolIndex = toolIndex
                    });
                }
                AppLog.Info("agent", $"内置搜索函数[{envelope.Tool}] bing 兜底 ok={builtinResult.Ok}");
                return builtinResult.Ok;
            }

            if (!registry.TryGet(envelope.Tool, out var tool))
            {
                var failure = new ToolResult
                {
                    Ok = false,
                    ErrorCode = "UNKNOWN_TOOL",
                    Data = $"未知工具：{envelope.Tool}"
                };
                options.Trace?.Invoke(envelope.Tool, false, "UNKNOWN_TOOL", failure.Data as string);
                AppendResult(messages, envelope, failure, nativeAppend);
                if (onProgress != null)
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.ToolResult,
                        Tool = envelope.Tool,
                        Ok = false,
                        Confirmed = true,
                        Summary = failure.Data?.ToString(),
                        ToolIndex = toolIndex
                    });
                }
                return false;
            }

            options.Trace?.Invoke(envelope.Tool, false, "start", null);

            // 确认与执行都必须兜底：任一 COM/运行时异常不得逃逸杀整轮，否则用户只看到一行
            // 「未指定的错误 (HRESULT…E_FAIL)」且日志无记录——必须转成工具级结果交给 agent
            // 继续处理，并写轨迹日志以便监控定位。
            var outcome = ConfirmationOutcome.Confirmed;
            var confirmThrew = false;
            try
            {
                outcome = await confirmationHandler(envelope, tool.RiskLevel);
            }
            catch (Exception ex) when (ShouldContain(ex))
            {
                confirmThrew = true;
                var failure = ExceptionResult(envelope.Tool, ex);
                options.Trace?.Invoke(envelope.Tool, false, "TOOL_EXCEPTION", SafeDetail(ex));
                result.ToolEvents.Add(new ToolEventRecord
                {
                    CallId = envelope.CallId,
                    Tool = envelope.Tool,
                    Args = envelope.Args,
                    Result = failure,
                    Confirmed = false
                });
                AppendResult(messages, envelope, failure, nativeAppend);
                if (onProgress != null)
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.ToolResult,
                        Tool = envelope.Tool,
                        Ok = false,
                        Confirmed = false,
                        Summary = failure.Data?.ToString(),
                        ToolIndex = toolIndex
                    });
                }
            }

            if (confirmThrew)
            {
                return false;
            }

            if (outcome != ConfirmationOutcome.Confirmed)
            {
                // 用户明确取消 与 确认超时（系统自动中止）是两种截然不同的结果，
                // 绝不能把超时误报成「用户取消」——否则用户在等待时会被冤枉。
                var declined = new ToolResult
                {
                    Ok = false,
                    ErrorCode = outcome == ConfirmationOutcome.Declined ? "USER_DENIED" : "CONFIRM_TIMEOUT",
                    Data = outcome == ConfirmationOutcome.Declined
                        ? "用户取消了本次操作。"
                        : "确认超时：等待您选择是否执行已超过时限，本次操作已被系统自动中止（并非用户取消）。请重试，并在弹出的确认框中点「允许执行」以放行。"
                };
                options.Trace?.Invoke(envelope.Tool, false, declined.ErrorCode, declined.Data?.ToString());
                result.ToolEvents.Add(new ToolEventRecord
                {
                    CallId = envelope.CallId,
                    Tool = envelope.Tool,
                    Args = envelope.Args,
                    Result = declined,
                    Confirmed = false
                });
                AppendResult(messages, envelope, declined, nativeAppend);
                if (onProgress != null)
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.ToolResult,
                        Tool = envelope.Tool,
                        Ok = false,
                        Confirmed = false,
                        Summary = declined.Data?.ToString(),
                        ToolIndex = toolIndex
                    });
                }
                return false;
            }

            ToolResult executed;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                executed = tool.Execute(envelope.Args);
                sw.Stop();
                var brief = ShortTrace(executed);
                AppLog.Info("tool", $"工具[{envelope.Tool}] {sw.ElapsedMilliseconds}ms {(executed.Ok ? "ok" : (executed.ErrorCode ?? "failed"))} {(brief.Length > 200 ? brief.Substring(0, 200) + "…" : brief)}");
            }
            catch (Exception ex) when (ShouldContain(ex))
            {
                sw.Stop();
                AppLog.Error("tool", $"工具[{envelope.Tool}] {sw.ElapsedMilliseconds}ms 异常: {SafeDetail(ex)}");
                var failure = ExceptionResult(envelope.Tool, ex);
                options.Trace?.Invoke(envelope.Tool, false, "TOOL_EXCEPTION", SafeDetail(ex));
                result.ToolEvents.Add(new ToolEventRecord
                {
                    CallId = envelope.CallId,
                    Tool = envelope.Tool,
                    Args = envelope.Args,
                    Result = failure,
                    Confirmed = true
                });
                AppendResult(messages, envelope, failure, nativeAppend);
                if (onProgress != null)
                {
                    onProgress(new AgentProgress
                    {
                        Kind = AgentProgressKind.ToolResult,
                        Tool = envelope.Tool,
                        Ok = false,
                        Confirmed = true,
                        Summary = failure.Data?.ToString(),
                        ToolIndex = toolIndex
                    });
                }
                return false;
            }

            options.Trace?.Invoke(envelope.Tool, true,
                executed.Ok ? "ok" : (executed.ErrorCode ?? "failed"),
                ShortTrace(executed));
            result.ToolEvents.Add(new ToolEventRecord
            {
                CallId = envelope.CallId,
                Tool = envelope.Tool,
                Args = envelope.Args,
                Result = executed,
                Confirmed = true
            });
            AppendResult(messages, envelope, executed, nativeAppend);
            if (onProgress != null)
            {
                onProgress(new AgentProgress
                {
                    Kind = AgentProgressKind.ToolResult,
                    Tool = envelope.Tool,
                    Ok = executed.Ok,
                    Confirmed = true,
                    Summary = ToolSummaryBuilder.Summarize(envelope.Tool, executed.Data),
                    Detail = ToolSummaryBuilder.SerializeDetail(executed.Data),
                    ToolIndex = toolIndex
                });
            }

            RunAutoVerify(envelope, registry, executed, result, autoVerifyCollector, autoVerifyMonitor);
            return true;
        }

        /// <summary>
        /// 执行层自动自检：主工具成功且有配套 verify_* 工具（只读）时自动回读核对（批量写入可产生多条计划），
        /// 事件记录进 ToolEvents / Steps（前端可见），观察消息由调用方在整批工具结束后统一下发，
        /// 避免在原生协议同一批 tool_calls 的多条 tool 消息之间插入 user 消息。
        /// </summary>
        private void RunAutoVerify(
            ToolEnvelope envelope,
            ToolRegistry registry,
            ToolResult executed,
            AgentRunResult result,
            List<ToolEventRecord> collector,
            AutoVerifyMonitor monitor)
        {
            if (!options.AutoVerify || collector == null)
            {
                return;
            }

            // M28 审查（第三轮）：主操作失败（如 OVERWRITE_DENIED / 参数错误）不产生自检计划——
            // 否则富化取不到结果会记 skipped，触发器误报「自检链路疑似失效」。
            if (executed == null || !executed.Ok)
            {
                return;
            }

            var plans = new List<AutoVerifyPlan>();
            if (!AutoVerifyPolicy.TryPlanAll(envelope.Tool, envelope.Args, plans) || plans.Count == 0)
            {
                return;
            }

            // 元自检：主操作有自检计划即计入可自检操作；执行/跳过结局进覆盖率统计。
            monitor?.RecordSelfCheckableOp();
            var executedCount = 0;
            var skippedCount = 0;

            foreach (var plan in plans)
            {
                if (!AutoVerifyPolicy.EnrichFromResult(plan, envelope.Tool, executed?.Data, executed?.AffectedRanges) ||
                    !registry.TryGet(plan.VerifyTool, out var verifyTool) ||
                    verifyTool.RiskLevel != RiskLevel.ReadOnly)
                {
                    skippedCount++;
                    continue;
                }

                ToolResult verifyResult = null;
                var verifySw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    verifyResult = verifyTool.Execute(plan.Args);
                }
                catch (Exception ex) when (ShouldContain(ex))
                {
                    verifyResult = ExceptionResult(plan.VerifyTool, ex);
                }
                finally
                {
                    verifySw.Stop();
                    var ok = verifyResult != null && verifyResult.Ok;
                    AppLog.Info("verify", $"自检[{plan.VerifyTool}] {verifySw.ElapsedMilliseconds}ms {(ok ? "ok" : (verifyResult != null ? verifyResult.ErrorCode : "failed"))} {plan.Reason}");
                    if (!ok)
                    {
                        // 自检失败明细（判定 issues/异常文案）落 Error 级日志：定位"哪一步核对没过"不再只靠错误码反推。
                        AppLog.Error("verify", $"自检未通过明细[{plan.VerifyTool}] 主工具={envelope.Tool} 原因={plan.Reason} 明细={DescribeVerifyFailure(verifyResult)}");
                    }
                }

                executedCount++;
                options.Trace?.Invoke(
                    plan.VerifyTool,
                    verifyResult.Ok,
                    "auto:" + (verifyResult.Ok ? (ExtractVerdict(verifyResult) ?? "done") : (verifyResult.ErrorCode ?? "failed")),
                    plan.Reason);
                // 同一条记录进 ToolEvents（前端步骤列表可见）与 collector（整批结束后合并为 auto_verify 观察消息）。
                var record = new ToolEventRecord
                {
                    CallId = envelope.CallId + ":autoverify",
                    Tool = plan.VerifyTool,
                    Args = plan.Args,
                    Result = verifyResult,
                    Confirmed = true
                };
                result.ToolEvents.Add(record);
                collector.Add(record);
            }

            monitor?.RecordOutcome(executedCount, skippedCount);
        }

        /// <summary>把本批自动自检结果合并成一条用户观察消息，让模型与用户都能看到回读核对结论。</summary>
        private static void FlushAutoVerify(List<ChatMessage> messages, List<ToolEventRecord> collector)
        {
            if (collector == null || collector.Count == 0)
            {
                return;
            }

            var payload = collector.Select(record => new
            {
                auto_verify = true,
                tool = record.Tool,
                result = record.Result
            }).ToList();
            messages.Add(ChatMessage.User(JsonConvert.SerializeObject(payload)));
            collector.Clear();
        }

        private static string ExtractVerdict(ToolResult result)
        {
            var data = result?.Data;
            if (data == null)
            {
                return null;
            }

            try
            {
                if (!(JToken.FromObject(data) is JObject token))
                {
                    return null;
                }

                return token.Value<string>("verdict");
            }
            catch (Newtonsoft.Json.JsonException)
            {
                AppLog.Warn("AgentRunner", "[LogCover] AgentRunner.cs:774  silent-catch covered (no exception variable)");
                return null;
        }
        }

        /// <summary>仅兜底真正的异常；取消/内存不足交给上层语义处理。</summary>
        private static bool ShouldContain(Exception ex)
        {
            return !(ex is OperationCanceledException) && !(ex is OutOfMemoryException);
        }

        /// <summary>异常 → 工具级失败结果，补齐运行时细节使 AI 能据此修正重试。</summary>
        private static ToolResult ExceptionResult(string tool, Exception ex)
        {
            return new ToolResult
            {
                Ok = false,
                ErrorCode = "TOOL_EXCEPTION",
                Data = $"工具 {tool} 执行异常，已拦截：{ex.Message}"
            };
        }

        private static string SafeDetail(Exception ex)
        {
            var message = ex.Message;
            if (ex is System.Runtime.InteropServices.COMException com)
            {
                message += $" (HRESULT:0x{unchecked((uint)com.HResult):X8})";
            }

            return message;
        }

        private static string ShortTrace(ToolResult executed)
        {
            if (executed.Ok)
            {
                var data = executed.Data;
                if (data == null)
                {
                    return string.Empty;
                }

                var text = JsonConvert.SerializeObject(data);
                return text.Length <= 200 ? text : text.Substring(0, 200) + "…";
            }

            return (executed.ErrorCode ?? "failed") + ": " + (executed.Data ?? string.Empty);
        }

        private static void AppendResult(
            List<ChatMessage> messages,
            ToolEnvelope envelope,
            ToolResult result,
            bool nativeAppend)
        {
            if (nativeAppend)
            {
                messages.Add(ChatMessage.ToolResult(
                    envelope.CallId ?? envelope.Tool,
                    SerializeObservation(result)));
                return;
            }

            messages.Add(ChatMessage.Assistant(
                SerializeAction("tool", envelope.Tool, envelope.Args, envelope.Reason, envelope.CallId)));
            messages.Add(ChatMessage.User(SerializeObservation(result)));
        }

        private static string SerializeAction(
            string action,
            string tool,
            Dictionary<string, object> args,
            string reason,
            string callId)
        {
            return JsonConvert.SerializeObject(new
            {
                action,
                tool,
                args,
                reason,
                call_id = callId
            });
        }

        private static string SerializeObservation(ToolResult result)
        {
            return JsonConvert.SerializeObject(result);
        }
    }
}
