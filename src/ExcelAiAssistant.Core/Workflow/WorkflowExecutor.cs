using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Core.Workflow
{
    public sealed class WorkflowStepRun
    {
        public int Index { get; set; }

        public string Tool { get; set; }

        public string Description { get; set; }

        public bool Confirmed { get; set; }

        public ToolResult Result { get; set; }

        /// <summary>该步骤注入的自检执行结果（步骤未配置 Verify 时为 null）。</summary>
        public WorkflowVerifyRun Verify { get; set; }
    }

    /// <summary>流程步骤注入自检的一次执行结果：工具产出 verdict/issues/selfcheck 三要素。</summary>
    public sealed class WorkflowVerifyRun
    {
        public string Tool { get; set; }

        public ToolResult Result { get; set; }

        /// <summary>自检 verdict：ok 表示达标，其余为未通过标记（number/source/count 等）。</summary>
        public string Verdict { get; set; }

        public List<string> Issues { get; set; } = new List<string>();

        /// <summary>自检是否达标：工具成功且 verdict 为 ok。</summary>
        public bool Passed => Result != null
            && Result.Ok
            && string.Equals(Verdict, "ok", StringComparison.OrdinalIgnoreCase);
    }

    public sealed class WorkflowRunResult
    {
        public List<WorkflowStepRun> Steps { get; } = new List<WorkflowStepRun>();

        /// <summary>全部步骤成功且无需用户额外输入时为 true。</summary>
        public bool Completed { get; set; }

        /// <summary>阻塞失败（未执行或失败）的步骤下标；全部完成时为 null。</summary>
        public int? FailedIndex { get; set; }

        /// <summary>被用户取消（危险确认拒绝）的步骤下标；无时为 null。</summary>
        public int? CancelledIndex { get; set; }

        public WorkflowStepRun StepAt(int index)
        {
            return Steps.FirstOrDefault(item => item.Index == index);
        }
    }

    /// <summary>
    /// 流程执行引擎（纯逻辑，可单测）：
    /// 逐条执行模板步骤，命中工具取其风险级别，经确认回调后执行；任一步失败即停在失败步骤，
    /// 返回下标供上层做单步重试（retry），不再重跑已成功步骤。
    /// 合成步骤参数：把模板占位符 {name} 用运行时覆盖值填充。
    /// </summary>
    public sealed class WorkflowExecutor
    {
        private readonly ToolRegistry registry;
        private readonly Func<ToolEnvelope, RiskLevel, Task<bool>> confirmHandler;

        public WorkflowExecutor(
            ToolRegistry registry,
            Func<ToolEnvelope, RiskLevel, Task<bool>> confirmHandler)
        {
            this.registry = registry ?? throw new ArgumentNullException(nameof(registry));
            this.confirmHandler = confirmHandler ?? throw new ArgumentNullException(nameof(confirmHandler));
        }

        /// <summary>
        /// 从 fromIndex 起顺序执行；fromIndex 之前的结果（如重试前已成功步骤）由 recovered 提供并合并。
        /// </summary>
        public async Task<WorkflowRunResult> RunAsync(
            WorkflowTemplate template,
            IReadOnlyDictionary<string, object> overrides,
            IReadOnlyList<WorkflowStepRun> recovered,
            int fromIndex,
            CancellationToken cancellationToken = default,
            Action<AgentProgress> onProgress = null)
        {
            var result = new WorkflowRunResult();
            if (template == null || template.Steps == null || template.Steps.Count == 0)
            {
                result.Completed = true;
                return result;
            }

            // 生效覆盖 = 调用方覆盖 叠加 已完成步骤的输出（以 stepN.属性 为键），使后续步骤能用
            // {step2.chart_type} 这类占位引用前一步成果，从而让 verify_chart 等自检步骤可以真正地“对照实际产物”兜底。
            var effectiveOverrides = new Dictionary<string, object>();
            if (overrides != null)
            {
                foreach (var kv in overrides)
                {
                    effectiveOverrides[kv.Key] = kv.Value;
                }
            }

            if (recovered != null)
            {
                foreach (var run in recovered)
                {
                    result.Steps.Add(run);
                    CaptureStepOutput(effectiveOverrides, run.Index, run.Result);
                }
            }

            var totalSteps = template.Steps.Count;
            // 让用户在实时「思考过程」里看到流程总体规模，以便判断进行到哪了。
            raiseReasoning(onProgress,
                recovered != null && recovered.Count > 0
                    ? $"继续执行流程「{template.Name}」，仍待完成 {totalSteps - fromIndex}/{totalSteps} 步…"
                    : $"开始执行流程「{template.Name}」，需要依次完成 {totalSteps} 步操作…");

            for (var index = fromIndex; index < template.Steps.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var step = template.Steps[index];
                var args = ResolveArguments(step.Arguments, effectiveOverrides);

                var run = new WorkflowStepRun
                {
                    Index = index,
                    Tool = step.Tool,
                    Description = step.Description
                };
                result.Steps.Add(run);

                if (!registry.TryGet(step.Tool, out var descriptor))
                {
                    run.Result = new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "UNKNOWN_TOOL",
                        Data = $"未知工具：{step.Tool}"
                    };
                    result.FailedIndex = index;
                    return result;
                }

                var envelope = new ToolEnvelope
                {
                    Tool = step.Tool,
                    Args = args,
                    CallId = "wf-" + index + "-" + Guid.NewGuid().ToString("N"),
                    Reason = "执行流程步骤（" + step.Description + "）"
                };

                var confirmed = await confirmHandler(envelope, descriptor.RiskLevel);
                run.Confirmed = confirmed;
                if (!confirmed)
                {
                    raiseProgress(onProgress, AgentProgressKind.ToolResult, index, step.Tool,
                        "步骤需确认，已被取消", false, false);
                    run.Result = new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "USER_CANCELLED",
                        Data = "该步骤需要确认，已被取消。"
                    };
                    result.CancelledIndex = index;
                    result.FailedIndex = index;
                    return result;
                }

                raiseReasoning(onProgress, recovered == null || index > fromIndex
                        ? $"第 {index + 1}/{totalSteps} 步：{step.Description}"
                        : $"第 {index + 1}/{totalSteps} 步（重试续跑）：{step.Description}");
                raiseProgress(onProgress, AgentProgressKind.ToolStart, index, step.Tool, step.Description, true, false);
                try
                {
                    run.Result = descriptor.Execute(args);
                }
                catch (Exception ex)
                {
                    AppLog.Warn("WorkflowExecutor", $"[LogCover] WorkflowExecutor.cs:189  {ex.GetType().Name}: {ex.Message}");
                    run.Result = new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "EXECUTION_ERROR",
                        Data = ex.Message
                    };
                }

                if (run.Result.Ok)
                {
                    CaptureStepOutput(effectiveOverrides, index, run.Result);
                    raiseProgress(onProgress, AgentProgressKind.ToolResult, index, step.Tool, SummaryOf(step.Tool, run.Result), true, true);

                    // 每步可选自检注入：主步骤成功后执行对应 verify 工具并校验 verdict；
                    // 未达标视同该步未通过，停在当前步供上层单步重试。
                    if (step.Verify != null)
                    {
                        var verifyRun = RunVerify(step.Verify, effectiveOverrides, index, onProgress);
                        run.Verify = verifyRun;
                        if (verifyRun != null && !verifyRun.Passed)
                        {
                            result.FailedIndex = index;
                            return result;
                        }
                    }
                }
                else
                {
                    raiseProgress(onProgress, AgentProgressKind.ToolResult, index, step.Tool,
                        run.Result.ErrorCode + ": " + run.Result.Data, false, false);
                    result.FailedIndex = index;
                    return result;
                }
            }

            // 成功完成整个模板时才置 true；任一步失败/取消都会提前 return，此时保持 false。
            result.Completed = true;
            return result;
        }

        private const string WorkflowNamespace = "workflow";
        private const string WorkflowVerifyNamespace = "workflow-verify";

        /// <summary>
        /// 把某个步骤成功产出的 Data（匿名对象/字典）的顶层字段注册为 {stepN.属性} 覆盖键，
        /// 供后续步骤的 ResolveArguments 以占位符引用（自检步骤对照“实际产物”兜底）。
        /// 反射失败或 Data 为空则静默跳过，不影响流程。
        /// </summary>
        private static void CaptureStepOutput(Dictionary<string, object> target, int index, ToolResult run)
        {
            if (target == null || run == null || !run.Ok || run.Data == null)
            {
                return;
            }

            try
            {
                var prefix = "step" + index + ".";
                switch (run.Data)
                {
                    case System.Collections.IDictionary dict:
                        foreach (DictionaryEntry e in dict)
                        {
                            if (e.Key != null && !target.ContainsKey(prefix + e.Key))
                            {
                                target[prefix + e.Key] = e.Value;
                            }
                        }

                        break;
                    default:
                        foreach (var p in run.Data.GetType().GetProperties())
                        {
                            // 已存在的键跳过（不覆盖既有覆盖值），其余注册为 {stepN.属性}。
                            if (!p.CanRead || target.ContainsKey(prefix + p.Name))
                            {
                                continue;
                            }

                            var v = p.GetValue(run.Data, null);
                            target[prefix + p.Name] = v ?? string.Empty;
                        }

                        break;
                }
            }
            catch
            {
                AppLog.Warn("WorkflowExecutor", "[LogCover] WorkflowExecutor.cs:277  silent-catch covered (no exception variable)");
                // 确保任何反射意外都不会拖垮主流程。
            }
        }

        /// <summary>
        /// 执行注入的自检步骤：解析参数（可引用 {stepN.属性} 前一步产物与 {selection}）、查注册表、执行，
        /// 提取工具产出的 verdict/issues 判定是否达标，并把过程以 workflow 命名空间上报前端。
        /// 不经过确认回调——verify 工具均为只读，注入即自动放行。
        /// </summary>
        private WorkflowVerifyRun RunVerify(
            WorkflowVerify spec,
            IReadOnlyDictionary<string, object> overrides,
            int index,
            Action<AgentProgress> onProgress)
        {
            if (spec == null)
            {
                return null;
            }

            // 配置了自检但未指定工具：视为配置错误而非静默放行，fail-closed 阻断当前步，
            // 避免模板作者少填 Tool 导致自检兜底被悄悄关闭。
            if (string.IsNullOrWhiteSpace(spec.Tool))
            {
                return new WorkflowVerifyRun
                {
                    Tool = "(未配置)",
                    Result = new ToolResult { Ok = false, ErrorCode = "VERIFY_NOT_CONFIGURED", Data = "自检步骤未指定 verify 工具" },
                    Verdict = "error",
                    Issues = { "自检步骤未指定 verify 工具" }
                };
            }

            var args = ResolveArguments(spec.Arguments, overrides);
            if (!registry.TryGet(spec.Tool, out var descriptor))
            {
                return new WorkflowVerifyRun
                {
                    Tool = spec.Tool,
                    Result = new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = "未知自检工具：" + spec.Tool },
                    Verdict = "error",
                    Issues = { "未知自检工具：" + spec.Tool }
                };
            }

            var label = string.IsNullOrWhiteSpace(spec.Description) ? spec.Tool : spec.Description;
            raiseReasoning(onProgress, $"第 {index + 1} 步自检：{label}");
            raiseProgress(onProgress, AgentProgressKind.ToolStart, index, spec.Tool, "步骤自检：" + label, true, false, WorkflowVerifyNamespace);

            ToolResult result;
            try
            {
                result = descriptor.Execute(args);
            }
            catch (Exception ex)
            {
                AppLog.Warn("WorkflowExecutor", $"[LogCover] WorkflowExecutor.cs:333  {ex.GetType().Name}: {ex.Message}");
                result = new ToolResult { Ok = false, ErrorCode = "EXECUTION_ERROR", Data = ex.Message };
            }

            var verdict = ExtractVerdict(result);
            var issues = ExtractIssues(result);
            var passed = result.Ok && string.Equals(verdict, "ok", StringComparison.OrdinalIgnoreCase);
            var summary = "自检" + (passed ? "通过" : "未通过") + "：" + label
                + (issues.Count > 0 ? "；" + string.Join("；", issues) : string.Empty);
            raiseProgress(onProgress, AgentProgressKind.ToolResult, index, spec.Tool, summary, passed, true, WorkflowVerifyNamespace);

            return new WorkflowVerifyRun
            {
                Tool = spec.Tool,
                Result = result,
                Verdict = verdict,
                Issues = issues
            };
        }

        /// <summary>从自检工具结果 Data 中提取 verdict（支持匿名对象属性与 IDictionary 键）。</summary>
        private static string ExtractVerdict(ToolResult result)
        {
            if (result == null || !result.Ok || result.Data == null)
            {
                return null;
            }

            var raw = TryGetField(result.Data, "verdict");
            return raw == null ? null : Convert.ToString(raw);
        }

        /// <summary>从自检工具结果 Data 中提取 issues 列表。</summary>
        private static List<string> ExtractIssues(ToolResult result)
        {
            var issues = new List<string>();
            if (result == null || result.Data == null)
            {
                return issues;
            }

            var raw = TryGetField(result.Data, "issues");
            if (raw is System.Collections.IEnumerable sequence && !(raw is string))
            {
                foreach (var item in sequence)
                {
                    if (item != null)
                    {
                        issues.Add(Convert.ToString(item));
                    }
                }
            }

            return issues;
        }

        /// <summary>按名称取匿名对象属性或 IDictionary 键值（大小写不敏感），无则 null。</summary>
        private static object TryGetField(object data, string name)
        {
            if (data is System.Collections.IDictionary dict)
            {
                foreach (System.Collections.DictionaryEntry e in dict)
                {
                    if (string.Equals(Convert.ToString(e.Key), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return e.Value;
                    }
                }

                return null;
            }

            foreach (var p in data.GetType().GetProperties())
            {
                if (p.CanRead && string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        return p.GetValue(data, null);
                    }
                    catch
                    {
                        AppLog.Warn("WorkflowExecutor", "[LogCover] WorkflowExecutor.cs:414  silent-catch covered (no exception variable)");
                        return null;
                    }
                }
            }

            return null;
        }

        /// <summary>内层流程步骤的进度事件，Namespace 恒为 "workflow"，与外围 agent 的工具序号隔离。</summary>
        private static void raiseProgress(Action<AgentProgress> onProgress, AgentProgressKind kind, int index, string tool, string summary, bool ok, bool confirmed, string ns = WorkflowNamespace)
        {
            if (onProgress == null) return;
            try
            {
                onProgress(new AgentProgress
                {
                    Kind = kind,
                    Namespace = ns,
                    Tool = tool,
                    Summary = summary,
                    Ok = ok,
                    Confirmed = confirmed,
                    ToolIndex = index,
                    ToolCount = index + 1
                });
            }
            catch { }
        }

        /// <summary>以「思考片段」形式向前端陈述当前执行位点（第几步/在做什么），使用户跟进进度。</summary>
        private static void raiseReasoning(Action<AgentProgress> onProgress, string text)
        {
            if (onProgress == null || string.IsNullOrWhiteSpace(text)) return;
            try
            {
                onProgress(new AgentProgress
                {
                    Kind = AgentProgressKind.Reasoning,
                    Namespace = WorkflowNamespace,
                    Reasoning = text
                });
            }
            catch { }
        }

        private static string SummaryOf(string tool, ToolResult r)
        {
            if (r == null) return null;
            var summary = Agent.ToolSummaryBuilder.Summarize(tool ?? string.Empty, r.Data);
            return summary.Length > 120 ? summary.Substring(0, 120) : summary;
        }

        /// <summary>浅层参数合成：遍历模板参数字典，把 {name} 字符串替换为覆盖值（支持嵌套字典/数组）。</summary>
        public static Dictionary<string, object> ResolveArguments(
            IReadOnlyDictionary<string, object> arguments,
            IReadOnlyDictionary<string, object> overrides)
        {
            var result = new Dictionary<string, object>();
            foreach (var pair in arguments ?? new Dictionary<string, object>())
            {
                result[pair.Key] = ResolveValue(pair.Value, overrides);
            }

            return result;
        }

        private static object ResolveValue(object value, IReadOnlyDictionary<string, object> overrides)
        {
            switch (value)
            {
                case string text when overrides != null:
                    if (text.Length > 1 && text[0] == '{' && text[text.Length - 1] == '}')
                    {
                        var key = text.Substring(1, text.Length - 2);
                        if (overrides.TryGetValue(key, out var direct))
                        {
                            return direct;
                        }
                    }

                    foreach (var pair in overrides)
                    {
                        if (text.Contains("{" + pair.Key + "}"))
                        {
                            text = text.Replace("{" + pair.Key + "}", Stringify(pair.Value));
                        }
                    }

                    return text;
                case IDictionary<string, object> map:
                    return ResolveArguments((IReadOnlyDictionary<string, object>)map, overrides);
                case System.Collections.IEnumerable sequence when !(value is string):
                    var list = new List<object>();
                    foreach (var item in sequence)
                    {
                        list.Add(ResolveValue(item, overrides));
                    }

                    return list;
                default:
                    return value;
            }
        }

        /// <summary>覆盖值注入字符串模板时统一为紧凑 JSON 文本：JToken 走 ToString(Formatting.None)，
        /// 否则 Convert.ToString。避免 JObject 默认格式化输出（多行缩进）污染被拼装的模板串。</summary>
        private static string Stringify(object value)
        {
            if (value is Newtonsoft.Json.Linq.JToken token)
            {
                return token.ToString(Newtonsoft.Json.Formatting.None);
            }

            return Convert.ToString(value);
        }
    }
}