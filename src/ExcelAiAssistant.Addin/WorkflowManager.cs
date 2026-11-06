using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 流程编排协作者（阶段 B 自 SidebarBridge 抽出）。
    /// 流程专属状态（workflowAttempts / activeWorkflowNames / WorkflowAttempt）整体迁入；
    /// ListConnectedBooks（跨簿协同）与 StopGeneration（agent 运行控制）不属于流程域，留在 Bridge。
    /// 桥接能力经委托注入：busy/异常壳、工具注册表、进度广播、确认门、会话工具记录、选区捕获。
    /// </summary>
    internal sealed class WorkflowManager
    {
        private readonly Func<Func<Task<string>>, Task<string>> executeAsync;
        private readonly Func<string, ToolRegistry> createRegistry;
        private readonly Func<string> currentTurnLabel;
        private readonly Action<string, object> publish;
        private readonly Func<ToolEnvelope, RiskLevel, Task<ConfirmationOutcome>> confirmTool;
        private readonly Func<IReadOnlyList<ToolEventRecord>> lastToolRecords;
        private readonly TryCaptureSelectionDelegate tryCaptureSelection;
        private readonly WorkflowStore workflowStore;
        private readonly Dictionary<string, WorkflowAttempt> workflowAttempts =
            new Dictionary<string, WorkflowAttempt>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> activeWorkflowNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal delegate bool TryCaptureSelectionDelegate(out string address, out string sheet);

        public WorkflowManager(
            Func<Func<Task<string>>, Task<string>> executeAsync,
            Func<string, ToolRegistry> createRegistry,
            Func<string> currentTurnLabel,
            Action<string, object> publish,
            Func<ToolEnvelope, RiskLevel, Task<ConfirmationOutcome>> confirmTool,
            Func<IReadOnlyList<ToolEventRecord>> lastToolRecords,
            TryCaptureSelectionDelegate tryCaptureSelection,
            WorkflowStore workflowStore)
        {
            this.executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
            this.createRegistry = createRegistry ?? throw new ArgumentNullException(nameof(createRegistry));
            this.currentTurnLabel = currentTurnLabel ?? throw new ArgumentNullException(nameof(currentTurnLabel));
            this.publish = publish ?? throw new ArgumentNullException(nameof(publish));
            this.confirmTool = confirmTool ?? throw new ArgumentNullException(nameof(confirmTool));
            this.lastToolRecords = lastToolRecords ?? throw new ArgumentNullException(nameof(lastToolRecords));
            this.tryCaptureSelection = tryCaptureSelection ?? throw new ArgumentNullException(nameof(tryCaptureSelection));
            this.workflowStore = workflowStore ?? throw new ArgumentNullException(nameof(workflowStore));
        }

        public Task<string> ListWorkflows()
        {
            return executeAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                workflows = workflowStore.GetAll().Select(item => (object)new
                {
                    name = item.Name,
                    description = item.Description,
                    is_built_in = item.IsBuiltIn,
                    step_count = item.Steps.Count,
                    steps = item.Steps.Select(step => (object)new
                    {
                        tool = step.Tool,
                        description = step.Description,
                        allow_retry = step.AllowRetry,
                        arguments = step.Arguments
                    }).ToArray()
                }).ToArray()
            })));
        }

        /// <summary>预览当前会话可提取为流程的工具步骤（不保存）。</summary>
        public Task<string> PreviewWorkflowSteps()
        {
            return executeAsync(() =>
            {
                var steps = ExtractWorkflowStepsFromConversation();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    step_count = steps.Count,
                    steps = steps.Select(step => (object)new
                    {
                        tool = step.Tool,
                        description = step.Description,
                        arguments = step.Arguments
                    }).ToArray()
                }));
            });
        }

        /// <summary>从当前会话的工具调用记录另存为可复用流程（B4）。</summary>
        public Task<string> SaveWorkflow(string name, string description, bool deScope = true)
        {
            return executeAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    throw new ArgumentException("流程名不能为空。", nameof(name));
                }

                var steps = ExtractWorkflowStepsFromConversation();
                if (steps.Count == 0)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "当前会话没有可提取的工具调用步骤。" }));
                }

                if (deScope)
                {
                    // 把提取步骤里固化的作用区域（如 VBA Range("C5:F22")）改写为 {selection} 占位符，
                    // 使流程执行时取"当前选区"动态解析，从而复用任意数据区域。
                    var stub = new WorkflowTemplate { Name = name.Trim(), Steps = steps };
                    FlowScopeNormalizer.NormalizeTemplate(stub);
                }

                var template = new WorkflowTemplate
                {
                    Name = name.Trim(),
                    Description = string.IsNullOrWhiteSpace(description) ? "从成功会话提取的流程" : description.Trim(),
                    CreatedAtUtc = DateTime.UtcNow,
                    Steps = steps
                };
                if (!workflowStore.Save(template))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "保存失败：内置流程名不可覆盖。" }));
                }

                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, name = template.Name, step_count = steps.Count }));
            });
        }

        /// <summary>对一条已保存的自定义流程执行"去作用域"：把固化区域改写为 {selection} 占位符，就地保存。</summary>
        public Task<string> DeScopeWorkflow(string name)
        {
            return executeAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    throw new ArgumentException("流程名不能为空。", nameof(name));
                }

                if (!workflowStore.TryGet(name, out var template))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = $"流程不存在：{name}" }));
                }

                if (template.IsBuiltIn)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "内置流程不可改写。" }));
                }

                // 复制对象以避免改写污染内存中的运行模板引用。
                var copy = new WorkflowTemplate
                {
                    Name = template.Name,
                    Description = template.Description,
                    IsBuiltIn = template.IsBuiltIn,
                    CreatedAtUtc = template.CreatedAtUtc,
                    Steps = JsonConvert.DeserializeObject<List<WorkflowStep>>(
                        JsonConvert.SerializeObject(template.Steps))
                };
                var replaced = FlowScopeNormalizer.NormalizeTemplate(copy);
                if (replaced == 0)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, changed = false, replaced = 0, msg = "未发现可去除的作用区域。" }));
                }

                workflowStore.Save(copy);
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, changed = true, replaced = replaced }));
            });
        }


        public Task<string> DeleteWorkflow(string name)
        {
            return executeAsync(() =>
            {
                if (workflowStore.Delete(name))
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = true }));
                }

                return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "删除失败：内置流程不可删除或不存在。" }));
            });
        }

        /// <summary>从设置页 / 流程页触发执行指定流程（携带运行时覆盖参数）。</summary>
        public Task<string> RunWorkflow(string name, string overridesJson)
        {
            return executeAsync(() =>
            {
                var overrides = ParseOverrides(overridesJson);
                var run = ExecuteWorkflowInternal(name, overrides, 0, null);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = run.Completed,
                    run = BuildRunPayload(run)
                }));
            });
        }

        /// <summary>单步重试：仅重跑失败步骤及其后步骤，不重跑已成功步骤（B3）。使用锁定的 attempt.Template，防止运行期间流程被编辑造成步骤错位。</summary>
        public Task<string> RetryWorkflowStep(string name, int index)
        {
            return executeAsync(() =>
            {
                if (!workflowAttempts.TryGetValue(name, out var attempt) || attempt == null || attempt.Template == null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "没有可重试的流程运行记录，请先执行一次。", cancelled = true }));
                }

                if (index < 0 || index >= attempt.Template.Steps.Count)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "无效的重试步骤下标。" }));
                }

                if (!attempt.Template.Steps[index].AllowRetry)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "该步骤标记为不可重试，请重新运行整个流程。" }));
                }

                var recovered = attempt.Steps
                    .Where(item => item.Index < index && item.Result?.Ok == true)
                    .ToList();
                var run = ExecuteWorkflowInternal(attempt.Template, name, attempt.Overrides, index, recovered);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = run.Completed,
                    run = BuildRunPayload(run)
                }));
            });
        }

        /// <summary>执行流程（同步封装，供 run_workflow 工具与桥接方法共用）。</summary>
        internal WorkflowRunResult ExecuteWorkflowInternal(
            string name,
            Dictionary<string, object> overrides,
            int fromIndex,
            IReadOnlyList<WorkflowStepRun> recovered)
        {
            if (!workflowStore.TryGet(name, out var template))
            {
                return new WorkflowRunResult
                {
                    FailedIndex = fromIndex,
                    Steps =
                    {
                        new WorkflowStepRun
                        {
                            Index = fromIndex,
                            Tool = string.Empty,
                            Result = new ToolResult { Ok = false, ErrorCode = "NO_WORKFLOW", Data = $"流程不存在：{name}" }
                        }
                    }
                };
            }

            return ExecuteWorkflowInternal(template, name, overrides, fromIndex, recovered);
        }

        /// <summary>以已解析模板执行流程。重试路径传入 attempt.Template 以锁定模板，避免运行期间流程被编辑导致的步骤错位。</summary>
        internal WorkflowRunResult ExecuteWorkflowInternal(
            WorkflowTemplate template,
            string name,
            Dictionary<string, object> overrides,
            int fromIndex,
            IReadOnlyList<WorkflowStepRun> recovered)
        {
            // 递归/环保护：同一命名流程在调用栈上再次进入即拒绝，杜绝 run_workflow 步骤造成的自引用/互引无界递归。
            if (!activeWorkflowNames.Add(name))
            {
                return new WorkflowRunResult
                {
                    FailedIndex = fromIndex,
                    Steps =
                    {
                        new WorkflowStepRun
                        {
                            Index = fromIndex,
                            Tool = template.Name,
                            Result = new ToolResult { Ok = false, ErrorCode = "WORKFLOW_CYCLE", Data = $"流程递归：{name} 已在执行中，禁止嵌套执行同一流程。" }
                        }
                    }
                };
            }

            try
            {
                var effectiveOverrides = overrides;
                if (FlowScopeNormalizer.TemplateReferencesSelection(template))
                {
                    // 流程步骤含 {selection}/{selection_sheet} 占位符：用 Excel 当前选区动态解析，
                    // 使用户选哪块就处理哪块，让"去作用域"后的流程复用任意数据区域。
                    // 用户显式覆盖的同名键优先；取不到选区且未覆盖时给出友好错误。
                    if (HasOverride(overrides, "selection") || HasOverride(overrides, "selection_sheet"))
                    {
                        // 用户已提供覆盖，保持原 overrides（占位符会由 ResolveValue 用用户值填充）
                    }
                    else if (tryCaptureSelection(out var selAddress, out var selSheet))
                    {
                        var merged = new Dictionary<string, object>(
                            overrides ?? new Dictionary<string, object>(), StringComparer.OrdinalIgnoreCase);
                        merged[FlowScopeNormalizer.SelectionMarker.Trim('{', '}')] = selAddress;
                        merged[FlowScopeNormalizer.SelectionSheetMarker.Trim('{', '}')] = selSheet;
                        effectiveOverrides = merged;
                    }
                    else
                    {
                        var selRun = new WorkflowRunResult();
                        var selStep = new WorkflowStepRun
                        {
                            Index = fromIndex,
                            Tool = template.Name,
                            Result = new ToolResult
                            {
                                Ok = false,
                                ErrorCode = "WORKFLOW_NO_SELECTION",
                                Data = "该流程引用了「当前选区」，请先在 Excel 中选中要处理的数据区域再运行。"
                            }
                        };
                        selRun.Steps.Add(selStep);
                        selRun.FailedIndex = fromIndex;
                        return selRun;
                    }
                }

                var registry = createRegistry(currentTurnLabel());
                var executor = new WorkflowExecutor(registry, (envelope, risk) => Task.FromResult(ConfirmToolSync(envelope, risk)));
                var run = executor.RunAsync(
                    template,
                    effectiveOverrides,
                    recovered,
                    fromIndex,
                    CancellationToken.None,
                    p => publish("agent-progress", SidebarBridge.ProgressPayload(p))).GetAwaiter().GetResult();

                workflowAttempts[name] = new WorkflowAttempt
                {
                    Template = template,
                    Overrides = effectiveOverrides ?? new Dictionary<string, object>(),
                    Steps = run.Steps
                };
                return run;
            }
            finally
            {
                activeWorkflowNames.Remove(name);
            }
        }

        private static bool HasOverride(Dictionary<string, object> overrides, string key)
        {
            if (overrides == null) return false;
            foreach (var k in overrides.Keys)
            {
                if (string.Equals(k, key, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }


        /// <summary>run_workflow 工具执行体（同步委托给内部引擎，结果作为工具观察返回）。</summary>
        internal ToolResult RunWorkflowTool(Dictionary<string, object> args)
        {
            var name = Convert.ToString(args.TryGetValue("workflow", out var n) ? n : null) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name))
            {
                return new ToolResult { Ok = false, ErrorCode = "NO_WORKFLOW", Data = "缺少 workflow 参数（流程名称）。" };
            }

            var overrides = new Dictionary<string, object>();
            foreach (var pair in args)
            {
                if (!string.Equals(pair.Key, "workflow", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(pair.Key, "description", StringComparison.OrdinalIgnoreCase))
                {
                    overrides[pair.Key] = pair.Value;
                }
            }

            var run = ExecuteWorkflowInternal(name, overrides, 0, null);
            return new ToolResult
            {
                Ok = run.Completed,
                Data = BuildWorkflowAnswer(run),
                AffectedRanges = run.Steps
                    .Where(item => item.Result?.AffectedRanges != null)
                    .SelectMany(item => item.Result.AffectedRanges)
                    .Distinct()
                    .ToList()
            };
        }

        private bool ConfirmToolSync(ToolEnvelope envelope, RiskLevel riskLevel)
        {
            // 非危工具/只读立即放行（ConfirmTool 内部短路），危工具与代码类会弹用户确认。
            return confirmTool(envelope, riskLevel).GetAwaiter().GetResult() == ConfirmationOutcome.Confirmed;
        }

        internal string BuildWorkflowAnswer(WorkflowRunResult run)
        {
            var lines = new List<string> { "已完成流程执行：" };
            foreach (var step in run.Steps)
            {
                var mark = step.Result?.Ok == true ? "✓" : "✗";
                var note = step.Result?.Ok == true
                    ? (step.Result.Data != null ? SidebarBridge.SummarizeToolResult(step.Result) : "成功")
                    : (step.Result?.Data?.ToString() ?? step.Result?.ErrorCode ?? "失败");
                var line = $"{mark} 步骤{step.Index + 1}（{step.Tool}）：{note}";
                if (step.Verify != null)
                {
                    line += step.Verify.Passed
                        ? $"\n　↳ 自检（{step.Verify.Tool}）通过"
                        : $"\n　↳ 自检（{step.Verify.Tool}）未通过：{(step.Verify.Result?.Data?.ToString() ?? step.Verify.Verdict ?? "失败")}" +
                          (step.Verify.Issues.Count > 0 ? "；" + string.Join("；", step.Verify.Issues) : "");
                }

                lines.Add(line);
            }

            if (!run.Completed && run.CancelledIndex.HasValue)
            {
                lines.Add("流程已中止：该步骤需要确认，已取消。可重试或重新批准。");
            }
            else if (!run.Completed && run.FailedIndex.HasValue)
            {
                var failedStep = run.StepAt(run.FailedIndex.Value);
                if (failedStep?.Verify != null)
                {
                    lines.Add($"流程已中止于步骤 {run.FailedIndex.Value + 1}：该步骤的自检（{failedStep.Verify.Tool}）未达标，请修正后对该步单独重试。");
                }
                else
                {
                    lines.Add($"流程已中止于步骤 {run.FailedIndex.Value + 1}，可对该步单独重试。");
                }
            }

            return string.Join("\n", lines);
        }

        internal object BuildRunPayload(WorkflowRunResult run)
        {
            return new
            {
                completed = run.Completed,
                failed_index = run.FailedIndex,
                cancelled_index = run.CancelledIndex,
                steps = run.Steps.Select(step => (object)new
                {
                    index = step.Index,
                    tool = step.Tool,
                    description = step.Description,
                    confirmed = step.Confirmed,
                    // 步骤整体成功 = 主工具成功 且（无自检 或 自检通过）；主工具成功但自检未达标也视为失败，
                    // 使详情面板标 ✗ 并可重试（retry.disabled 依赖 ok）。
                    ok = step.Result?.Ok == true && (step.Verify == null || step.Verify.Passed),
                    error_code = step.Result?.Ok == true ? null : step.Result?.ErrorCode,
                    error = step.Result?.Ok == true
                        ? (step.Verify != null && !step.Verify.Passed
                            ? "自检未通过：" + (step.Verify.Verdict ?? "?") + (step.Verify.Issues.Count > 0 ? "；" + string.Join("；", step.Verify.Issues) : "")
                            : null)
                        : (step.Result?.Data?.ToString() ?? step.Result?.ErrorCode),
                    summary = step.Result?.Ok == true ? SidebarBridge.SummarizeToolResult(step.Result) : null,
                    verify = step.Verify == null ? null : (object)new
                    {
                        tool = step.Verify.Tool,
                        verdict = step.Verify.Verdict,
                        passed = step.Verify.Passed,
                        issues = step.Verify.Issues,
                        error = step.Verify.Result?.Ok == true ? null : (step.Verify.Result?.Data?.ToString() ?? step.Verify.Result?.ErrorCode)
                    }
                }).ToArray()
            };
        }

        internal Dictionary<string, object> ParseOverrides(string text)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text))
            {
                return result;
            }

            foreach (var token in text.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var eq = token.IndexOf('=');
                if (eq <= 0)
                {
                    continue;
                }

                var key = token.Substring(0, eq).Trim();
                var value = token.Substring(eq + 1).Trim();
                if (key.Length > 0)
                {
                    result[key] = value;
                }
            }

            return result;
        }

        private List<WorkflowStep> ExtractWorkflowStepsFromConversation()
        {
            return lastToolRecords()
                .Select(record => new WorkflowStep
                {
                    Tool = record.Tool,
                    Description = record.Tool,
                    Arguments = record.Args != null
                        ? record.Args.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase)
                        : new Dictionary<string, object>()
                })
                .ToList();
        }


        private sealed class WorkflowAttempt
        {
            public WorkflowTemplate Template { get; set; }

            public Dictionary<string, object> Overrides { get; set; }

            public List<WorkflowStepRun> Steps { get; set; }
        }
    }
}
