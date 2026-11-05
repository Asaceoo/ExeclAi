using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public class WorkflowCoreTests : IDisposable
    {
        private readonly string filePath;
        private readonly WorkflowStore store;

        public WorkflowCoreTests()
        {
            filePath = Path.Combine(Path.GetTempPath(), $"xai-wf-{Guid.NewGuid():N}.json");
            store = new WorkflowStore(filePath);
        }

        [Fact]
        public void Built_in_library_has_at_least_three_workflows()
        {
            var all = store.GetAll();
            Assert.True(all.Count >= 3, "内置流程至少 3 条");
            Assert.Contains(all, w => w.Name.Contains("数据体检"));
            Assert.Contains(all, w => w.Name.Contains("出图"));
            Assert.Contains(all, w => w.Name.Contains("比对"));
            Assert.All(all, w => Assert.True(w.IsBuiltIn));
        }

        [Fact]
        public void Built_in_library_expanded_with_clean_pivot_filter_and_more()
        {
            var all = store.GetAll();
            Assert.True(all.Count >= 12, "内置流程至少 12 条");
            Assert.Contains(all, w => w.Name.Contains("数据清洗"));
            Assert.Contains(all, w => w.Name.Contains("透视"));
            Assert.Contains(all, w => w.Name.Contains("筛选"));
            Assert.Contains(all, w => w.Name.Contains("迷你图"));
            Assert.Contains(all, w => w.Name.Contains("工作报告导出"));
            Assert.Contains(all, w => w.Name.Contains("数值结构摘要"));
            Assert.Contains(all, w => w.Name.Contains("双表差异报告"));
            Assert.Contains(all, w => w.Name.Contains("清洗后一键出图"));
            // 每个内置流程的步骤都必须有具体工具名，运行才不会因空工具中断
            Assert.All(all, w => Assert.All(w.Steps, s => Assert.False(string.IsNullOrWhiteSpace(s.Tool))));
        }

        [Fact]
        public void Save_and_load_user_workflow_round_trips()
        {
            var wf = new WorkflowTemplate
            {
                Name = "我的清洗流程",
                Description = "清洗并出图",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "clean_excel_data", Description = "清洗数据" },
                    new WorkflowStep { Tool = "create_excel_chart", Description = "出图" }
                }
            };

            Assert.True(store.Save(wf));
            var reloaded = new WorkflowStore(filePath);
            Assert.True(reloaded.TryGet("我的清洗流程", out var got));
            Assert.Equal(2, got.Steps.Count);
            Assert.Equal("clean_excel_data", got.Steps[0].Tool);
            Assert.False(got.IsBuiltIn);
        }

        [Fact]
        public void Cannot_override_or_delete_built_in_workflow()
        {
            var builtIn = store.GetAll()[0];
            var cloned = new WorkflowTemplate
            {
                Name = builtIn.Name,
                Description = "试图覆盖"
            };
            Assert.False(store.Save(cloned));
            Assert.False(store.Delete(builtIn.Name));
        }

        [Fact]
        public void Delete_removes_only_user_workflow()
        {
            var wf = new WorkflowTemplate { Name = "临时流程" };
            Assert.True(store.Save(wf));
            Assert.True(store.Delete("临时流程"));
            Assert.False(store.TryGet("临时流程", out _));
        }

        [Fact]
        public void BuiltIn_chart_workflow_verify_chart_step_carries_expectations()
        {
            // 清单约束兜底：内置“选区一键出图”的 verify_chart 必须引用前一步 create_excel_chart 的实际产物，
            // 而不是空跑；否则新加的标题/数据源/类型断言在此规范流程里永不生效。
            Assert.True(store.TryGet("选区一键出图", out var tmpl));
            var verify = tmpl.Steps.FirstOrDefault(s => s.Tool == "verify_chart");
            Assert.NotNull(verify);

            Assert.True(verify.Arguments.TryGetValue("expected_chart_type", out var ct) && Convert.ToString(ct) == "{step2.chart_type}");
            Assert.True(verify.Arguments.TryGetValue("expected_title", out var tl) && Convert.ToString(tl) == "{step2.title}");
            Assert.True(verify.Arguments.TryGetValue("expected_data_source", out var ds) && Convert.ToString(ds) == "{step2.source_address}");
            Assert.True(verify.Arguments.TryGetValue("chart_name", out var cn) && Convert.ToString(cn) == "{step2.chart_name}");
        }

        [Fact]
        public void Persist_is_atomic_and_leaves_no_temp_file()
        {
            var rounds = new List<string>();
            for (var i = 0; i < 20; i++)
            {
                var name = "原子流程" + i;
                Assert.True(store.Save(new WorkflowTemplate { Name = name, Steps = new List<WorkflowStep> { new WorkflowStep { Tool = "probe" } } }));
                rounds.Add(name);
            }

            var leftovers = Directory.GetFiles(Path.GetDirectoryName(filePath), "*.tmp.*")
                .Where(f => f.StartsWith(filePath, StringComparison.OrdinalIgnoreCase))
                .ToList();
            Assert.Empty(leftovers);

            var reloaded = new WorkflowStore(filePath);
            foreach (var name in rounds)
            {
                Assert.True(reloaded.TryGet(name, out _), "原子落盘后应能读回 " + name);
            }
        }

        [Fact]
        public async Task Executor_feeds_prior_step_output_into_later_step_placeholders()
        {
            var registry = new ToolRegistry();
            // step0: 产出可被后续 {step0.xxx} 占位引用的匿名对象（模拟 create_excel_chart 返回）。
            registry.Register(new ToolDescriptor(
                "produce", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult
                {
                    Ok = true,
                    Data = new { chart_type = "Column", title = "销售", source_address = "A2:B8", chart_name = "c1" }
                }));
            // step1: 把收到的 args 原样回吐，便于断言占位符是否真的被替换成 step0 输出。
            registry.Register(new ToolDescriptor(
                "consume", "reads args", RiskLevel.ReadOnly,
                args => new ToolResult { Ok = true, Data = args }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "chain",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "produce" },
                    new WorkflowStep
                    {
                        Tool = "consume",
                        Arguments = new Dictionary<string, object>
                        {
                            { "expected_chart_type", "{step0.chart_type}" },
                            { "expected_title", "{step0.title}" },
                            { "expected_data_source", "{step0.source_address}" },
                            { "chart_name", "{step0.chart_name}" }
                        }
                    }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.True(result.Completed);
            Assert.Equal(2, result.Steps.Count);

            var echoed = Assert.IsType<Dictionary<string, object>>(result.Steps[1].Result.Data);
            Assert.Equal("Column", Convert.ToString(echoed["expected_chart_type"]));
            Assert.Equal("销售", Convert.ToString(echoed["expected_title"]));
            Assert.Equal("A2:B8", Convert.ToString(echoed["expected_data_source"]));
            Assert.Equal("c1", Convert.ToString(echoed["chart_name"]));
        }

        [Fact]
        public void Workflow_verify_spec_round_trips_through_store()
        {
            var wf = new WorkflowTemplate
            {
                Name = "带自检流程",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "create_excel_pivot_summary",
                        Verify = new WorkflowVerify
                        {
                            Tool = "verify_pivot",
                            Description = "透视自检",
                            Arguments = new Dictionary<string, object> { ["target_address"] = "{step1.target_address}" }
                        }
                    }
                }
            };

            Assert.True(store.Save(wf));
            var reloaded = new WorkflowStore(filePath);
            Assert.True(reloaded.TryGet("带自检流程", out var got));
            var step = got.Steps[0];
            Assert.NotNull(step.Verify);
            Assert.Equal("verify_pivot", step.Verify.Tool);
            Assert.Equal("透视自检", step.Verify.Description);
            Assert.Equal("{step1.target_address}", Convert.ToString(step.Verify.Arguments["target_address"]));
        }

        [Fact]
        public async Task Executor_injects_verify_and_passes_when_verdict_ok()
        {
            var registry = new ToolRegistry();
            Dictionary<string, object> capturedVerifyArgs = null;
            // step0: 产出 {step0.target_address}/{step0.row_count} 供自检引用（模拟 create_excel_pivot_summary）。
            registry.Register(new ToolDescriptor(
                "act", "op", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = new { target_address = "B2", row_count = 3 } }));
            registry.Register(new ToolDescriptor(
                "verify_pivot", "verify", RiskLevel.ReadOnly,
                args => { capturedVerifyArgs = args; return new ToolResult { Ok = true, Data = new { verdict = "ok", issues = new List<string>() } }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "act",
                        Verify = new WorkflowVerify
                        {
                            Tool = "verify_pivot",
                            Description = "透视汇总自检",
                            Arguments = new Dictionary<string, object>
                            {
                                { "target_address", "{step0.target_address}" },
                                { "expected_count", "{step0.row_count}" }
                            }
                        }
                    }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.True(result.Completed);
            Assert.NotNull(result.Steps[0].Verify);
            Assert.True(result.Steps[0].Verify.Passed);
            Assert.Equal("ok", result.Steps[0].Verify.Verdict);

            // 占位符必须被前一步真实产物替换，而不是空跑。
            Assert.Equal("B2", Convert.ToString(capturedVerifyArgs["target_address"]));
            Assert.Equal("3", Convert.ToString(capturedVerifyArgs["expected_count"]));
        }

        [Fact]
        public async Task Executor_stops_on_verify_failure_and_does_not_run_following_steps()
        {
            var registry = new ToolRegistry();
            var neverCalls = 0;
            registry.Register(new ToolDescriptor("act", "op", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));
            registry.Register(new ToolDescriptor(
                "verify_write", "verify", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = new { verdict = "count", issues = new List<string> { "区域非空单元格数与期望不符" } } }));
            registry.Register(new ToolDescriptor("never", "not reached", RiskLevel.ReadOnly,
                _ => { neverCalls++; return new ToolResult { Ok = true }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "act", Verify = new WorkflowVerify { Tool = "verify_write", Description = "写入自检" } },
                    new WorkflowStep { Tool = "never" }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.False(result.Completed);
            Assert.Equal(0, result.FailedIndex);
            Assert.NotNull(result.Steps[0].Verify);
            Assert.False(result.Steps[0].Verify.Passed);
            Assert.Equal("count", result.Steps[0].Verify.Verdict);
            Assert.Contains("单元格数", result.Steps[0].Verify.Issues[0]);
            Assert.Single(result.Steps); // never 未执行
            Assert.Equal(0, neverCalls);
        }

        [Fact]
        public async Task Executor_stops_when_verify_tool_itself_errors()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("act", "op", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));
            registry.Register(new ToolDescriptor(
                "verify", "verify", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = false, ErrorCode = "VERIFY_FAILED", Data = "自检执行失败" }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "act", Verify = new WorkflowVerify { Tool = "verify" } }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.False(result.Completed);
            Assert.Equal(0, result.FailedIndex);
            Assert.False(result.Steps[0].Verify.Passed);
            Assert.Equal("VERIFY_FAILED", result.Steps[0].Verify.Result.ErrorCode);
        }

        [Fact]
        public async Task Executor_skips_verify_when_not_configured()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("act", "op", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep> { new WorkflowStep { Tool = "act" } }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.True(result.Completed);
            Assert.Null(result.Steps[0].Verify);
        }

        [Fact]
        public async Task Executor_blocks_when_verify_configured_but_tool_empty()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("act", "op", RiskLevel.ReadOnly, _ => new ToolResult { Ok = true }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "act", Verify = new WorkflowVerify { Description = "未指定自检工具" } }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.False(result.Completed);
            Assert.Equal(0, result.FailedIndex);
            Assert.NotNull(result.Steps[0].Verify);
            Assert.False(result.Steps[0].Verify.Passed);
            Assert.Equal("VERIFY_NOT_CONFIGURED", result.Steps[0].Verify.Result.ErrorCode);
        }

        [Fact]
        public async Task Executor_unresolved_placeholder_stays_untouched_for_detection()
        {
            // 若前一步无输出（缺 key），{missing.x} 应原样保留，印证自检工具的“占位符泄漏”守卫有必要。
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "probe", "reads", RiskLevel.ReadOnly,
                args => new ToolResult { Ok = true, Data = args }));
            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "leak",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "probe",
                        Arguments = new Dictionary<string, object> { { "expected_title", "{missing.title}" } }
                    }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            var echoed = Assert.IsType<Dictionary<string, object>>(result.Steps[0].Result.Data);
            Assert.Equal("{missing.title}", Convert.ToString(echoed["expected_title"]));
        }

        [Fact]
        public async Task Executor_runs_all_steps_and_stops_on_failure()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "good", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = "fine" }));
            registry.Register(new ToolDescriptor(
                "bad", "fails", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = false, ErrorCode = "X", Data = "boom" }));
            registry.Register(new ToolDescriptor(
                "never", "not reached", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "good" },
                    new WorkflowStep { Tool = "bad" },
                    new WorkflowStep { Tool = "never" }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);

            Assert.False(result.Completed);
            Assert.Equal(2, result.Steps.Count);
            Assert.Equal(1, result.FailedIndex);
            Assert.Equal("boom", result.Steps[1].Result.Data);
        }

        [Fact]
        public async Task Executor_progress_is_namespaced_with_narration_for_live_ui()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "good", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = "fine" }));
            registry.Register(new ToolDescriptor(
                "other", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = "ok" }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "流程X",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "good", Description = "第一步" },
                    new WorkflowStep { Tool = "other", Description = "第二步" }
                }
            };

            var events = new List<AgentProgress>();
            await executor.RunAsync(template, null, null, 0, CancellationToken.None,
                onProgress: events.Add);

            // 每个流程工具步骤都必须带 workflow 命名空间，避免与外围 agent 的工具序号冲突覆盖。
            Assert.All(events.Where(e => e.Tool != null), e => Assert.Equal("workflow", e.Namespace));

            // 首事件为流程总览叙述（Reasoning），让用户看到整体分几步。
            Assert.NotEmpty(events);
            Assert.Equal(AgentProgressKind.Reasoning, events[0].Kind);
            Assert.Contains("流程X", events[0].Reasoning);
            Assert.Contains("2 步", events[0].Reasoning);

            // 每个步骤前都有「第 x/N 步」叙述，陈述当前执行位点。
            var narrations = events.Where(e => e.Kind == AgentProgressKind.Reasoning && e.Reasoning.Contains("步")).ToList();
            Assert.Contains(narrations, n => n.Reasoning.Contains("1/2"));
            Assert.Contains(narrations, n => n.Reasoning.Contains("2/2"));

            // 内层步骤索引从 0 起且与叙述配对：第 0 步 ToolStart → index=0。
            var start0 = events.First(e => e.Tool == "good" && e.Kind == AgentProgressKind.ToolStart);
            Assert.Equal(0, start0.ToolIndex);
            var start1 = events.First(e => e.Tool == "other" && e.Kind == AgentProgressKind.ToolStart);
            Assert.Equal(1, start1.ToolIndex);
        }

        [Fact]
        public async Task Executor_retries_single_step_without_rerunning_successful_ones()
        {
            var registry = new ToolRegistry();
            var goodCalls = 0;
            var flakyCalls = 0;
            var lastCalls = 0;
            registry.Register(new ToolDescriptor(
                "good", "ok", RiskLevel.ReadOnly,
                _ => { goodCalls++; return new ToolResult { Ok = true }; }));
            registry.Register(new ToolDescriptor(
                "flaky", "retryable", RiskLevel.ReadOnly,
                _ =>
                {
                    flakyCalls++;
                    return flakyCalls < 2
                        ? new ToolResult { Ok = false, ErrorCode = "R", Data = "first fail" }
                        : new ToolResult { Ok = true };
                }));
            registry.Register(new ToolDescriptor(
                "last", "ok", RiskLevel.ReadOnly,
                _ => { lastCalls++; return new ToolResult { Ok = true }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "good" },
                    new WorkflowStep { Tool = "flaky" },
                    new WorkflowStep { Tool = "last" }
                }
            };

            // 首轮：good 成功，flaky 失败 → 停在 1
            var first = await executor.RunAsync(template, null, null, 0, CancellationToken.None);
            Assert.Equal(2, first.Steps.Count);
            Assert.Equal(1, first.FailedIndex);

            // 单步重试：已成功步骤（Index < fromIndex）作为 recovered 传入，从 1 继续
            var recovered = first.Steps.Where(s => s.Index < 1).ToList();
            var retry = await executor.RunAsync(template, null, recovered, 1, CancellationToken.None);

            Assert.True(retry.Completed);
            Assert.Equal(3, retry.Steps.Count);
            // good 只执行了一次（未重跑）；flaky 第二次成功；last 执行一次
            Assert.Equal(1, goodCalls);
            Assert.Equal(2, flakyCalls);
            Assert.Equal(1, lastCalls);
        }

        [Fact]
        public async Task Executor_resolves_placeholders_from_overrides()
        {
            var registry = new ToolRegistry();
            Dictionary<string, object> captured = null;
            registry.Register(new ToolDescriptor(
                "read", "read", RiskLevel.ReadOnly,
                args => { captured = args; return new ToolResult { Ok = true }; }));

            var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "read", Arguments = new Dictionary<string, object>
                    {
                        { "source", "{range}" },
                        { "message", "分析 {range} 与 {sheet}" },
                        { "nested", new Dictionary<string, object> { { "inner", "{sheet}" } } }
                    } }
                }
            };

            var result = await executor.RunAsync(template, new Dictionary<string, object>
            {
                { "range", "A1:C9" },
                { "sheet", "汇总" }
            }, null, 0, CancellationToken.None);

            Assert.True(result.Completed);
            Assert.Equal("A1:C9", captured["source"]);
            Assert.Equal("分析 A1:C9 与 汇总", captured["message"]);
            Assert.Equal("汇总", ((Dictionary<string, object>)captured["nested"])["inner"]);
        }

        [Fact]
        public async Task Executor_cancels_on_unconfirmed_dangerous_step()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "safe", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true }));
            registry.Register(new ToolDescriptor(
                "danger", "risky", RiskLevel.Dangerous,
                _ => new ToolResult { Ok = true }));

            var executor = new WorkflowExecutor(registry, (_, risk) => Task.FromResult(risk != RiskLevel.Dangerous));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "safe" },
                    new WorkflowStep { Tool = "danger" },
                    new WorkflowStep { Tool = "safe" }
                }
            };

            var result = await executor.RunAsync(template, null, null, 0, CancellationToken.None);

            Assert.False(result.Completed);
            Assert.Equal(1, result.CancelledIndex);
            Assert.False(result.Steps[1].Confirmed);
        }

        [Fact]
        public void Domain_workflows_cover_analysis_hr_attendance_finance_with_known_tools()
        {
            var known = new HashSet<string>
            {
                "get_current_excel_context", "get_current_selection_values_and_formulas",
                "get_excel_sheet_list", "get_excel_range_values_and_formulas",
                "analyze_range_statistics", "analyze_value_distribution", "analyze_column_correlation",
                "generate_analysis_report", "recommend_excel_chart", "create_excel_chart",
                "export_excel_chart_image", "compare_two_sheets", "clean_excel_data",
                "create_excel_pivot_summary", "format_excel_range", "filter_excel_data",
                "create_excel_sparkline", "export_sheets_to_pdf",
                "sort_range", "dedup_range", "find_replace_range", "transpose_range", "format_excel_conditional",
                "text_clean", "text_split", "text_merge", "date_normalize", "number_pad", "fill_formula", "merge_sheets"
            };

            var all = store.GetAll();
            Assert.True(all.Count >= 31, "内置流程至少 31 条");
            Assert.Contains(all, w => w.Name.Contains("数据分析"));
            Assert.Contains(all, w => w.Name.Contains("人事"));
            Assert.Contains(all, w => w.Name.Contains("考勤"));
            Assert.Contains(all, w => w.Name.Contains("财务"));
            // 每个领域流程步骤都必须引用已知工具，运行才不会因未知工具中断。
            var domain = all.Where(w => w.Name.Contains("数据分析") || w.Name.Contains("人事") ||
                                        w.Name.Contains("考勤") || w.Name.Contains("财务") ||
                                        w.Name.Contains("数据·")).ToList();
            Assert.NotEmpty(domain);
            Assert.All(domain.SelectMany(w => w.Steps), s => Assert.True(known.Contains(s.Tool), $"未知工具：{s.Tool}"));
        }

        // ---- 流程"去作用区域"（FlowScopeNormalizer）----

        [Fact]
        public void Descope_rewrites_vba_range_literal_to_selection_marker()
        {
            var wf = NewCustom("改区域", new WorkflowStep
            {
                Tool = "run_vba_snippet",
                Arguments = new Dictionary<string, object>
                {
                    ["snippet"] = "Sub F()\n Range(\"C5:F22\").Font.Bold = True\nEnd Sub"
                }
            });

            var changed = FlowScopeNormalizer.NormalizeTemplate(wf);
            var snippet = (string)wf.Steps[0].Arguments["snippet"];

            Assert.True(changed >= 1);
            Assert.Contains("Range(\"{selection}\")", snippet);
            Assert.DoesNotContain("C5:F22", snippet);
        }

        [Fact]
        public void Descope_rewrites_sheet_qualified_and_dollar_ranges()
        {
            var wf = NewCustom("跨表", new WorkflowStep
            {
                Tool = "run_vba_snippet",
                Arguments = new Dictionary<string, object>
                {
                    ["snippet"] = "Set r = .Range(\"'销售表'!$C$5:$F$22\")\nb = Range(\"$A$1:$D$9\")"
                }
            });

            FlowScopeNormalizer.NormalizeTemplate(wf);
            var snippet = (string)wf.Steps[0].Arguments["snippet"];

            Assert.Contains("Range(\"{selection}\")", snippet);
            Assert.DoesNotContain("$C$5:$F$22", snippet);
            Assert.DoesNotContain("$A$1:$D$9", snippet);
        }

        [Fact]
        public void Descope_leaves_cells_and_named_refs_untouched()
        {
            var wf = NewCustom("不碰", new WorkflowStep
            {
                Tool = "run_vba_snippet",
                Arguments = new Dictionary<string, object>
                {
                    ["snippet"] = "Cells(5,3).Value = 1\nRange(\"DataTable\").ClearContents\nDebug.Print \"C5:F22 是描述文字\""
                }
            });

            FlowScopeNormalizer.NormalizeTemplate(wf);
            var snippet = (string)wf.Steps[0].Arguments["snippet"];

            Assert.Contains("Cells(5,3)", snippet);
            Assert.Contains("DataTable", snippet);
            // VBA Range 上下文只替换 Range("区域") 写法；无匹配时不应把描述里的 C5:F22 当区域改掉
            Assert.Contains("C5:F22 是描述文字", snippet);
        }

        [Fact]
        public void Descope_rewrites_plain_address_arg_and_nested_containers()
        {
            var wf = NewCustom("地址参数", new WorkflowStep
            {
                Tool = "format_excel_range",
                Arguments = new Dictionary<string, object>
                {
                    ["range"] = "D4:H10",
                    ["nested"] = new Dictionary<string, object>
                    {
                        ["source_address"] = "D4:H10",
                        ["list"] = new List<object> { "D4:H10", 3, "x" }
                    }
                }
            });

            var changed = FlowScopeNormalizer.NormalizeTemplate(wf);
            var args = wf.Steps[0].Arguments;

            Assert.True(changed >= 3);
            Assert.Equal("{selection}", Convert.ToString(args["range"]));
            Assert.Equal("{selection}", Convert.ToString(((IDictionary<string, object>)args["nested"])["source_address"]));
            var list = (List<object>)((IDictionary<string, object>)args["nested"])["list"];
            Assert.Equal("{selection}", Convert.ToString(list[0]));
            Assert.Equal("x", Convert.ToString(list[2]));
        }

        [Fact]
        public void Descope_also_rewrites_verify_arguments_and_selection_is_detected()
        {
            // 自检步骤的固化区域参数同样要被去作用域，且 TemplateReferencesSelection 要能看到 Verify 里的选区。
            var wf = NewCustom("自检作用域", new WorkflowStep
            {
                Tool = "create_excel_pivot_summary",
                Arguments = new Dictionary<string, object> { ["source_address"] = "C5:F22" },
                Verify = new WorkflowVerify
                {
                    Tool = "verify_pivot",
                    Arguments = new Dictionary<string, object>
                    {
                        ["target_address"] = "B2:D9",
                        ["expected_rows"] = new List<object> { new Dictionary<string, object> { ["label"] = "华东" } }
                    }
                }
            });

            var changed = FlowScopeNormalizer.NormalizeTemplate(wf);
            var verifyArgs = wf.Steps[0].Verify.Arguments;

            Assert.True(changed >= 2);
            Assert.Equal("{selection}", Convert.ToString(wf.Steps[0].Arguments["source_address"]));
            Assert.Equal("{selection}", Convert.ToString(verifyArgs["target_address"]));
            // 仅整串为地址才改写：对象里的 label 文本不应被误伤
            var rows = (List<object>)verifyArgs["expected_rows"];
            Assert.Equal("华东", Convert.ToString(((IDictionary<string, object>)rows[0])["label"]));
            Assert.True(FlowScopeNormalizer.TemplateReferencesSelection(wf));

            var noSel = NewCustom("无选区自检", new WorkflowStep
            {
                Tool = "act",
                Verify = new WorkflowVerify { Tool = "verify_pivot", Arguments = new Dictionary<string, object> { ["target_address"] = "B2:D9" } }
            });
            Assert.False(FlowScopeNormalizer.TemplateReferencesSelection(noSel));
        }

        [Fact]
        public void Descope_ignores_builtin_flows()
        {
            var wf = new WorkflowTemplate
            {
                Name = "内置",
                IsBuiltIn = true,
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "run_vba_snippet",
                        Arguments = new Dictionary<string, object> { ["snippet"] = "Range(\"C5:F22\")" }
                    }
                }
            };

            var changed = FlowScopeNormalizer.NormalizeTemplate(wf);
            Assert.Equal(0, changed);
            Assert.Contains("C5:F22", (string)wf.Steps[0].Arguments["snippet"]);
        }

        [Fact]
        public void Executor_resolves_injected_selection_marker_and_user_override_wins()
        {
            string snippet = "Range(\"{selection}\").Select";
            var overrides = new Dictionary<string, object> { ["selection"] = "$C$5:$F$22" };
            var resolved = WorkflowExecutor.ResolveArguments(
                new Dictionary<string, object> { ["snippet"] = snippet, ["t"] = "x" }, overrides);

            Assert.Equal("Range(\"$C$5:$F$22\").Select", Convert.ToString(resolved["snippet"]));
            Assert.Equal("x", Convert.ToString(resolved["t"]));
        }

        [Fact]
        public void Template_references_selection_is_detected()
        {
            var wf = NewCustom("含选区", new WorkflowStep
            {
                Tool = "run_vba_snippet",
                Arguments = new Dictionary<string, object> { ["snippet"] = "Range(\"{selection}\")" }
            });
            Assert.True(FlowScopeNormalizer.TemplateReferencesSelection(wf));

            var plain = NewCustom("无选区", new WorkflowStep
            {
                Tool = "run_vba_snippet",
                Arguments = new Dictionary<string, object> { ["snippet"] = "Range(\"C5:F22\")" }
            });
            Assert.False(FlowScopeNormalizer.TemplateReferencesSelection(plain));
        }

        private static WorkflowTemplate NewCustom(string name, params WorkflowStep[] steps)
        {
            return new WorkflowTemplate
            {
                Name = name,
                IsBuiltIn = false,
                Steps = steps.ToList()
            };
        }

        public void Dispose()
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}