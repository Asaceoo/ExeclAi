using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;

namespace VerifyPhaseH
{
    /// <summary>
    /// 自动冒烟（第 H 批）：P5-B 流程自动化与编排。
    /// 纯文件级 + 反射校验，不依赖 Excel COM。断言：
    ///  1) 内置流程库 ≥3 条且含数据体检 / 出图 / 比对；
    ///  2) WorkflowStore 保存/加载往返、不可覆盖内置、仅删用户流程；
    ///  3) WorkflowExecutor 顺序执行、失败即停、单步重试不重跑成功步骤；
    ///  4) 参数覆盖 {placeholder} 注入；
    ///  5) 危险步骤被确认回调拦截（CancelledIndex 命中，接到受控执行确认）。
    ///  6) SidebarBridge 暴露流程桥接方法 + ExcelToolHost 暴露 run_workflow 注入钩子。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        // 已注册工具全集：所有内置流程步骤必须落到这些已注册工具上，运行才不会因未知工具中断。
        private static readonly HashSet<string> knownTools = new HashSet<string>
        {
            "get_current_excel_context", "get_current_selection_values_and_formulas",
            "get_excel_sheet_list", "get_excel_range_values_and_formulas",
            "analyze_range_statistics", "analyze_value_distribution", "analyze_column_correlation",
            "generate_analysis_report", "recommend_excel_chart", "create_excel_chart",
            "export_excel_chart_image", "compare_two_sheets", "clean_excel_data",
            "create_excel_pivot_summary", "format_excel_range", "filter_excel_data",
            "create_excel_sparkline", "export_sheets_to_pdf",
            "sort_range", "dedup_range", "find_replace_range", "transpose_range",
            "format_excel_conditional", "text_clean", "text_split", "text_merge",
            "date_normalize", "number_pad", "fill_formula", "merge_sheets"
        };

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL PHASE-H WORKFLOW CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            var filePath = Path.Combine(Path.GetTempPath(), "verifyphaseh-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                var store = new WorkflowStore(filePath);
                CheckBuiltInLibrary(store);
                CheckStorePersistence(store, filePath);
                CheckExecutorAndRetry();
                CheckParameterInjection();
                CheckDangerousGate();
                CheckBridgeIntegration();
            }
            finally
            {
                try { if (File.Exists(filePath)) { File.Delete(filePath); } }
                catch { }
            }
        }

        private static void CheckBuiltInLibrary(WorkflowStore store)
        {
            var all = store.GetAll();
            Assert("builtin-count-ge-3", all.Count >= 3, "内置流程至少 3 条");
            Assert("builtin-schema", all.All(w => w.IsBuiltIn), "内置流程均标记 IsBuiltIn");
            Assert("builtin-health-report", all.Any(w => w.Name.Contains("数据体检")), "含「数据体检并生成报告」");
            Assert("builtin-chart", all.Any(w => w.Name.Contains("出图")), "含「选区一键出图」");
            Assert("builtin-compare", all.Any(w => w.Name.Contains("比对")), "含「两工作表比对」");
            Assert("builtin-steps-ordered", all.First(w => w.Name.Contains("数据体检")).Steps[0].Tool == "get_current_excel_context",
                "数据体检流程首步为读取选区");
            Assert("builtin-count-ge-8", all.Count >= 8, "内置流程至少 8 条");
            Assert("builtin-clean", all.Any(w => w.Name.Contains("数据清洗")), "含「数据清洗并复检」");
            Assert("builtin-pivot", all.Any(w => w.Name.Contains("透视")), "含「透视汇总表」");
            Assert("builtin-filter", all.Any(w => w.Name.Contains("筛选")), "含「数据筛选提炼」");
            Assert("builtin-correlation", all.Any(w => w.Name.Contains("迷你图")), "含「相关性+迷你图」");
            Assert("builtin-pdf", all.Any(w => w.Name.Contains("工作报告导出")), "含「工作报告导出」");
            Assert("builtin-count-ge-12", all.Count >= 12, "内置流程至少 12 条");
            Assert("builtin-struct-summary", all.Any(w => w.Name.Contains("数值结构摘要")), "含「数值结构摘要」");
            Assert("builtin-diff-report", all.Any(w => w.Name.Contains("差异报告")), "含「双表差异报告」");
            Assert("builtin-sparkline-glance", all.Any(w => w.Name.Contains("趋势速览")), "含「迷你图趋势速览」");
            Assert("builtin-clean-chart", all.Any(w => w.Name.Contains("清洗后一键出图")), "含「清洗后一键出图」");
            Assert("builtin-count-ge-20", all.Count >= 20, "内置流程至少 20 条");
            // 领域模板：数据分析 / 人事 / 考勤 / 财务
            Assert("builtin-domain-analysis", all.Any(w => w.Name.Contains("数据分析")), "含数据分析领域流程");
            Assert("builtin-domain-hr", all.Any(w => w.Name.Contains("人事")), "含人事领域流程");
            Assert("builtin-domain-attendance", all.Any(w => w.Name.Contains("考勤")), "含考勤领域流程");
            Assert("builtin-domain-finance", all.Any(w => w.Name.Contains("财务")), "含财务领域流程");
            Assert("builtin-steps-known", all.SelectMany(w => w.Steps).All(s => knownTools.Contains(s.Tool)),
                "所有内置流程步骤均引用已注册工具（运行不会因未知工具中断）");
        }

        private static void CheckStorePersistence(WorkflowStore store, string filePath)
        {
            var custom = new WorkflowTemplate
            {
                Name = "我的清洗",
                Description = "清洗并出图",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "clean_excel_data", Description = "清洗" },
                    new WorkflowStep { Tool = "create_excel_chart", Description = "出图" }
                }
            };

            Assert("store-save", store.Save(custom), "保存用户流程成功");
            var reloaded = new WorkflowStore(filePath);
            Assert("store-roundtrip", reloaded.TryGet("我的清洗", out var got) && got.Steps.Count == 2 && got.Steps[1].Tool == "create_excel_chart",
                "保存后可跨实例读回");
            Assert("store-user-not-builtin", !got.IsBuiltIn, "用户流程 IsBuiltIn=false");

            var clone = new WorkflowTemplate { Name = "数据体检并生成报告", Description = "试图覆盖内置" };
            Assert("store-no-override-builtin", !store.Save(clone), "不可覆盖内置流程");
            Assert("store-no-delete-builtin", !store.Delete("数据体检并生成报告"), "不可删除内置流程");
            Assert("store-delete-user", store.Delete("我的清洗") && !store.TryGet("我的清洗", out _), "可删除用户流程");
        }

        private static void CheckExecutorAndRetry()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("good", "ok", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true, Data = "fine" }));
            registry.Register(new ToolDescriptor("flaky", "retryable", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = false, ErrorCode = "E", Data = "boom" }));
            registry.Register(new ToolDescriptor("never", "not reached", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true }));

            var executor = new WorkflowExecutor(registry, (e, r) => Task.FromResult(true));
            var template = new WorkflowTemplate
            {
                Name = "t",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "good" },
                    new WorkflowStep { Tool = "flaky" },
                    new WorkflowStep { Tool = "never" }
                }
            };

            var run = executor.RunAsync(template, null, null, 0).GetAwaiter().GetResult();
            Assert("executor-stops-on-failure", !run.Completed, "流程在失败步骤停止");
            Assert("executor-failed-index", run.FailedIndex == 1, "失败下标指向第 2 步");

            // 单步重试：第 1 次完整运行 flaky 失败，随后从失败步开始重试，
            // 只重跑失败步骤及其后，不重跑已成功步骤 good。
            var goodCalls = 0;
            var flakyCalls = 0;
            var r2 = new ToolRegistry();
            r2.Register(new ToolDescriptor("good", "ok", RiskLevel.ReadOnly,
                _ => { goodCalls++; return new ToolResult { Ok = true }; }));
            r2.Register(new ToolDescriptor("flaky", "retryable", RiskLevel.ReadOnly,
                _ => { flakyCalls++; return flakyCalls >= 2 ? new ToolResult { Ok = true } : new ToolResult { Ok = false }; }));
            r2.Register(new ToolDescriptor("never", "last", RiskLevel.ReadOnly,
                _ => new ToolResult { Ok = true }));

            var e2 = new WorkflowExecutor(r2, (en, rr) => Task.FromResult(true));
            var template2 = new WorkflowTemplate
            {
                Name = "t2",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "good" },
                    new WorkflowStep { Tool = "flaky" },
                    new WorkflowStep { Tool = "never" }
                }
            };

            var first = e2.RunAsync(template2, null, null, 0).GetAwaiter().GetResult();
            Assert("retry-first-fails", !first.Completed && first.FailedIndex == 1 && goodCalls == 1,
                "初次运行在第 2 步失败且 good 执行 1 次");

            var recovered = new List<WorkflowStepRun> { first.Steps[0] };
            var retried = e2.RunAsync(template2, null, recovered, 1).GetAwaiter().GetResult();
            Assert("retry-completes", retried.Completed, "重试后流程完成");
            Assert("retry-no-rerun-success", goodCalls == 1, "已成功步骤未重跑（good 计数=1）");
            Assert("retry-rerun-failed", flakyCalls == 2, "失败步骤第二次调用成功（flaky 计数=2）");
        }

        private static void CheckParameterInjection()
        {
            object received = null;
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("probe", "probe", RiskLevel.ReadOnly,
                args => { received = args.TryGetValue("range", out var v) ? v : null; return new ToolResult { Ok = true }; }));
            var executor = new WorkflowExecutor(registry, (e, r) => Task.FromResult(true));

            var template = new WorkflowTemplate
            {
                Name = "p",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep
                    {
                        Tool = "probe",
                        Arguments = new Dictionary<string, object> { ["range"] = "{target}", ["mode"] = "auto" }
                    }
                }
            };
            var overrides = new Dictionary<string, object> { ["target"] = "A1:C9" };
            var run = executor.RunAsync(template, overrides, null, 0).GetAwaiter().GetResult();
            Assert("param-inject", run.Completed && Convert.ToString(received) == "A1:C9", "占位符 {target} 被覆盖为 A1:C9");
        }

        private static void CheckDangerousGate()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor("del_sheet", "删除工作表", RiskLevel.Dangerous,
                _ => new ToolResult { Ok = true }));

            RiskLevel observed = RiskLevel.ReadOnly;
            var executor = new WorkflowExecutor(registry, (e, r) => { observed = r; return Task.FromResult(false); });
            var template = new WorkflowTemplate
            {
                Name = "d",
                Steps = new List<WorkflowStep>
                {
                    new WorkflowStep { Tool = "del_sheet" }
                }
            };
            var run = executor.RunAsync(template, null, null, 0).GetAwaiter().GetResult();
            Assert("danger-gate-confirm-asked", observed == RiskLevel.Dangerous, "危险工具触发确认回调");
            Assert("danger-gate-cancelled", run.CancelledIndex == 0 && !run.Completed, "确认被拒后流程中止并标记取消");
        }

        private static void CheckBridgeIntegration()
        {
            var bridgeType = Type.GetType("ExcelAiAssistant.Addin.SidebarBridge, ExcelAiAssistant.Addin");
            if (bridgeType == null)
            {
                Fail("bridge-type-found", "未加载 SidebarBridge 类型");
                return;
            }

            var expected = new[] { "ListWorkflows", "SaveWorkflow", "DeleteWorkflow", "RunWorkflow", "RetryWorkflowStep" };
            foreach (var name in expected)
            {
                Assert("bridge-" + name, bridgeType.GetMethod(name) != null, "桥接方法 " + name + " 存在");
            }

            var cycleField = bridgeType.GetField("activeWorkflowNames",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert("bridge-cycle-guard", cycleField != null, "SidebarBridge 持有递归环保护字段 activeWorkflowNames");

            var templateOverload = bridgeType.GetMethod("ExecuteWorkflowInternal",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance,
                null,
                new[] { typeof(WorkflowTemplate), typeof(string), typeof(Dictionary<string, object>), typeof(int), typeof(System.Collections.Generic.List<WorkflowStepRun>) },
                null);

            Assert("bridge-template-lock", templateOverload != null, "SidebarBridge 提供锁定模板的执行过载（重试防步骤错位）");

            var hostType = Type.GetType("ExcelAiAssistant.Excel.ExcelToolHost, ExcelAiAssistant.Excel");
            if (hostType == null)
            {
                Fail("host-type-found", "未加载 ExcelToolHost 类型");
                return;
            }

            var hook = hostType.GetProperty("RunWorkflowTool");
            Assert("host-runworkflow-hook", hook != null,
                "ExcelToolHost 暴露 run_workflow 注入钩子 RunWorkflowTool 属性");

            var schemasType = Type.GetType("ExcelAiAssistant.Excel.WorkflowToolSchemas, ExcelAiAssistant.Excel");
            Assert("schemas-type", schemasType != null, "存在 WorkflowToolSchemas");
        }

        private static void Assert(string name, bool condition, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("PASS " + name + " — " + detail);
            }
            else
            {
                _failed++;
                Failures.Add(name + ": " + detail);
                Console.WriteLine("FAIL " + name + " — " + detail);
            }
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + ": " + detail);
            Console.WriteLine("FAIL " + name + " — " + detail);
        }

        private static string Safe(Exception ex)
        {
            return ex == null ? string.Empty : ex.Message;
        }
    }
}