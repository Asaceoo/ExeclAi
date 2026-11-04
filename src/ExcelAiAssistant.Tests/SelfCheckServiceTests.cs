using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M29：AI 自主闭环三件套的单测（不依赖 COM 的判定与归一化逻辑）：
    /// - VerifyPassed/ExtractVerdict 中央判定（AutoVerifyPolicy 新公共面）
    /// - run_self_check 的计划推导与执行汇总（fake registry，无 COM）
    /// - apply_correction 的修正阶梯语义（含 Dangerous 拒绝、verify 非只读拒绝）
    /// </summary>
    public class SelfCheckServiceTests
    {
        // ===================== 中央判定（AutoVerifyPolicy.VerifyPassed） =====================

        [Fact]
        public void Verify_passed_requires_ok()
        {
            Assert.False(AutoVerifyPolicy.VerifyPassed(null));
            Assert.False(AutoVerifyPolicy.VerifyPassed(new ToolResult { Ok = false }));
            Assert.True(AutoVerifyPolicy.VerifyPassed(new ToolResult { Ok = true }));
        }

        [Fact]
        public void Verify_passed_accepts_ok_verdict_only()
        {
            Assert.True(AutoVerifyPolicy.VerifyPassed(new ToolResult
            {
                Ok = true,
                Data = new JObject { ["verdict"] = "ok" }
            }));

            // 非 ok verdict（no_expected/mismatch/…）一律未通过
            Assert.False(AutoVerifyPolicy.VerifyPassed(new ToolResult
            {
                Ok = true,
                Data = new JObject { ["verdict"] = "no_expected" }
            }));
            Assert.False(AutoVerifyPolicy.VerifyPassed(new ToolResult
            {
                Ok = true,
                Data = new JObject { ["verdict"] = "mismatch" }
            }));
        }

        [Fact]
        public void Extract_verdict_handles_json_string_and_object()
        {
            Assert.Equal("ok", AutoVerifyPolicy.ExtractVerdict(new ToolResult { Ok = true, Data = new JObject { ["verdict"] = "ok" } }));
            Assert.Equal("mismatch", AutoVerifyPolicy.ExtractVerdict(new ToolResult
            {
                Ok = true,
                Data = "{\"verdict\":\"mismatch\"}"
            }));
            // 纯文本 / null：无 verdict
            Assert.Null(AutoVerifyPolicy.ExtractVerdict(new ToolResult { Ok = true, Data = "plain text" }));
            Assert.Null(AutoVerifyPolicy.ExtractVerdict(new ToolResult { Ok = true, Data = null }));
        }

        /// <summary>
        /// M29 真机闭环抓出的缺陷回归锁：verify 家族结果的 Data 是**匿名对象**（new { address, verdict, ... }），
        /// 不是 JObject。ExtractVerdict 必须经 JToken.FromObject 覆盖匿名形态，否则 verdict 取不到 →
        /// VerifyPassed 退回 Ok=true → mismatch 被误判通过（apply_correction 第一轮就假通过）。
        /// </summary>
        [Fact]
        public void Extract_verdict_handles_anonymous_object_data()
        {
            // 真实 verify_write_result 的 Data 形态（匿名对象）
            var data = new
            {
                address = "A1",
                actual_count = 1,
                verdict = "value",
                issues = new[] { "A1：实际值 LEFT，期望值 RIGHT" },
                selfcheck = "自检未通过：写入结果"
            };

            Assert.Equal("value", AutoVerifyPolicy.ExtractVerdict(new ToolResult { Ok = true, Data = data }));
            Assert.False(AutoVerifyPolicy.VerifyPassed(new ToolResult { Ok = true, Data = data }), "匿名对象 verdict=value 必须判未通过");

            // ok verdict 的匿名对象：通过
            var okData = new { address = "A1", verdict = "ok", issues = new string[0] };
            Assert.True(AutoVerifyPolicy.VerifyPassed(new ToolResult { Ok = true, Data = okData }));

            // 无 verdict 的匿名对象：退回 Ok 语义（兼容探测型 verify）
            var noVerdict = new { address = "A1", issues = new string[0] };
            Assert.True(AutoVerifyPolicy.VerifyPassed(new ToolResult { Ok = true, Data = noVerdict }));
        }

        // ===================== run_self_check（fake registry） =====================

        private static ToolRegistry BuildRegistry(params ToolDescriptor[] extra)
        {
            var registry = new ToolRegistry();
            // 测试替身 verify：只要拿到 expected_values 就判 ok（真实比对逻辑在 VerifyJudgment* 测试族覆盖）
            registry.Register(new ToolDescriptor(
                "verify_write_result",
                "回读核对（测试替身）",
                RiskLevel.ReadOnly,
                args =>
                {
                    var hasExpected = args != null && args.ContainsKey("expected_values") && args["expected_values"] != null;
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new JObject
                        {
                            ["verdict"] = hasExpected ? "ok" : "no_expected",
                            ["expected_present"] = hasExpected
                        }
                    };
                }));
            foreach (var tool in extra)
            {
                registry.Register(tool);
            }

            return registry;
        }

        private static Dictionary<string, object> WriteCellArgs(string address, object value)
        {
            return new Dictionary<string, object>
            {
                ["sheet"] = "数据",
                ["address"] = address,
                ["value"] = value
            };
        }

        [Fact]
        public void Run_self_check_executes_plans_and_aggregates()
        {
            var registry = BuildRegistry();
            var service = new ExcelSelfCheckService(registry);

            var result = service.RunSelfCheck(new Dictionary<string, object>
            {
                ["tool"] = "write_excel_cell",
                ["args"] = new Dictionary<string, object>
                {
                    ["sheet"] = "数据",
                    ["address"] = "A1",
                    ["value"] = "v"
                }
            });

            Assert.True(result.Ok);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            Assert.Equal(1, data["total"]);
            Assert.Equal(1, data["passed"]);
            Assert.Equal(0, data["failed"]);
        }

        [Fact]
        public void Run_self_check_rejects_verify_tool_as_primary()
        {
            var service = new ExcelSelfCheckService(BuildRegistry());
            var result = service.RunSelfCheck(new Dictionary<string, object> { ["tool"] = "verify_write_result" });
            Assert.False(result.Ok);
            Assert.Equal("INVALID_ARGS", result.ErrorCode);
        }

        [Fact]
        public void Run_self_check_rejects_unplanned_tool()
        {
            var service = new ExcelSelfCheckService(BuildRegistry());
            var result = service.RunSelfCheck(new Dictionary<string, object>
            {
                ["tool"] = "run_vba_snippet", // 豁免清单成员：无确定性计划
                ["args"] = new Dictionary<string, object>()
            });
            Assert.False(result.Ok);
            Assert.Equal("NO_VERIFY_PLAN", result.ErrorCode);
        }

        [Fact]
        public void Run_self_check_skips_non_readonly_verify()
        {
            // verify 工具被标成 NormalWrite（异常形态）：必须跳过而非执行
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "verify_write_result",
                "被污染的 verify",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = true, Data = new JObject { ["verdict"] = "ok" } }));
            var service = new ExcelSelfCheckService(registry);

            var result = service.RunSelfCheck(new Dictionary<string, object>
            {
                ["tool"] = "write_excel_cell",
                ["args"] = WriteCellArgs("A1", "v")
            });

            Assert.False(result.Ok);
            Assert.Equal("SELF_CHECK_SKIPPED", result.ErrorCode);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            Assert.Equal(1, data["skipped"]);
        }

        // ===================== apply_correction（fake registry） =====================

        private static ToolDescriptor OkEchoTool(string name, RiskLevel risk = RiskLevel.NormalWrite)
        {
            return new ToolDescriptor(
                name,
                "测试替身",
                risk,
                args => new ToolResult { Ok = true, Data = new JObject { ["done"] = true } });
        }

        private static ToolDescriptor FailingTool(string name, string errorCode = "TOOL_FAILED")
        {
            return new ToolDescriptor(
                name,
                "测试替身（总失败）",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = false, ErrorCode = errorCode, Data = "always fails" });
        }

        [Fact]
        public void Apply_correction_stops_on_first_pass()
        {
            var registry = BuildRegistry(
                FailingTool("write_tool_a"),
                OkEchoTool("write_tool_b"));
            var service = new ExcelSelfCheckService(registry);

            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>
                {
                    new Dictionary<string, object> { ["tool"] = "write_tool_a" },
                    new Dictionary<string, object> { ["tool"] = "write_tool_b" }
                }
            });

            Assert.True(result.Ok);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            Assert.Equal(2, data["passed_attempt"]); // 第 2 条通过
            var attempts = System.Linq.Enumerable.ToList(Assert.IsAssignableFrom<System.Collections.Generic.IEnumerable<object>>(data["attempts"]));
            Assert.Equal(2, attempts.Count); // stop_on_first_pass：不再执行第 3 条（本来也没有）
        }

        [Fact]
        public void Apply_correction_reports_failure_when_all_fail()
        {
            var registry = BuildRegistry(FailingTool("write_tool_a"), FailingTool("write_tool_b"));
            var service = new ExcelSelfCheckService(registry);

            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>
                {
                    new Dictionary<string, object> { ["tool"] = "write_tool_a" },
                    new Dictionary<string, object> { ["tool"] = "write_tool_b" }
                }
            });

            Assert.False(result.Ok);
            Assert.Equal("CORRECTION_FAILED", result.ErrorCode);
        }

        [Fact]
        public void Apply_correction_rejects_dangerous_sub_tool()
        {
            var registry = BuildRegistry(OkEchoTool("dangerous_tool", RiskLevel.Dangerous));
            var service = new ExcelSelfCheckService(registry);

            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>
                {
                    new Dictionary<string, object> { ["tool"] = "dangerous_tool" }
                }
            });

            Assert.False(result.Ok);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            var attempts = System.Linq.Enumerable.ToList(Assert.IsAssignableFrom<System.Collections.Generic.IEnumerable<object>>(data["attempts"]));
            var first = Assert.IsAssignableFrom<Dictionary<string, object>>(attempts[0]);
            Assert.Equal("CONFIRMATION_REQUIRED", first["error_code"]);
        }

        [Fact]
        public void Apply_correction_rejects_non_readonly_verify()
        {
            var registry = new ToolRegistry();
            registry.Register(OkEchoTool("write_tool_a"));
            registry.Register(new ToolDescriptor(
                "verify_fake",
                "非只读 verify",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = true, Data = new JObject { ["verdict"] = "ok" } }));
            var service = new ExcelSelfCheckService(registry);

            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["tool"] = "write_tool_a",
                        ["verify_tool"] = "verify_fake"
                    }
                }
            });

            Assert.False(result.Ok); // 主工具 ok 但 verify 被拒 → 未通过
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            var attempts = System.Linq.Enumerable.ToList(Assert.IsAssignableFrom<System.Collections.Generic.IEnumerable<object>>(data["attempts"]));
            var first = Assert.IsAssignableFrom<Dictionary<string, object>>(attempts[0]);
            Assert.Equal("verify 工具非只读，拒绝执行", first["verify_skipped"]);
        }

        [Fact]
        public void Apply_correction_respects_max_attempts()
        {
            var registry = BuildRegistry(OkEchoTool("write_tool_a"));
            var service = new ExcelSelfCheckService(registry);

            var attempts = new List<object>();
            for (var i = 0; i < 8; i++)
            {
                attempts.Add(new Dictionary<string, object> { ["tool"] = "write_tool_a" });
            }

            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = attempts,
                ["max_attempts"] = 3
            });

            // 第 1 条就通过且 stop_on_first_pass=true → 只执行 1 条
            Assert.True(result.Ok);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            Assert.Equal(1, data["passed_attempt"]);
        }

        [Fact]
        public void Apply_correction_rejects_empty_attempts()
        {
            var service = new ExcelSelfCheckService(BuildRegistry());
            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>()
            });
            Assert.False(result.Ok);
            Assert.Equal("INVALID_ARGS", result.ErrorCode);
        }

        [Fact]
        public void Apply_correction_unknown_tool_is_recorded_not_thrown()
        {
            var service = new ExcelSelfCheckService(BuildRegistry());
            var result = service.ApplyCorrection(new Dictionary<string, object>
            {
                ["attempts"] = new List<object>
                {
                    new Dictionary<string, object> { ["tool"] = "no_such_tool" }
                }
            });

            Assert.False(result.Ok);
            var data = Assert.IsAssignableFrom<Dictionary<string, object>>(result.Data);
            var attempts = System.Linq.Enumerable.ToList(Assert.IsAssignableFrom<System.Collections.Generic.IEnumerable<object>>(data["attempts"]));
            var first = Assert.IsAssignableFrom<Dictionary<string, object>>(attempts[0]);
            Assert.Equal("UNKNOWN_TOOL", first["error_code"]);
        }

        // ===================== JObject 归一化（MCP 入参形态直通） =====================

        [Fact]
        public void Run_self_check_accepts_jobject_args_from_mcp()
        {
            // MCP 层 args 会以 JObject/JToken 形态传入（HandleToolsCall 的 ToObject<object>()）
            var registry = BuildRegistry();
            var service = new ExcelSelfCheckService(registry);

            var result = service.RunSelfCheck(new Dictionary<string, object>
            {
                ["tool"] = "write_excel_cell",
                ["args"] = new JObject
                {
                    ["sheet"] = "数据",
                    ["address"] = "A1",
                    ["value"] = "v"
                }
            });

            Assert.True(result.Ok);
        }
    }
}