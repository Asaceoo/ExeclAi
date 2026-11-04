using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>AgentRunner 粘合层集成测试：用脚本化假模型客户端驱动完整回合，
    /// 回归断言「主工具 → 自动自检 → 观察消息 → 元监控摘要」编排不被改坏。</summary>
    public sealed class AgentRunnerAutoVerifyIntegrationTests
    {
        private sealed class FakeCompletionClient : IChatCompletionClient
        {
            private readonly Queue<ChatCompletionResult> scripted;

            public FakeCompletionClient(params ChatCompletionResult[] completions)
            {
                scripted = new Queue<ChatCompletionResult>(completions);
            }

            public IReadOnlyList<ChatMessage> LastMessages { get; private set; } =
                Array.Empty<ChatMessage>();

            public Task<ChatCompletionResult> CompleteAsync(
                IReadOnlyList<ChatMessage> messages,
                IReadOnlyList<ToolDescriptor> tools,
                CancellationToken cancellationToken,
                Action<AgentProgress> onProgress)
            {
                LastMessages = messages.ToList();
                return Task.FromResult(scripted.Dequeue());
            }

            public void Dispose()
            {
            }
        }

        private static Dictionary<string, object> Args(params object[] keyValues)
        {
            var map = new Dictionary<string, object>();
            for (var i = 0; i < keyValues.Length; i += 2)
            {
                map[(string)keyValues[i]] = keyValues[i + 1];
            }

            return map;
        }

        private static ChatCompletionResult ToolCallTurn(string callId, string tool, Dictionary<string, object> args)
        {
            return new ChatCompletionResult
            {
                ToolCalls = new[] { new ToolCallRequest { Id = callId, Name = tool, Arguments = args } }
            };
        }

        private static ToolRegistry RegistryWithVerify()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "write_excel_range",
                "写入",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = true, Data = new { written = true } }));
            registry.Register(new ToolDescriptor(
                "verify_write_result",
                "写入自检",
                RiskLevel.ReadOnly,
                args => new ToolResult
                {
                    Ok = true,
                    Data = new { verdict = "ok", issues = Array.Empty<string>(), selfcheck = "自检通过" }
                }));
            return registry;
        }

        [Fact]
        public async Task Write_tool_call_triggers_auto_verify_and_flushes_observation()
        {
            var verifyArgs = new List<Dictionary<string, object>>();
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "write_excel_range",
                "写入",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = true, Data = new { written = true } }));
            registry.Register(new ToolDescriptor(
                "verify_write_result",
                "写入自检",
                RiskLevel.ReadOnly,
                args =>
                {
                    verifyArgs.Add(new Dictionary<string, object>(args));
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { verdict = "ok", issues = Array.Empty<string>(), selfcheck = "自检通过" }
                    };
                }));

            var client = new FakeCompletionClient(
                ToolCallTurn("c1", "write_excel_range", Args("address", "A1:B2", "values", new[] { new[] { 1, 2 } })),
                new ChatCompletionResult { Content = "已写入并自检" });
            var runner = new AgentRunner(
                client,
                registry,
                (envelope, risk) => Task.FromResult(ConfirmationOutcome.Confirmed),
                new AgentRunnerOptions { AutoVerify = true });

            var result = await runner.RunAsync(
                new[] { ChatMessage.User("写入数据") },
                CancellationToken.None);

            // 主工具事件 + 自动自检事件都进 ToolEvents / Steps。
            Assert.Equal(2, result.ToolEvents.Count);
            Assert.Equal("write_excel_range", result.ToolEvents[0].Tool);
            Assert.Equal("verify_write_result", result.ToolEvents[1].Tool);
            Assert.True(result.ToolEvents[1].Result.Ok);
            Assert.Equal(2, result.Steps.Count(step => step.Type == "tool"));

            // 自检预期参数来自主工具实参：地址透传、类型感知默认开启。
            var planArgs = Assert.Single(verifyArgs);
            Assert.Equal("A1:B2", planArgs["address"]);
            Assert.True((bool)planArgs["type_aware"]);

            // 自动自检观察消息在下一跳前下发给模型。
            var observation = client.LastMessages.Last(message => message.Role == "user");
            Assert.Contains("auto_verify", observation.Content);
            Assert.Contains("verify_write_result", observation.Content);

            // 元监控摘要：1 个可自检操作、1 条计划执行。
            Assert.NotNull(result.AutoVerifySummary);
            Assert.Contains("1 个可自检写操作", result.AutoVerifySummary);
            Assert.Contains("执行 1 条", result.AutoVerifySummary);
        }

        [Fact]
        public async Task Missing_verify_tool_raises_link_suspect_summary()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "write_excel_range",
                "写入",
                RiskLevel.NormalWrite,
                args => new ToolResult { Ok = true, Data = new { written = true } }));
            // 刻意不注册 verify_write_result：自检链路失效场景。

            var client = new FakeCompletionClient(
                ToolCallTurn("c1", "write_excel_range", Args("address", "A1", "values", new[] { new[] { 1 } })),
                new ChatCompletionResult { Content = "完成" });
            var runner = new AgentRunner(
                client,
                registry,
                (envelope, risk) => Task.FromResult(ConfirmationOutcome.Confirmed),
                new AgentRunnerOptions { AutoVerify = true });

            var result = await runner.RunAsync(
                new[] { ChatMessage.User("写入数据") },
                CancellationToken.None);

            Assert.Single(result.ToolEvents);
            Assert.NotNull(result.AutoVerifySummary);
            Assert.Contains("自检链路疑似失效", result.AutoVerifySummary);
        }

        [Fact]
        public async Task Read_only_turn_has_no_auto_verify_summary()
        {
            var registry = new ToolRegistry();
            registry.Register(new ToolDescriptor(
                "preview_range",
                "预览",
                RiskLevel.ReadOnly,
                args => new ToolResult { Ok = true, Data = "ok" }));

            var client = new FakeCompletionClient(
                ToolCallTurn("c1", "preview_range", Args("address", "A1")),
                new ChatCompletionResult { Content = "看完 了" });
            var runner = new AgentRunner(
                client,
                registry,
                (envelope, risk) => Task.FromResult(ConfirmationOutcome.Confirmed),
                new AgentRunnerOptions { AutoVerify = true });

            var result = await runner.RunAsync(
                new[] { ChatMessage.User("预览") },
                CancellationToken.None);

            Assert.Single(result.ToolEvents);
            Assert.Null(result.AutoVerifySummary);
        }

        [Fact]
        public async Task Auto_verify_disabled_produces_suspect_free_summary()
        {
            // AutoVerify 关闭时：无自检计划 → 不计可自检操作 → 摘要为 null（不告警，属显式关闭）。
            var registry = RegistryWithVerify();
            var client = new FakeCompletionClient(
                ToolCallTurn("c1", "write_excel_range", Args("address", "A1", "values", new[] { new[] { 1 } })),
                new ChatCompletionResult { Content = "完成" });
            var runner = new AgentRunner(
                client,
                registry,
                (envelope, risk) => Task.FromResult(ConfirmationOutcome.Confirmed),
                new AgentRunnerOptions { AutoVerify = false });

            var result = await runner.RunAsync(
                new[] { ChatMessage.User("写入数据") },
                CancellationToken.None);

            Assert.Single(result.ToolEvents);
            Assert.Null(result.AutoVerifySummary);
        }
    }
}
