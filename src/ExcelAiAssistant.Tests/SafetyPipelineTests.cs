using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Safety;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class SafetyPipelineTests
    {
        [Fact]
        public void Read_tools_are_allowed_without_confirmation()
        {
            var pipeline = new SafetyPipeline(new SafetyOptions());
            var tool = Create(RiskLevel.ReadOnly);

            var decision = pipeline.Evaluate(tool);

            Assert.True(decision.Allowed);
            Assert.False(decision.RequiresConfirmation);
        }

        [Fact]
        public void Normal_write_requires_confirmation_by_default()
        {
            var pipeline = new SafetyPipeline(new SafetyOptions());
            var decision = pipeline.Evaluate(Create(RiskLevel.NormalWrite));

            Assert.True(decision.Allowed);
            Assert.True(decision.RequiresConfirmation);
        }

        [Fact]
        public void Dangerous_tools_are_blocked_until_enabled()
        {
            var disabled = new SafetyPipeline(new SafetyOptions());
            Assert.False(disabled.Evaluate(Create(RiskLevel.Dangerous)).Allowed);

            var enabled = new SafetyPipeline(new SafetyOptions { EnableHighRiskTools = true });
            var decision = enabled.Evaluate(Create(RiskLevel.Dangerous));
            Assert.True(decision.Allowed);
            Assert.True(decision.RequiresConfirmation);
        }

        private static ToolDescriptor Create(RiskLevel level)
        {
            return new ToolDescriptor(
                "test-tool",
                "测试",
                level,
                args => new Core.Models.ToolResult());
        }
    }
}
