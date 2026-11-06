using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>本轮修复与新增工具的注册级/纯逻辑测试（不依赖真实 Excel）。</summary>
    public sealed class EnhancementToolTests
    {
        private static ToolDescriptor Get(string name)
        {
            var registry = new ExcelToolHost(new object()).CreateRegistry();
            Assert.True(registry.TryGet(name, out var descriptor), $"{name} 已注册");
            return descriptor;
        }

        [Theory]
        [InlineData("rate_change_analysis")]
        [InlineData("cell_comment")]
        [InlineData("data_validation")]
        public void EnhancementTools_are_registered(string toolName)
        {
            var d = Get(toolName);
            Assert.Equal(RiskLevel.NormalWrite, d.RiskLevel);
        }

        [Fact]
        public void DataValidation_requires_address()
        {
            var d = Get("data_validation");
            var result = d.Execute(new Dictionary<string, object>());
            Assert.False(result.Ok);
        }

        [Fact]
        public void CellComment_requires_address()
        {
            var d = Get("cell_comment");
            var result = d.Execute(new Dictionary<string, object> { ["action"] = "get" });
            Assert.False(result.Ok);
        }
    }
}