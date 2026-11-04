using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class WebFetchToolTests
    {
        private static ToolDescriptor GetFetchTool()
        {
            var registry = new ExcelToolHost(new object()).CreateRegistry();
            Assert.True(registry.TryGet("fetch_url", out var descriptor), "fetch_url 已注册");
            return descriptor;
        }

        [Fact]
        public void FetchUrl_is_registered_as_readonly()
        {
            var descriptor = GetFetchTool();
            Assert.Equal(RiskLevel.ReadOnly, descriptor.RiskLevel);
        }

        [Theory]
        [InlineData("file:///c:/windows/win.ini")]
        [InlineData("file://localhost/etc/passwd")]
        [InlineData("ftp://example.com/file.txt")]
        [InlineData("data:text/plain,hello")]
        [InlineData("javascript:alert(1)")]
        public void FetchUrl_rejects_non_http_protocol_without_network(string url)
        {
            var descriptor = GetFetchTool();
            var result = descriptor.Execute(new Dictionary<string, object> { ["url"] = url });
            Assert.False(result.Ok);
            Assert.Equal("INVALID_URL", result.ErrorCode);
        }

        [Fact]
        public void FetchUrl_requires_url_argument()
        {
            var descriptor = GetFetchTool();
            var result = descriptor.Execute(new Dictionary<string, object>());
            Assert.False(result.Ok);
            Assert.Equal("INVALID_VALUE", result.ErrorCode);
        }
    }
}