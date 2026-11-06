using ExcelAiAssistant.Core.Agent;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class AgentProtocolParserTests
    {
        [Fact]
        public void Parses_plain_tool_json()
        {
            const string output = "{\"action\":\"tool\",\"tool\":\"preview_excel_range\",\"args\":{\"address\":\"A1:C3\"}}";

            Assert.True(AgentProtocolParser.TryParse(output, out var action));
            Assert.Equal("tool", action.Action);
            Assert.Equal("preview_excel_range", action.ToolCall.Tool);
            Assert.Equal("A1:C3", action.ToolCall.Args["address"]);
        }

        [Fact]
        public void Extracts_tool_json_from_markdown_text()
        {
            const string output = "```json\n{\"action\":\"tool\",\"tool\":\"write_excel_range\"}\n```";

            Assert.True(AgentProtocolParser.TryParse(output, out var action));
            Assert.Equal("write_excel_range", action.ToolCall.Tool);
        }

        [Fact]
        public void Treats_non_protocol_output_as_final()
        {
            Assert.True(AgentProtocolParser.TryParse("这是最终回答。", out var action));
            Assert.Equal("final", action.Action);
            Assert.Equal("这是最终回答。", action.Content);
        }

        [Fact]
        public void Rejects_empty_output()
        {
            Assert.False(AgentProtocolParser.TryParse(" ", out _));
        }
    }
}
