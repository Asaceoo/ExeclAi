using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class VbaAutoSinkPolicyTests
    {
        private static ToolEventRecord Ev(string tool, bool ok)
        {
            return new ToolEventRecord
            {
                Tool = tool,
                Result = new ToolResult { Ok = ok }
            };
        }

        [Fact]
        public void Auto_sinks_after_successful_run_vba_without_save_vba_tool()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("run_vba_snippet", true)
            };
            Assert.True(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }

        [Fact]
        public void Does_not_sink_when_run_vba_failed()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("run_vba_snippet", false)
            };
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }

        [Fact]
        public void Does_not_sink_when_no_run_vba_this_turn()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("write_excel_range", true)
            };
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }

        [Fact]
        public void Does_not_sink_when_last_snippet_is_blank()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("run_vba_snippet", true)
            };
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, "   "));
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, null));
        }

        [Fact]
        public void Does_not_sink_when_save_vba_tool_was_called_even_if_failed()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("run_vba_snippet", true),
                Ev("save_vba_tool", false)
            };
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }

        [Fact]
        public void Does_not_sink_when_save_vba_tool_succeeded()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("run_vba_snippet", true),
                Ev("save_vba_tool", true)
            };
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }

        [Fact]
        public void Null_events_never_sink()
        {
            Assert.False(VbaAutoSinkPolicy.ShouldAutoSink(null, "Sub x(): End Sub"));
        }

        [Fact]
        public void Tool_name_comparison_is_case_insensitive()
        {
            var events = new List<ToolEventRecord>
            {
                Ev("RUN_VBA_SNIPPET", true)
            };
            Assert.True(VbaAutoSinkPolicy.ShouldAutoSink(events, "Sub x(): End Sub"));
        }
    }
}
