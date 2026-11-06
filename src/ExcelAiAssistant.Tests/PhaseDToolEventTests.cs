using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ToolEventViewBuilderTests
    {
        [Fact]
        public void Success_is_not_retryable()
        {
            var view = Build(new ToolEventRecord
            {
                Tool = "write_excel_cell",
                Args = Args("address", "A1"),
                Result = new ToolResult { Ok = true, Data = "ok" },
                Confirmed = true
            });

            Assert.True(view["ok"].Value<bool>());
            Assert.False(view["retryable"].Value<bool>());
        }

        [Fact]
        public void Failure_with_args_is_retryable()
        {
            var view = Build(new ToolEventRecord
            {
                Tool = "write_excel_cell",
                Args = Args("address", "A1"),
                Result = new ToolResult { Ok = false, ErrorCode = "TOOL_FAILED", Data = "failed" },
                Confirmed = true
            });

            Assert.False(view["ok"].Value<bool>());
            Assert.True(view["retryable"].Value<bool>());
            Assert.Equal("write_excel_cell", view["tool"].Value<string>());
        }

        [Fact]
        public void Failure_without_args_is_not_retryable()
        {
            var view = Build(new ToolEventRecord
            {
                Tool = "unknown",
                Args = null,
                Result = new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL" },
                Confirmed = true
            });

            Assert.False(view["retryable"].Value<bool>());
        }

        [Fact]
        public void Null_record_returns_null()
        {
            Assert.Null(ToolEventViewBuilder.Build(null));
        }

        private static JObject Build(ToolEventRecord record)
        {
            return JObject.FromObject(ToolEventViewBuilder.Build(record));
        }

        private static Dictionary<string, object> Args(params object[] keyValues)
        {
            var map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < keyValues.Length; i += 2)
            {
                map[Convert.ToString(keyValues[i])] = keyValues[i + 1];
            }

            return map;
        }
    }
}