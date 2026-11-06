using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>跨程序集日志门面 AppLog 的纯逻辑测试：sink 注入、级别/类别透传、接收方异常隔离。</summary>
    public sealed class AppLogTests
    {
        // 每个测试前重置 sink，避免用例间串扰。
        public AppLogTests()
        {
            AppLog.SetSink(null);
        }

        [Fact]
        public void No_sink_is_noop_and_does_not_throw()
        {
            AppLog.Info("hello");
            AppLog.Warn("w");
            AppLog.Error("e");
            AppLog.Info("cat", "m");
            // 不应抛异常即通过。
        }

        [Fact]
        public void Info_writes_with_info_level_and_message()
        {
            LogEntry captured = null;
            // 并行竞态防护：AppLog sink 进程级单例，其他测试类（工具链）可能后台并发写日志，
            // sink 内按目标消息过滤，只捕获本用例自己的条目。
            AppLog.SetSink(e => { if (e.Message == "调试信息") { captured = e; } });
            AppLog.Info("调试信息");
            Assert.NotNull(captured);
            Assert.Equal(CoreLogLevel.Info, captured.Level);
            Assert.Equal("调试信息", captured.Message);
            Assert.Null(captured.Category);
            Assert.True((DateTime.Now - captured.Time).TotalMinutes < 1);
        }

        [Fact]
        public void Warn_and_Error_levels_route_correctly()
        {
            var levels = new List<CoreLogLevel>();
            // 并行竞态防护：同上，按消息过滤只统计本用例条目。
            AppLog.SetSink(e => { if (e.Message == "x" || e.Message == "y") { levels.Add(e.Level); } });
            AppLog.Warn("x");
            AppLog.Error("y");
            Assert.Equal(2, levels.Count);
            Assert.Equal(CoreLogLevel.Warn, levels[0]);
            Assert.Equal(CoreLogLevel.Error, levels[1]);
        }

        [Fact]
        public void Category_flows_through()
        {
            LogEntry captured = null;
            // 并行竞态防护：同上，按 category 过滤。
            AppLog.SetSink(e => { if (e.Category == "python" && e.Message.Contains("boom")) { captured = e; } });
            AppLog.Warn("python", "worker stderr: boom");
            Assert.Equal("python", captured.Category);
            Assert.Equal("worker stderr: boom", captured.Message);
        }

        [Fact]
        public void Throwing_sink_does_not_escape_to_caller()
        {
            AppLog.SetSink(_ => throw new InvalidOperationException("sink 故障"));
            // 接收方抛异常不得向外传播，日志失败的副作用必须被隔离。
            AppLog.Error("故障路径");
        }

        [Fact]
        public void SetSink_null_restores_noop()
        {
            AppLog.SetSink(null);
            AppLog.Info("should not throw");
        }
    }
}