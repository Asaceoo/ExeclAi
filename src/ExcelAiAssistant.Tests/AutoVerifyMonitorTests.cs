using ExcelAiAssistant.Core.Agent;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>自动自检的元自检：覆盖率统计与链路失效告警。</summary>
    public sealed class AutoVerifyMonitorTests
    {
        [Fact]
        public void No_self_checkable_ops_yields_no_summary()
        {
            var monitor = new AutoVerifyMonitor();
            Assert.Null(monitor.BuildSummary());
        }

        [Fact]
        public void Healthy_round_reports_coverage()
        {
            var monitor = new AutoVerifyMonitor();
            monitor.RecordSelfCheckableOp();
            monitor.RecordOutcome(executed: 2, skipped: 0);
            monitor.RecordSelfCheckableOp();
            monitor.RecordOutcome(executed: 1, skipped: 1);

            var summary = monitor.BuildSummary();
            Assert.NotNull(summary);
            Assert.Contains("2 个可自检写操作", summary);
            Assert.Contains("执行 3 条", summary);
            Assert.Contains("跳过 1 条", summary);
            Assert.DoesNotContain("疑似失效", summary);
        }

        [Fact]
        public void Zero_executed_plans_raises_link_suspect()
        {
            var monitor = new AutoVerifyMonitor();
            monitor.RecordSelfCheckableOp();
            monitor.RecordOutcome(executed: 0, skipped: 2);

            var summary = monitor.BuildSummary();
            Assert.NotNull(summary);
            Assert.Contains("自检链路疑似失效", summary);
            Assert.Contains("0 条自检计划", summary);
        }

        [Fact]
        public void Skipped_only_still_counts_as_suspect()
        {
            // 全部计划被跳过（verify 工具缺失/富化丢弃）等于自检没跑，必须告警。
            var monitor = new AutoVerifyMonitor();
            monitor.RecordSelfCheckableOp();
            monitor.RecordOutcome(executed: 0, skipped: 1);
            Assert.Contains("疑似失效", monitor.BuildSummary());
        }
    }
}
