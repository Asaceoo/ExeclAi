using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public class ToolSummaryBuilderTests
    {
        [Fact]
        public void Verify_chart_ok_renders_human_readable_summary()
        {
            var data = new
            {
                chart_found = true,
                chart_name = "Chart 1",
                chart_type = 51,
                chart_type_name = "柱形图",
                series_count = 6,
                empty_series_count = 0,
                point_total = 90,
                verdict = "ok",
                title_ok = true,
                data_source_ok = true,
                data_rendered = true
            };
            var s = ToolSummaryBuilder.Summarize("verify_chart", data);
            Assert.Contains("自检通过", s);
            Assert.Contains("柱形图", s);
            Assert.Contains("6 系列", s);
            Assert.Contains("90 点", s);
            Assert.DoesNotContain("{", s);
        }

        [Fact]
        public void Verify_chart_failure_flags_title_and_source()
        {
            var data = new
            {
                chart_found = true,
                chart_type_name = "柱形图",
                series_count = 1,
                point_total = 7,
                verdict = "title+source",
                title_ok = false,
                data_source_ok = false,
                data_rendered = true
            };
            var s = ToolSummaryBuilder.Summarize("verify_chart", data);
            Assert.Contains("自检未通过(title+source)", s);
            Assert.Contains("标题不符", s);
            Assert.Contains("数据源不符", s);
        }

        [Fact]
        public void New_verify_tools_render_verdict_head_with_issue_count()
        {
            var data = new { verdict = "ok", issues = new List<object>() };
            Assert.Equal("写入自检通过", ToolSummaryBuilder.Summarize("verify_write_result", data));

            var fail = new { verdict = "date+blank", issues = new List<object> { "a", "b" } };
            var s = ToolSummaryBuilder.Summarize("verify_data_clean", fail);
            Assert.Contains("清洗自检未通过(date+blank)", s);
            Assert.Contains("2 项问题", s);

            // verify_pivot 命中专用摘要而非「N 行汇总」的 pivot 创建样式
            var pivot = ToolSummaryBuilder.Summarize("verify_pivot", new { verdict = "count", issues = new List<object> { "x" } });
            Assert.Contains("透视自检未通过(count)", pivot);
            Assert.DoesNotContain("行汇总", pivot);

            Assert.Contains("标记自检通过", ToolSummaryBuilder.Summarize("verify_marker", data));
        }

        [Fact]
        public void Create_chart_renders_name_type_and_source()
        {
            var data = new
            {
                chart_name = "Chart 3",
                chart_type = "ColumnClustered",
                source_address = "A2:B8",
                title = "销售趋势"
            };
            var s = ToolSummaryBuilder.Summarize("create_excel_chart", data);
            Assert.Contains("已创建图表", s);
            Assert.Contains("Chart 3", s);
            Assert.Contains("ColumnClustered", s);
            Assert.Contains("数据源 A2:B8", s);
        }

        [Fact]
        public void Sparkline_renders_type_and_dest()
        {
            var data = new { type = "line", data_range = "A2:A9", dest_range = "B2", created = true };
            var s = ToolSummaryBuilder.Summarize("create_excel_sparkline", data);
            Assert.Contains("已插入", s);
            Assert.Contains("line 迷你图", s);
            Assert.Contains("B2", s);
        }

        [Fact]
        public void Generic_fallback_prefers_scope_field()
        {
            var data = new { address = "A1:B10", rows = 10, columns = 2, ok = true };
            var s = ToolSummaryBuilder.Summarize("write_excel_range", data);
            Assert.Equal("A1:B10", s);
        }

        [Fact]
        public void Generic_fallback_uses_short_scalars_when_no_scope()
        {
            var data = new { total = 42, matched = 7, note = "done" };
            var s = ToolSummaryBuilder.Summarize("some_unknown_tool", data);
            Assert.Contains("total=42", s);
            Assert.DoesNotContain("{", s);
        }

        [Fact]
        public void String_data_passes_through()
        {
            Assert.Equal("完成", ToolSummaryBuilder.Summarize("any", null));
            Assert.Equal("hello", ToolSummaryBuilder.Summarize("any", "hello"));
            Assert.Equal("完成", ToolSummaryBuilder.Summarize("any", "   "));
        }

        [Fact]
        public void Serialize_detail_returns_full_json()
        {
            var data = new { a = 1, b = "x" };
            var d = ToolSummaryBuilder.SerializeDetail(data);
            Assert.Contains("\"a\":1", d);
            Assert.Contains("\"b\":\"x\"", d);
        }

        [Fact]
        public void ToolEventView_summary_uses_builder_and_detail_carries_json()
        {
            var record = new ToolEventRecord
            {
                Tool = "verify_chart",
                Args = new Dictionary<string, object>(),
                Confirmed = true,
                Result = new ToolResult
                {
                    Ok = true,
                    Data = new { chart_found = true, chart_type_name = "柱形图", series_count = 6, point_total = 90, verdict = "ok" }
                }
            };
            var view = ToolEventView.FromRecord(record);
            Assert.Contains("自检通过", view.Summary);
            Assert.DoesNotContain("{", view.Summary);
            Assert.NotNull(view.Detail);
            Assert.Contains("chart_type_name", view.Detail);
        }
    }
}
