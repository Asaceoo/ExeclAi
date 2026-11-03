using ExcelAiAssistant.Core.Workflow;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>「插入选区 → 作为代码作用区域」解析逻辑单测。</summary>
    public class FlowScopeNormalizerTests
    {
        [Fact]
        public void ScopeMarker_extracts_multi_sheet_union()
        {
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6,Sheet2!$A$1:$E$5#\nSub F()\n Range(\"C5:F22\")\nEnd Sub";
            Assert.True(FlowScopeNormalizer.TryExtractScopeAddress(code, out var u));
            Assert.Equal("Sheet1!$B$2:$D$6,Sheet2!$A$1:$E$5", u);
        }

        [Fact]
        public void ScopeMarker_replaces_largest_block_only_keeps_smaller_and_single_cell()
        {
            // C5:F22 = 18*4=72；J5:P30 = 26*7=182 → 替换 J5:P30（较大），C5:F22 不动，A1 单格跳过。
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6#\nSub F()\n" +
                                " Range(\"C5:F22\").ClearContents\n" +
                                " Range(\"A1\").Value = 1\n" +
                                " Range(\"J5:P30\").Value = 2\n" +
                                "End Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.True(o.HasMarker);
            Assert.False(o.NoDataBlock);
            Assert.Contains("Range(\"Sheet1!$B$2:$D$6\")", o.Code);
            Assert.Contains("Range(\"C5:F22\")", o.Code);
            Assert.Contains("Range(\"A1\")", o.Code);
            Assert.DoesNotContain("#XAI_SCOPE:", o.Code);
        }

        [Fact]
        public void ScopeMarker_with_only_single_cell_reports_no_data_block()
        {
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6#\nSub F()\n Range(\"A1\").Value = 1\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.True(o.HasMarker);
            Assert.True(o.NoDataBlock);
        }

        [Fact]
        public void ScopeMarker_no_range_reports_no_data_block()
        {
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6#\nSub F()\n  MsgBox \"hi\"\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.True(o.HasMarker);
            Assert.True(o.NoDataBlock);
        }

        [Fact]
        public void ScopeMarker_overrides_selection_placeholder()
        {
            // 去作用域模板里 Range("{selection}") 即作用区域 → 标记优先，用并集覆盖，并移除标记。
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6#\nSub F()\n Range(\"{selection}\").Font.Bold = True\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.True(o.HasMarker);
            Assert.False(o.NoDataBlock);
            Assert.Contains("Range(\"Sheet1!$B$2:$D$6\")", o.Code);
            Assert.DoesNotContain("{selection}", o.Code);
            Assert.DoesNotContain("#XAI_SCOPE:", o.Code);
        }

        [Fact]
        public void ScopeMarker_marker_line_removed_after_apply()
        {
            const string code = "'#XAI_SCOPE: Sheet1!$B$2:$D$6#\nSub F()\n Range(\"C5:F22\")\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.True(o.Applied);
            Assert.StartsWith("Sub F()", o.Code.TrimStart());
            Assert.DoesNotContain("#XAI_SCOPE:", o.Code);
        }

        [Fact]
        public void No_marker_returns_code_untouched()
        {
            const string code = "Sub F()\n Range(\"C5:F22\")\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.False(o.HasMarker);
            Assert.Null(o.UnionAddress);
            Assert.Equal(code, o.Code);
        }

        [Fact]
        public void ScopeMarker_preserves_dot_range_and_qualified_prefix()
        {
            // ActiveSheet.Range("...")：.Range 与限定前缀都须保留，仅替换数据块地址。
            const string code = "'#XAI_SCOPE: 'Sheet1'!$B$2:$D$6#\nSub F()\n ActiveSheet.Range(\"C5:F22\").ClearContents\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.False(o.NoDataBlock);
            Assert.Contains("ActiveSheet.Range(\"'Sheet1'!$B$2:$D$6\")", o.Code);
        }

        [Fact]
        public void ScopeMarker_escapes_double_quote_in_sheet_name()
        {
            // Excel 允许表名含双引号；注入 VBA 双引号字符串前须转义为 ""，避免拆断字符串边界。
            const string code = "'#XAI_SCOPE: 'Sheet\"1'!$B$2:$D$6#\nSub F()\n Range(\"C5:F22\")\nEnd Sub";
            var o = FlowScopeNormalizer.ResolveScope(code);
            Assert.False(o.NoDataBlock);
            Assert.Contains("Range(\"'Sheet\"\"1'!$B$2:$D$6\")", o.Code);
        }
    }
}