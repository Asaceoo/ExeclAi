using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>公式解释器 FormulaAnalyzer 纯逻辑测试（可脱离 Excel 运行）。</summary>
    public sealed class FormulaAnalyzerTests
    {
        // ---------- 参数拆分（深度1逗号，忽略字符串与嵌套括号） ----------
        [Fact]
        public void Split_IF_three_arguments()
        {
            var args = FormulaAnalyzer.SplitArgs("E2=\"\",\"\",D2*E2");
            Assert.Equal(3, args.Count);
            Assert.Equal("E2=\"\"", args[0]);
            Assert.Equal("\"\"", args[1]);
            Assert.Equal("D2*E2", args[2]);
        }

        [Fact]
        public void Split_ignores_comma_inside_string_literal()
        {
            var args = FormulaAnalyzer.SplitArgs("A,\"a,b\",2");
            Assert.Equal(3, args.Count);
            Assert.Equal("\"a,b\"", args[1]);
        }

        [Fact]
        public void Split_ignores_parentheses_inside_nested_call()
        {
            var args = FormulaAnalyzer.SplitArgs("A1:B10,SUM(B2:B5)*2");
            Assert.Equal(2, args.Count);
            Assert.Equal("SUM(B2:B5)*2", args[1]);
        }

        [Fact]
        public void Split_handles_escaped_quotes_inside_string()
        {
            var args = FormulaAnalyzer.SplitArgs("\"a\"\"b\",1");
            Assert.Equal(2, args.Count);
            Assert.Equal("\"a\"\"b\"", args[0]);
        }

        // ---------- 参数类型分类 ----------
        [Theory]
        [InlineData("\"text\"", ArgKind.Constant)]
        [InlineData("123", ArgKind.Constant)]
        [InlineData("-5.5", ArgKind.Constant)]
        [InlineData("TRUE", ArgKind.Constant)]
        [InlineData("#N/A", ArgKind.Constant)]
        [InlineData("A1", ArgKind.CellRef)]
        [InlineData("$A$1", ArgKind.CellRef)]
        [InlineData("A1:C3", ArgKind.RangeRef)]
        [InlineData("$A$1:$C$3", ArgKind.RangeRef)]
        [InlineData("SUM(B2:B5)", ArgKind.NestedFunc)]
        [InlineData("Sheet2!A1", ArgKind.ExternalRef)]
        [InlineData("[Book1]Sheet1!A1", ArgKind.ExternalRef)]
        [InlineData("Sales", ArgKind.NamedRange)]
        [InlineData("D2*E2", ArgKind.Expression)]
        [InlineData("E2=\"\"", ArgKind.Expression)]
        public void Classify_argument(string text, ArgKind expected)
        {
            Assert.Equal(expected, FormulaAnalyzer.Classify(text));
        }

        // ---------- 顶层函数解析 ----------
        [Fact]
        public void Parse_top_level_IF()
        {
            var node = FormulaAnalyzer.ParseTopLevel("=IF(E2=\"\",\"\",D2*E2)");
            Assert.Equal("IF", node.Name);
            Assert.Equal(3, node.Args.Count);
            Assert.Equal(ArgKind.Expression, node.Args[0].Kind);
            Assert.Equal(ArgKind.Constant, node.Args[1].Kind);
            Assert.Equal(ArgKind.Expression, node.Args[2].Kind);
        }

        [Fact]
        public void Parse_named_range_as_arg()
        {
            var node = FormulaAnalyzer.ParseTopLevel("=SUM(Sales)");
            Assert.Equal("SUM", node.Name);
            Assert.Single(node.Args);
            Assert.Equal(ArgKind.NamedRange, node.Args[0].Kind);
        }

        [Fact]
        public void Parse_range_reference_as_arg()
        {
            var node = FormulaAnalyzer.ParseTopLevel("=SUM(A1:C3)");
            Assert.Equal("SUM", node.Name);
            Assert.Equal(ArgKind.RangeRef, node.Args[0].Kind);
        }

        [Fact]
        public void Parse_empty_pseudo_root_when_no_function()
        {
            var node = FormulaAnalyzer.ParseTopLevel("=A1+B1");
            Assert.Null(node.Name);
            Assert.Single(node.Args);
            Assert.Equal(ArgKind.Expression, node.Args[0].Kind);
        }

        [Fact]
        public void Parse_nested_function_keeps_text_for_drill_down()
        {
            var node = FormulaAnalyzer.ParseTopLevel("=SUMIFS(SUM(B2:B5),A:A,\">10\")");
            Assert.Equal("SUMIFS", node.Name);
            Assert.Equal(ArgKind.NestedFunc, node.Args[0].Kind);
            Assert.Equal("SUM(B2:B5)", node.Args[0].Text);
        }

        // ---------- 引用地址提取（供关联公式追溯/循环检测） ----------
        [Fact]
        public void Extract_cell_refs_returns_simple_addresses()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=D2*E2");
            Assert.Contains(refs, r => string.Equals(r, "D2", System.StringComparison.OrdinalIgnoreCase));
            Assert.Contains(refs, r => string.Equals(r, "E2", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Extract_cell_refs_sees_condition_argument()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=IF(E2=\"\",\"\",D2*E2)");
            Assert.Contains(refs, r => string.Equals(r, "E2", System.StringComparison.OrdinalIgnoreCase));
            Assert.Contains(refs, r => string.Equals(r, "D2", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Extract_cell_refs_ignores_string_literals()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=IF(A1=\"see B2\" ,\"\",1)");
            Assert.DoesNotContain(refs, r => string.Equals(r, "B2", System.StringComparison.OrdinalIgnoreCase));
            Assert.Contains(refs, r => string.Equals(r, "A1", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Canonical_key_normalizes_dollar_and_case()
        {
            Assert.Equal("A1", FormulaAnalyzer.CanonicalKey("$A$1"));
            Assert.Equal("B10", FormulaAnalyzer.CanonicalKey("b10"));
            Assert.Equal("A1:C3", FormulaAnalyzer.CanonicalKey("$A$1:$C$3"));
            Assert.Null(FormulaAnalyzer.CanonicalKey("SUM(1)"));
        }

        [Theory]
        [InlineData("A1", true)]
        [InlineData("$A$1", true)]
        [InlineData("A1:C3", false)]
        [InlineData("SUM(1)", false)]
        public void Is_cell_ref_detects_single_cell(string text, bool expected)
        {
            Assert.Equal(expected, FormulaAnalyzer.IsCellRef(text));
        }

        [Fact]
        public void Cyclic_pair_addresses_are_extracted()
        {
            // A1=B1、B1=A1 的循环引用：两式各自能提取到对方地址，供上层 visited 判断循环
            var refsA = FormulaAnalyzer.ExtractCellRefs("=B1");
            var refsB = FormulaAnalyzer.ExtractCellRefs("=A1");
            Assert.Contains(refsA, r => string.Equals(r, "B1", System.StringComparison.OrdinalIgnoreCase));
            Assert.Contains(refsB, r => string.Equals(r, "A1", System.StringComparison.OrdinalIgnoreCase));
        }

        // ---------- E6 引用变体（真实公式形态） ----------
        [Fact]
        public void Extract_absolute_reference_keeps_dollar_prefix()
        {
            Assert.Contains(FormulaAnalyzer.ExtractCellRefs("=$A$1"), r => r == "$A$1");
        }

        [Fact]
        public void Extract_range_reference_full_region()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=SUM($A$1:$C$3)");
            Assert.Contains(refs, r => r.Replace("$", "") == "A1:C3");
        }

        [Fact]
        public void Extract_whole_column_reference()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=SUM(C:C)");
            Assert.Contains(refs, r => r.Replace("$", "") == "C:C");
        }

        [Fact]
        public void Extract_quoted_sheet_reference_keeps_prefix()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("='带空格表'!B2");
            Assert.Contains(refs, r => r == "'带空格表'!B2");
        }

        [Fact]
        public void Extract_unquoted_sheet_reference()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=Sheet2!A1");
            Assert.Contains(refs, r => r.Replace("$", "") == "Sheet2!A1");
        }

        [Fact]
        public void Extract_cross_book_reference()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=[Book1]Sheet1!A1");
            Assert.Contains(refs, r => r == "[Book1]Sheet1!A1");
        }

        [Fact]
        public void Extract_cross_book_quoted_sheet()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=[Book1]'My Sheet'!B3");
            Assert.Contains(refs, r => r == "[Book1]'My Sheet'!B3");
        }

        [Fact]
        public void Extract_ignores_function_name_not_reference()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=SUM(A1,ROUND(B2,2))");
            Assert.DoesNotContain(refs, r => r == "SUM");
            Assert.Contains(refs, r => r == "A1");
            Assert.Contains(refs, r => r == "B2");
        }

        [Fact]
        public void Extract_r1c1_style_returns_sentinel_not_silent()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=R[1]C[2]*R[-1]C[-1]");
            Assert.Contains(refs, r => r.StartsWith("UNPARSED:"));
        }

        [Fact]
        public void Extract_mixture_keeps_all_valid_and_skips_string()
        {
            var refs = FormulaAnalyzer.ExtractCellRefs("=IF(A1=\"SUM(C3)\",'表'!B1,C:C)");
            Assert.Contains(refs, r => r == "A1");
            Assert.DoesNotContain(refs, r => r == "C3");
            Assert.Contains(refs, r => r == "'表'!B1");
            Assert.Contains(refs, r => r == "C:C");
        }
    }
}