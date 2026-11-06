using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// FormulaDiff 真值锁定（D2 公式历史 diff，纯逻辑确定性）：
    /// 判定优先级：未变化 &gt; 新增/清除 &gt; IFERROR 增删 &gt; 仅引用方式调整 &gt; 函数增删 &gt; 参数/条件调整。
    /// </summary>
    public class FormulaDiffTests
    {
        [Fact]
        public void Same_formula_reports_unchanged()
        {
            Assert.Equal("未变化", FormulaDiff.DescribeChange("=SUM(A1:A10)", "=SUM(A1:A10)"));
            Assert.Equal("未变化", FormulaDiff.DescribeChange("=SUM(A1:A10)", "=sum(a1:a10)")); // 大小写不敏感
        }

        [Fact]
        public void Empty_to_formula_is_added()
        {
            Assert.Equal("新增公式", FormulaDiff.DescribeChange(null, "=SUM(A1:A10)"));
            Assert.Equal("新增公式", FormulaDiff.DescribeChange("", "=SUM(A1:A10)"));
        }

        [Fact]
        public void Formula_to_empty_is_cleared()
        {
            Assert.Equal("清除公式", FormulaDiff.DescribeChange("=SUM(A1:A10)", null));
            Assert.Equal("清除公式", FormulaDiff.DescribeChange("=SUM(A1:A10)", ""));
        }

        [Fact]
        public void Wrap_iferror_detected()
        {
            Assert.Equal(
                "增加 IFERROR 错误兜底",
                FormulaDiff.DescribeChange("=AVERAGE(B2:B10)", "=IFERROR(AVERAGE(B2:B10),0)"));
        }

        [Fact]
        public void Unwrap_iferror_detected()
        {
            Assert.Equal(
                "移除 IFERROR 错误兜底",
                FormulaDiff.DescribeChange("=IFERROR(AVERAGE(B2:B10),0)", "=AVERAGE(B2:B10)"));
        }

        [Fact]
        public void Iferror_with_changed_inner_is_not_unwrap()
        {
            // 壳内公式也变了 → 不是单纯增删 IFERROR，应落函数/参数分支
            var label = FormulaDiff.DescribeChange("=AVERAGE(B2:B10)", "=IFERROR(SUM(B2:B10),0)");
            Assert.NotEqual("增加 IFERROR 错误兜底", label);
        }

        [Fact]
        public void Anchor_only_change_detected()
        {
            Assert.Equal(
                "仅调整引用方式（相对/绝对）",
                FormulaDiff.DescribeChange("=B2*C2", "=$B$2*$C$2"));
            Assert.Equal(
                "仅调整引用方式（相对/绝对）",
                FormulaDiff.DescribeChange("=SUM(B2:B10)", "=SUM($B$2:$B$10)"));
        }

        [Fact]
        public void Function_change_detected()
        {
            var label = FormulaDiff.DescribeChange("=SUM(B2:B10)", "=AVERAGE(B2:B10)");
            Assert.Contains("函数变化", label);
            Assert.Contains("AVERAGE", label);
            Assert.Contains("SUM", label);
        }

        [Fact]
        public void Criteria_change_detected()
        {
            // 函数集相同、内容不同 → 参数/条件调整
            Assert.Equal(
                "参数/条件调整",
                FormulaDiff.DescribeChange(
                    "=SUMIFS(C2:C10,B2:B10,\"华东\")",
                    "=SUMIFS(C2:C10,B2:B10,\"华南\")"));
        }

        [Fact]
        public void Nested_iferror_fallback_change_is_criteria()
        {
            // fallback 变化：函数集相同（IFERROR/AVERAGE 都在）→ 参数/条件调整
            Assert.Equal(
                "参数/条件调整",
                FormulaDiff.DescribeChange("=IFERROR(AVERAGE(B2:B10),0)", "=IFERROR(AVERAGE(B2:B10),\"-\")"));
        }

        [Fact]
        public void String_literal_space_difference_is_criteria_change_not_unchanged()
        {
            // 字符串字面量内的空格是内容差异，不得被空白归一化抹掉
            Assert.NotEqual(
                "未变化",
                FormulaDiff.DescribeChange("=SUMIFS(C2:C10,B2:B10,\"华东 区\")", "=SUMIFS(C2:C10,B2:B10,\"华东区\")"));
        }

        [Fact]
        public void String_literal_anchor_not_touched_by_to_absolute()
        {
            // ToAbsolute 只动引用，不得篡改字符串字面量里的类引用文本（如编码 "B2"）
            Assert.Equal(
                "IF($A$1=\"B2\",\"yes\",\"no\")",
                FormulaBuilder.ToAbsolute("IF(A1=\"B2\",\"yes\",\"no\")"));
        }

        [Fact]
        public void Sheet_quoted_name_not_touched_by_to_absolute()
        {
            // 工具分类审查第六批复现（v3.8.1.96 残留缺口）：单引号工作表名内的类引用文本
            // （如 'Sales S2' 的 S2）不得被 to_absolute 篡改成 $S$2 损坏公式。
            // 注意：B:C 整列引用无行号，不在 ToAbsolute 转换域（既有语义），保持原样。
            Assert.Equal(
                "VLOOKUP($A$1,'Sales S2'!B:C,2,0)",
                FormulaBuilder.ToAbsolute("VLOOKUP(A1,'Sales S2'!B:C,2,0)"));
        }

        [Fact]
        public void Apostrophe_inside_double_quotes_is_content_not_delimiter()
        {
            // ' 在双引号内是文本内容（="it's" 合法），不是工作表名定界符——不得误开单引号保护段。
            Assert.Equal(
                "IF($A$1=\"it's\",\"yes\",\"no\")",
                FormulaBuilder.ToAbsolute("IF(A1=\"it's\",\"yes\",\"no\")"));
        }

        [Fact]
        public void Whitespace_and_equals_prefix_normalized()
        {
            Assert.Equal("未变化", FormulaDiff.DescribeChange("= SUM( A1 : A10 )", "=SUM(A1:A10)"));
        }
    }
}
