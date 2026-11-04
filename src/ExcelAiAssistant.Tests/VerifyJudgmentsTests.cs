using System.Collections.Generic;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>四类操作结果自检的纯静态判据单测：单元格格式 / 条件格式 / 排序筛选 / 公式。</summary>
    public class VerifyJudgmentsTests
    {
        // ===================== 单元格格式 =====================

        [Theory]
        [InlineData("0.00", "0.00", true)]
        [InlineData(" 0.00 ", "0.00", true)]
        [InlineData("#,##0", "#,##0", true)] // 忽略大小写
        [InlineData("0.00", "#,##0", false)]
        [InlineData(null, "0.00", false)]
        [InlineData("0.00", null, true)] // 未给预期 → 不断言
        [InlineData("", "", true)]
        public void Number_format_matches(string actual, string expected, bool ok)
        {
            Assert.Equal(ok, ExcelVerifyJudgments.NumberFormatMatches(actual, expected));
        }

        [Fact]
        public void Bool_matches_only_asserts_when_expected_given()
        {
            Assert.True(ExcelVerifyJudgments.BoolMatches(true, null));
            Assert.True(ExcelVerifyJudgments.BoolMatches(false, null));
            Assert.True(ExcelVerifyJudgments.BoolMatches(true, true));
            Assert.False(ExcelVerifyJudgments.BoolMatches(false, true));
            Assert.True(ExcelVerifyJudgments.BoolMatches(false, false));
        }

        [Theory]
        [InlineData(65535, "#FFFF00")] // 黄 (BGR: R=FF,G=FF,B=00 → 65535)
        [InlineData(255, "#FF0000")]       // 红 BGR=255 → R=FF
        [InlineData(16711680, "#0000FF")]  // 蓝
        [InlineData(65280, "#00FF00")]     // 绿
        [InlineData(0, "#000000")]         // 黑
        [InlineData(255, "255")]           // int 形式
        public void Color_matches_expected_hex_and_int(int actualBgr, string expected)
        {
            Assert.True(ExcelVerifyJudgments.ColorMatches(actualBgr, expected));
            Assert.False(ExcelVerifyJudgments.ColorMatches(actualBgr == 255 ? 0 : 255, expected));
        }

        [Fact]
        public void Color_no_assert_when_no_expected()
        {
            Assert.True(ExcelVerifyJudgments.ColorMatches(255, null));
            Assert.True(ExcelVerifyJudgments.ColorMatches(255, ""));
        }

        [Theory]
        [InlineData(-4131, "left", true)]
        [InlineData(-4152, "右对齐", true)]
        [InlineData(-4108, "center", true)]
        [InlineData(-4160, "top", true)]
        [InlineData(-4107, "bottom", true)]
        [InlineData(-4130, "justify", true)]
        [InlineData(-4131, "center", false)]
        [InlineData(-4107, "top", false)]
        public void Alignment_matches_maps_constants(int actual, string expected, bool ok)
        {
            Assert.Equal(ok, ExcelVerifyJudgments.AlignmentMatches(actual, expected));
        }

        [Theory]
        [InlineData(1, "continuous", true)]
        [InlineData(1, "实线", true)]
        [InlineData(-4142, "none", true)]
        [InlineData(-4142, "无边框", true)]
        [InlineData(-4115, "dashed", true)]
        [InlineData(1, "none", false)]
        [InlineData(1, "dashed", false)]
        public void Border_style_matches_maps_constants(int actual, string expected, bool ok)
        {
            Assert.Equal(ok, ExcelVerifyJudgments.BorderStyleMatches(actual, expected));
        }

        [Fact]
        public void Compute_cell_format_verdict_ok_when_no_mismatch()
        {
            var r = ExcelVerifyJudgments.ComputeCellFormatVerdict(
                "0.00", true, false, "微软雅黑", 12, 255, 16777215, -4131, -4107, 1, true,
                "0.00", true, false, "微软雅黑", 12, "#FF0000", "#FFFFFF", "left", "bottom", "实线", true);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_cell_format_verdict_flags_number_bold_and_fill_mismatch()
        {
            var r = ExcelVerifyJudgments.ComputeCellFormatVerdict(
                "#,##0", false, false, "宋体", 12, 255, 0, -4131, -4107, 1, true,
                "0.00", true, false, "微软雅黑", 12, "#00FF00", "#FFFFFF", "center", "bottom", "实线", true);
            Assert.Contains("number", r.verdict);
            Assert.Contains("bold", r.verdict);
            Assert.Contains("fill", r.verdict);
            Assert.Contains("font", r.verdict);
            Assert.Contains("align", r.verdict);
            Assert.Contains("数字格式与预期不符", r.issues[0]);
            Assert.Contains("自检未通过", r.selfcheck);
        }

        // ===================== 条件格式 =====================

        [Theory]
        [InlineData("大于", "greater")]
        [InlineData("GREATER", "greater")]
        [InlineData("数据条", "data_bar")]
        [InlineData("重复值", "duplicates")]
        [InlineData("介于", "between")]
        [InlineData("less", "less")]
        [InlineData("4", "data_bar")]    // xlDatabar 常量
        [InlineData("8", "duplicates")]  // xlUniqueValues 常量
        [InlineData("1", "cell_value")]  // xlCellValue 常量
        [InlineData(null, null)]
        [InlineData("bogus", null)]
        public void Normalize_rule_type_supports_multiple_spellings(string input, string expected)
        {
            Assert.Equal(expected, ExcelVerifyJudgments.NormalizeRuleType(input));
        }

        [Fact]
        public void Compute_conditional_format_verdict_ok_on_match()
        {
            var r = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                "greater", "=A1>100", null, 5, 255, 16777215,
                "greater", "=A1>100", null, 5, "#FF0000", "#FFFFFF");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_conditional_format_verdict_flags_rule_and_fill_mismatch()
        {
            var r = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                "less", "=A1<50", null, 6, 0, 0,
                "greater", "=A1>100", null, 5, "#FF0000", "#FFFFFF");
            Assert.Contains("rule", r.verdict);
            Assert.Contains("fill", r.verdict);
            Assert.Contains("规则类型与预期不符", r.issues[0]);
        }

        [Fact]
        public void Compute_conditional_format_verdict_no_assert_without_expected_rule()
        {
            var r = ExcelVerifyJudgments.ComputeConditionalFormatVerdict(
                "greater", "=A1>100", null, 5, 255, 0,
                null, null, null, null, null, null);
            Assert.Equal("ok", r.verdict);
        }

        // ===================== 排序与筛选 =====================

        [Theory]
        [InlineData("升序", "asc")]
        [InlineData("DESC", "desc")]
        [InlineData("descending", "desc")]
        [InlineData("ascending", "asc")]
        [InlineData(null, null)]
        [InlineData("bogus", null)]
        public void Normalize_sort_direction(string input, string expected)
        {
            Assert.Equal(expected, ExcelVerifyJudgments.NormalizeSortDirection(input));
        }

        [Fact]
        public void Compute_sort_verdict_ok_on_match()
        {
            var matrix = new object[,] { { "city", "val" }, { "Beijing", 20 }, { "Beijing", 30 } };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, 2, "asc", null, null);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_sort_verdict_flags_wrong_direction()
        {
            // 实际第 2 列为升序，预期降序 → 触发 sort 告警
            var matrix = new object[,] { { "city", "val" }, { "a", 10 }, { "b", 20 } };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, 2, "desc", null, null);
            Assert.Contains("sort", r.verdict);
            Assert.Contains("并未按预期排序", r.issues[0]);
        }

        [Fact]
        public void Compute_sort_verdict_no_assert_when_no_expectation()
        {
            var matrix = new object[,] { { "city", "val" }, { "a", 10 }, { "b", 20 } };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, null, null, null, null);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_filter_verdict_flags_not_satisfied()
        {
            // 区域内混有 Beijing/Shanghai，预期筛选 Beijing 未完全生效 → 触发 filter 告警
            var matrix = new object[,] { { "city", "val" }, { "Shanghai", 10 }, { "Beijing", 20 }, { "Shanghai", 30 } };
            var r = ExcelVerifyJudgments.ComputeFilterVerdict(matrix, 1, 1, "Beijing", null);
            Assert.Contains("filter", r.verdict);
            Assert.Contains("并未完全生效", r.issues[0]);
        }

        [Fact]
        public void Compute_filter_verdict_no_assert_when_no_filter_expectation()
        {
            // 未传筛选预期 → 不断言筛选状态，不应误报
            var matrix = new object[,] { { "city", "val" }, { "Beijing", 20 }, { "Shanghai", 30 } };
            var r = ExcelVerifyJudgments.ComputeFilterVerdict(matrix, 1, null, null, null);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_filter_verdict_ok_on_match()
        {
            var matrix = new object[,] { { "city", "val" }, { "Beijing", 20 }, { "Beijing", 30 } };
            var r = ExcelVerifyJudgments.ComputeFilterVerdict(matrix, 1, 1, "Beijing", null);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Theory]
        [InlineData("#DIV/0!", true)]
        [InlineData("#REF!", true)]
        [InlineData("Error 2007", true)]   // .NET COM 互操作错误码文本形态
        [InlineData("120", false)]
        [InlineData("普通文本", false)]
        public void Is_error_value_detects_error_texts(string value, bool expected)
        {
            Assert.Equal(expected, ExcelVerifyJudgments.IsErrorValue(value));
        }

        [Fact]
        public void Compute_sort_verdict_flags_invalid_column()
        {
            // 2 列区域，预期按第 5 列排序 → 不应静默假通过，须打 sort 标
            var matrix = new object[,] { { "a", 10 }, { "b", 20 } };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 0, 5, "asc", null, null);
            Assert.Contains("sort", r.verdict);
            Assert.Contains("列号非法", r.issues[0]);
        }

        [Fact]
        public void Compute_sort_verdict_composite_ok_when_secondary_orders_ties()
        {
            // 主列(2)升序有并列，次列(3)升序补齐 → 复合字典序满足 → ok
            var matrix = new object[,]
            {
                { "city", "val", "idx" },
                { "Beijing", 10, 1 },
                { "Shanghai", 10, 2 },
                { "Beijing", 20, 3 }
            };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, 2, "asc", 3, "asc");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_sort_verdict_composite_flags_when_secondary_breaks_order()
        {
            // 主列(1)升序有并列(a,a)，次列(2)期望升序但实际降序(10,5) → 不满足复合排序 → sort 告警
            var matrix = new object[,]
            {
                { "city", "val" },
                { "a", 10 },
                { "a", 5 },
                { "b", 1 }
            };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, 1, "asc", 2, "asc");
            Assert.Contains("sort", r.verdict);
            Assert.Contains("并未按预期排序", r.issues[0]);
        }

        [Fact]
        public void Compute_sort_verdict_composite_flags_invalid_secondary_column()
        {
            // 3 列区域，次排序列号 9 越界 → 不打假通过，须报次列非法
            var matrix = new object[,]
            {
                { "a", "b", "c" },
                { 1, 2, 3 }
            };
            var r = ExcelVerifyJudgments.ComputeSortVerdict(matrix, 1, 1, "asc", 9, "asc");
            Assert.Contains("sort", r.verdict);
            Assert.Contains("次排序列号非法", r.issues[0]);
        }

        [Fact]
        public void Compute_filter_verdict_flags_invalid_field()
        {
            // 2 列区域，预期筛选第 9 列 → 不应越界抛异常，须打 filter 标
            var matrix = new object[,] { { "city", "val" }, { "Beijing", 10 } };
            var r = ExcelVerifyJudgments.ComputeFilterVerdict(matrix, 1, 9, "Beijing", null);
            Assert.Contains("filter", r.verdict);
            Assert.Contains("列号非法", r.issues[0]);
        }

        // ===================== 公式 =====================

        [Theory]
        [InlineData("=B2*C2", "=b2*c2", true)]       // 忽略大小写
        [InlineData("=B2*C2", " = B2*C2 ", true)]     // 去空白
        [InlineData("=B2*C2", "=B2+C2", false)]
        [InlineData("=SUM(A1:A5)", "=sum(a1:a5)", true)]
        public void Normalize_and_compare_formula(string actual, string expected, bool ok)
        {
            Assert.Equal(ok, string.Equals(ExcelVerifyJudgments.NormalizeFormula(actual), ExcelVerifyJudgments.NormalizeFormula(expected), System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Compute_formula_verdict_ok_on_valid_formula_with_matching_value()
        {
            var r = ExcelVerifyJudgments.ComputeFormulaVerdict(true, "=B2*C2", "120", "=B2*C2", "120", null);
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_formula_verdict_detects_error_value()
        {
            var r = ExcelVerifyJudgments.ComputeFormulaVerdict(true, "=A1/B1", "#DIV/0!", null, null, null);
            Assert.Equal("error", r.verdict);
            Assert.Contains("错误值", r.issues[0]);
        }

        [Fact]
        public void Compute_formula_verdict_flags_no_formula_and_value_mismatch()
        {
            var r = ExcelVerifyJudgments.ComputeFormulaVerdict(false, null, "文本", "=B2*C2", "120", null);
            Assert.Contains("noformula", r.verdict);
            Assert.Contains("value", r.verdict);
        }

        [Fact]
        public void Values_match_numeric_with_tolerance()
        {
            Assert.True(ExcelVerifyJudgments.ValuesMatch("3.14", "3.1400000001", null)); // 默认 1e-9 容差内
            Assert.False(ExcelVerifyJudgments.ValuesMatch("3.14", "3.25", null));
            Assert.True(ExcelVerifyJudgments.ValuesMatch("3.14", "3.2", 0.1));
            Assert.True(ExcelVerifyJudgments.ValuesMatch("北京", "北京", null)); // 文本
            Assert.False(ExcelVerifyJudgments.ValuesMatch("北京", "上海", null));
            Assert.True(ExcelVerifyJudgments.ValuesMatch("120", "120", null));
        }

        [Theory]
        [InlineData("#REF!", true)]
        [InlineData("#DIV/0!", true)]
        [InlineData("#VALUE!", true)]
        [InlineData("#NAME?", true)]
        [InlineData("#N/A", true)]
        [InlineData("#NUM!", true)]
        [InlineData("#NULL!", true)]
        [InlineData("120", false)]
        [InlineData("北京", false)]
        [InlineData(null, false)]
        [InlineData("", false)]
        [InlineData("前缀#REF!后缀", false)]
        public void Is_error_value_detects_excel_errors(string value, bool expected)
        {
            Assert.Equal(expected, ExcelVerifyJudgments.IsErrorValue(value));
        }

        // ===================== 写入结果自检 =====================

        [Fact]
        public void Compute_write_result_verdict_ok_on_exact_match()
        {
            var actual = new object[,] { { "A", 10 }, { "B", 20 } };
            var expected = new object[,] { { "A", 10 }, { "B", 20 } };
            var r = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, 4, null, "A2:B3");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_write_result_verdict_flags_value_and_count_mismatch()
        {
            // 实际 2x2 全有值（count=4），期望 2x2 但 B2 是 99；期望 count=5 → value 与 count 双标
            var actual = new object[,] { { "A", 10 }, { "B", 20 } };
            var expected = new object[,] { { "A", 99 }, { "B", 20 } };
            var r = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, 5, null, "A2:B3");
            Assert.Contains("value", r.verdict);
            Assert.Contains("count", r.verdict);
            Assert.Contains("B2", r.issues[0]); // 偏移地址定位到具体错位格
            Assert.Contains("自检未通过", r.selfcheck);
        }

        [Fact]
        public void Compute_write_result_verdict_treats_blank_expected_as_blank_required()
        {
            // 期望为空的格，实际必须也为空；期望 A=1/B=空，实际 A=1/B=5 → B 格报 value
            var actual = new object[,] { { 1, 5 } };
            var expected = new object[,] { { 1, null } };
            var r = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, null, "C2:D2");
            Assert.Contains("value", r.verdict);
            Assert.Contains("D2", r.issues[0]);
        }

        [Fact]
        public void Compute_write_result_verdict_numeric_with_tolerance()
        {
            var actual = new object[,] { { 3.14159 } };
            var expected = new object[,] { { 3.1416 } };
            Assert.Equal("ok", ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, 0.001, "A1").verdict);
            Assert.Contains("value", ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, null, "A1").verdict);
        }

        [Fact]
        public void Compute_write_result_verdict_no_assert_without_expected()
        {
            var r = ExcelVerifyJudgments.ComputeWriteResultVerdict(new object[,] { { "x" } }, null, null, null, "A1");
            Assert.Contains("no_expected", r.verdict);
            Assert.Contains(r.issues, i => i.Contains("无判定依据"));
        }

        [Fact]
        public void Compute_write_result_verdict_handles_actual_shorter_than_expected()
        {
            // 实际 1 格，期望 2 格 → 缺格视为 null，期望非空 → 报 value
            var r = ExcelVerifyJudgments.ComputeWriteResultVerdict(
                new object[,] { { 1 } }, new object[,] { { 1, 2 } }, null, null, "A1:B1");
            Assert.Contains("value", r.verdict);
        }

        // ===================== 数据清洗自检 =====================

        [Fact]
        public void Compute_data_clean_verdict_ok_when_all_checks_pass()
        {
            // 表头 + 2 行干净数据：无空行/无空列/无重复/日期合法
            var values = new object[,]
            {
                { "name", "date", "qty" },
                { "A", "2026-01-01", "12" },
                { "B", "2026-02-01", "34" }
            };
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_blank_rows" },
                new ExcelVerifyJudgments.CleanCheck { Type = "no_duplicates" },
                new ExcelVerifyJudgments.CleanCheck { Type = "dates_valid", Columns = new[] { 2 } }
            };
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 1, "A1:C3");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_data_clean_verdict_flags_blank_row_and_duplicate()
        {
            var values = new object[,]
            {
                { "name", "qty" },
                { "A", 1 },
                { null, null },
                { "A", 1 }
            };
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_blank_rows" },
                new ExcelVerifyJudgments.CleanCheck { Type = "no_duplicates" }
            };
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 1, "A1:B4");
            Assert.Contains("blank_row", r.verdict);
            Assert.Contains("duplicate", r.verdict);
            Assert.Contains("整行空白", r.issues[0]);
            Assert.Contains("重复记录", r.issues[1]);
        }

        [Fact]
        public void Compute_data_clean_verdict_flags_text_number()
        {
            // qty 列 body 含文本型数字 "12" → no_text_numbers 命中
            var values = new object[,] { { "qty" }, { "12" }, { "34" } };
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_text_numbers", Columns = new[] { 1 } }
            };
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 1, "A1:A3");
            Assert.Contains("text_number", r.verdict);
            Assert.Contains("文本型数字", r.issues[0]);
        }

        [Fact]
        public void Compute_data_clean_verdict_blank_in_specific_column()
        {
            // 第 2 列 body 有空值 → no_blanks 命中；但未指定列时全列扫描也会命中
            var values = new object[,] { { "a", "b" }, { "x", null } };
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_blanks", Columns = new[] { 2 } }
            };
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 1, "A1:B2");
            Assert.Contains("blank", r.verdict);
            Assert.Contains("空白单元格", r.issues[0]);
        }

        [Fact]
        public void Compute_data_clean_verdict_no_assert_without_checks()
        {
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(
                new object[,] { { "a" }, { "b" } }, null, 1, "A1:A2");
            Assert.Contains("no_checks", r.verdict);
            Assert.Contains(r.issues, i => i.Contains("无判定依据"));
        }

        [Fact]
        public void Compute_data_clean_verdict_ignores_unknown_check_type()
        {
            var r = ExcelVerifyJudgments.ComputeDataCleanVerdict(
                new object[,] { { "a" }, { "b" } },
                new List<ExcelVerifyJudgments.CleanCheck> { new ExcelVerifyJudgments.CleanCheck { Type = "bogus" } },
                1, "A1:A2");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        // ===================== 透视表自检 =====================

        [Fact]
        public void Compute_pivot_verdict_ok_on_match()
        {
            // 首行表头 + 数据行：按第一列标签匹配第二列汇总值
            var values = new object[,]
            {
                { "类别", "合计" },
                { "华东", 120 },
                { "华北", 80 }
            };
            var expected = new List<ExcelVerifyJudgments.PivotExpectedRow>
            {
                new ExcelVerifyJudgments.PivotExpectedRow { Label = "华东", Value = 120 },
                new ExcelVerifyJudgments.PivotExpectedRow { Label = "华北", Value = 80 }
            };
            var r = ExcelVerifyJudgments.ComputePivotVerdict(values, expected, 2, null, "B2");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_pivot_verdict_flags_count_and_missing_label()
        {
            var values = new object[,]
            {
                { "类别", "合计" },
                { "华东", 120 }
            };
            var expected = new List<ExcelVerifyJudgments.PivotExpectedRow>
            {
                new ExcelVerifyJudgments.PivotExpectedRow { Label = "华东", Value = 120 },
                new ExcelVerifyJudgments.PivotExpectedRow { Label = "华南", Value = 60 }
            };
            var r = ExcelVerifyJudgments.ComputePivotVerdict(values, expected, 3, null, "B2");
            Assert.Contains("count", r.verdict);
            Assert.Contains("missing", r.verdict);
            Assert.Contains(r.issues, i => i.Contains("数据行数") && i.Contains("3")); // count 问题先列
            Assert.Contains(r.issues, i => i.Contains("未找到标签行") && i.Contains("华南"));
        }

        [Fact]
        public void Compute_pivot_verdict_flags_value_mismatch_with_tolerance()
        {
            var values = new object[,] { { "类别", "合计" }, { "华东", 120.5 } };
            var expected = new List<ExcelVerifyJudgments.PivotExpectedRow>
            {
                new ExcelVerifyJudgments.PivotExpectedRow { Label = "华东", Value = 121 }
            };
            Assert.Contains("value", ExcelVerifyJudgments.ComputePivotVerdict(values, expected, null, 0.01, "B2").verdict);
            Assert.Equal("ok", ExcelVerifyJudgments.ComputePivotVerdict(values, expected, null, 1.0, "B2").verdict);
        }

        [Fact]
        public void Compute_pivot_verdict_no_assert_without_expectations()
        {
            var r = ExcelVerifyJudgments.ComputePivotVerdict(
                new object[,] { { "类别", "合计" }, { "华东", 120 } }, null, null, null, "B2");
            Assert.Contains("no_expected", r.verdict);
            Assert.Contains(r.issues, i => i.Contains("无判定依据"));
        }

        // ===================== 标记自检 =====================

        [Fact]
        public void Compute_marker_verdict_ok_on_match()
        {
            var marked = new List<ExcelVerifyJudgments.MarkedCell>
            {
                new ExcelVerifyJudgments.MarkedCell { Address = "$B$2", Color = 255 }, // #FF0000 红 → BGR 255
                new ExcelVerifyJudgments.MarkedCell { Address = "$B$3", Color = 255 }
            };
            var colors = new Dictionary<string, string> { ["red"] = "#FF0000" };
            var expectedMarked = new List<string> { "$B$2", "B3" }; // $ 忽略
            var r = ExcelVerifyJudgments.ComputeMarkerVerdict(marked, 2, colors, expectedMarked, "B2:B3");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }

        [Fact]
        public void Compute_marker_verdict_flags_count_color_and_missing_address()
        {
            var marked = new List<ExcelVerifyJudgments.MarkedCell>
            {
                new ExcelVerifyJudgments.MarkedCell { Address = "$B$2", Color = 65280 } // 绿
            };
            var colors = new Dictionary<string, string> { ["red"] = "#FF0000" };
            var expectedMarked = new List<string> { "B2", "B5" };
            var r = ExcelVerifyJudgments.ComputeMarkerVerdict(marked, 3, colors, expectedMarked, "B2:B3");
            Assert.Contains("count", r.verdict);
            Assert.Contains("color", r.verdict);
            Assert.Contains("mark", r.verdict);
            Assert.Contains("被标记单元格数", r.issues[0]);
        }

        [Fact]
        public void Compute_marker_verdict_flags_expected_color_missing()
        {
            // 实际标记只有绿，期望红+绿 → 期望色集合中存在实际未出现的颜色 → color 标
            var marked = new List<ExcelVerifyJudgments.MarkedCell>
            {
                new ExcelVerifyJudgments.MarkedCell { Address = "B2", Color = 65280 }
            };
            var colors = new Dictionary<string, string> { ["red"] = "#FF0000", ["green"] = "#00FF00" };
            var r = ExcelVerifyJudgments.ComputeMarkerVerdict(marked, null, colors, null, "B2");
            Assert.Contains("color", r.verdict);
            Assert.Contains("未出现", r.issues[0]);
        }

        [Fact]
        public void Compute_marker_verdict_no_assert_without_expectations()
        {
            var marked = new List<ExcelVerifyJudgments.MarkedCell>
            {
                new ExcelVerifyJudgments.MarkedCell { Address = "B2", Color = 65280 }
            };
            var r = ExcelVerifyJudgments.ComputeMarkerVerdict(marked, null, null, null, "B2");
            Assert.Contains("no_expected", r.verdict);
            Assert.Contains(r.issues, i => i.Contains("无判定依据"));
        }

        [Fact]
        public void Compute_marker_verdict_ignores_unparseable_expected_color()
        {
            // 期望色解析失败 → 不做颜色断言（known=false），不误报 color
            var marked = new List<ExcelVerifyJudgments.MarkedCell>
            {
                new ExcelVerifyJudgments.MarkedCell { Address = "B2", Color = 65280 }
            };
            var colors = new Dictionary<string, string> { ["red"] = "not-a-color" };
            var r = ExcelVerifyJudgments.ComputeMarkerVerdict(marked, null, colors, null, "B2");
            Assert.Equal("ok", r.verdict);
            Assert.Empty(r.issues);
        }
    }
}