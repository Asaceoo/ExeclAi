using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>判定层扩展：类型感知回读、透视数量守恒、结构自检。</summary>
    public sealed class VerifyJudgmentExtensionTests
    {
        [Fact]
        public void Issue_addresses_offset_with_dollar_addresses()
        {
            // 回归：带 $ 的地址（"$N$1:$O$2"）此前只去前导 $，OffsetAddress 解析行号失败，
            // 所有 issue 地址错误地显示为左上角 "N$1"。
            var actual = new object[,] { { "甲", "乙" }, { "丙", "丁" } };
            var expected = new object[,] { { "壹", "贰" }, { "叁", "肆" } };
            var result = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, null, "$N$1:$O$2");
            Assert.Equal(4, result.issues.Count);
            Assert.Contains(result.issues, i => i.StartsWith("N1："));
            Assert.Contains(result.issues, i => i.StartsWith("O1："));
            Assert.Contains(result.issues, i => i.StartsWith("N2："));
            Assert.Contains(result.issues, i => i.StartsWith("O2："));
            Assert.All(result.issues, i => Assert.False(i.StartsWith("N$1："), "issue 地址不应停留在左上角原文。"));
        }

        [Fact]
        public void Type_aware_flags_text_number_that_plain_compare_accepts()
        {
            // 同一矩阵：常规比对判 ok（文本 "3" 与数值 3 数值相等），type_aware 开启后必须报 type。
            var actual = new object[,] { { "苹果", "3" } };
            var expected = new object[,] { { "苹果", 3 } };

            var plain = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, null, "A1");
            Assert.Equal("ok", plain.verdict);

            var aware = ExcelVerifyJudgments.ComputeWriteResultVerdict(actual, expected, null, null, "A1", typeAware: true);
            Assert.Contains("type", aware.verdict);
            Assert.Contains(aware.issues, i => i.Contains("文本型数字"));
        }

        [Fact]
        public void Type_aware_passes_real_numeric_and_text_expectations()
        {
            var ok = ExcelVerifyJudgments.ComputeWriteResultVerdict(
                new object[,] { { 3.0 } }, new object[,] { { 3 } }, null, null, "A1", typeAware: true);
            Assert.Equal("ok", ok.verdict);

            // 期望本身是文本时不做类型断言（类型语义只有模型知道）。
            var textExpected = ExcelVerifyJudgments.ComputeWriteResultVerdict(
                new object[,] { { "3" } }, new object[,] { { "3" } }, null, null, "A1", typeAware: true);
            Assert.Equal("ok", textExpected.verdict);
        }

        [Fact]
        public void Pivot_conservation_flags_sum_mismatch()
        {
            var pivot = new object[,] { { "Row Labels", "Total" }, { "苹果", 5.0 } };

            var ok = ExcelVerifyJudgments.ComputePivotVerdict(
                pivot, null, null, null, "E1", sourceSum: 5.0, targetSum: 5.0);
            Assert.Equal("ok", ok.verdict);

            var broken = ExcelVerifyJudgments.ComputePivotVerdict(
                pivot, null, null, null, "E1", sourceSum: 12.0, targetSum: 5.0);
            Assert.Contains("conservation", broken.verdict);
            Assert.Contains(broken.issues, i => i.Contains("数量守恒不符"));
        }

        [Fact]
        public void Pivot_without_any_expectation_still_reports_no_expected()
        {
            var pivot = new object[,] { { "Row Labels", "Total" }, { "苹果", 5.0 } };
            var none = ExcelVerifyJudgments.ComputePivotVerdict(pivot, null, null, null, "E1");
            Assert.Contains("no_expected", none.verdict);
        }

        [Fact]
        public void Structure_verdict_checks_sheet_existence_and_absence()
        {
            var names = new[] { "Sheet1", "汇总" };

            var ok = ExcelVerifyJudgments.ComputeStructureVerdict(
                names, "汇总", "草稿", null, null, null, null, null);
            Assert.Equal("ok", ok.verdict);

            var missing = ExcelVerifyJudgments.ComputeStructureVerdict(
                names, "报表", null, null, null, null, null, null);
            Assert.Contains("missing_sheet", missing.verdict);

            var stillThere = ExcelVerifyJudgments.ComputeStructureVerdict(
                names, null, "汇总", null, null, null, null, null);
            Assert.Contains("sheet_present", stillThere.verdict);
        }

        [Fact]
        public void Structure_verdict_checks_region_dimensions()
        {
            var ok = ExcelVerifyJudgments.ComputeStructureVerdict(
                new[] { "Sheet1" }, null, null, 3, 2, 3, 2, "A1:B3");
            Assert.Equal("ok", ok.verdict);

            var mismatch = ExcelVerifyJudgments.ComputeStructureVerdict(
                new[] { "Sheet1" }, null, null, 5, 2, 3, 2, "A1:B3");
            Assert.NotEqual("ok", mismatch.verdict);
            Assert.Contains(mismatch.issues, i => i.Contains("行数"));

            // 期望行列数但没给地址 → 如实报无法核对。
            var noAddress = ExcelVerifyJudgments.ComputeStructureVerdict(
                new[] { "Sheet1" }, null, null, 5, null, null, null, null);
            Assert.Contains("no_address", noAddress.verdict);
        }
    }
}
