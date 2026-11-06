using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 列质量画像引擎（ColumnProfileEngine，M14）的纯计算用例：
    /// 类型判定顺序（bool 先于 number）、空值率/错误率、质量提示 flag、数值/文本剖析、边界（空列/单行）。
    /// 矩阵约定：第 0 行为标题行，Compute 只统计第 1 行起的数据。
    /// </summary>
    public sealed class ColumnProfileEngineTests
    {
        private static object[,] Matrix(int dataRows, int columns)
        {
            return new object[dataRows + 1, columns];
        }

        private static object[,] OneColumnMatrix(params object[] data)
        {
            var m = new object[data.Length + 1, 1];
            m[0, 0] = "列A";
            for (var i = 0; i < data.Length; i++)
            {
                m[i + 1, 0] = data[i];
            }

            return m;
        }

        [Fact]
        public void Numeric_column_yields_full_stats()
        {
            var profile = ColumnProfileEngine.Compute(OneColumnMatrix(10d, 20d, 30d, 40d), 1, 10);
            Assert.Equal(4, profile.TotalRows);
            Assert.Equal(0, profile.BlankCount);
            Assert.Equal(4, profile.Numbers.Count);
            Assert.Equal(10d, profile.Numbers.Min);
            Assert.Equal(40d, profile.Numbers.Max);
            Assert.Equal(25d, profile.Numbers.Mean);
            Assert.Equal(25d, profile.Numbers.Median);
            // 样本标准差：sqrt(((15)^2+(5)^2+(5)^2+(15)^2)/3) = sqrt(500/3) ≈ 12.9099
            Assert.Equal(12.9099, profile.Numbers.StdDev, 4);
            Assert.Equal(4, profile.DistinctCount);
            Assert.Contains("number", profile.TypeCounts.Keys);
            Assert.Equal(4, profile.TypeCounts["number"]);
        }

        [Fact]
        public void Even_count_median_is_mean_of_middle_two()
        {
            var profile = ColumnProfileEngine.Compute(OneColumnMatrix(1d, 2d, 3d, 100d), 1, 10);
            Assert.Equal(2.5, profile.Numbers.Median);
        }

        [Fact]
        public void Errors_and_blanks_are_counted_separately()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix("#N/A", "#VALUE!", null, "  ", 5d), 1, 10);
            Assert.Equal(5, profile.TotalRows);
            Assert.Equal(2, profile.BlankCount); // null + 纯空白
            Assert.Equal(2, profile.ErrorCount);
            Assert.Equal(0.4, profile.BlankRate);
            Assert.Equal(0.4, profile.ErrorRate);
            Assert.Equal(2, profile.TypeCounts["error"]);
            Assert.Equal(2, profile.TypeCounts["blank"]);
            Assert.Contains("has_errors", profile.QualityFlags);
        }

        [Fact]
        public void Cverr_int32_cells_are_errors_not_numbers()
        {
            // M14 真机实证（RMV 用例25 + COM 探针）："#N/A" 写入后被 Excel 解析为错误值，
            // Value2 读回 Int32 -2146826246，不得计入数值统计。
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix(10d, 20d, 20d, -2146826246, ""), 1, 10);
            Assert.Equal(1, profile.ErrorCount);
            Assert.Equal(3, profile.Numbers.Count);
            Assert.Equal((10d + 20d + 20d) / 3, profile.Numbers.Mean, 4);
            Assert.Equal(1, profile.TypeCounts["blank"]);
            Assert.Contains("has_errors", profile.QualityFlags);
            // 错误码计入唯一值与分布（"10"/"20"/"#N/A"）
            Assert.Equal(3, profile.DistinctCount);
        }

        [Fact]
        public void Blank_heavy_flag_when_over_half_blank()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix(null, "  ", 1d), 1, 10);
            Assert.Equal(2.0 / 3, profile.BlankRate, 4);
            Assert.Contains("blank_heavy", profile.QualityFlags);
        }

        [Fact]
        public void Mixed_types_flagged_and_bool_takes_precedence_over_number()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix(true, 3d, "文本"), 1, 10);
            Assert.Contains("mixed_types", profile.QualityFlags);
            Assert.Equal(1, profile.TypeCounts["bool"]);
            Assert.Equal(1, profile.TypeCounts["number"]);
            Assert.Equal(1, profile.TypeCounts["text"]);
        }

        [Fact]
        public void Text_stats_and_padded_flag()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix(" 苹果", "苹果", "banana"), 1, 10);
            Assert.NotNull(profile.Text);
            Assert.Equal(3, profile.Text.Count);
            Assert.Equal(2, profile.Text.MinLength); // "苹果"（trim 后）=2，"banana"=6
            Assert.Equal(6, profile.Text.MaxLength);
            Assert.Equal(1, profile.Text.PaddedCount);
            Assert.Contains("padded_text", profile.QualityFlags);
        }

        [Fact]
        public void Top_values_sorted_by_count_desc_then_value_asc()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix("b", "a", "b", "a", "c", "b"), 1, 2);
            Assert.Equal(2, profile.TopValues.Count);
            Assert.Equal("b", profile.TopValues[0].Value);
            Assert.Equal(3, profile.TopValues[0].Count);
            Assert.Equal("a", profile.TopValues[1].Value);
            Assert.Equal(2, profile.TopValues[1].Count);
            Assert.Equal(3, profile.DistinctCount);
        }

        [Fact]
        public void Empty_and_out_of_range_inputs_are_safe()
        {
            var empty = ColumnProfileEngine.Compute(OneColumnMatrix(), 1, 10);
            Assert.Equal(0, empty.TotalRows);
            Assert.Empty(empty.TopValues);
            Assert.Null(empty.Numbers);
            Assert.Null(empty.Text);

            var matrix = Matrix(2, 2);
            var oor = ColumnProfileEngine.Compute(matrix, 5, 10);
            Assert.Equal(0, oor.TotalRows);

            var nullInput = ColumnProfileEngine.Compute(null, 1, 10);
            Assert.Equal(0, nullInput.TotalRows);
        }

        [Fact]
        public void Date_and_constant_flags()
        {
            var profile = ColumnProfileEngine.Compute(
                OneColumnMatrix(new DateTime(2026, 1, 1), new DateTime(2026, 1, 2)), 1, 10);
            Assert.Equal(2, profile.TypeCounts["date"]);

            var constant = ColumnProfileEngine.Compute(
                OneColumnMatrix("同值", "同值", "同值"), 1, 10);
            Assert.Contains("constant", constant.QualityFlags);
            Assert.Equal(1, constant.DistinctCount);
        }
    }
}
