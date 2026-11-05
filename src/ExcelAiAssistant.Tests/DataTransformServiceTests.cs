using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class DataTransformServiceTests
    {
        [Fact]
        public void Removes_blank_rows_and_reports_count()
        {
            var values = new object[,]
            {
                { "A", 1D },
                { null, null },
                { "B", 2D }
            };

            var result = DataTransformService.RemoveBlankRows(values);

            Assert.Equal(1, result.ChangedCount);
            Assert.Equal(2, result.Values.GetLength(0));
            Assert.Equal("B", result.Values[1, 0]);
        }

        [Fact]
        public void Fills_only_blank_cells_in_selected_column()
        {
            var values = new object[,]
            {
                { "A", null },
                { "B", "x" }
            };

            var result = DataTransformService.FillBlanks(values, 2, "N/A");

            Assert.Equal(1, result.ChangedCount);
            Assert.Equal("N/A", result.Values[0, 1]);
            Assert.Equal("x", result.Values[1, 1]);
        }

        [Fact]
        public void Trims_whitespace_on_string_cells_only()
        {
            var values = new object[,]
            {
                { "  a  ", "b " },
                { "c", 1D },
                { null, null }
            };

            var result = DataTransformService.TrimWhitespace(values);

            Assert.Equal(2, result.ChangedCount);
            Assert.Equal("a", result.Values[0, 0]);
            Assert.Equal("b", result.Values[0, 1]);
            Assert.Equal("c", result.Values[1, 0]);
            Assert.Equal(1D, result.Values[1, 1]);
            Assert.Null(result.Values[2, 0]);
        }

        [Fact]
        public void Removes_non_printable_control_chars_but_keeps_visible_text()
        {
            var values = new object[,] { { "a\u0001b", "x\t", "你好" } };

            var result = DataTransformService.RemoveNonPrintable(values);

            Assert.Equal(2, result.ChangedCount);
            Assert.Equal("ab", result.Values[0, 0]);
            Assert.Equal("x", result.Values[0, 1]);
            Assert.Equal("你好", result.Values[0, 2]);
        }

        [Fact]
        public void Filters_numeric_rows_by_operator()
        {
            var values = new object[,]
            {
                { "A", 1D },
                { "B", 2D },
                { "C", 3D }
            };

            var result = DataTransformService.ApplyFilter(values, 2, "gte", "2");

            Assert.Equal(2, result.MatchedRows);
            Assert.Equal(2, result.Values.GetLength(0));
            Assert.Equal("B", result.Values[0, 0]);
        }

        // 回归：修复前 FixDates 用 0 基下标访问 values[row, columnIndex] 而非 columnIndex-1，
        // 当 columnIndex == 列数（如列数=2、columnIndex=2）且该列全为可解析日期文本时越界抛 IndexOutOfRangeException。
        [Fact]
        public void FixDates_of_last_column_does_not_throw_and_converts()
        {
            var values = new object[,]
            {
                { "名称", "日期" },
                { "甲", "2023/06/05" },
                { "丙", "2023/06/06" }
            };

            var result = DataTransformService.FixDates(values, 2);

            Assert.Equal(2, result.ChangedCount);
            Assert.IsType<DateTime>(result.Values[1, 1]);
            Assert.IsType<DateTime>(result.Values[2, 1]);
            Assert.Equal(new DateTime(2023, 6, 5), result.Values[1, 1]);
        }

        // 回归：FixDates 对单列数据（列数=1）也须正确处理 columnIndex=1，不越界。
        [Fact]
        public void FixDates_single_column_works()
        {
            var values = new object[,]
            {
                { "2023/06/05" },
                { "2023/06/07" }
            };

            var result = DataTransformService.FixDates(values, 1);

            Assert.Equal(2, result.ChangedCount);
            Assert.Equal(new DateTime(2023, 6, 5), result.Values[0, 0]);
        }
    }

    public sealed class ExcelRangeHelperTests
    {
        [Theory]
        [InlineData("A1:C3", 2, 3, "A1:C2")]
        [InlineData("Sheet1!B2:D5", 1, 2, "'Sheet1'!B2:C2")]
        public void Resizes_addresses(string source, int rows, int columns, string expected)
        {
            Assert.Equal(expected, ExcelRangeHelper.Resize(source, rows, columns));
        }
    }
}
