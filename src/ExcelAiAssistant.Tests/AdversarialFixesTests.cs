using System;
using System.Linq;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>对抗性审查修复的纯逻辑回归测试（可脱离 Excel 运行）。</summary>
    public sealed class AdversarialFixesTests
    {
        // ===================== 公式引用改写（ConvertRefsInFormula） =====================
        [Theory]
        [InlineData("=A1+B2", "to_absolute", "=$A$1+$B$2")]
        [InlineData("=$A$1+$B$2", "to_relative", "=A1+B2")]
        [InlineData("=SUM(A1:B2)", "to_absolute", "=SUM($A$1:$B$2)")]
        public void Convert_refs_toggles_absolute_relative(string input, string mode, string expected)
        {
            Assert.Equal(expected, ExcelAdvancedEditService.ConvertRefsInFormula(input, mode));
        }

        [Fact]
        public void Convert_refs_does_not_corrupt_string_literals()
        {
            // ="abc123" 里的 abc123 是文本字面量，绝不能改写成 $abc$123
            Assert.Equal("=\"abc123\"&A1", ExcelAdvancedEditService.ConvertRefsInFormula("=\"abc123\"&A1", "to_relative"));
            Assert.Equal("=\"abc123\"&$A$1", ExcelAdvancedEditService.ConvertRefsInFormula("=\"abc123\"&A1", "to_absolute"));
        }

        [Fact]
        public void Convert_refs_skips_sheet_name_single_quotes_but_rewrites_the_ref()
        {
            Assert.Equal("='Sheet 1'!$B$2",
                ExcelAdvancedEditService.ConvertRefsInFormula("='Sheet 1'!B2", "to_absolute"));
        }

        [Fact]
        public void Convert_refs_handles_lowercase_and_ignores_word_boundary()
        {
            Assert.Equal("=$a$1", ExcelAdvancedEditService.ConvertRefsInFormula("=a1", "to_absolute"));
            // Total1 不是单元格引用（词内字母紧跟数字），不应被改写
            Assert.Equal("=SUM(Total1)", ExcelAdvancedEditService.ConvertRefsInFormula("=SUM(Total1)", "to_absolute"));
        }

        [Fact]
        public void Convert_refs_does_not_corrupt_numeric_function_names()
        {
            // LOG10 / BIN2DEC 等函数名含数字，尾随 ( 时绝不能改写（否则成错误引用直接损坏公式）。
            Assert.Equal("=LOG10($A$1)", ExcelAdvancedEditService.ConvertRefsInFormula("=LOG10(A1)", "to_absolute"));
            Assert.Equal("=BIN2DEC($A$1)+$C$1", ExcelAdvancedEditService.ConvertRefsInFormula("=BIN2DEC(A1)+C1", "to_absolute"));
            Assert.Equal("=LOG10(A1)", ExcelAdvancedEditService.ConvertRefsInFormula("=LOG10(A1)", "to_relative"));
        }

        // ===================== 分隔符探测（DetectDelimiter） =====================
        [Fact]
        public void Detect_delimiter_picks_space_when_space_wins()
        {
            Assert.Equal(" ", ExcelImportService.DetectDelimiter(new[] { "a b c", "d e f", "g h i" }));
        }

        [Fact]
        public void Detect_delimiter_picks_comma_tab_semicolon()
        {
            Assert.Equal(",", ExcelImportService.DetectDelimiter(new[] { "a,b,c", "d,e,f" }));
            Assert.Equal("\t", ExcelImportService.DetectDelimiter(new[] { "a\tb\tc", "d\te\tf" }));
            Assert.Equal(";", ExcelImportService.DetectDelimiter(new[] { "a;b;c", "d;e;f" }));
        }

        [Fact]
        public void Detect_delimiter_falls_back_to_comma_when_no_strong_candidate()
        {
            Assert.Equal(",", ExcelImportService.DetectDelimiter(new[] { "abcdef", "ghijkl" }));
        }

        // ===================== 日期标准化防误写（TryParseDate / IsDateLikeFormat） =====================
        [Fact]
        public void Numeric_non_date_cell_is_not_silently_turned_into_a_date()
        {
            // 年份/ID 列：General 格式的纯数字绝不能被误识别为序列日期
            Assert.Null(ExcelUtilityToolsService.TryParseDate(2023.0, null, "General"));
            Assert.Null(ExcelUtilityToolsService.TryParseDate(45298.0, null, "0.00"));
        }

        [Fact]
        public void Numeric_cell_with_date_format_is_converted()
        {
            var d = ExcelUtilityToolsService.TryParseDate(45298.0, null, "yyyy/m/d");
            Assert.True(d.HasValue);
            Assert.InRange(d.Value.Year, 2023, 2026);
        }

        [Fact]
        public void Numeric_cell_with_explicit_input_format_is_converted()
        {
            // 用户显式传 input_format 视为明确意图，允许整列按序列日期处理
            Assert.True(ExcelUtilityToolsService.TryParseDate(45298.0, "yyyy-mm-dd", null).HasValue);
        }

        [Fact]
        public void Text_date_parses_regardless_of_number_format()
        {
            Assert.True(ExcelUtilityToolsService.TryParseDate("2024-01-05", null, null).HasValue);
        }

        [Theory]
        [InlineData("yyyy/m/d", true)]
        [InlineData("m/d/yy", true)]
        [InlineData("h:mm", true)]
        [InlineData("General", false)]
        [InlineData("0.00", false)]
        [InlineData("0%", false)]
        [InlineData("#,##0", false)]
        public void Date_like_format_detection(string fmt, bool expected)
        {
            Assert.Equal(expected, ExcelUtilityToolsService.IsDateLikeFormat(fmt));
        }

        // ===================== 报告 max_columns 真正生效（ReportTextBuilder） =====================
        [Fact]
        public void Report_respects_max_columns_cap()
        {
            var matrix = new object[20, 5];
            for (var r = 0; r < 20; r++)
            {
                matrix[r, 0] = "类别" + r;
                matrix[r, 1] = 100.0 + r;
                matrix[r, 2] = 200.0 + r;
                matrix[r, 3] = 300.0 + r;
                matrix[r, 4] = 400.0 + r;
            }

            var report = ReportTextBuilder.Build(matrix, "Sheet1", "A1:E20", 2, maxNumericColumns: 1);
            // 概览章节应呈现「1 列」（max_columns=1 生效，而非默认 50）
            var overview = report.Sections.First(s => s.Title == "概览");
            Assert.Contains("1 列", overview.Lines.Single(l => l.Contains("维度")));
        }

        // ===================== dedup_range 判重列越界护栏（防整表坍缩） =====================
        [Fact]
        public void Dedup_out_of_range_column_is_rejected()
        {
            // columns=[99]（列不存在）会让 DedupKey 跳过所有列 → 每行键退化为空串 → 整表坍缩为单行。
            // 护栏必须显式揪出越界列号，而非静默接受。
            Assert.NotNull(ExcelDataToolsService.FirstOutOfRangeColumn(new[] { 99 }, 5));
            Assert.NotNull(ExcelDataToolsService.FirstOutOfRangeColumn(new[] { 0 }, 5));
            Assert.NotNull(ExcelDataToolsService.FirstOutOfRangeColumn(new[] { 1, 6 }, 5));
            // 全部合法 + 空集 → 无越界列
            Assert.Null(ExcelDataToolsService.FirstOutOfRangeColumn(new[] { 1, 5 }, 5));
            Assert.Null(ExcelDataToolsService.FirstOutOfRangeColumn(new int[] { }, 5));
        }
    }
}