using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 工作表名预校验锁定（工具分类审查第四批）：Excel 工作表名 ≤31 字符且不含 : \ / ? * [ ]——
    /// AI 起名带斜杠（「2026/09」）/超长此前抛裸 COM 异常无法自纠，现人话 INVALID_SHEET_NAME。
    /// </summary>
    public class SheetNameValidatorTests
    {
        [Theory]
        [InlineData("销售数据")]
        [InlineData("2026-09 月报")]
        [InlineData("Sheet1")]
        [InlineData("a")] // 单字符
        public void Valid_names_pass(string name)
        {
            Assert.Null(ExcelStructureService.ValidateSheetName(name));
        }

        [Theory]
        [InlineData("2026/09")]
        [InlineData("报表:汇总")]
        [InlineData("a\\b")]
        [InlineData("什么是?")]
        [InlineData("乘*法")]
        [InlineData("区[间]")]
        public void Illegal_characters_rejected_with_human_message(string name)
        {
            var error = ExcelStructureService.ValidateSheetName(name);
            Assert.NotNull(error);
            Assert.Contains("不能包含", error);
        }

        [Fact]
        public void Over_31_chars_rejected()
        {
            var error = ExcelStructureService.ValidateSheetName(new string('数', 32));
            Assert.NotNull(error);
            Assert.Contains("31", error);
        }

        [Fact]
        public void Exactly_31_chars_pass()
        {
            Assert.Null(ExcelStructureService.ValidateSheetName(new string('数', 31)));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Empty_rejected(string name)
        {
            Assert.NotNull(ExcelStructureService.ValidateSheetName(name));
        }
    }
}
