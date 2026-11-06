using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 智能表（ListObject）名预校验锁定（工具分类审查第八批，第四批命名预校验家族）：
    /// 规则——非空、≤255、仅字母/数字/下划线、禁数字开头、禁与 A1 引用形态冲突。
    /// </summary>
    public class TableNameValidatorTests
    {
        [Theory]
        [InlineData("销售数据")]
        [InlineData("Sales_2026")]
        [InlineData("_private")]
        [InlineData("T")]
        public void Valid_names_pass(string name)
        {
            Assert.Null(ExcelStructureService.ValidateTableName(name));
        }

        [Theory]
        [InlineData("销售 数据")]   // 含空格
        [InlineData("表名!")]       // 含符号
        [InlineData("Rate%")]       // 含符号
        public void Illegal_characters_rejected(string name)
        {
            var error = ExcelStructureService.ValidateTableName(name);
            Assert.NotNull(error);
            Assert.Contains("字母/数字/下划线", error);
        }

        [Theory]
        [InlineData("2026报表")]
        [InlineData("9x")]
        public void Digit_start_rejected(string name)
        {
            var error = ExcelStructureService.ValidateTableName(name);
            Assert.NotNull(error);
            Assert.Contains("数字开头", error);
        }

        [Theory]
        [InlineData("B2")]
        [InlineData("xfd1048576")]
        [InlineData("A1")]
        public void Cell_reference_shape_rejected(string name)
        {
            var error = ExcelStructureService.ValidateTableName(name);
            Assert.NotNull(error);
            Assert.Contains("单元格引用", error);
        }

        [Fact]
        public void Over_255_chars_rejected()
        {
            Assert.NotNull(ExcelStructureService.ValidateTableName(new string('表', 256)));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Empty_rejected(string name)
        {
            Assert.NotNull(ExcelStructureService.ValidateTableName(name));
        }
    }
}
