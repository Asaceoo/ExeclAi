using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ExcelAddressParserTests
    {
        [Theory]
        [InlineData("A1", 1, 1)]
        [InlineData("$AB$12", 12, 28)]
        [InlineData("XFD1048576", 1048576, 16384)]
        public void Parses_supported_addresses(string address, int expectedRow, int expectedColumn)
        {
            Assert.True(ExcelAddressParser.TryParseCell(address, out var row, out var column));
            Assert.Equal(expectedRow, row);
            Assert.Equal(expectedColumn, column);
        }

        [Theory]
        [InlineData("")]
        [InlineData("R1C1")]
        [InlineData("A0")]
        [InlineData("Sheet1!A1")]
        [InlineData("A1048577")]
        public void Rejects_unsupported_addresses(string address)
        {
            Assert.False(ExcelAddressParser.TryParseCell(address, out _, out _));
        }
    }
}
