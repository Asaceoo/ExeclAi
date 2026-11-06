using System;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// TextDateParser 回归锁定（XLL UDF M1 提公：ExcelUtilityToolsService.TryParseDate/IsDateLikeFormat
    /// 上移 Core 单一来源，Excel 侧转发；本测试锁提公后行为与原实现一致）。
    /// </summary>
    public class TextDateParserTests
    {
        [Fact]
        public void DateTime_passthrough()
        {
            var dt = new DateTime(2026, 9, 19);
            Assert.Equal(dt, TextDateParser.TryParseDate(dt, null, null));
        }

        [Fact]
        public void Serial_double_requires_date_like_format()
        {
            // 45298≈2024 年的 OADate；无格式/无 inputFormat 时不转换（防 ID/年份误判）
            Assert.Null(TextDateParser.TryParseDate(45298.0, null, "General"));
            Assert.Null(TextDateParser.TryParseDate(45298.0, null, null));
            // 日期格式 → 转 OADate
            var withFmt = TextDateParser.TryParseDate(45298.0, null, "yyyy\\-m\\-d");
            Assert.NotNull(withFmt);
            Assert.Equal(2024, withFmt.Value.Year);
            // 显式 inputFormat 视为明确意图
            var withInput = TextDateParser.TryParseDate(45298.0, "yyyy-mm-dd", null);
            Assert.NotNull(withInput);
            Assert.Equal(2024, withInput.Value.Year);
        }

        [Fact]
        public void Serial_double_out_of_range_rejected()
        {
            // 上限 2958466（9999-12-31 之外）与负值/零拒绝
            Assert.Null(TextDateParser.TryParseDate(2958466.0, null, "yyyy-mm-dd"));
            Assert.Null(TextDateParser.TryParseDate(0.0, null, "yyyy-mm-dd"));
            Assert.Null(TextDateParser.TryParseDate(-5.0, null, "yyyy-mm-dd"));
        }

        [Fact]
        public void Text_with_explicit_format_parses_exact()
        {
            var parsed = TextDateParser.TryParseDate("2026/09/19", "yyyy/MM/dd", null);
            Assert.NotNull(parsed);
            Assert.Equal(new DateTime(2026, 9, 19), parsed.Value);
            // 格式不符 → fail-closed 返 null
            Assert.Null(TextDateParser.TryParseDate("19-09-2026", "yyyy/MM/dd", null));
        }

        [Fact]
        public void Null_empty_whitespace_rejected()
        {
            Assert.Null(TextDateParser.TryParseDate(null, null, null));
            Assert.Null(TextDateParser.TryParseDate("", null, null));
            Assert.Null(TextDateParser.TryParseDate("   ", null, null));
        }

        [Fact]
        public void IsDateLikeFormat_detection()
        {
            Assert.True(TextDateParser.IsDateLikeFormat("yyyy-mm-dd"));
            Assert.True(TextDateParser.IsDateLikeFormat("m月d日"));
            Assert.True(TextDateParser.IsDateLikeFormat("h:mm"));
            Assert.False(TextDateParser.IsDateLikeFormat("General"));
            Assert.False(TextDateParser.IsDateLikeFormat("0.00"));
            Assert.False(TextDateParser.IsDateLikeFormat("@"));
            Assert.False(TextDateParser.IsDateLikeFormat(null));
            Assert.False(TextDateParser.IsDateLikeFormat(""));
        }
    }
}
