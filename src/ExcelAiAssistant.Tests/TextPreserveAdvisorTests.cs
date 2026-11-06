using ExcelAiAssistant.Core.Diagnostics;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// v3.8.1.76 写入文本保护回归锁（审查批次真机实证驱动）：
    /// 18 位身份证被 COM Value2 写入重解析为 5.787E+17（精度丢失）、"007"→7（前导零丢失）。
    /// TextPreserveAdvisor 判定口径：纯数字 ≥11 位或前导零纯数字 → 锁文本；日期串不拦但重解析要上报。
    /// </summary>
    public sealed class TextPreserveAdvisorTests
    {
        // ===================== ShouldForceText =====================

        [Theory]
        [InlineData("578712345678901234", true)]   // 18 位身份证（真机实证损坏形态）
        [InlineData("91310000MA1K35X79H" /* 含字母统一社会信用代码 */, false)]
        [InlineData("13800138000", true)]           // 11 位手机号
        [InlineData("1380013800", false)]           // 10 位普通数值串
        [InlineData("007", true)]                   // 前导零编码
        [InlineData("00", true)]
        [InlineData("0", false)]                    // 单字符不锁
        [InlineData("1e5", false)]                  // 非纯数字
        [InlineData(" 12345678901", false)]         // 含空格
        [InlineData("2026-01", false)]              // 文本月份不拦（看板锁型另有先例）
        [InlineData("1987-07-04", false)]           // 日期串不拦（Excel 原生日期化与手工输入一致）
        [InlineData("", false)]
        public void ShouldForceText_classifies_digit_strings(object input, bool expected)
        {
            Assert.Equal(expected, TextPreserveAdvisor.ShouldForceText(input));
        }

        [Fact]
        public void ShouldForceText_non_string_scalars_never_lock()
        {
            Assert.False(TextPreserveAdvisor.ShouldForceText(12345));
            Assert.False(TextPreserveAdvisor.ShouldForceText(123.45));
            Assert.False(TextPreserveAdvisor.ShouldForceText(null));
            Assert.False(TextPreserveAdvisor.ShouldForceText(new object()));
        }

        [Fact]
        public void ShouldForceText_jvalue_string_is_recognized()
        {
            // 踩坑同族：JSON 反序列化到 C# 侧是 JValue 而非 System.String（write_ranges「项格式无效」教训）
            Assert.True(TextPreserveAdvisor.ShouldForceText(new JValue("578712345678901234")));
            Assert.False(TextPreserveAdvisor.ShouldForceText(new JValue(12345)));
        }

        // ===================== FindTextLockCells =====================

        [Fact]
        public void FindTextLockCells_returns_one_based_coordinates()
        {
            var values = new object[2, 2]
            {
                { "578712345678901234", "abc" },
                { 123, "007" }
            };
            var cells = TextPreserveAdvisor.FindTextLockCells(values);
            Assert.Equal(2, cells.Count);
            Assert.Contains(cells, kv => kv.Key == 1 && kv.Value == 1);
            Assert.Contains(cells, kv => kv.Key == 2 && kv.Value == 2);
        }

        [Fact]
        public void FindTextLockCells_empty_and_null_are_safe()
        {
            Assert.Empty(TextPreserveAdvisor.FindTextLockCells(null));
            Assert.Empty(TextPreserveAdvisor.FindTextLockCells(new object[1, 1] { { "abc" } }));
        }

        // ===================== ShouldReportCoercion =====================

        [Theory]
        [InlineData("1987-07-04", true)]    // 日期串重解析必须可见（真机实证 F15 漂移）
        [InlineData("578712345678901234", true)]
        [InlineData("123", false)]           // 短数字串数值化属无歧义等价
        [InlineData("abcdefg", false)]       // <8 字符不上报防噪声
        [InlineData("abcdefgh", true)]
        public void ShouldReportCoercion_filters_short_strings(object input, bool expected)
        {
            Assert.Equal(expected, TextPreserveAdvisor.ShouldReportCoercion(input));
        }

        [Fact]
        public void ShouldReportCoercion_handles_null_and_jvalue()
        {
            Assert.False(TextPreserveAdvisor.ShouldReportCoercion(null));
            Assert.True(TextPreserveAdvisor.ShouldReportCoercion(new JValue("2026-01-01")));
        }
    }
}
