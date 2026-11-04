using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// TextEncodingDetector 编码探测锁定（工具分类审查第九批：encoding=auto 硬返 UTF-8 乱码真缺陷）。
    /// 探测链：BOM 优先 → UTF-8 严格解码 → GBK 兜底。
    /// </summary>
    public class TextEncodingDetectorTests
    {
        [Fact]
        public void Utf8_bom_detected()
        {
            var bytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("月份,金额\n1月,100")).ToArray();
            var result = TextEncodingDetector.Detect(bytes);
            Assert.Equal("utf-8-bom", result.Item1);
            var lines = TextEncodingDetector.AutoDecodeLines(bytes, -1);
            Assert.Equal("utf-8-bom", lines.Item2);
            Assert.Contains("月份,金额", lines.Item1[0]);
        }

        [Fact]
        public void Utf16_le_bom_detected()
        {
            var bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("月份,金额\n1月,100")).ToArray();
            var result = TextEncodingDetector.Detect(bytes);
            Assert.Equal("utf-16-le", result.Item1);
            var lines = TextEncodingDetector.AutoDecodeLines(bytes, -1);
            Assert.Contains("月份,金额", lines.Item1[0]);
        }

        [Fact]
        public void Utf8_no_bom_still_decodes()
        {
            var bytes = Encoding.UTF8.GetBytes("月份,金额\n1月,100");
            var result = TextEncodingDetector.Detect(bytes);
            Assert.Equal("utf-8", result.Item1);
            var lines = TextEncodingDetector.AutoDecodeLines(bytes, -1);
            Assert.Contains("月份,金额", lines.Item1[0]);
        }

        [Fact]
        public void Gbk_no_bom_falls_back_and_decodes_chinese()
        {
            // 中文记事本「ANSI」另存 = GBK 无 BOM——此前硬返 UTF-8 全乱码，现 GBK 兜底正确解码
            var gbk = Encoding.GetEncoding(936);
            var bytes = gbk.GetBytes("月份,金额\n1月,100");
            var result = TextEncodingDetector.Detect(bytes);
            Assert.Equal("gbk", result.Item1);
            var lines = TextEncodingDetector.AutoDecodeLines(bytes, -1);
            Assert.Contains("月份,金额", lines.Item1[0]); // 中文正确解码（不乱码）
        }

        [Fact]
        public void Ascii_decodes_as_utf8()
        {
            var bytes = Encoding.ASCII.GetBytes("month,amount\nJan,100");
            var result = TextEncodingDetector.Detect(bytes);
            Assert.Equal("utf-8", result.Item1);
        }

        [Fact]
        public void MaxLines_limit_works()
        {
            var bytes = Encoding.UTF8.GetBytes("a\nb\nc\nd\ne");
            var lines = TextEncodingDetector.AutoDecodeLines(bytes, 2);
            Assert.Equal(2, lines.Item1.Length);
        }

        [Fact]
        public void Empty_bytes_return_empty()
        {
            var lines = TextEncodingDetector.AutoDecodeLines(new byte[0], -1);
            Assert.Empty(lines.Item1);
        }
    }

}
