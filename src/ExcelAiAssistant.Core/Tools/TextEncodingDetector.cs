using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 文本文件编码自动探测（纯逻辑，可单测；工具分类审查第九批：
    /// ExcelImportService 的 encoding=auto 此前硬返 UTF-8——无 BOM 的 GBK/ANSI 中文文件
    /// （中文记事本「ANSI」另存常见）预览+导入全乱码。探测链：BOM 优先 → UTF-8 严格解码 → GBK 兜底。
    /// </summary>
    public static class TextEncodingDetector
    {
        /// <summary>
        /// 自动探测编码并解码文本（限幅 maxLines 行，-1 全读）。
        /// 返回 (行列表, 实际使用的编码名)。
        /// </summary>
        public static Tuple<string[], string> AutoDecodeLines(byte[] bytes, int maxLines)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return Tuple.Create(new string[0], "utf-8");
            }

            var encoding = Detect(bytes);
            var text = encoding.Item2.GetString(bytes);
            var lines = SplitLines(text, maxLines);
            return Tuple.Create(lines, encoding.Item1);
        }

        /// <summary>
        /// 编码探测：BOM 优先（UTF-32/UTF-16/UTF-8 BOM）；
        /// 无 BOM 时 UTF-8 严格解码（DecoderExceptionFallback 抛错即非 UTF-8）→ GBK(936) 兜底
        /// （中文记事本 ANSI=GBK，与 read_business_glossary 口径一致）。
        /// 返回 (编码名, Encoding)。
        /// </summary>
        public static Tuple<string, Encoding> Detect(byte[] bytes)
        {
            // BOM 优先
            if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00)
                return Tuple.Create("utf-32-le", (Encoding)Encoding.UTF32);
            if (bytes.Length >= 4 && bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF)
                return Tuple.Create("utf-32-be", (Encoding)Encoding.GetEncoding("utf-32-be"));
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Tuple.Create("utf-16-le", (Encoding)Encoding.Unicode);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Tuple.Create("utf-16-be", (Encoding)Encoding.BigEndianUnicode);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Tuple.Create("utf-8-bom", (Encoding)new UTF8Encoding(true));

            // 无 BOM：UTF-8 严格解码尝试（非法序列抛异常 → 非 UTF-8）
            var strictUtf8 = new UTF8Encoding(false, true);
            try
            {
                strictUtf8.GetString(bytes);
                return Tuple.Create("utf-8", (Encoding)new UTF8Encoding(false));
            }
            catch (DecoderFallbackException)
            {
                // 非 UTF-8 → GBK 兜底（中文 Windows ANSI）；GBK 无法注册时退 Default
                try
                {
                    return Tuple.Create("gbk", (Encoding)Encoding.GetEncoding(936));
                }
                catch (ArgumentException)
                {
                    return Tuple.Create("default", (Encoding)Encoding.Default);
                }
            }
        }

        private static string[] SplitLines(string text, int maxLines)
        {
            var lines = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return lines.ToArray();
            }

            using (var reader = new StringReader(text))
            {
                string line;
                var count = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                    count++;
                    if (maxLines > 0 && count >= maxLines)
                    {
                        break;
                    }
                }
            }

            return lines.ToArray();
        }
    }
}
