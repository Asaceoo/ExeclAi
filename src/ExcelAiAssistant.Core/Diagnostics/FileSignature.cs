using System;
using System.IO;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// M28：文件魔数嗅探（纯逻辑，无 COM）——export/生成类工具的 verify_file 判定基础。
    /// 只读文件头若干字节；不解析文件内容，不提供"格式完整性"结论。
    /// </summary>
    public static class FileSignature
    {
        public const int ProbeBytes = 512;

        /// <summary>嗅探文件头：pdf / zip / png / jpeg / gif / text / binary / empty / missing / unknown。</summary>
        public static string Detect(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) { return "missing"; }
                var buffer = new byte[ProbeBytes];
                int read;
                using (var stream = File.OpenRead(path))
                {
                    read = stream.Read(buffer, 0, buffer.Length);
                }

                if (read == 0) { return "empty"; }
                if (StartsWith(buffer, read, new byte[] { 0x25, 0x50, 0x44, 0x46 })) { return "pdf"; } // %PDF
                if (StartsWith(buffer, read, new byte[] { 0x50, 0x4B, 0x03, 0x04 }) ||
                    StartsWith(buffer, read, new byte[] { 0x50, 0x4B, 0x05, 0x06 })) { return "zip"; } // PK (xlsx/docx/zip)
                if (StartsWith(buffer, read, new byte[] { 0x89, 0x50, 0x4E, 0x47 })) { return "png"; }
                if (StartsWith(buffer, read, new byte[] { 0xFF, 0xD8, 0xFF })) { return "jpeg"; }
                if (StartsWith(buffer, read, new byte[] { 0x47, 0x49, 0x46, 0x38 })) { return "gif"; } // GIF8
                if (StartsWith(buffer, read, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 })) { return "ole2"; } // 老式 OLE2（.xls/.doc）
                // M28 审查：UTF-16/UTF-8 BOM 文本（PowerShell -Encoding Unicode 等）含 NUL，不能判 binary
                if (StartsWith(buffer, read, new byte[] { 0xFF, 0xFE }) ||
                    StartsWith(buffer, read, new byte[] { 0xFE, 0xFF }) ||
                    StartsWith(buffer, read, new byte[] { 0xEF, 0xBB, 0xBF })) { return "text"; }
                return LooksLikeText(buffer, read) ? "text" : "binary";
            }
            catch
            {
                return "unknown";
            }
        }

        /// <summary>按扩展名推导期望魔数（大小写不敏感）；未知扩展名返回 null（只做存在性/大小核验）。</summary>
        public static string ExpectedFromExtension(string path)
        {
            var ext = Path.GetExtension(path ?? string.Empty).ToLowerInvariant();
            switch (ext)
            {
                case ".pdf": return "pdf";
                case ".xlsx":
                case ".xlsm":
                case ".xlsb":
                case ".docx":
                case ".pptx":
                case ".zip": return "xlsx";
                case ".xls":
                case ".doc": return "ole2";
                case ".png": return "png";
                case ".jpg":
                case ".jpeg": return "jpeg";
                case ".gif": return "gif";
                case ".csv":
                case ".txt":
                case ".md":
                case ".json":
                case ".log":
                case ".xml":
                case ".html": return "text";
                default: return null;
            }
        }

        /// <summary>期望魔数与实测魔数是否匹配（xlsx 实为 zip 容器；大小写不敏感）。</summary>
        public static bool Matches(string expected, string detected)
        {
            if (string.IsNullOrWhiteSpace(expected) ||
                string.Equals(expected, "auto", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(expected, "any", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var exp = expected.ToLowerInvariant().TrimStart('.');
            var det = (detected ?? string.Empty).ToLowerInvariant();
            if (exp == "jpg") { exp = "jpeg"; } // M28 审查：jpg/jpeg 别名归一
            if (string.Equals(exp, det, StringComparison.Ordinal)) { return true; }
            return string.Equals(exp, "xlsx", StringComparison.Ordinal) && string.Equals(det, "zip", StringComparison.Ordinal);
        }

        private static bool StartsWith(byte[] buffer, int length, byte[] signature)
        {
            if (length < signature.Length) { return false; }
            for (var i = 0; i < signature.Length; i++)
            {
                if (buffer[i] != signature[i]) { return false; }
            }
            return true;
        }

        private static bool LooksLikeText(byte[] buffer, int length)
        {
            // 控制字符（除制表符/换行/回车外）一律视为二进制
            for (var i = 0; i < length; i++)
            {
                var b = buffer[i];
                if (b == 0) { return false; }
                if (b < 0x09 || (b > 0x0D && b < 0x20)) { return false; }
            }
            return true;
        }
    }
}
