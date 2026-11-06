using System;
using System.IO;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>M28：verify_file 的魔数嗅探纯逻辑回归（无 COM，CI 可跑）。</summary>
    public class FileSignatureTests
    {
        private static string TempFile(byte[] bytes)
        {
            var path = Path.Combine(Path.GetTempPath(), "sig-" + Guid.NewGuid().ToString("N") + ".bin");
            File.WriteAllBytes(path, bytes);
            return path;
        }

        [Fact]
        public void Detect_pdf_zip_png_jpeg_gif_text()
        {
            var cases = new[]
            {
                new { Bytes = new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D, 0x31 }, Magic = "pdf", Ext = ".pdf" },
                new { Bytes = new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00 }, Magic = "zip", Ext = ".xlsx" },
                new { Bytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A }, Magic = "png", Ext = ".png" },
                new { Bytes = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }, Magic = "jpeg", Ext = ".jpg" },
                new { Bytes = new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, Magic = "gif", Ext = ".gif" },
                new { Bytes = System.Text.Encoding.UTF8.GetBytes("品名,数量"), Magic = "text", Ext = ".csv" }
            };
            foreach (var item in cases)
            {
                var path = TempFile(item.Bytes);
                try
                {
                    Assert.Equal(item.Magic, FileSignature.Detect(path));
                    Assert.True(FileSignature.Matches(FileSignature.ExpectedFromExtension("x" + item.Ext), item.Magic),
                        item.Ext + " 的期望魔数应与实测匹配（xlsx 特判 zip 容器）");
                }
                finally
                {
                    try { File.Delete(path); } catch { }
                }
            }
        }

        [Fact]
        public void Detect_empty_missing_binary()
        {
            var empty = TempFile(new byte[0]);
            try
            {
                Assert.Equal("empty", FileSignature.Detect(empty));
                Assert.Equal("missing", FileSignature.Detect(Path.Combine(Path.GetTempPath(), "no-such-" + Guid.NewGuid().ToString("N"))));
                var binary = TempFile(new byte[] { 0x00, 0x01, 0x02, 0x03 });
                try { Assert.Equal("binary", FileSignature.Detect(binary)); }
                finally { try { File.Delete(binary); } catch { } }
            }
            finally
            {
                try { File.Delete(empty); } catch { }
            }
        }

        [Fact]
        public void Ole2_and_extension_aliases_are_supported()
        {
            var ole2 = TempFile(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00 });
            try
            {
                Assert.Equal("ole2", FileSignature.Detect(ole2));
                Assert.Equal("ole2", FileSignature.ExpectedFromExtension("x.xls"));
                Assert.Equal("xlsx", FileSignature.ExpectedFromExtension("x.docx"));
                Assert.True(FileSignature.Matches("jpg", "jpeg"));
                Assert.True(FileSignature.Matches("any", "binary"));
                Assert.False(FileSignature.Matches("empty", "ole2"));
                Assert.True(FileSignature.Matches("ole2", "ole2"));
            }
            finally
            {
                try { File.Delete(ole2); } catch { }
            }
        }

        [Fact]
        public void Utf16_and_utf8_bom_text_are_text_not_binary()
        {
            var utf16 = TempFile(new byte[] { 0xFF, 0xFE, (byte)'a', 0x00, (byte)'b', 0x00 });
            var utf16Be = TempFile(new byte[] { 0xFE, 0xFF, 0x00, (byte)'a', 0x00, (byte)'b' });
            var utf8Bom = TempFile(new byte[] { 0xEF, 0xBB, 0xBF, (byte)'h', (byte)'i' });
            try
            {
                Assert.Equal("text", FileSignature.Detect(utf16));
                Assert.Equal("text", FileSignature.Detect(utf16Be));
                Assert.Equal("text", FileSignature.Detect(utf8Bom));
                Assert.True(FileSignature.Matches(FileSignature.ExpectedFromExtension("x.csv"), "text"));
            }
            finally
            {
                foreach (var p in new[] { utf16, utf16Be, utf8Bom }) { try { File.Delete(p); } catch { } }
            }
        }

        [Fact]
        public void Matches_rejects_mismatch_and_accepts_auto()
        {
            Assert.False(FileSignature.Matches("pdf", "zip"));
            Assert.True(FileSignature.Matches("auto", "binary"));
            Assert.True(FileSignature.Matches(null, "binary"));
            Assert.True(FileSignature.Matches("xlsx", "zip"));
            Assert.True(FileSignature.Matches("PNG", "png"));
        }
    }
}
