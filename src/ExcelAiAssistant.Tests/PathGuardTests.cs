using System;
using Xunit;
using ExcelAiAssistant.Core.Security;

namespace ExcelAiAssistant.Tests
{
    public sealed class PathGuardTests
    {
        [Fact]
        public void Normal_absolute_path_passes()
        {
            var result = PathGuard.GuardOutputPath(@"C:\Users\me\Desktop\rpt.md");
            Assert.Equal(@"C:\Users\me\Desktop\rpt.md", result);
        }

        [Fact]
        public void Extension_completed_when_missing()
        {
            var result = PathGuard.GuardAndComplete(@"C:\Users\me\Desktop\chart", ".png");
            Assert.EndsWith(".png", result, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Relative_path_is_absolutized_against_server_cwd()
        {
            // M28 审查：Excel 与 server 的 CWD 可能不同——PathGuard 必须返回绝对路径，
            // 否则覆盖门/自检与真实落盘目录不一致（误拒或静默覆盖）。
            var result = PathGuard.GuardOutputPath(@"out\rel.xlsx");
            Assert.True(System.IO.Path.IsPathRooted(result), "必须返回绝对路径：" + result);
            Assert.EndsWith("rel.xlsx", result, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Unc_path_is_rejected_before_any_io()
        {
            // M28（第三轮）：UNC/设备路径一律拒绝（防 SMB/NTLM 外连），与 verify_file 同口径
            var ex = Assert.Throws<ArgumentException>(() => PathGuard.GuardOutputPath(@"\\server\share\a.pdf"));
            Assert.Contains("UNC", ex.Message);
        }

        [Fact]
        public void Backup_copies_existing_file_and_is_unique()
        {
            var dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pg_bak_" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(dir);
            var file = System.IO.Path.Combine(dir, "a.txt");
            System.IO.File.WriteAllText(file, "v1");
            try
            {
                var first = PathGuard.CreateBackup(file);
                Assert.True(System.IO.File.Exists(first));
                Assert.Equal("v1", System.IO.File.ReadAllText(first));
                System.IO.File.WriteAllText(file, "v2");
                var second = PathGuard.CreateBackup(file);
                Assert.NotEqual(first, second);
                Assert.Equal("v2", System.IO.File.ReadAllText(second));
                Assert.Null(PathGuard.CreateBackup(System.IO.Path.Combine(dir, "missing.txt")));

                // M28（第四轮自查）：陈旧源文件（写入时间 >30 天）刚建的备份不得被保留策略误删
                var oldFile = System.IO.Path.Combine(dir, "old.txt");
                System.IO.File.WriteAllText(oldFile, "legacy");
                System.IO.File.SetLastWriteTimeUtc(oldFile, DateTime.UtcNow.AddDays(-45));
                var oldBackup = PathGuard.CreateBackup(oldFile);
                Assert.True(System.IO.File.Exists(oldBackup), "陈旧源文件的新备份必须存活（备份时间=创建时刻）");
                var oldBackup2 = PathGuard.CreateBackup(oldFile);
                Assert.True(System.IO.File.Exists(oldBackup), "第二次备份后第一份仍在");
                Assert.True(System.IO.File.Exists(oldBackup2));
            }
            finally
            {
                try { System.IO.Directory.Delete(dir, true); } catch { }
            }
        }

        [Fact]
        public void Traversal_segment_is_rejected()
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                PathGuard.GuardOutputPath(@"C:\Users\me\Desktop\..\..\..\Windows\evil.txt"));
            Assert.Contains("..", ex.Message);
        }

        [Fact]
        public void By_forward_slash_traversal_is_rejected()
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                PathGuard.GuardOutputPath(@"C:/Users/me/../../Windows/evil.txt"));
            Assert.Contains("..", ex.Message);
        }

        [Fact]
        public void Nul_character_is_rejected()
        {
            var ex = Assert.Throws<ArgumentException>(() =>
                PathGuard.GuardOutputPath("C:\\Users\\me\\evil\0.txt"));
            Assert.Contains("NUL", ex.Message);
        }

        [Fact]
        public void Empty_path_is_rejected()
        {
            Assert.Throws<ArgumentException>(() => PathGuard.GuardOutputPath("   "));
            Assert.Throws<ArgumentException>(() => PathGuard.GuardOutputPath(null));
        }
    }
}