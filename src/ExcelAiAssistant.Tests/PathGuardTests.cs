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