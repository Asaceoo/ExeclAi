using System.Linq;
using ExcelAiAssistant.Core.Errors;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>M36-3：错误码目录必须每条都有可执行的下一步动作，不允许只写"失败"。</summary>
    public class ErrorCatalogTests
    {
        [Fact]
        public void EveryEntryHasMessageAndNextAction()
        {
            Assert.NotEmpty(ErrorCatalog.All);
            foreach (var e in ErrorCatalog.All)
            {
                Assert.False(string.IsNullOrWhiteSpace(e.Message), e.Code + " 缺 message");
                Assert.False(string.IsNullOrWhiteSpace(e.NextAction), e.Code + " 缺 next_action");
            }
        }

        [Fact]
        public void CoreCodesAreRegistered()
        {
            foreach (var code in new[] { "EXCEL_WARMING", "PROFILE_FILTERED", "DUPLICATE_WRITE_SUPPRESSED", "UNKNOWN_TOOL", "SELF_CHECK_FAILED" })
            {
                Assert.True(ErrorCatalog.TryGet(code, out _), code + " 未登记");
            }
        }

        [Fact]
        public void UnknownCodeReturnsFalseWithoutThrowing()
        {
            Assert.False(ErrorCatalog.TryGet("NOT_A_REAL_CODE", out var e));
            Assert.Null(e);
            Assert.Null(ErrorCatalog.MessageOf("NOT_A_REAL_CODE"));
        }
    }
}
