using System;
using System.IO;
using ExcelAiAssistant.Core.Configuration;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class MemoryStoreTests : IDisposable
    {
        private readonly string filePath;
        private readonly MemoryStore store;
        private const string BookA = "BookA.xlsx";
        private const string BookB = "BookB.xlsx";

        public MemoryStoreTests()
        {
            filePath = Path.Combine(Path.GetTempPath(), $"excel-ai-memory-{Guid.NewGuid():N}.json");
            store = new MemoryStore(filePath);
        }

        [Fact]
        public void Defaults_are_empty()
        {
            Assert.Equal(string.Empty, store.Preferences);
            Assert.Equal(string.Empty, store.Notes);
            Assert.Empty(store.RecentSummaries(BookA));
            Assert.Equal(string.Empty, store.BuildInjectionText(BookA));
        }

        [Fact]
        public void Preferences_and_notes_persist_across_instances()
        {
            store.Preferences = "图表默认中文标题；写入优先公式。";
            store.Notes = "用户偏好简洁回复。";

            var reloaded = new MemoryStore(filePath);
            Assert.Equal("图表默认中文标题；写入优先公式。", reloaded.Preferences);
            Assert.Equal("用户偏好简洁回复。", reloaded.Notes);
        }

        [Fact]
        public void BuildInjectionText_contains_preferences_and_notes()
        {
            store.Preferences = "默认 HTML 报告";
            store.Notes = "备注内容";

            var text = store.BuildInjectionText(BookA);
            Assert.Contains("用户长期偏好：默认 HTML 报告", text);
            Assert.Contains("用户备注：备注内容", text);
        }

        [Fact]
        public void AddRecentSummary_keeps_latest_and_caps_at_20()
        {
            for (var i = 0; i < 25; i++)
            {
                store.AddRecentSummary(BookA, $"摘要 {i}");
            }

            Assert.Equal(20, store.RecentSummaries(BookA).Count);
            Assert.Equal("摘要 5", store.RecentSummaries(BookA)[0]);
            Assert.Equal("摘要 24", store.RecentSummaries(BookA)[19]);
        }

        [Fact]
        public void AddRecentSummary_ignores_blank()
        {
            store.AddRecentSummary(BookA, "   ");
            store.AddRecentSummary(BookA, null);
            Assert.Empty(store.RecentSummaries(BookA));
        }

        [Fact]
        public void BuildInjectionText_limits_recent_summaries()
        {
            store.AddRecentSummary(BookA, "摘要 一");
            store.AddRecentSummary(BookA, "摘要 二");
            store.AddRecentSummary(BookA, "摘要 三");
            store.AddRecentSummary(BookA, "摘要 四");

            var text = store.BuildInjectionText(BookA, recentSummaryCount: 2);
            Assert.Contains("摘要 三", text);
            Assert.Contains("摘要 四", text);
            Assert.DoesNotContain("摘要 一", text);
        }

        [Fact]
        public void ClearRecentSummaries_empties_list()
        {
            store.AddRecentSummary(BookA, "某摘要");
            store.ClearRecentSummaries(BookA);
            Assert.Empty(store.RecentSummaries(BookA));
        }

        [Fact]
        public void RecentSummaries_are_isolated_per_workbook()
        {
            store.AddRecentSummary(BookA, "A 的摘要");
            store.AddRecentSummary(BookB, "B 的摘要");

            Assert.Equal("A 的摘要", Assert.Single(store.RecentSummaries(BookA)));
            Assert.Equal("B 的摘要", Assert.Single(store.RecentSummaries(BookB)));

            var textA = store.BuildInjectionText(BookA);
            Assert.Contains("A 的摘要", textA);
            Assert.DoesNotContain("B 的摘要", textA);
        }

        public void Dispose()
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}
