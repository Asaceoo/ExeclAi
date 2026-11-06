using System;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class SessionStoreTests : IDisposable
    {
        private readonly string directory;
        private readonly SessionStore store;
        private const string Book = "Book1.xlsx";

        public SessionStoreTests()
        {
            directory = Path.Combine(Path.GetTempPath(), $"excel-ai-session-{Guid.NewGuid():N}");
            store = new SessionStore(directory);
        }

        [Fact]
        public void Round_trips_messages_across_instances()
        {
            var sessionId = store.Create(Book);
            var messages = new[]
            {
                ChatMessage.User("第一轮问题"),
                ChatMessage.Assistant("第一轮回答"),
                ChatMessage.User("第二轮问题")
            };

            store.Save(Book, sessionId, messages, "第一批操作");

            var reloaded = new SessionStore(directory).Load(Book, sessionId);
            Assert.Equal(3, reloaded.Count);
            Assert.Equal("user", reloaded[0].Role);
            Assert.Equal("第一轮问题", reloaded[0].Content);
            Assert.Equal("第二轮问题", reloaded[2].Content);
        }

        [Fact]
        public void Load_returns_empty_for_unknown_session()
        {
            Assert.Empty(store.Load(Book, "missing-session"));
        }

        [Fact]
        public void Round_trips_tool_events_on_assistant_message()
        {
            var sessionId = store.Create(Book);
            var assistant = ChatMessage.Assistant("已写入汇总。");
            assistant.ToolEvents = new[]
            {
                new ToolEventView { Tool = "write_excel_range", Ok = true, Summary = "写入 F28:G36" },
                new ToolEventView { Tool = "create_excel_chart", Ok = false, ErrorCode = "EXCEL_COM_ERROR", ErrorMessage = "图表创建失败" }
            };
            var messages = new[] { ChatMessage.User("做汇总"), assistant };

            store.Save(Book, sessionId, messages, "工具事件回环");

            var reloaded = new SessionStore(directory).Load(Book, sessionId);
            var loaded = reloaded[1];
            Assert.Equal("assistant", loaded.Role);
            Assert.NotNull(loaded.ToolEvents);
            Assert.Equal(2, loaded.ToolEvents.Count);
            Assert.Equal("write_excel_range", loaded.ToolEvents[0].Tool);
            Assert.True(loaded.ToolEvents[0].Ok);
            Assert.Equal("写入 F28:G36", loaded.ToolEvents[0].Summary);
            Assert.False(loaded.ToolEvents[1].Ok);
            Assert.Equal("EXCEL_COM_ERROR", loaded.ToolEvents[1].ErrorCode);
        }

        [Fact]
        public void Round_trips_auto_sinked_vba_flag_on_assistant_message()
        {
            var sessionId = store.Create(Book);
            var assistant = ChatMessage.Assistant("已执行 VBA。");
            assistant.AutoSinkedVba = true;
            var messages = new[] { ChatMessage.User("执行 VBA"), assistant };

            store.Save(Book, sessionId, messages, "自动沉淀标记回环");

            var reloaded = new SessionStore(directory).Load(Book, sessionId);
            Assert.True(reloaded[1].AutoSinkedVba);
        }

        [Fact]
        public void Auto_sinked_vba_defaults_to_false_when_not_set()
        {
            var sessionId = store.Create(Book);
            store.Save(Book, sessionId, new[] { ChatMessage.Assistant("普通回复") }, "默认值");

            var reloaded = new SessionStore(directory).Load(Book, sessionId);
            Assert.False(reloaded[0].AutoSinkedVba);
        }

        [Fact]
        public void Load_respects_max_messages_and_keeps_tail()
        {
            var sessionId = store.Create(Book);
            var messages = Enumerable.Range(0, 10)
                .Select(i => ChatMessage.User($"问题{i}"))
                .ToList();

            store.Save(Book, sessionId, messages, "tail");

            var loaded = store.Load(Book, sessionId, maxMessages: 4);
            Assert.Equal(4, loaded.Count);
            Assert.Equal("问题6", loaded[0].Content);
            Assert.Equal("问题9", loaded[3].Content);
        }

        [Fact]
        public void Multiple_sessions_are_isolated()
        {
            var first = store.Create(Book);
            var second = store.Create(Book);
            store.Save(Book, first, new[] { ChatMessage.User("会话A问题") }, "A");
            store.Save(Book, second, new[] { ChatMessage.User("会话B问题") }, "B");

            Assert.Equal("会话A问题", store.Load(Book, first)[0].Content);
            Assert.Equal("会话B问题", store.Load(Book, second)[0].Content);
            Assert.Equal(2, store.List(Book).Count);
        }

        [Fact]
        public void Delete_removes_session()
        {
            var sessionId = store.Create(Book);
            store.Save(Book, sessionId, new[] { ChatMessage.User("你好") }, "待删除");

            Assert.True(store.Delete(Book, sessionId));
            Assert.Empty(store.Load(Book, sessionId));
            Assert.Empty(store.List(Book));
        }

        [Fact]
        public void List_orders_by_most_recent_and_infers_title_from_first_user_message()
        {
            var old = store.Create(Book);
            store.Save(Book, old, new[] { ChatMessage.User("这是很早的问题") }, "旧会话");
            var fresh = store.Create(Book);
            store.Save(Book, fresh, new[] { ChatMessage.User("今天的新需求，需要做一张汇总表") }, null);

            var sessions = store.List(Book);
            Assert.Equal(2, sessions.Count);
            Assert.Equal(fresh, sessions[0].Id);
            Assert.Equal(old, sessions[1].Id);
            Assert.StartsWith("今天的新需求", sessions[0].Title);
            Assert.Equal(0, sessions[1].TurnCount);
        }

        [Fact]
        public void LastUpdated_reflects_save()
        {
            var sessionId = store.Create(Book);
            Assert.NotNull(store.LastUpdated(Book, sessionId));

            store.Save(Book, sessionId, new[] { ChatMessage.User("你好") }, null);
            Assert.NotNull(store.LastUpdated(Book, sessionId));
        }

        [Fact]
        public void Summary_round_trips_and_persists_across_instances()
        {
            var sessionId = store.Create(Book);
            store.SaveSummary(Book, sessionId, "此前对话摘要：清洗了 A2:D20 并生成汇总公式。");

            Assert.Equal("此前对话摘要：清洗了 A2:D20 并生成汇总公式。", store.LoadSummary(Book, sessionId));

            var reloaded = new SessionStore(directory).LoadSummary(Book, sessionId);
            Assert.Equal("此前对话摘要：清洗了 A2:D20 并生成汇总公式。", reloaded);
        }

        [Fact]
        public void LoadSummary_returns_null_when_absent()
        {
            var sessionId = store.Create(Book);
            Assert.Null(store.LoadSummary(Book, sessionId));
        }

        [Fact]
        public void ClearSummary_resets_to_empty()
        {
            var sessionId = store.Create(Book);
            store.SaveSummary(Book, sessionId, "某摘要");
            store.ClearSummary(Book, sessionId);

            Assert.Equal(string.Empty, store.LoadSummary(Book, sessionId));
        }

        [Fact]
        public void Save_does_not_erase_existing_summary()
        {
            var sessionId = store.Create(Book);
            store.SaveSummary(Book, sessionId, "保留的摘要");
            store.Save(Book, sessionId, new[] { ChatMessage.User("压缩后第一问") }, null);

            Assert.Equal("保留的摘要", store.LoadSummary(Book, sessionId));
            Assert.Equal("压缩后第一问", store.Load(Book, sessionId)[0].Content);
        }

        [Fact]
        public void Rename_updates_title_and_persists()
        {
            var sessionId = store.Create(Book);
            store.Save(Book, sessionId, new[] { ChatMessage.User("旧标题内容") }, "旧标题");

            Assert.True(store.Rename(Book, sessionId, "新标题"));
            Assert.Equal("新标题", store.List(Book)[0].Title);
            Assert.Equal("新标题", new SessionStore(directory).Title(Book, sessionId));
        }

        [Fact]
        public void Rename_rejects_empty_title()
        {
            var sessionId = store.Create(Book);
            Assert.False(store.Rename(Book, sessionId, "   "));
            Assert.False(store.Rename(Book, sessionId, null));
        }

        [Fact]
        public void Archive_flag_round_trips()
        {
            var sessionId = store.Create(Book);
            Assert.False(store.List(Book)[0].Archived);

            Assert.True(store.SetArchived(Book, sessionId, true));
            Assert.True(store.List(Book)[0].Archived);

            Assert.True(store.SetArchived(Book, sessionId, false));
            Assert.False(store.List(Book)[0].Archived);
        }

        [Fact]
        public void Search_matches_title_and_content()
        {
            var sessionA = store.Create(Book);
            store.Save(Book, sessionA, new[] { ChatMessage.User("把销售数据做成汇总表") }, "销售汇总");
            var sessionB = store.Create(Book);
            store.Save(Book, sessionB, new[] { ChatMessage.User("分析库存周转率") }, "库存分析");

            var titleHits = store.Search(Book, "销售");
            Assert.Single(titleHits);
            Assert.Equal(sessionA, titleHits[0].SessionId);
            Assert.True(titleHits[0].TitleMatched);

            var contentHits = store.Search(Book, "周转率");
            Assert.Single(contentHits);
            Assert.Equal(sessionB, contentHits[0].SessionId);
            Assert.False(contentHits[0].TitleMatched);
            Assert.Contains("周转率", contentHits[0].Snippet);
        }

        [Fact]
        public void Search_returns_empty_for_missing_keyword_or_no_match()
        {
            var sessionId = store.Create(Book);
            store.Save(Book, sessionId, new[] { ChatMessage.User("普通内容") }, "标题");

            Assert.Empty(store.Search(Book, "不存在的关键词"));
            Assert.Empty(store.Search(Book, "   "));
            Assert.Empty(store.Search(Book, null));
        }

        [Fact]
        public void Search_orders_by_recency()
        {
            var old = store.Create(Book);
            store.Save(Book, old, new[] { ChatMessage.User("共享关键词 甲") }, "旧");
            var fresh = store.Create(Book);
            store.Save(Book, fresh, new[] { ChatMessage.User("共享关键词 乙") }, "新");

            var hits = store.Search(Book, "共享关键词");
            Assert.Equal(2, hits.Count);
            Assert.Equal(fresh, hits[0].SessionId);
            Assert.Equal(old, hits[1].SessionId);
        }

        public void Dispose()
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
    }
}