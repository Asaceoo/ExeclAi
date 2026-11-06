using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>自检历史存储：落库、查询与趋势统计。</summary>
    public sealed class SelfCheckHistoryStoreTests : IDisposable
    {
        private readonly string dbPath = Path.Combine(
            Path.GetTempPath(), "xai-selfcheck-test-" + Guid.NewGuid().ToString("N") + ".db");

        public void Dispose()
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }
        }

        [Fact]
        public void Append_then_recent_returns_records_in_order()
        {
            var store = new SelfCheckHistoryStore(dbPath);
            store.Append(new SelfCheckRecord { Kind = "turn", Healthy = true, Ops = 2, Executed = 3, Summary = "回合A" });
            store.Append(new SelfCheckRecord { Kind = "adjudication", Healthy = false, IssueCount = 2, Summary = "研判B", Issues = new List<string> { "问题一", "问题二" } });

            var recent = store.Recent(10);
            Assert.Equal(2, recent.Count);
            Assert.Equal("研判B", recent[0].Summary); // 新→旧
            Assert.Equal("回合A", recent[1].Summary);
            Assert.False(recent[0].Healthy);
            Assert.Equal(2, recent[0].IssueCount);
        }

        [Fact]
        public void Stats_computes_pass_rate_and_top_issues()
        {
            var store = new SelfCheckHistoryStore(dbPath);
            // 80+ 字符公共前缀 + 结尾细节差异：60 字符截断后应聚合同一失败模式。
            var longIssue = "筛选输出区域有 3 行不满足筛选条件：包含具体地址、行号、列号、阈值、时间戳等大量细节信息" +
                "用于把文本撑过聚合截断线的填充内容填充内容填充内容";
            Assert.True(longIssue.Length > 60);
            for (var i = 0; i < 3; i++)
            {
                store.Append(new SelfCheckRecord { Kind = "adjudication", Healthy = false, IssueCount = 1, Summary = "异常", Issues = new List<string> { longIssue + " #细节" + i } });
            }
            store.Append(new SelfCheckRecord { Kind = "adjudication", Healthy = true, Summary = "通过" });
            store.Append(new SelfCheckRecord { Kind = "turn", Healthy = true, Ops = 2, Executed = 2, Summary = "回合" });

            var stats = store.Stats(20);
            Assert.Equal(5, stats.Total);
            Assert.Equal(2, stats.Healthy);
            Assert.Equal(0.4, stats.PassRate, 3);
            var top = Assert.Single(stats.TopIssues);
            Assert.Equal(3, top.Value);
            Assert.StartsWith("筛选输出区域", top.Key);
        }

        [Fact]
        public void Stats_on_empty_store_is_zeroed()
        {
            var store = new SelfCheckHistoryStore(dbPath);
            var stats = store.Stats(20);
            Assert.Equal(0, stats.Total);
            Assert.Equal(0, stats.PassRate);
            Assert.Empty(stats.TopIssues);
            Assert.Empty(stats.Recent);
        }

        [Fact]
        public void Corrupt_issues_json_degrades_to_empty_list()
        {
            var store = new SelfCheckHistoryStore(dbPath);
            store.Append(new SelfCheckRecord { Kind = "adjudication", Healthy = false, Summary = "坏数据" });

            // 直接改库注入非法 JSON，验证读取不炸、按空清单降级。
            using (var connection = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
            {
                connection.Open();
                using var command = connection.CreateCommand();
                command.CommandText = "UPDATE selfcheck_history SET issues_json = '{not-json' WHERE id = 1;";
                command.ExecuteNonQuery();
            }

            var recent = store.Recent(5);
            Assert.Single(recent);
            Assert.Empty(recent[0].Issues);
        }

        [Fact]
        public void Store_recreates_table_on_existing_file()
        {
            // 同一文件重复构造（进程重启场景）不得报错、历史可读。
            var store = new SelfCheckHistoryStore(dbPath);
            store.Append(new SelfCheckRecord { Kind = "turn", Healthy = true, Summary = "首条" });
            var reopened = new SelfCheckHistoryStore(dbPath);
            Assert.Single(reopened.Recent(10));
        }

        [Fact]
        public void Append_trims_history_to_cap()
        {
            var store = new SelfCheckHistoryStore(dbPath);
            for (var i = 0; i < 505; i++)
            {
                store.Append(new SelfCheckRecord { Kind = "turn", Healthy = true, Summary = "r" + i });
            }

            Assert.Equal(500, store.Recent(10000).Count);
            // 最老的被裁掉、最新的保留。
            var newest = store.Recent(1)[0];
            Assert.Equal("r504", newest.Summary);
        }
    }
}
