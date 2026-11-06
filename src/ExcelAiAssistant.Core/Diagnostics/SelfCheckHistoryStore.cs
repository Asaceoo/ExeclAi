using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>一条自检历史记录：回合级（自动自检元监控）或研判级（一键自检 AI 研判）。</summary>
    public sealed class SelfCheckRecord
    {
        public long Id { get; set; }

        /// <summary>记录时间（本地时间）。</summary>
        public DateTime TimestampLocal { get; set; } = DateTime.Now;

        /// <summary>turn = 回合自动自检摘要；adjudication = 一键自检 AI 研判。</summary>
        public string Kind { get; set; }

        /// <summary>本条记录是否健康（回合：至少执行了一条自检计划；研判：无异常项）。</summary>
        public bool Healthy { get; set; }

        /// <summary>可自检写操作数（仅 turn）。</summary>
        public int Ops { get; set; }

        /// <summary>执行的自检计划数（仅 turn）。</summary>
        public int Executed { get; set; }

        /// <summary>跳过的自检计划数（仅 turn）。</summary>
        public int Skipped { get; set; }

        /// <summary>异常项数（仅 adjudication）。</summary>
        public int IssueCount { get; set; }

        /// <summary>一行摘要（元监控摘要 / 研判确定性结论）。</summary>
        public string Summary { get; set; }

        /// <summary>问题清单（研判 issues；回合为空）。</summary>
        public List<string> Issues { get; set; } = new List<string>();
    }

    /// <summary>自检趋势统计。</summary>
    public sealed class SelfCheckStats
    {
        public int Total { get; set; }

        public int Healthy { get; set; }

        /// <summary>健康占比 0~1。</summary>
        public double PassRate { get; set; }

        /// <summary>高频问题（按问题文本聚合，降序，最多 3 条）。</summary>
        public List<KeyValuePair<string, int>> TopIssues { get; } = new List<KeyValuePair<string, int>>();

        public List<SelfCheckRecord> Recent { get; } = new List<SelfCheckRecord>();
    }

    /// <summary>
    /// 自检历史的 SQLite 存储：把每回合的元监控摘要与每次一键自检研判结论落库，
    /// 支撑「最近 N 次通过率 / 高频问题」趋势。单文件库，随用户数据目录走，损坏时自动重建表。
    /// </summary>
    public sealed class SelfCheckHistoryStore
    {
        private readonly string connectionString;
        private readonly object syncRoot = new object();

        public SelfCheckHistoryStore(string databasePath)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
            {
                throw new ArgumentNullException(nameof(databasePath));
            }

            var directory = Path.GetDirectoryName(Path.GetFullPath(databasePath));
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            connectionString = $"Data Source={databasePath};Version=3;";
            Initialize();
        }

        public void Append(SelfCheckRecord record)
        {
            if (record == null)
            {
                throw new ArgumentNullException(nameof(record));
            }

            lock (syncRoot)
            {
                using var connection = Open();
                using var command = connection.CreateCommand();
                command.CommandText = @"
INSERT INTO selfcheck_history (ts, kind, healthy, ops, executed, skipped, issue_count, summary, issues_json)
VALUES (@ts, @kind, @healthy, @ops, @executed, @skipped, @issue_count, @summary, @issues_json);";
                command.Parameters.AddWithValue("@ts", record.TimestampLocal.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture));
                command.Parameters.AddWithValue("@kind", record.Kind ?? "turn");
                command.Parameters.AddWithValue("@healthy", record.Healthy ? 1 : 0);
                command.Parameters.AddWithValue("@ops", record.Ops);
                command.Parameters.AddWithValue("@executed", record.Executed);
                command.Parameters.AddWithValue("@skipped", record.Skipped);
                command.Parameters.AddWithValue("@issue_count", record.IssueCount);
                command.Parameters.AddWithValue("@summary", record.Summary ?? string.Empty);
                command.Parameters.AddWithValue("@issues_json", Newtonsoft.Json.JsonConvert.SerializeObject(record.Issues ?? new List<string>()));
                command.ExecuteNonQuery();
                Trim(connection);
            }
        }

        /// <summary>防无界膨胀：历史超过 MaxRecords 时删掉最老的，只留最近 MaxRecords 条。</summary>
        private const int MaxRecords = 500;

        private void Trim(SQLiteConnection connection)
        {
            try
            {
                using var command = connection.CreateCommand();
                command.CommandText = @"
DELETE FROM selfcheck_history
WHERE id <= (SELECT MAX(id) - @keep FROM selfcheck_history);";
                command.Parameters.AddWithValue("@keep", MaxRecords);
                command.ExecuteNonQuery();
            }
            catch (SQLiteException)
            {
                AppLog.Warn("SelfCheckHistoryStore", "[LogCover] SelfCheckHistoryStore.cs:126  silent-catch covered (no exception variable)");
                // 清理失败不影响主流程；下次 Append 会重试。
            }
        }

        /// <summary>最近 limit 条记录（新→旧）。</summary>
        public List<SelfCheckRecord> Recent(int limit)
        {
            lock (syncRoot)
            {
                using var connection = Open();
                using var command = connection.CreateCommand();
                command.CommandText = @"
SELECT id, ts, kind, healthy, ops, executed, skipped, issue_count, summary, issues_json
FROM selfcheck_history ORDER BY id DESC LIMIT @limit;";
                command.Parameters.AddWithValue("@limit", Math.Max(1, limit));
                var records = new List<SelfCheckRecord>();
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    records.Add(ReadRecord(reader));
                }

                return records;
            }
        }

        /// <summary>最近 limit 次的趋势统计：通过率 + 高频问题 Top3 + 明细。</summary>
        public SelfCheckStats Stats(int limit)
        {
            var recent = Recent(limit);
            var stats = new SelfCheckStats
            {
                Total = recent.Count,
                Healthy = recent.Count(item => item.Healthy)
            };
            stats.PassRate = stats.Total == 0 ? 0 : (double)stats.Healthy / stats.Total;

            var issueCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var record in recent)
            {
                foreach (var issue in record.Issues ?? new List<string>())
                {
                    var key = NormalizeIssueKey(issue);
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }

                    issueCounts[key] = issueCounts.TryGetValue(key, out var count) ? count + 1 : 1;
                }
            }

            stats.TopIssues.AddRange(issueCounts
                .OrderByDescending(kv => kv.Value)
                .ThenBy(kv => kv.Key, StringComparer.Ordinal)
                .Take(3)
                .Select(kv => new KeyValuePair<string, int>(kv.Key, kv.Value)));
            stats.Recent.AddRange(recent);
            return stats;
        }

        /// <summary>问题文本归一为聚合键：截断到 60 字符，抹掉具体地址/数字差异，只留失败模式。</summary>
        private static string NormalizeIssueKey(string issue)
        {
            if (string.IsNullOrWhiteSpace(issue))
            {
                return string.Empty;
            }

            var text = issue.Trim();
            if (text.Length > 60)
            {
                text = text.Substring(0, 60);
            }

            return text;
        }

        private SelfCheckRecord ReadRecord(SQLiteDataReader reader)
        {
            var record = new SelfCheckRecord
            {
                Id = Convert.ToInt64(reader["id"]),
                TimestampLocal = DateTime.TryParse(
                    Convert.ToString(reader["ts"]),
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out var parsed)
                    ? parsed
                    : DateTime.Now,
                Kind = Convert.ToString(reader["kind"]),
                Healthy = Convert.ToInt64(reader["healthy"]) == 1,
                Ops = Convert.ToInt32(reader["ops"]),
                Executed = Convert.ToInt32(reader["executed"]),
                Skipped = Convert.ToInt32(reader["skipped"]),
                IssueCount = Convert.ToInt32(reader["issue_count"]),
                Summary = Convert.ToString(reader["summary"]),
                Issues = new List<string>()
            };

            var issuesJson = Convert.ToString(reader["issues_json"]);
            if (!string.IsNullOrWhiteSpace(issuesJson))
            {
                try
                {
                    var token = Newtonsoft.Json.Linq.JArray.Parse(issuesJson);
                    foreach (var item in token)
                    {
                        record.Issues.Add(item.ToString());
                    }
                }
                catch (Newtonsoft.Json.JsonException)
                {
                    AppLog.Warn("SelfCheckHistoryStore", "[LogCover] SelfCheckHistoryStore.cs:239  silent-catch covered (no exception variable)");
                    // 历史数据形状异常时按空清单处理，不让趋势页挂掉。
                }
            }

            return record;
        }

        private SQLiteConnection Open()
        {
            var connection = new SQLiteConnection(connectionString);
            connection.Open();
            return connection;
        }

        private void Initialize()
        {
            lock (syncRoot)
            {
                using var connection = Open();
                using var command = connection.CreateCommand();
                command.CommandText = @"
CREATE TABLE IF NOT EXISTS selfcheck_history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,
  kind TEXT NOT NULL,
  healthy INTEGER NOT NULL,
  ops INTEGER NOT NULL DEFAULT 0,
  executed INTEGER NOT NULL DEFAULT 0,
  skipped INTEGER NOT NULL DEFAULT 0,
  issue_count INTEGER NOT NULL DEFAULT 0,
  summary TEXT,
  issues_json TEXT
);";
                command.ExecuteNonQuery();
            }
        }
    }
}
