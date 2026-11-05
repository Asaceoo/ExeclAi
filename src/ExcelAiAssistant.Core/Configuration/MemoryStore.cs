using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Configuration
{
    /// <summary>
    /// 跨会话长期记忆。持久化用户偏好（默认风险容忍度、常用表结构、写作风格、高频指令套路）
    /// 与按工作簿组织的最近会话摘要，供系统提示注入；设置页可查看 / 编辑。
    /// </summary>
    public sealed class MemoryStore
    {
        private readonly string filePath;
        private MemoryRecord record;

        public MemoryStore(string filePath)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        public string Preferences
        {
            get => Load().Preferences ?? string.Empty;
            set
            {
                var current = Load();
                current.Preferences = value ?? string.Empty;
                Save(current);
            }
        }

        public string Notes
        {
            get => Load().Notes ?? string.Empty;
            set
            {
                var current = Load();
                current.Notes = value ?? string.Empty;
                Save(current);
            }
        }

        public IReadOnlyList<string> RecentSummaries(string workbookKey)
        {
            var current = Load();
            var map = current.RecentSummariesByBook;
            if (map != null && map.TryGetValue(workbookKey ?? string.Empty, out var list))
            {
                return list;
            }

            return new List<string>();
        }

        public void AddRecentSummary(string workbookKey, string summary)
        {
            if (string.IsNullOrWhiteSpace(summary))
            {
                return;
            }

            var current = Load();
            var map = current.RecentSummariesByBook ?? new Dictionary<string, List<string>>();
            var key = workbookKey ?? string.Empty;
            if (!map.TryGetValue(key, out var list))
            {
                list = new List<string>();
                map[key] = list;
            }

            list.Add(summary.Trim());
            if (list.Count > 20)
            {
                list.RemoveRange(0, list.Count - 20);
            }

            current.RecentSummariesByBook = map;
            Save(current);
        }

        public void ClearRecentSummaries(string workbookKey)
        {
            var current = Load();
            if (current.RecentSummariesByBook != null)
            {
                current.RecentSummariesByBook.Remove(workbookKey ?? string.Empty);
                Save(current);
            }
        }

        /// <summary>生成注入系统提示的记忆摘要文本；无内容返回空串。</summary>
        public string BuildInjectionText(string workbookKey, int recentSummaryCount = 3)
        {
            var current = Load();
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(current.Preferences))
            {
                parts.Add("用户长期偏好：" + current.Preferences.Trim());
            }

            if (!string.IsNullOrWhiteSpace(current.Notes))
            {
                parts.Add("用户备注：" + current.Notes.Trim());
            }

            var recent = RecentSummaries(workbookKey)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToList();
            if (recent.Count > recentSummaryCount)
            {
                recent = recent.Skip(recent.Count - recentSummaryCount).ToList();
            }

            if (recent.Count > 0)
            {
                parts.Add("最近会话摘要：" + string.Join(" | ", recent));
            }

            return parts.Count == 0 ? string.Empty : string.Join("\n", parts);
        }

        private MemoryRecord Load()
        {
            if (record != null)
            {
                return record;
            }

            if (!File.Exists(filePath))
            {
                record = new MemoryRecord();
                return record;
            }

            try
            {
                record = JsonConvert.DeserializeObject<MemoryRecord>(File.ReadAllText(filePath, Encoding.UTF8)) ??
                         new MemoryRecord();
            }
            catch (JsonException)
            {
                AppLog.Warn("MemoryStore", "[LogCover] MemoryStore.cs:143  silent-catch covered (no exception variable)");
                record = new MemoryRecord();
            }

            return record;
        }

        private void Save(MemoryRecord value)
        {
            record = value;
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // 原子写：先写临时文件再替换，崩溃中途不会损坏 memory.json（跨会话长期记忆整库）。
            var tmp = filePath + ".tmp." + Guid.NewGuid().ToString("N");
            File.WriteAllText(tmp, JsonConvert.SerializeObject(value, Formatting.Indented), Encoding.UTF8);
            try
            {
                if (File.Exists(filePath))
                {
                    File.Replace(tmp, filePath, null);
                }
                else
                {
                    File.Move(tmp, filePath);
                }
            }
            finally
            {
                if (File.Exists(tmp))
                {
                    File.Delete(tmp);
                }
            }
        }

        private sealed class MemoryRecord
        {
            [JsonProperty("preferences")]
            public string Preferences { get; set; }

            [JsonProperty("notes")]
            public string Notes { get; set; }

            [JsonProperty("recent_summaries_by_book")]
            public Dictionary<string, List<string>> RecentSummariesByBook { get; set; }
        }
    }
}
