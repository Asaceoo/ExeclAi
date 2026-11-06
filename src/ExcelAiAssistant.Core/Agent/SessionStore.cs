using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 多会话持久化存储。每个工作簿（workbookKey）下可存在多个独立会话，
    /// 每个会话保存自身多轮消息（标题 + 消息列表 + 时间戳），支持新建 / 删除 / 切换。
    /// </summary>
    public sealed class SessionStore
    {
        private readonly string rootDirectory;

        public SessionStore(string directory)
        {
            rootDirectory = directory ?? throw new ArgumentNullException(nameof(directory));
            Directory.CreateDirectory(rootDirectory);
        }

        public IReadOnlyList<SessionInfo> List(string workbookKey)
        {
            var bookDir = BookDirectory(workbookKey);
            if (!Directory.Exists(bookDir))
            {
                return Array.Empty<SessionInfo>();
            }

            return Directory
                .EnumerateFiles(bookDir, "*.json")
                .Select(ReadRecord)
                .Where(record => record != null)
                .OrderByDescending(record => record.UpdatedAt)
                .Select(record => new SessionInfo
                {
                    Id = record.SessionId,
                    Title = GetTitle(record),
                    TurnCount = (record.Messages?.Count ?? 0) / 2,
                    UpdatedAt = record.UpdatedAt,
                    Archived = record.Archived
                })
                .ToList();
        }

        public string Create(string workbookKey, string title = null)
        {
            var sessionId = Guid.NewGuid().ToString("N");
            var record = new SessionRecord
            {
                SessionId = sessionId,
                WorkbookKey = NormalizeKey(workbookKey),
                Title = title ?? "新会话",
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow,
                Messages = new List<PersistedMessage>()
            };
            WriteRecord(workbookKey, record);
            return sessionId;
        }

        public IReadOnlyList<ChatMessage> Load(string workbookKey, string sessionId, int maxMessages = 24)
        {
            var record = ReadRecord(PathFor(workbookKey, sessionId));
            if (record?.Messages == null)
            {
                return Array.Empty<ChatMessage>();
            }

            if (record.Messages.Count > maxMessages)
            {
                return record.Messages
                    .Skip(record.Messages.Count - maxMessages)
                    .Select(PersistedMessage.ToChatMessage)
                    .ToList();
            }

            return record.Messages.Select(PersistedMessage.ToChatMessage).ToList();
        }

        public void Save(
            string workbookKey,
            string sessionId,
            IReadOnlyList<ChatMessage> messages,
            string title = null)
        {
            var record = ReadRecord(PathFor(workbookKey, sessionId)) ?? new SessionRecord
            {
                SessionId = sessionId,
                WorkbookKey = NormalizeKey(workbookKey)
            };
            record.SessionId = sessionId;
            record.WorkbookKey = NormalizeKey(workbookKey);
            if (!string.IsNullOrWhiteSpace(title))
            {
                record.Title = title;
            }

            record.UpdatedAt = DateTime.UtcNow;
            record.Messages = messages.Select(PersistedMessage.FromChatMessage).ToList();
            WriteRecord(workbookKey, record);
        }

        public bool Delete(string workbookKey, string sessionId)
        {
            var path = PathFor(workbookKey, sessionId);
            if (!File.Exists(path))
            {
                return false;
            }

            File.Delete(path);
            return true;
        }

        public bool Rename(string workbookKey, string sessionId, string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return false;
            }

            var record = ReadRecord(PathFor(workbookKey, sessionId));
            if (record == null)
            {
                return false;
            }

            record.Title = title.Trim();
            record.UpdatedAt = DateTime.UtcNow;
            WriteRecord(workbookKey, record);
            return true;
        }

        public bool SetArchived(string workbookKey, string sessionId, bool archived)
        {
            var record = ReadRecord(PathFor(workbookKey, sessionId));
            if (record == null)
            {
                return false;
            }

            record.Archived = archived;
            record.UpdatedAt = DateTime.UtcNow;
            WriteRecord(workbookKey, record);
            return true;
        }

        /// <summary>按标题与消息内容关键词搜索会话，返回命中会话及其匹配片段。</summary>
        public IReadOnlyList<SessionSearchHit> Search(string workbookKey, string keyword)
        {
            var bookDir = BookDirectory(workbookKey);
            if (!Directory.Exists(bookDir) || string.IsNullOrWhiteSpace(keyword))
            {
                return Array.Empty<SessionSearchHit>();
            }

            var needle = keyword.Trim();
            if (needle.Length == 0)
            {
                return Array.Empty<SessionSearchHit>();
            }

            var hits = new List<SessionSearchHit>();
            foreach (var file in Directory.EnumerateFiles(bookDir, "*.json"))
            {
                var record = ReadRecord(file);
                if (record?.Messages == null)
                {
                    continue;
                }

                var titleHit = !string.IsNullOrWhiteSpace(record.Title)
                    && record.Title.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
                string contentHit = null;
                foreach (var message in record.Messages)
                {
                    var content = message.Content ?? string.Empty;
                    if (content.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        contentHit = Snippet(content, needle);
                        break;
                    }
                }

                if (titleHit || contentHit != null)
                {
                    hits.Add(new SessionSearchHit
                    {
                        SessionId = record.SessionId,
                        Title = GetTitle(record),
                        UpdatedAt = record.UpdatedAt,
                        Archived = record.Archived,
                        TitleMatched = titleHit,
                        Snippet = contentHit
                    });
                }
            }

            return hits.OrderByDescending(hit => hit.UpdatedAt).ToList();
        }

        private static string Snippet(string text, string needle)
        {
            var flattened = text.Replace("\r", " ").Replace("\n", " ");
            var index = flattened.IndexOf(needle, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                return Truncate(flattened, 80);
            }

            var start = Math.Max(0, index - 24);
            var length = Math.Min(flattened.Length - start, 72);
            var prefix = start > 0 ? "…" : string.Empty;
            var suffix = start + length < flattened.Length ? "…" : string.Empty;
            return prefix + flattened.Substring(start, length) + suffix;
        }

        private static string Truncate(string text, int max)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return text.Length > max ? text.Substring(0, max) + "…" : text;
        }

        public DateTime? LastUpdated(string workbookKey, string sessionId)
        {
            return ReadRecord(PathFor(workbookKey, sessionId))?.UpdatedAt;
        }

        public string Title(string workbookKey, string sessionId)
        {
            return GetTitle(ReadRecord(PathFor(workbookKey, sessionId)));
        }

        /// <summary>读取该会话的压缩摘要；没有则返回 null。</summary>
        public string LoadSummary(string workbookKey, string sessionId)
        {
            return ReadRecord(PathFor(workbookKey, sessionId))?.Summary;
        }

        /// <summary>保存该会话的压缩摘要（不清空消息）。</summary>
        public void SaveSummary(string workbookKey, string sessionId, string summary)
        {
            if (string.IsNullOrWhiteSpace(sessionId) || string.IsNullOrWhiteSpace(workbookKey))
            {
                return;
            }

            var record = ReadRecord(PathFor(workbookKey, sessionId)) ?? new SessionRecord
            {
                SessionId = sessionId,
                WorkbookKey = NormalizeKey(workbookKey)
            };
            record.SessionId = sessionId;
            record.WorkbookKey = NormalizeKey(workbookKey);
            record.UpdatedAt = DateTime.UtcNow;
            record.Summary = summary ?? string.Empty;
            WriteRecord(workbookKey, record);
        }

        /// <summary>清空该会话的压缩摘要（例如新建/切换会话时）。</summary>
        public void ClearSummary(string workbookKey, string sessionId)
        {
            SaveSummary(workbookKey, sessionId, string.Empty);
        }

        private static string GetTitle(SessionRecord record)
        {
            if (record == null)
            {
                return "新会话";
            }

            if (!string.IsNullOrWhiteSpace(record.Title) && record.Title != "新会话")
            {
                return record.Title;
            }

            var firstUser = record.Messages?.FirstOrDefault(message =>
                string.Equals(message.Role, "user", StringComparison.OrdinalIgnoreCase));
            if (firstUser != null)
            {
                return Truncate(HumanText(firstUser.Content));
            }

            return "新会话";
        }

        // 用户消息 content 是给模型打的 {request, excel_context, ...} 负载，取标题时还原为 request 纯文本，
        // 否则标题会显示成整段内部 JSON（乱码）。
        private static string HumanText(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return content;
            }

            if (content.TrimStart().StartsWith("{"))
            {
                try
                {
                    using (var reader = new JsonTextReader(new System.IO.StringReader(content)))
                    {
                        var obj = Newtonsoft.Json.Linq.JObject.Load(reader);
                        if (obj["request"] != null && obj["request"].Type == Newtonsoft.Json.Linq.JTokenType.String)
                        {
                            var text = obj.Value<string>("request");
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                return text;
                            }
                        }
                    }
                }
                catch
                {
                    AppLog.Warn("SessionStore", "[LogCover] SessionStore.cs:321  silent-catch covered (no exception variable)");
                    // 不是可解析的 bundle，退化为原文
                }
            }

            return content;
        }

        private static string Truncate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "新会话";
            }

            var flattened = text.Replace("\r", " ").Replace("\n", " ");
            return flattened.Length > 14 ? flattened.Substring(0, 14) + "…" : flattened;
        }

        private string BookDirectory(string workbookKey)
        {
            return Path.Combine(rootDirectory, Safe(NormalizeKey(workbookKey)));
        }

        private string PathFor(string workbookKey, string sessionId)
        {
            return Path.Combine(BookDirectory(workbookKey), $"{sessionId}.json");
        }

        private void WriteRecord(string workbookKey, SessionRecord record)
        {
            if (record == null || string.IsNullOrEmpty(record.SessionId))
            {
                return;
            }

            Directory.CreateDirectory(BookDirectory(workbookKey));
            File.WriteAllText(
                PathFor(workbookKey, record.SessionId),
                JsonConvert.SerializeObject(record, Formatting.None));
        }

        private SessionRecord ReadRecord(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                return JsonConvert.DeserializeObject<SessionRecord>(File.ReadAllText(path));
            }
            catch (JsonException)
            {
                AppLog.Warn("SessionStore", "[LogCover] SessionStore.cs:375  silent-catch covered (no exception variable)");
                return null;
            }
        }

        private static string NormalizeKey(string key)
        {
            return string.IsNullOrWhiteSpace(key) ? "default" : key.Trim();
        }

        private static string Safe(string key)
        {
            return key.Replace(':', '_').Replace('/', '_').Replace('\\', '_');
        }

        public sealed class SessionInfo
        {
            public string Id { get; set; }

            public string Title { get; set; }

            public int TurnCount { get; set; }

            public DateTime UpdatedAt { get; set; }

            public bool Archived { get; set; }
        }

        public sealed class SessionSearchHit
        {
            public string SessionId { get; set; }

            public string Title { get; set; }

            public DateTime UpdatedAt { get; set; }

            public bool Archived { get; set; }

            public bool TitleMatched { get; set; }

            public string Snippet { get; set; }
        }

        private sealed class SessionRecord
        {
            [JsonProperty("session_id")]
            public string SessionId { get; set; }

            [JsonProperty("workbook_key")]
            public string WorkbookKey { get; set; }

            [JsonProperty("title")]
            public string Title { get; set; }

            [JsonProperty("created_at")]
            public DateTime CreatedAt { get; set; }

            [JsonProperty("updated_at")]
            public DateTime UpdatedAt { get; set; }

            [JsonProperty("summary")]
            public string Summary { get; set; }

            [JsonProperty("archived")]
            public bool Archived { get; set; }

            [JsonProperty("messages")]
            public List<PersistedMessage> Messages { get; set; }
        }

        private sealed class PersistedMessage
        {
            [JsonProperty("role")]
            public string Role { get; set; }

            [JsonProperty("content")]
            public string Content { get; set; }

            [JsonProperty("tool_events")]
            public List<ToolEventView> ToolEvents { get; set; }

            [JsonProperty("reasoning")]
            public string Reasoning { get; set; }

            [JsonProperty("steps")]
            public List<AgentStepView> Steps { get; set; }

            [JsonProperty("turn_label")]
            public string TurnLabel { get; set; }

            [JsonProperty("auto_sinked_vba")]
            public bool AutoSinkedVba { get; set; }

            public static PersistedMessage FromChatMessage(ChatMessage message)
            {
                return new PersistedMessage
                {
                    Role = message.Role,
                    Content = message.Content ?? string.Empty,
                    ToolEvents = message.ToolEvents?.ToList(),
                    Reasoning = message.Reasoning,
                    Steps = message.Steps?.ToList(),
                    TurnLabel = message.TurnLabel,
                    AutoSinkedVba = message.AutoSinkedVba
                };
            }

            public static ChatMessage ToChatMessage(PersistedMessage message)
            {
                return new ChatMessage
                {
                    Role = message.Role,
                    Content = message.Content ?? string.Empty,
                    ToolEvents = message.ToolEvents,
                    Reasoning = message.Reasoning,
                    Steps = message.Steps,
                    TurnLabel = message.TurnLabel,
                    AutoSinkedVba = message.AutoSinkedVba
                };
            }
        }
    }
}