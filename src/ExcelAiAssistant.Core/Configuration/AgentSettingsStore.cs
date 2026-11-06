using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.IO;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Configuration
{
    /// <summary>
    /// 助手运行时设置独立 JSON 持久化：
    /// %LocalAppData%\XaiAssistant\agent-settings.json。
    /// 目前仅包含"单轮工具调用上限"，后续可扩展更多运行时偏好。
    /// </summary>
    public sealed class AgentSettingsStore
    {
        private const int DefaultMaxToolCalls = 15;

        private readonly string filePath;

        public AgentSettingsStore()
            : this(DefaultPath())
        {
        }

        public AgentSettingsStore(string filePath)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        public static string DefaultPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(baseDir, "XaiAssistant", "agent-settings.json");
        }

        public string FilePath => filePath;

        public AgentSettings Load()
        {
            if (!File.Exists(filePath))
            {
                return AgentSettings.Default();
            }

            try
            {
                var json = File.ReadAllText(filePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return AgentSettings.Default();
                }

                var settings = JsonConvert.DeserializeObject<AgentSettings>(json);
                return settings != null && settings.IsValid ? settings : AgentSettings.Default();
            }
            catch (Exception)
            {
                return AgentSettings.Default();
        }
        }

        public void Save(AgentSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            File.WriteAllText(filePath, JsonConvert.SerializeObject(settings, Formatting.Indented));
        }
    }

    public sealed class AgentSettings
    {
        public int MaxToolCallsPerTurn { get; set; } = 10000;

        /// <summary>联网搜索总开关（v3.8.1.83）：关=不注入服务端搜索参数、web_search 工具明确拒绝执行。</summary>
        public bool WebSearchEnabled { get; set; } = true;

        public static AgentSettings Default() => new AgentSettings { MaxToolCallsPerTurn = 10000, WebSearchEnabled = true };

        public bool IsValid => MaxToolCallsPerTurn >= 1 && MaxToolCallsPerTurn <= 10000;
    }
}