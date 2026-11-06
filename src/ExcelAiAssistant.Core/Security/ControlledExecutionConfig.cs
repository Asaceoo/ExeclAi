using System;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Security
{
    /// <summary>
    /// 受控执行（脚本 / CLI / EXE）的持久化配置。
    /// 采用独立 JSON 文件存储，字段名保持小写下划线，便于前端与文件直观对应。
    /// </summary>
    public sealed class ControlledExecutionConfig
    {
        /// <summary>总开关：关闭时三类受控能力一律拒绝。</summary>
        [JsonProperty("enabled")]
        public bool Enabled { get; set; }

        /// <summary>CLI 命令名（可含前缀，例如 dir、ping）白名单；空表示一律不通过。</summary>
        [JsonProperty("cli_allowlist")]
        public string[] CliAllowlist { get; set; } = new string[0];

        /// <summary>CLI 命令黑名单：命中直接拒绝。</summary>
        [JsonProperty("cli_denylist")]
        public string[] CliDenylist { get; set; } = new string[0];

        /// <summary>可执行文件完整路径白名单；空表示一律不通过。</summary>
        [JsonProperty("exe_allowlist")]
        public string[] ExeAllowlist { get; set; } = new string[0];

        /// <summary>可执行文件完整路径黑名单：命中直接拒绝。</summary>
        [JsonProperty("exe_denylist")]
        public string[] ExeDenylist { get; set; } = new string[0];

        /// <summary>脚本文件路径白名单（用于以 script_path 方式运行脚本）；空表示不允许以脚本文件方式执行。</summary>
        [JsonProperty("script_allowlist")]
        public string[] ScriptAllowlist { get; set; } = new string[0];

        /// <summary>脚本文件路径黑名单：命中直接拒绝。</summary>
        [JsonProperty("script_denylist")]
        public string[] ScriptDenylist { get; set; } = new string[0];

        /// <summary>进程超时（秒）。</summary>
        [JsonProperty("timeout_seconds")]
        public int TimeoutSeconds { get; set; } = 30;

        /// <summary>输出截断上限（KB）。</summary>
        [JsonProperty("output_limit_kb")]
        public int OutputLimitKb { get; set; } = 64;

        /// <summary>解释器命令名（python 或 python3 等）。</summary>
        [JsonProperty("python_path")]
        public string PythonPath { get; set; } = "python";

        public static ControlledExecutionConfig CreateDefault()
        {
            return new ControlledExecutionConfig();
        }
    }
}