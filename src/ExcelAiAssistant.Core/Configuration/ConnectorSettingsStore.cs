using System;
using System.IO;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Configuration
{
    /// <summary>
    /// MCP / CLI 外部接口面的开关设置（与 agent-settings.json 同级独立文件，避免相互覆盖）。
    /// 语义：关闭后对应入口在启动时直接拒绝（exit 6），不再创建 Excel 实例——
    /// 面板与命令行读到的是同一份真值，不存在“面板关了但 exe 还在跑”的分裂状态。
    /// </summary>
    public sealed class ConnectorSettingsStore
    {
        private readonly string filePath;

        public ConnectorSettingsStore()
            : this(DefaultPath())
        {
        }

        public ConnectorSettingsStore(string filePath)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        public static string DefaultPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(baseDir, "XaiAssistant", "connectors.json");
        }

        public string FilePath => filePath;

        /// <summary>读取设置；文件缺失/损坏一律回落到默认（两者开启）——外部接口面默认可用。
        /// 审查 A8/S3：新增 state 出参（ok/missing/invalid/denied），面板可如实提示"配置损坏已按默认开启"。</summary>
        public ConnectorSettings Load()
        {
            return Load(out _);
        }

        public ConnectorSettings Load(out string state)
        {
            state = "ok";
            if (!File.Exists(filePath))
            {
                state = "missing";
                return ConnectorSettings.Default();
            }

            try
            {
                var json = File.ReadAllText(filePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    state = "invalid";
                    return ConnectorSettings.Default();
                }

                var settings = JsonConvert.DeserializeObject<ConnectorSettings>(json);
                if (settings == null)
                {
                    state = "invalid";
                    return ConnectorSettings.Default();
                }

                return settings;
            }
            catch (UnauthorizedAccessException)
            {
                state = "denied";
                return ConnectorSettings.Default();
            }
            catch (Exception)
            {
                // 损坏：按默认开启处理，避免误把接口面锁死（面板仍可写入修正）
                state = "invalid";
                return ConnectorSettings.Default();
            }
        }

        public void Save(ConnectorSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // 审查 A2：临时文件必须唯一——固定 ".tmp" 会让两个 Excel 实例互相覆盖并静默丢更新。
            var temp = filePath + "." + Guid.NewGuid().ToString("N").Substring(0, 8) + ".tmp";
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            try
            {
                File.WriteAllText(temp, json, new System.Text.UTF8Encoding(false));

                // 审查 A1：File.Replace 可能因目标被短暂读取（共享冲突）失败——退避重试后降级为直接写。
                // 审查 B5：原 `catch when (attempt < 4)` 在第 5 次失败时会把异常抛穿，降级写永远不可达。
                // 改为无条件吞异常 + 条件退避，让循环自然结束，由下方降级写兜底。
                var replaced = false;
                for (var attempt = 0; attempt < 5 && !replaced; attempt++)
                {
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            File.Replace(temp, filePath, null, true);
                        }
                        else
                        {
                            File.Move(temp, filePath);
                        }

                        replaced = true;
                    }
                    catch (Exception)
                    {
                        if (attempt < 4)
                        {
                            System.Threading.Thread.Sleep(50 * (attempt + 1));
                        }
                    }
                }

                if (!replaced)
                {
                    // 最终降级：非原子，但保证开关可用；损坏时 Load 会报 invalid 并提示用户
                    File.WriteAllText(filePath, json, new System.Text.UTF8Encoding(false));
                }

                // 审查 B6：降级写后回读校验，损坏则抛出让调用方提示（不静默留下半截文件）
                var verify = JsonConvert.DeserializeObject<ConnectorSettings>(
                    File.ReadAllText(filePath));
                if (verify == null)
                {
                    throw new IOException("连接开关写入后回读校验失败：" + filePath);
                }
            }
            finally
            {
                try { if (File.Exists(temp)) { File.Delete(temp); } } catch { }
            }
        }
    }

    public sealed class ConnectorSettings
    {
        /// <summary>MCP 服务入口开关：关闭后 ExcelAiMcpServer 启动即拒绝。</summary>
        public bool McpEnabled { get; set; } = true;

        /// <summary>CLI 命令行入口开关：关闭后 ExcelAiCli 的 Excel 类命令启动即拒绝（version/help 恒可用）。</summary>
        public bool CliEnabled { get; set; } = true;

        public static ConnectorSettings Default()
        {
            return new ConnectorSettings { McpEnabled = true, CliEnabled = true };
        }
    }
}