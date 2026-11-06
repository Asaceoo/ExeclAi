using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.IO;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Security
{
    /// <summary>
    /// 受控执行配置的独立 JSON 持久化。默认写入
    /// %LocalAppData%\XaiAssistant\controlled-execution.json；
    /// 测试或自定义宿主可传入目录覆盖。
    /// </summary>
    public sealed class ControlledExecutionStore
    {
        private readonly string filePath;

        public ControlledExecutionStore()
            : this(DefaultPath())
        {
        }

        public ControlledExecutionStore(string filePath)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        public static string DefaultPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(baseDir, "XaiAssistant", "controlled-execution.json");
        }

        public string FilePath => filePath;

        /// <summary>文件不存在时返回默认配置（Enabled=false，各白名单为空）。</summary>
        public ControlledExecutionConfig Load()
        {
            if (!File.Exists(filePath))
            {
                return ControlledExecutionConfig.CreateDefault();
            }

            try
            {
                var json = File.ReadAllText(filePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return ControlledExecutionConfig.CreateDefault();
                }

                return JsonConvert.DeserializeObject<ControlledExecutionConfig>(json)
                       ?? ControlledExecutionConfig.CreateDefault();
            }
            catch (Exception)
            {
                // 配置损坏时以默认值兜底，避免把错误传染给宿主。
                return ControlledExecutionConfig.CreateDefault();
        }
        }

        public void Save(ControlledExecutionConfig config)
        {
            if (config == null)
            {
                throw new ArgumentNullException(nameof(config));
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            // 原子写：中途崩溃不损坏策略 JSON；损坏会触发 Load 回退默认(Enabled=false)，等于静默放开受控执行。
            var tmp = filePath + ".tmp." + Guid.NewGuid().ToString("N");
            File.WriteAllText(tmp, JsonConvert.SerializeObject(config, Formatting.Indented));
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
    }
}