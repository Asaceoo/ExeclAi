using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.IO;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Configuration
{
    /// <summary>
    /// 写入偏好独立 JSON 持久化。默认写入
    /// %LocalAppData%\XaiAssistant\write-preference.json。
    /// ValueMode 取值：auto（默认，跟随模型判断）、formula（公式优先）、value（直接写值）。
    /// </summary>
    public sealed class WritePreferenceStore
    {
        private readonly string filePath;

        public WritePreferenceStore()
            : this(DefaultPath())
        {
        }

        public WritePreferenceStore(string filePath)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
        }

        public static string DefaultPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(baseDir, "XaiAssistant", "write-preference.json");
        }

        public string FilePath => filePath;

        public WritePreference Load()
        {
            if (!File.Exists(filePath))
            {
                return WritePreference.Auto();
            }

            try
            {
                var json = File.ReadAllText(filePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return WritePreference.Auto();
                }

                var pref = JsonConvert.DeserializeObject<WritePreference>(json);
                return pref != null && pref.IsValid ? pref : WritePreference.Auto();
            }
            catch (Exception)
            {
                return WritePreference.Auto();
        }
        }

        public void Save(WritePreference preference)
        {
            if (preference == null)
            {
                throw new ArgumentNullException(nameof(preference));
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            File.WriteAllText(filePath, JsonConvert.SerializeObject(preference, Formatting.Indented));
        }
    }

    public sealed class WritePreference
    {
        public string ValueMode { get; set; } = "auto";

        public static WritePreference Auto() => new WritePreference { ValueMode = "auto" };

        public static WritePreference Formula() => new WritePreference { ValueMode = "formula" };

        public static WritePreference Value() => new WritePreference { ValueMode = "value" };

        public bool IsValid =>
            string.Equals(ValueMode, "auto", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ValueMode, "formula", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ValueMode, "value", StringComparison.OrdinalIgnoreCase);
    }
}