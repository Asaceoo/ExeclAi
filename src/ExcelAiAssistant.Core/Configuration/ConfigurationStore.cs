using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Configuration
{
    public sealed class ConfigurationStore
    {
        private readonly string filePath;
        private readonly Func<string, byte[]> protectData;
        private readonly Func<byte[], string> unprotectData;

        public ConfigurationStore(
            string filePath,
            Func<string, byte[]> protectData,
            Func<byte[], string> unprotectData)
        {
            this.filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            this.protectData = protectData ?? throw new ArgumentNullException(nameof(protectData));
            this.unprotectData = unprotectData ?? throw new ArgumentNullException(nameof(unprotectData));
        }

        public IReadOnlyList<ModelProfile> Load()
        {
            if (!File.Exists(filePath))
            {
                return new List<ModelProfile>().AsReadOnly();
            }

            var json = File.ReadAllText(filePath, Encoding.UTF8);
            var records = JsonConvert.DeserializeObject<List<StoredModelProfile>>(json) ??
                          new List<StoredModelProfile>();
            return records.Select(Decrypt).ToList().AsReadOnly();
        }

        public void Save(IEnumerable<ModelProfile> profiles)
        {
            if (profiles == null)
            {
                throw new ArgumentNullException(nameof(profiles));
            }

            var list = profiles.ToList();
            foreach (var profile in list)
            {
                profile.Validate();
            }

            var duplicates = list.GroupBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                                 .Where(group => group.Count() > 1)
                                 .Select(group => group.Key)
                                 .ToList();
            if (duplicates.Count > 0)
            {
                throw new InvalidOperationException($"方案名称重复：{string.Join(", ", duplicates)}");
            }

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonConvert.SerializeObject(
                list.Select(Encrypt).ToList(),
                Formatting.Indented);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        private StoredModelProfile Encrypt(ModelProfile profile)
        {
            return new StoredModelProfile
            {
                Id = profile.Id,
                Name = profile.Name,
                BaseUrl = profile.BaseUrl,
                ModelName = profile.ModelName,
                Temperature = profile.Temperature,
                MaxTokens = profile.MaxTokens,
                ProtocolMode = profile.ProtocolMode.ToString(),
                AuthScheme = profile.AuthScheme.ToString(),
                ApiKeyCipher = Convert.ToBase64String(protectData(profile.ApiKey))
            };
        }

        private ModelProfile Decrypt(StoredModelProfile record)
        {
            var cipher = Convert.FromBase64String(record.ApiKeyCipher);
            return new ModelProfile
            {
                Id = record.Id,
                Name = record.Name,
                BaseUrl = record.BaseUrl,
                ModelName = record.ModelName,
                Temperature = record.Temperature,
                MaxTokens = record.MaxTokens,
                ProtocolMode = Enum.TryParse<AgentProtocolMode>(record.ProtocolMode, true, out var mode)
                    ? mode
                    : AgentProtocolMode.Auto,
                AuthScheme = Enum.TryParse<AuthScheme>(record.AuthScheme, true, out var auth)
                    ? auth
                    : AuthScheme.Bearer,
                ApiKey = unprotectData(cipher)
            };
        }

        private sealed class StoredModelProfile
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string BaseUrl { get; set; }
            public string ApiKeyCipher { get; set; }
            public string ModelName { get; set; }
            public double Temperature { get; set; }
            public int? MaxTokens { get; set; }
            public string ProtocolMode { get; set; }
            public string AuthScheme { get; set; }
        }
    }
}
