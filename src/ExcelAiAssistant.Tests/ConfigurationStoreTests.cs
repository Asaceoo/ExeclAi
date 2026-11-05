using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Models;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ConfigurationStoreTests : IDisposable
    {
        private readonly string filePath;

        public ConfigurationStoreTests()
        {
            filePath = Path.Combine(Path.GetTempPath(), $"excel-ai-{Guid.NewGuid():N}.json");
        }

        [Fact]
        public void Saves_and_loads_profiles_with_protected_api_key()
        {
            var store = CreateStore();
            var profile = new ModelProfile
            {
                Name = "DeepSeek",
                BaseUrl = "https://api.deepseek.com/v1",
                ApiKey = "sk-secret",
                ModelName = "deepseek-chat"
            };

            store.Save(new List<ModelProfile> { profile });
            var loaded = Assert.Single(store.Load());

            Assert.Equal("DeepSeek", loaded.Name);
            Assert.Equal("https://api.deepseek.com/v1", loaded.BaseUrl);
            Assert.Equal("deepseek-chat", loaded.ModelName);
            Assert.Equal("sk-secret", loaded.ApiKey);
            Assert.Contains("ApiKeyCipher", File.ReadAllText(filePath));
            Assert.DoesNotContain("sk-secret", File.ReadAllText(filePath));
        }

        [Fact]
        public void Rejects_duplicate_profile_names()
        {
            var store = CreateStore();
            var profiles = new List<ModelProfile>
            {
                CreateProfile("Same"),
                CreateProfile("same")
            };

            Assert.Throws<InvalidOperationException>(() => store.Save(profiles));
        }

        [Fact]
        public void Validates_required_fields()
        {
            var profile = new ModelProfile { Name = "Bad" };
            Assert.Throws<ArgumentException>(() => profile.Validate());
        }

        [Fact]
        public void Round_trips_real_dpapi_api_key()
        {
            var store = new ConfigurationStore(
                filePath,
                apiKey => ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(apiKey),
                    Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                    DataProtectionScope.CurrentUser),
                cipher => Encoding.UTF8.GetString(
                    ProtectedData.Unprotect(
                        cipher,
                        Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                        DataProtectionScope.CurrentUser)));

            store.Save(new List<ModelProfile>
            {
                new ModelProfile
                {
                    Name = "sensenova",
                    BaseUrl = "https://token.sensenova.cn/v1",
                    ApiKey = "sk-LWxQ5BTj8qy5Bu3KReYmNwSW6CA4M3jD",
                    ModelName = "sensenova-6.8-flash-lite",
                    Temperature = 0.1
                }
            });

            var loaded = Assert.Single(store.Load());
            Assert.Equal("sk-LWxQ5BTj8qy5Bu3KReYmNwSW6CA4M3jD", loaded.ApiKey);
            Assert.DoesNotContain("sk-LWxQ5BTj8qy5Bu3KReYmNwSW6CA4M3jD", File.ReadAllText(filePath));
        }

        private static ModelProfile CreateProfile(string name)
        {
            return new ModelProfile
            {
                Name = name,
                BaseUrl = "https://example.com/v1",
                ApiKey = "key",
                ModelName = "test-model"
            };
        }

        private ConfigurationStore CreateStore()
        {
            return new ConfigurationStore(
                filePath,
                value => Encoding.UTF8.GetBytes($"protected:{value}:protected"),
                cipher => Encoding.UTF8.GetString(cipher).Replace("protected:", string.Empty).Replace(":protected", string.Empty));
        }

        public void Dispose()
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }
}
