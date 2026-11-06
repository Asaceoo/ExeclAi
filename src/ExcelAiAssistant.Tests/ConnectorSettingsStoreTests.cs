using System;
using System.IO;
using ExcelAiAssistant.Core.Configuration;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M29 连接面板：MCP / CLI 开关设置存储的行为锁（默认开启、往返、损坏回落、原子写入）。
    /// 该文件是面板与两个入口的共同真值，任何一处漂移都会造成“面板显示已关闭但入口仍在跑”。
    /// </summary>
    public class ConnectorSettingsStoreTests : IDisposable
    {
        private readonly string tempFile;
        private readonly ConnectorSettingsStore store;

        public ConnectorSettingsStoreTests()
        {
            tempFile = Path.Combine(Path.GetTempPath(), "jyyj-connectors-" + Guid.NewGuid().ToString("N") + ".json");
            store = new ConnectorSettingsStore(tempFile);
        }

        public void Dispose()
        {
            try { if (File.Exists(tempFile)) { File.Delete(tempFile); } } catch { }
            try { if (File.Exists(tempFile + ".tmp")) { File.Delete(tempFile + ".tmp"); } } catch { }
        }

        [Fact]
        public void Missing_file_defaults_to_enabled()
        {
            var settings = store.Load();
            Assert.True(settings.McpEnabled);
            Assert.True(settings.CliEnabled);
        }

        [Fact]
        public void Save_and_load_round_trips()
        {
            store.Save(new ConnectorSettings { McpEnabled = false, CliEnabled = true });
            var loaded = store.Load();
            Assert.False(loaded.McpEnabled);
            Assert.True(loaded.CliEnabled);

            store.Save(new ConnectorSettings { McpEnabled = true, CliEnabled = false });
            var second = store.Load();
            Assert.True(second.McpEnabled);
            Assert.False(second.CliEnabled);
        }

        [Fact]
        public void Corrupt_json_falls_back_to_enabled()
        {
            File.WriteAllText(tempFile, "{ this is not json ");
            var settings = store.Load();
            Assert.True(settings.McpEnabled);
            Assert.True(settings.CliEnabled);
        }

        [Fact]
        public void Empty_json_falls_back_to_enabled()
        {
            File.WriteAllText(tempFile, "   ");
            var settings = store.Load();
            Assert.True(settings.McpEnabled);
            Assert.True(settings.CliEnabled);
        }

        [Fact]
        public void Save_leaves_no_temp_file_behind()
        {
            store.Save(new ConnectorSettings { McpEnabled = false, CliEnabled = false });
            Assert.True(File.Exists(tempFile));
            Assert.False(File.Exists(tempFile + ".tmp"), "原子写入后不应残留 .tmp 文件");
        }

        [Fact]
        public void Save_rejects_null_settings()
        {
            Assert.Throws<ArgumentNullException>(() => store.Save(null));
        }

        [Fact]
        public void Load_reports_missing_state()
        {
            var settings = store.Load(out var state);
            Assert.Equal("missing", state);
            Assert.True(settings.McpEnabled);
        }

        [Fact]
        public void Load_reports_invalid_state_for_corrupt_file()
        {
            File.WriteAllText(tempFile, "not json at all");
            var settings = store.Load(out var state);
            Assert.Equal("invalid", state);
            Assert.True(settings.McpEnabled);
        }

        [Fact]
        public void Load_reports_ok_state_for_valid_file()
        {
            store.Save(new ConnectorSettings { McpEnabled = false, CliEnabled = true });
            store.Load(out var state);
            Assert.Equal("ok", state);
        }

        [Fact]
        public void Save_fails_loudly_when_target_locked_instead_of_silently_corrupting()
        {
            store.Save(new ConnectorSettings { McpEnabled = true, CliEnabled = true });
            var original = File.ReadAllText(tempFile);

            // 独占打开目标文件：Replace 与降级写都会失败 → Save 必须抛异常，且原内容不得被破坏
            using (var lockStream = new FileStream(tempFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Assert.ThrowsAny<Exception>(() => store.Save(new ConnectorSettings { McpEnabled = false, CliEnabled = false }));
            }

            Assert.Equal(original, File.ReadAllText(tempFile));
        }
        [Fact]
        public void Default_path_is_under_local_appdata()
        {
            var path = ConnectorSettingsStore.DefaultPath();
            Assert.Contains("XaiAssistant", path);
            Assert.EndsWith("connectors.json", path);
        }
    }
}