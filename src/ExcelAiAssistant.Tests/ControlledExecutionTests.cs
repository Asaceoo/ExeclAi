using System;
using System.IO;
using Xunit;
using ExcelAiAssistant.Core.Security;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Tests
{
    public sealed class ControlledExecutionPolicyTests
    {
        [Fact]
        public void Empty_allowlist_means_not_allowlisted()
        {
            Assert.Equal(PolicyDecision.NotAllowlisted,
                ControlledExecutionPolicy.Evaluate("ping", new string[0], new string[0]));
        }

        [Fact]
        public void Allowlist_grants_matching_identity_case_insensitively()
        {
            Assert.Equal(PolicyDecision.Allow,
                ControlledExecutionPolicy.Evaluate("PING", new[] { "ping" }, new string[0]));
        }

        [Fact]
        public void Missing_allowlist_member_is_not_allowlisted()
        {
            Assert.Equal(PolicyDecision.NotAllowlisted,
                ControlledExecutionPolicy.Evaluate("dir", new[] { "ping" }, new string[0]));
        }

        [Fact]
        public void Denylist_always_wins_over_allowlist()
        {
            Assert.Equal(PolicyDecision.Deny,
                ControlledExecutionPolicy.Evaluate("ping", new[] { "ping" }, new[] { "ping" }));
        }

        [Fact]
        public void Full_path_matching_normalizes_case_and_separators()
        {
            var exe = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows", "System32", "cmd.exe");
            Assert.Equal(PolicyDecision.Allow,
                ControlledExecutionPolicy.Evaluate(exe.ToUpperInvariant().Replace('\\', '/'), new[] { exe }, new string[0], fullPath: true));
        }

        [Fact]
        public void Null_or_blank_identity_is_not_allowlisted()
        {
            Assert.Equal(PolicyDecision.NotAllowlisted,
                ControlledExecutionPolicy.Evaluate(null, new[] { "x" }, new string[0]));
            Assert.Equal(PolicyDecision.NotAllowlisted,
                ControlledExecutionPolicy.Evaluate("  ", new[] { "x" }, new string[0]));
        }
    }

    public sealed class ControlledExecutionStoreTests : IDisposable
    {
        private readonly string dir;
        private readonly string file;
        private readonly ControlledExecutionStore store;

        public ControlledExecutionStoreTests()
        {
            dir = Path.Combine(Path.GetTempPath(), "xai_security_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            file = Path.Combine(dir, "controlled-execution.json");
            store = new ControlledExecutionStore(file);
        }

        public void Dispose()
        {
            try { Directory.Delete(dir, recursive: true); } catch { }
        }

        [Fact]
        public void Missing_file_returns_default_disabled_config()
        {
            var config = store.Load();
            Assert.False(config.Enabled);
            Assert.Equal(30, config.TimeoutSeconds);
            Assert.Equal(64, config.OutputLimitKb);
        }

        [Fact]
        public void Save_and_load_round_trips_config()
        {
            var config = new ControlledExecutionConfig
            {
                Enabled = true,
                CliAllowlist = new[] { "ping", "dir" },
                CliDenylist = new[] { "format" },
                ExeAllowlist = new[] { @"C:\Tools\app.exe" },
                ScriptDenylist = new[] { @"D:\scripts\evil.py" },
                TimeoutSeconds = 15,
                OutputLimitKb = 8,
                PythonPath = "python3"
            };

            store.Save(config);
            var loaded = store.Load();

            Assert.True(loaded.Enabled);
            Assert.Equal(new[] { "ping", "dir" }, loaded.CliAllowlist);
            Assert.Equal(new[] { "format" }, loaded.CliDenylist);
            Assert.Equal(15, loaded.TimeoutSeconds);
            Assert.Equal(8, loaded.OutputLimitKb);
            Assert.Equal("python3", loaded.PythonPath);

            var raw = File.ReadAllText(file);
            Assert.Contains("\"enabled\": true", raw);
            Assert.Contains("\"timeout_seconds\": 15", raw);
        }

        [Fact]
        public void Corrupt_file_falls_back_to_defaults()
        {
            File.WriteAllText(file, "{ not-valid-json ");
            var config = store.Load();
            Assert.False(config.Enabled);
        }
    }

    /// <summary>P5-C4 受控策略导出/导入：JSON 序列化应完整还原白名单/黑名单与开关。</summary>
    public sealed class PolicyJsonRoundTripTests
    {
        [Fact]
        public void Export_import_round_trip_preserves_all_lists_and_flags()
        {
            var original = new ControlledExecutionConfig
            {
                Enabled = true,
                CliAllowlist = new[] { "ping", "dir" },
                CliDenylist = new[] { "shutdown" },
                ExeAllowlist = new[] { @"C:\Tools\1.exe", @"C:\Tools\2.exe" },
                ExeDenylist = new[] { @"C:\Windows\System32\cmd.exe" },
                ScriptAllowlist = new[] { @"D:\scripts\safe.py" },
                ScriptDenylist = new[] { @"D:\scripts\1.py", @"D:\scripts\2.py" },
                TimeoutSeconds = 20,
                OutputLimitKb = 128,
                PythonPath = "C:\\Python311\\python.exe"
            };

            // 导出路径：JsonConvert.SerializeObject(config)
            var json = JsonConvert.SerializeObject(original, Formatting.Indented);
            // 导入路径：JsonConvert.DeserializeObject<ControlledExecutionConfig>(json)
            var restored = JsonConvert.DeserializeObject<ControlledExecutionConfig>(json);

            Assert.NotNull(restored);
            Assert.True(restored.Enabled);
            Assert.Equal(20, restored.TimeoutSeconds);
            Assert.Equal(128, restored.OutputLimitKb);
            Assert.Equal(@"C:\Python311\python.exe", restored.PythonPath);
            Assert.Equal(original.CliAllowlist, restored.CliAllowlist);
            Assert.Equal(original.CliDenylist, restored.CliDenylist);
            Assert.Equal(original.ExeAllowlist, restored.ExeAllowlist);
            Assert.Equal(original.ExeDenylist, restored.ExeDenylist);
            Assert.Equal(original.ScriptAllowlist, restored.ScriptAllowlist);
            Assert.Equal(original.ScriptDenylist, restored.ScriptDenylist);
        }

        [Fact]
        public void Import_invalid_json_throws()
        {
            var threw = false;
            try { JsonConvert.DeserializeObject<ControlledExecutionConfig>("{ not-json "); }
            catch { threw = true; }
            Assert.True(threw);
        }
    }
}