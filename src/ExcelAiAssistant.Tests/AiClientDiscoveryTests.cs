using System;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Connectors;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M30 对抗性审查回归：AI 客户端配置登记的核心不变量——
    /// 别人的条目不能丢、用户自定义字段不能丢、备份基线不能被覆盖、TOML 后面的节不能被吃掉。
    /// </summary>
    public class AiClientDiscoveryTests : IDisposable
    {
        private readonly string _dir;

        public AiClientDiscoveryTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "jyyj_m30_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        private string FakeServer()
        {
            var path = Path.Combine(_dir, "ExcelAiMcpServer.exe");
            if (!File.Exists(path)) { File.WriteAllText(path, "stub"); }
            return path;
        }

        private string Write(string name, string content)
        {
            var path = Path.Combine(_dir, name);
            File.WriteAllText(path, content, new System.Text.UTF8Encoding(false));
            return path;
        }

        [Fact]
        public void Json_register_preserves_other_servers_and_user_fields()
        {
            var path = Write("mcp.json", @"{
  ""mcpServers"": {
    ""other"": { ""command"": ""other.exe"", ""args"": [""--x""] },
    ""jyyj-excel"": { ""command"": ""old.exe"", ""args"": [], ""env"": { ""A"": ""1"" }, ""disabled"": true }
  }
}");

            var result = AiClientDiscovery.ConnectAt("json-mcpServers", path, FakeServer(), new[] { "--allow-write" }, true);

            Assert.True(string.IsNullOrEmpty(result.Error), result.Error);
            Assert.True(result.Changed);
            var root = JObject.Parse(File.ReadAllText(path));
            Assert.NotNull(root["mcpServers"]["other"]);
            var entry = root["mcpServers"]["jyyj-excel"];
            Assert.Equal(FakeServer(), (string)entry["command"]);
            Assert.Equal("--allow-write", (string)entry["args"][0]);
            // 用户自定义字段必须保留（不能整条覆盖）
            Assert.Equal("1", (string)entry["env"]["A"]);
            Assert.True((bool)entry["disabled"]);
        }

        [Fact]
        public void Json_second_apply_is_idempotent_and_keeps_original_backup()
        {
            var path = Write("mcp2.json", @"{ ""mcpServers"": { ""other"": { ""command"": ""o.exe"" } } }");

            var first = AiClientDiscovery.ConnectAt("json-mcpServers", path, FakeServer(), new string[0], true);
            Assert.True(first.BackupCreated);
            var backupText = File.ReadAllText(first.BackupPath);

            var second = AiClientDiscovery.ConnectAt("json-mcpServers", path, FakeServer(), new string[0], true);
            Assert.False(second.Changed, "同样参数第二次写入不应产生变化");
            Assert.False(second.BackupCreated, "备份只在首次创建，不能被覆盖");
            Assert.Equal(backupText, File.ReadAllText(second.BackupPath));
        }

        [Fact]
        public void Dsh_store_keeps_other_connections_and_timestamps()
        {
            var path = Write("mcp_connector.json", @"{
  ""unit"": { ""name"": ""mcp_connector"", ""version"": 1 },
  ""tables"": { ""connections"": {
      ""keep-me"": { ""key"": ""keep-me"", ""command"": ""x.exe"", ""enabled"": false },
      ""jyyj-excel"": { ""key"": ""jyyj-excel"", ""command"": ""old.exe"", ""enabled"": false, ""createdAt"": 111 }
  } }
}");

            var result = AiClientDiscovery.ConnectAt("dsh-store", path, FakeServer(), new string[0], true);
            Assert.True(string.IsNullOrEmpty(result.Error), result.Error);

            var root = JObject.Parse(File.ReadAllText(path));
            Assert.NotNull(root["tables"]["connections"]["keep-me"]);
            Assert.Equal(111L, (long)root["tables"]["connections"]["jyyj-excel"]["createdAt"]);
            Assert.False((bool)root["tables"]["connections"]["jyyj-excel"]["enabled"], "用户禁用决定不能被撤销");
            Assert.Equal("keep-me", (string)root["tables"]["connections"]["keep-me"]["key"]);
        }

        [Fact]
        public void Json_register_refuses_when_container_type_is_wrong()
        {
            var path = Write("bad.json", @"{ ""mcpServers"": [ 1, 2, 3 ] }");
            var before = File.ReadAllText(path);

            var result = AiClientDiscovery.ConnectAt("json-mcpServers", path, FakeServer(), new string[0], true);

            Assert.False(string.IsNullOrEmpty(result.Error));
            Assert.Equal(before, File.ReadAllText(path));
        }

        [Fact]
        public void Codex_replace_keeps_following_sections()
        {
            var path = Write("config.toml", @"model = ""x""

[mcp_servers.jyyj-excel]
command = 'old.exe'
args = []

[mcp_servers.node_repl]
args = []
command = 'node.exe'

[features]
goals = true
");

            var result = AiClientDiscovery.ConnectAt("toml-codex", path, FakeServer(), new[] { "--allow-write" }, true);
            Assert.True(string.IsNullOrEmpty(result.Error), result.Error);

            var text = File.ReadAllText(path);
            Assert.Contains("[mcp_servers.jyyj-excel]", text);
            Assert.Contains("command = ", text);
            Assert.Contains("[mcp_servers.node_repl]", text);
            Assert.Contains("[features]", text);
            // 后续节不能挂到 jyyj-excel 下面：node_repl 节必须仍带自己的 command
            var idx = text.IndexOf("[mcp_servers.node_repl]", StringComparison.Ordinal);
            Assert.Contains("command = 'node.exe'", text.Substring(idx));
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, @"\[mcp_servers\.jyyj-excel\]").Cast<object>());
        }

        [Fact]
        public void Codex_read_entry_roundtrips_quoted_args()
        {
            var server = FakeServer();
            var toml = "[mcp_servers.jyyj-excel]\n" +
                "command = \"" + server.Replace("\\", "\\\\") + "\"\n" +
                "args = [\"--workbook\", \"C:\\\\My Docs\\\\a.xlsx\"]\n" +
                "startup_timeout_sec = 120\n";
            var path = Write("config2.toml", toml);
            var entry = AiClientDiscovery.ReadEntry("toml-codex", path);

            Assert.NotNull(entry);
            Assert.Equal(FakeServer(), entry.Item1);
            Assert.Equal(new[] { "--workbook", @"C:\My Docs\a.xlsx" }, entry.Item2);
        }
    }
}