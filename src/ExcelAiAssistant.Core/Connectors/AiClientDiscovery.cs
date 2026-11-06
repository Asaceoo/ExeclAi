using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Connectors
{
    /// <summary>本机可接入的 AI 客户端（MCP host）探测结果。</summary>
    public sealed class AiClientInfo
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string ConfigPath { get; set; }
        public string Schema { get; set; }          // json-mcpServers / toml-codex / dsh-store
        public bool Installed { get; set; }
        public bool Configured { get; set; }        // 是否已登记 jyyj-excel
        public string Command { get; set; }         // 已登记时的启动命令
        public string[] Arguments { get; set; }     // 已登记时的启动参数（保持数组，避免空格路径被拆错）
        public string Detail { get; set; }
    }

    public sealed class AiClientConnectResult
    {
        public string Client { get; set; }
        public bool Applied { get; set; }
        public bool Changed { get; set; }
        public string Action { get; set; }          // create / update / unchanged
        public string ConfigPath { get; set; }
        public string BackupPath { get; set; }
        public bool BackupCreated { get; set; }
        public string Preview { get; set; }
        public string Error { get; set; }
    }

    public sealed class AiClientTestResult
    {
        public string Client { get; set; }
        public bool Configured { get; set; }
        public bool Reachable { get; set; }
        public string ProtocolVersion { get; set; }
        public long ElapsedMs { get; set; }
        public string Reason { get; set; }          // not_configured / spawn_failed / timeout / no_response / server_error / protocol_mismatch
        public string Error { get; set; }
    }

    /// <summary>
    /// 本机 AI 客户端（DSH / WorkBuddy / Codex / Claude / Cursor / Windsurf / Trae）的 MCP 配置探测、
    /// 一键登记与连通性自测。三个能力同源，MCP 工具与 CLI 共用：
    ///   Detect  —— 只读扫描：装了没、配了没、指到哪个 exe
    ///   Connect —— 预览 / 落地（写前备份、只增改 jyyj-excel 一项、保留用户在该条目里的其他字段）
    ///   Test    —— 用该客户端配置里的真实 command+args 启动一次握手
    /// 安全边界：只读用户配置；写入前备份为 *.jyyj-bak（只在首次创建，不覆盖原始基线）；
    /// 解析不合法或结构类型不符时拒绝改写，交给人工处理。
    /// </summary>
    public static class AiClientDiscovery
    {
        public const string ServerName = "jyyj-excel";

        /// <summary>合法客户端 id（供 schema enum 与错误文案复用，避免模型猜别名）。</summary>
        public static readonly string[] ClientIds =
        {
            "workbuddy", "dsh", "codex", "claude-desktop", "claude-code", "cursor", "windsurf", "trae"
        };

        private sealed class ClientSpec
        {
            public string Id;
            public string DisplayName;
            public string Schema;
            public Func<string> ResolveConfig;
            public Func<bool> ResolveInstalled;
        }

        private static IEnumerable<ClientSpec> Specs()
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            yield return new ClientSpec
            {
                Id = "claude-desktop", DisplayName = "Claude Desktop", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(appData, "Claude", "claude_desktop_config.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(appData, "Claude"))
            };
            yield return new ClientSpec
            {
                Id = "claude-code", DisplayName = "Claude Code", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(home, ".claude.json"),
                ResolveInstalled = () => File.Exists(Path.Combine(home, ".claude.json")) || Directory.Exists(Path.Combine(home, ".claude"))
            };
            yield return new ClientSpec
            {
                Id = "codex", DisplayName = "Codex CLI", Schema = "toml-codex",
                ResolveConfig = () => Path.Combine(home, ".codex", "config.toml"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(home, ".codex"))
            };
            yield return new ClientSpec
            {
                Id = "workbuddy", DisplayName = "WorkBuddy", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(home, ".workbuddy", "mcp.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(home, ".workbuddy"))
            };
            yield return new ClientSpec
            {
                Id = "dsh", DisplayName = "DSH", Schema = "dsh-store",
                ResolveConfig = () => Path.Combine(home, ".dsh", "storages", "mcp_connector.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(home, ".dsh"))
            };
            yield return new ClientSpec
            {
                Id = "cursor", DisplayName = "Cursor", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(home, ".cursor", "mcp.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(home, ".cursor"))
            };
            yield return new ClientSpec
            {
                Id = "windsurf", DisplayName = "Windsurf", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(home, ".codeium", "windsurf", "mcp_config.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(home, ".codeium", "windsurf"))
            };
            yield return new ClientSpec
            {
                Id = "trae", DisplayName = "Trae CN", Schema = "json-mcpServers",
                ResolveConfig = () => Path.Combine(appData, "Trae CN", "User", "mcp.json"),
                ResolveInstalled = () => Directory.Exists(Path.Combine(appData, "Trae CN"))
            };
        }

        public static List<AiClientInfo> Detect()
        {
            var list = new List<AiClientInfo>();
            foreach (var spec in Specs())
            {
                var info = new AiClientInfo { Id = spec.Id, DisplayName = spec.DisplayName, Schema = spec.Schema };
                try
                {
                    info.ConfigPath = spec.ResolveConfig();
                    info.Installed = SafeBool(spec.ResolveInstalled) || File.Exists(info.ConfigPath);

                    if (!File.Exists(info.ConfigPath))
                    {
                        info.Detail = info.Installed ? "已安装，尚无 MCP 配置文件（connect 会新建）" : "未检测到安装";
                    }
                    else if (new FileInfo(info.ConfigPath).Length == 0)
                    {
                        info.Detail = "配置文件为空（connect 会写入初始结构）";
                    }
                    else
                    {
                        var entry = ReadEntry(spec.Schema, info.ConfigPath);
                        if (entry != null)
                        {
                            info.Configured = true;
                            info.Command = entry.Item1;
                            info.Arguments = entry.Item2;
                            info.Detail = "已登记 jyyj-excel";
                        }
                        else
                        {
                            info.Detail = "已安装，未登记 jyyj-excel";
                        }
                    }
                }
                catch (Exception ex)
                {
                    info.Detail = "配置无法解析（" + info.ConfigPath + "）：" + ex.Message;
                }
                list.Add(info);
            }
            return list;
        }

        private static bool SafeBool(Func<bool> f)
        {
            try { return f(); } catch { return false; }
        }

        private static JObject ParseJson(string text)
        {
            // DateParseHandling=None 保 "2024-01-02" 原样；Ignore 注释（Newtonsoft 无法原样写回注释，
            // 因此这类文件在备份后会被重排——已在文档中声明）。
            var settings = new JsonLoadSettings
            {
                CommentHandling = CommentHandling.Ignore,
                DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Replace
            };
            using (var reader = new JsonTextReader(new StringReader(text))
            {
                DateParseHandling = DateParseHandling.None,
                FloatParseHandling = FloatParseHandling.Decimal
            })
            {
                return JObject.Load(reader, settings);
            }
        }

        /// <summary>读取该客户端里 jyyj-excel 的启动命令与参数；未登记返回 null。</summary>
        public static Tuple<string, string[]> ReadEntry(string schema, string configPath)
        {
            if (string.IsNullOrEmpty(configPath) || !File.Exists(configPath)) { return null; }
            var text = File.ReadAllText(configPath, Encoding.UTF8);
            if (string.IsNullOrWhiteSpace(text)) { return null; }
            if (schema == "toml-codex") { return ReadCodexEntry(text); }

            var root = ParseJson(text);
            JToken node = schema == "dsh-store"
                ? root["tables"]?["connections"]?[ServerName]
                : root["mcpServers"]?[ServerName];
            if (node == null) { return null; }
            var cmd = Convert.ToString(node["command"]);
            var argsArr = node["args"] as JArray;
            var args = argsArr == null
                ? new string[0]
                : argsArr.Select(a => Convert.ToString(a)).ToArray();
            return Tuple.Create(cmd, args);
        }

        private static Tuple<string, string[]> ReadCodexEntry(string text)
        {
            var header = "[mcp_servers." + ServerName + "]";
            string command = null;
            var args = new List<string>();
            var inside = false;
            var inArgs = false;
            var argsBuffer = new StringBuilder();

            foreach (var raw in text.Split('\n'))
            {
                var line = raw.Trim().TrimEnd('\r');
                if (!inside)
                {
                    if (string.Equals(line, header, StringComparison.Ordinal)) { inside = true; }
                    continue;
                }
                if (line.StartsWith("[", StringComparison.Ordinal))
                {
                    break; // 下一个表开始
                }

                if (inArgs)
                {
                    argsBuffer.Append(' ').Append(StripComment(line));
                    if (line.Contains("]")) { inArgs = false; ParseTomlArray(argsBuffer.ToString(), args); }
                    continue;
                }

                var eq = line.IndexOf('=');
                if (eq <= 0) { continue; }
                var key = line.Substring(0, eq).Trim();
                var val = StripComment(line.Substring(eq + 1).Trim());
                if (key == "command")
                {
                    command = UnescapeToml(UnquoteToml(val));
                }
                else if (key == "args")
                {
                    if (val.Contains("]")) { ParseTomlArray(val, args); }
                    else { argsBuffer.Clear().Append(val); inArgs = true; }
                }
            }
            if (command == null) { return null; }
            return Tuple.Create(command, args.ToArray());
        }

        private static string StripComment(string value)
        {
            // 去掉引号外的 # 注释
            var inQuote = '\0';
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (inQuote != '\0')
                {
                    if (c == '\\') { i++; continue; }
                    if (c == inQuote) { inQuote = '\0'; }
                    continue;
                }
                if (c == '"' || c == '\'') { inQuote = c; continue; }
                if (c == '#') { return value.Substring(0, i).TrimEnd(); }
            }
            return value.TrimEnd();
        }

        private static string UnquoteToml(string value)
        {
            if (value.Length >= 2 && (value[0] == '"' || value[0] == '\'') && value[value.Length - 1] == value[0])
            {
                return value.Substring(1, value.Length - 2);
            }
            return value;
        }

        private static void ParseTomlArray(string value, List<string> args)
        {
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(value, "\"([^\"]*)\"|'([^']*)'"))
            {
                args.Add(UnescapeToml(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value));
            }
        }

        private static string UnescapeToml(string value) =>
            value.Replace("\\\\", "\\").Replace("\\\"", "\"").Replace("\\'", "'");

        /// <summary>解析要登记的 MCP server 可执行文件：环境变量 → 安装目录 → 本程序目录 → 仓库开发目录。</summary>
        public static string ResolveServerCommand()
        {
            var env = Environment.GetEnvironmentVariable("JYYJ_MCP_COMMAND");
            if (!string.IsNullOrWhiteSpace(env))
            {
                var full = TryFullPath(env);
                if (full != null && File.Exists(full)) { return full; }
            }

            var candidates = new List<string>();
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrEmpty(localAppData))
            {
                candidates.Add(Path.Combine(localAppData, "XaiAssistant", "app", "ExcelAiMcpServer.exe"));
                candidates.Add(Path.Combine(localAppData, "XaiAssistant", "mcp", "ExcelAiMcpServer.exe"));
            }
            if (!string.IsNullOrEmpty(programFiles))
            {
                candidates.Add(Path.Combine(programFiles, "JYYJ助手", "mcp", "ExcelAiMcpServer.exe"));
            }
            try
            {
                var selfDir = Path.GetDirectoryName(typeof(AiClientDiscovery).Assembly.Location);
                if (!string.IsNullOrEmpty(selfDir))
                {
                    candidates.Add(Path.Combine(selfDir, "ExcelAiMcpServer.exe"));
                    candidates.Add(Path.Combine(selfDir, "..", "mcp", "ExcelAiMcpServer.exe"));
                }
            }
            catch { }
            candidates.AddRange(FindDevBuilds());

            foreach (var candidate in candidates)
            {
                var full = TryFullPath(candidate);
                if (full != null && File.Exists(full)) { return full; }
            }
            return null;
        }

        private static string TryFullPath(string path)
        {
            try { return Path.GetFullPath(path); } catch { return null; }
        }

        private static IEnumerable<string> FindDevBuilds()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            for (var i = 0; i < 6 && dir != null; i++, dir = dir.Parent)
            {
                yield return Path.Combine(dir.FullName, "tools", "McpServer", "bin", "Release", "net48", "ExcelAiMcpServer.exe");
            }
        }

        /// <summary>生成登记预览；apply=true 才落盘。</summary>
        public static AiClientConnectResult Connect(string clientId, string command, string[] args, bool apply)
        {
            var spec = Specs().FirstOrDefault(s => string.Equals(s.Id, clientId, StringComparison.OrdinalIgnoreCase));
            if (spec == null)
            {
                return new AiClientConnectResult { Client = clientId, Error = "未知客户端：" + clientId };
            }
            var result = ConnectAt(spec.Schema, spec.ResolveConfig(), command, args, apply);
            result.Client = spec.Id;
            return result;
        }

        /// <summary>按给定 schema + 配置文件路径执行登记（单测与自定义路径复用）。</summary>
        public static AiClientConnectResult ConnectAt(string schema, string configPath, string command, string[] args, bool apply)
        {
            var result = new AiClientConnectResult { Applied = apply };
            try
            {
                result.ConfigPath = configPath;
                var cmd = string.IsNullOrWhiteSpace(command) ? ResolveServerCommand() : TryFullPath(command);
                if (string.IsNullOrWhiteSpace(cmd))
                {
                    result.Error = "找不到 ExcelAiMcpServer.exe：请把 mcp-server 包解压到 %LOCALAPPDATA%\\XaiAssistant\\mcp，"
                        + "或设置环境变量 JYYJ_MCP_COMMAND 指向它，或在调用时传 command。";
                    return result;
                }
                if (!File.Exists(cmd))
                {
                    result.Error = "指定的 server 可执行文件不存在：" + cmd;
                    return result;
                }

                var argList = args ?? new string[0];
                var parent = Path.GetDirectoryName(result.ConfigPath);
                result.Action = File.Exists(result.ConfigPath) ? "update" : "create";
                result.Preview = schema == "toml-codex"
                    ? BuildTomlSnippet(cmd, argList)
                    : JsonConvert.SerializeObject(BuildEntryNode(schema, cmd, argList, null), Formatting.Indented);

                if (!apply) { return result; }

                if (!string.IsNullOrEmpty(parent) && !Directory.Exists(parent))
                {
                    result.Error = "客户端尚未创建配置目录：" + parent + "（可能未安装或从未启动过）";
                    return result;
                }

                // 超大配置文件拒绝改写：整树重排风险高于收益
                if (File.Exists(result.ConfigPath) && new FileInfo(result.ConfigPath).Length > 8 * 1024 * 1024)
                {
                    result.Error = "配置文件超过 8MB，为避免整体重排风险拒绝改写：" + result.ConfigPath + "；请手动添加 jyyj-excel 条目。";
                    return result;
                }

                // 备份只在首次创建：模型/用户重试 apply 不会把"原始基线"覆盖成"已改过的版本"
                var backup = result.ConfigPath + ".jyyj-bak";
                if (File.Exists(result.ConfigPath) && !File.Exists(backup))
                {
                    File.Copy(result.ConfigPath, backup, false);
                    result.BackupCreated = true;
                }
                result.BackupPath = File.Exists(backup) ? backup : null;

                if (schema == "toml-codex")
                {
                    result.Changed = WriteCodexEntry(result.ConfigPath, cmd, argList);
                }
                else
                {
                    result.Changed = WriteJsonEntry(schema, result.ConfigPath, cmd, argList);
                }
                result.Action = result.Changed ? (result.Action == "create" ? "create" : "update") : "unchanged";
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
            }
            return result;
        }

        /// <summary>构造要写入的条目：已有条目则在其基础上只改 command/args，保留用户加的 env/cwd/enabled 等。</summary>
        private static JObject BuildEntryNode(string schema, string command, string[] args, JToken existing)
        {
            var argsArr = new JArray(args ?? new string[0]);

            if (existing is JObject old)
            {
                var merged = (JObject)old.DeepClone();
                merged["command"] = command;
                merged["args"] = argsArr;
                if (schema == "dsh-store")
                {
                    if (merged["key"] == null) { merged["key"] = ServerName; }
                    if (merged["serverKey"] == null) { merged["serverKey"] = ServerName; }
                    if (merged["serverName"] == null) { merged["serverName"] = ServerName; }
                    if (merged["enabled"] == null) { merged["enabled"] = true; }
                    merged["updatedAt"] = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                }
                return merged;
            }

            if (schema == "dsh-store")
            {
                var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                return new JObject
                {
                    ["key"] = ServerName,
                    ["connectorId"] = "__json__",
                    ["kind"] = "json",
                    ["name"] = "JYYJ助手",
                    ["serverKey"] = ServerName,
                    ["serverName"] = ServerName,
                    ["transport"] = "stdio",
                    ["command"] = command,
                    ["args"] = argsArr,
                    ["env"] = new JObject(),
                    ["cwd"] = "",
                    ["headers"] = new JObject(),
                    ["enabled"] = true,
                    ["injectionMode"] = "always",
                    ["createdAt"] = now,
                    ["updatedAt"] = now
                };
            }

            return new JObject
            {
                ["command"] = command,
                ["args"] = argsArr,
                ["env"] = new JObject()
            };
        }

        private static bool WriteJsonEntry(string schema, string configPath, string command, string[] args)
        {
            JObject root;
            if (File.Exists(configPath) && new FileInfo(configPath).Length > 0)
            {
                root = ParseJson(File.ReadAllText(configPath, Encoding.UTF8));
            }
            else
            {
                root = new JObject();
            }

            var path = schema == "dsh-store" ? new[] { "tables", "connections" } : new[] { "mcpServers" };
            JObject container = root;
            foreach (var segment in path)
            {
                if (container[segment] is JObject child) { container = child; }
                else if (container[segment] == null)
                {
                    var created = new JObject();
                    container[segment] = created;
                    container = created;
                }
                else
                {
                    // 结构类型不符（被改成数组/字符串）：拒绝覆盖，避免抹掉用户其他条目
                    throw new InvalidOperationException(
                        "配置结构异常：" + segment + " 不是对象（" + TypeName(container[segment]) + "），已拒绝改写以免覆盖既有配置：" + configPath);
                }
            }

            var existing = container[ServerName];
            var entry = BuildEntryNode(schema, command, args, existing);
            if (JToken.DeepEquals(existing, entry)) { return false; }

            container[ServerName] = entry;
            WriteAtomic(configPath, root.ToString(Formatting.Indented));
            return true;
        }

        private static string TypeName(JToken token) =>
            token == null ? "null" : token.Type.ToString();

        private static string BuildTomlSnippet(string command, string[] args)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[mcp_servers." + ServerName + "]");
            sb.AppendLine("command = " + TomlString(command));
            sb.AppendLine("args = [" + string.Join(", ", (args ?? new string[0]).Select(TomlString)) + "]");
            sb.AppendLine("startup_timeout_sec = 120");
            return sb.ToString();
        }

        /// <summary>TOML basic string：单引号字面量无法表达内含单引号，一律用双引号 + 转义。</summary>
        private static string TomlString(string value)
        {
            var sb = new StringBuilder("\"");
            foreach (var c in value ?? "")
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.Append('"').ToString();
        }

        private static bool WriteCodexEntry(string configPath, string command, string[] args)
        {
            var text = File.Exists(configPath) ? File.ReadAllText(configPath, Encoding.UTF8) : "";
            var block = BuildTomlSnippet(command, args);
            var header = "[mcp_servers." + ServerName + "]";

            var lines = text.Replace("\r\n", "\n").Split('\n').ToList();
            var start = -1;
            for (var i = 0; i < lines.Count; i++)
            {
                if (string.Equals(lines[i].Trim(), header, StringComparison.Ordinal)) { start = i; break; }
            }

            if (start >= 0)
            {
                var end = lines.Count;
                for (var i = start + 1; i < lines.Count; i++)
                {
                    var t = lines[i].Trim();
                    if (t.StartsWith("[", StringComparison.Ordinal)) { end = i; break; }
                }
                var replacement = block.TrimEnd().Split('\n').ToList();
                lines.RemoveRange(start, end - start);
                lines.InsertRange(start, replacement);
                // 保留原节后的空行分隔
                if (end < lines.Count && lines.Count > 0 && lines[start + replacement.Count - 1].Length > 0)
                {
                    lines.Insert(start + replacement.Count, "");
                }
            }
            else
            {
                if (lines.Count > 0 && lines[lines.Count - 1].Length > 0) { lines.Add(""); }
                if (!lines.Any(l => l.Trim() == "[mcp_servers]")) { lines.Add("[mcp_servers]"); }
                lines.AddRange(block.TrimEnd().Split('\n'));
            }

            WriteAtomic(configPath, string.Join("\n", lines) + "\n");
            return true;
        }

        private static void WriteAtomic(string path, string content)
        {
            var tmp = path + ".jyyj-tmp";
            try
            {
                File.WriteAllText(tmp, content, new UTF8Encoding(false));
                if (File.Exists(path))
                {
                    var attrs = File.GetAttributes(path);
                    if ((attrs & FileAttributes.ReadOnly) != 0 || (attrs & FileAttributes.ReparsePoint) != 0)
                    {
                        throw new IOException("目标文件为只读或符号链接，已拒绝改写：" + path);
                    }
                    File.Replace(tmp, path, null);
                }
                else { File.Move(tmp, path); }
            }
            finally
            {
                try { if (File.Exists(tmp)) { File.Delete(tmp); } } catch { }
            }
        }

        /// <summary>用客户端配置里的真实 command+args 启动一次 MCP 握手。</summary>
        public static AiClientTestResult Test(string clientId, int timeoutMs)
        {
            var result = new AiClientTestResult { Client = clientId };
            AiClientInfo info;
            try { info = Detect().FirstOrDefault(c => string.Equals(c.Id, clientId, StringComparison.OrdinalIgnoreCase)); }
            catch (Exception ex) { result.Reason = "not_configured"; result.Error = "探测失败：" + ex.Message; return result; }

            if (info == null) { result.Reason = "not_configured"; result.Error = "未知客户端：" + clientId; return result; }
            if (!info.Configured || string.IsNullOrWhiteSpace(info.Command))
            {
                result.Reason = "not_configured";
                result.Error = "未登记 jyyj-excel（" + info.Detail + "）";
                return result;
            }
            result.Configured = true;

            var sw = Stopwatch.StartNew();
            Process proc = null;
            try
            {
                var psi = new ProcessStartInfo(info.Command)
                {
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    Arguments = BuildArguments(info.Arguments)
                };

                proc = Process.Start(psi);

                // stderr 必须排空：MCP server 会往 stderr 写日志，4KB 管道写满后子进程会卡死，
                // 表现为"握手超时"的假故障。
                var stderrTask = proc.StandardError.ReadToEndAsync();
                var stdout = proc.StandardOutput;

                var request = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{\"protocolVersion\":\"2025-06-18\",\"capabilities\":{},\"clientInfo\":{\"name\":\"jyyj-selftest\",\"version\":\"1\"}}}";
                proc.StandardInput.WriteLine(request);
                proc.StandardInput.Flush();

                // 逐行扫描直到拿到 id=1 的响应或超时：服务端可能先发 notifications/*。
                var readTask = System.Threading.Tasks.Task.Run(() =>
                {
                    while (true)
                    {
                        var line = stdout.ReadLine();
                        if (line == null) { return (string)null; }
                        var trimmed = line.Trim();
                        if (!trimmed.StartsWith("{", StringComparison.Ordinal)) { continue; }
                        try
                        {
                            var msg = JObject.Parse(trimmed);
                            if (msg["id"] != null && Convert.ToString(msg["id"]) == "1") { return trimmed; }
                        }
                        catch { /* 非 JSON 行忽略 */ }
                    }
                });

                if (!readTask.Wait(timeoutMs))
                {
                    result.Reason = "timeout";
                    result.Error = "握手超时（" + timeoutMs + "ms）";
                }
                else
                {
                    var line = readTask.Result;
                    if (line == null)
                    {
                        result.Reason = "no_response";
                        var errText = SafeResult(stderrTask, 300);
                        result.Error = "进程未返回响应即退出" + (string.IsNullOrWhiteSpace(errText) ? "" : "；stderr: " + errText);
                    }
                    else
                    {
                        var msg = JObject.Parse(line);
                        if (msg["error"] is JObject err)
                        {
                            result.Reason = "server_error";
                            result.Error = "服务器返回错误：" + Convert.ToString(err["message"]);
                        }
                        else if (msg["result"] is JObject res)
                        {
                            result.ProtocolVersion = Convert.ToString(res["protocolVersion"]);
                            result.Reachable = !string.IsNullOrEmpty(result.ProtocolVersion);
                            if (!result.Reachable) { result.Reason = "protocol_mismatch"; result.Error = "响应缺少 protocolVersion"; }
                        }
                        else
                        {
                            result.Reason = "protocol_mismatch";
                            result.Error = "响应结构不符合 MCP；（" + line.Substring(0, Math.Min(160, line.Length)) + "）";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Reason = "spawn_failed";
                result.Error = "启动失败：" + ex.Message;
            }
            finally
            {
                result.ElapsedMs = sw.ElapsedMilliseconds;
                KillTree(proc);
            }
            return result;
        }

        private static string SafeResult(System.Threading.Tasks.Task<string> task, int max)
        {
            try
            {
                if (!task.Wait(500)) { return null; }
                var text = task.Result;
                if (string.IsNullOrWhiteSpace(text)) { return null; }
                return text.Length <= max ? text : text.Substring(0, max);
            }
            catch { return null; }
        }

        /// <summary>Windows 命令行参数引用规则（MSVC 运行时）：含空格/引号时加引号并转义反斜杠。</summary>
        private static string BuildArguments(string[] args)
        {
            if (args == null || args.Length == 0) { return ""; }
            var sb = new StringBuilder();
            foreach (var arg in args)
            {
                if (sb.Length > 0) { sb.Append(' '); }
                if (arg == null) { continue; }
                if (arg.Length > 0 && arg.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '"' }) < 0)
                {
                    sb.Append(arg);
                    continue;
                }
                sb.Append('"');
                var backslashes = 0;
                foreach (var c in arg)
                {
                    if (c == '\\') { backslashes++; continue; }
                    if (c == '"')
                    {
                        sb.Append('\\', backslashes * 2 + 1).Append('"');
                        backslashes = 0;
                        continue;
                    }
                    sb.Append('\\', backslashes).Append(c);
                    backslashes = 0;
                }
                sb.Append('\\', backslashes * 2).Append('"');
            }
            return sb.ToString();
        }

        /// <summary>结束进程树：npx/cmd 垫片会再拉起真实 server，只 Kill 直接子进程会留孤儿。</summary>
        private static void KillTree(Process proc)
        {
            if (proc == null) { return; }
            try
            {
                if (!proc.HasExited)
                {
                    try
                    {
                        var killer = Process.Start(new ProcessStartInfo("taskkill", "/PID " + proc.Id + " /T /F")
                        {
                            CreateNoWindow = true,
                            UseShellExecute = false,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true
                        });
                        killer?.WaitForExit(3000);
                    }
                    catch { try { proc.Kill(); } catch { } }
                }
            }
            catch { }
            finally
            {
                try { proc.Dispose(); } catch { }
            }
        }
    }
}