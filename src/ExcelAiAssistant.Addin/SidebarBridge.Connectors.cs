using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// SidebarBridge 分域 partial：MCP / CLI 外部接口面（面板 + 开关 + 连通性自检）。
    /// 公共 API 即 WebView2 前端契约（hostObjects.excelAiBridge.*），签名不得改动。
    ///
    /// 设计要点：
    /// - 开关真值只存一处（ConnectorSettingsStore → %LocalAppData%\XaiAssistant\connectors.json），
    ///   MCP/CLI 入口启动时读同一文件并硬拒绝，不存在“面板关了 exe 仍能跑”的分裂。
    /// - 连通性自检优先做**真实握手**：CLI 跑 `--version`，MCP 走一遍 initialize JSON-RPC；
    ///   找不到 exe 时如实报告“未随包分发”，并附可执行的替代路径，不假装成功。
    /// </summary>
    public sealed partial class SidebarBridge
    {
        /// <summary>面板初始化：返回开关状态 + 可执行文件探测结果 + 当前版本。</summary>
        public Task<string> GetConnectors()
        {
            return ExecuteAsync(() =>
            {
                var store = new ConnectorSettingsStore();
                var settings = store.Load(out var settingsState);
                var cliPath = ResolveConnectorExecutable("ExcelAiCli.exe");
                var mcpPath = ResolveConnectorExecutable("ExcelAiMcpServer.exe");

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    mcp_enabled = settings.McpEnabled,
                    cli_enabled = settings.CliEnabled,
                    settings_path = store.FilePath,
                    settings_state = settingsState, // ok/missing/invalid/denied：面板据此提示配置是否损坏
                    cli_exe = cliPath,
                    mcp_exe = mcpPath,
                    cli_available = cliPath != null,
                    mcp_available = mcpPath != null,
                    version = AddinVersionText(),
                    hints = BuildConnectorHints(cliPath, mcpPath)
                }));
            });
        }

        /// <summary>切换开关：body 形如 {"mcp_enabled":true,"cli_enabled":false}；未提供的字段保持不变。</summary>
        public Task<string> SetConnectors(string json)
        {
            return ExecuteAsync(() =>
            {
                var store = new ConnectorSettingsStore();
                var settings = store.Load();

                JObject patch = null;
                if (!string.IsNullOrWhiteSpace(json))
                {
                    patch = JObject.Parse(json);
                }

                if (patch != null)
                {
                    if (patch["mcp_enabled"] != null)
                    {
                        settings.McpEnabled = patch.Value<bool>("mcp_enabled");
                    }

                    if (patch["cli_enabled"] != null)
                    {
                        settings.CliEnabled = patch.Value<bool>("cli_enabled");
                    }
                }

                store.Save(settings);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    mcp_enabled = settings.McpEnabled,
                    cli_enabled = settings.CliEnabled,
                    settings_path = store.FilePath
                }));
            });
        }

        /// <summary>
        /// 连通性自检：对 CLI 与 MCP 各做一次真实握手（找不到 exe 则如实报告不可用）。
        /// 返回每个接口面的 ok / detail / elapsed_ms，供面板逐行显示。
        /// </summary>
        public Task<string> TestConnectors()
        {
            return ExecuteAsync(async () =>
            {
                // 审查 A3/S8：探测必须离开 UI 线程——同步 Wait 会把 Excel 卡成"未响应"（最坏 40s）。
                var result = await Task.Run(() =>
                {
                    var store = new ConnectorSettingsStore();
                    var settings = store.Load();
                    var probes = new List<object>();

                    var cliPath = ResolveConnectorExecutable("ExcelAiCli.exe");
                    probes.Add(ProbeCli(cliPath, settings.CliEnabled));

                    var mcpPath = ResolveConnectorExecutable("ExcelAiMcpServer.exe");
                    probes.Add(ProbeMcp(mcpPath, settings.McpEnabled));

                    // 审查 U4：主动关闭的入口是 skipped，不算失败——否则关掉开关后测试会误报"未连通"。
                    var anyFailed = probes.Any(p =>
                    {
                        var o = JObject.FromObject(p);
                        return !o.Value<bool>("ok") && !o.Value<bool>("skipped");
                    });

                    return new
                    {
                        ok = !anyFailed,
                        all_skipped = probes.All(p => JObject.FromObject(p).Value<bool>("skipped")),
                        mcp_enabled = settings.McpEnabled,
                        cli_enabled = settings.CliEnabled,
                        probes
                    };
                });

                return JsonConvert.SerializeObject(result, Formatting.Indented);
            });
        }

        // ===================== 探测实现 =====================

        /// <summary>审查 B3/B4：观察任务异常，避免未观察任务在 .NET Framework 下触发进程级未处理异常。</summary>
        private static void Observe(Task task)
        {
            if (task == null)
            {
                return;
            }

            task.ContinueWith(
                t => { var ignored = t.Exception; },
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously);
        }

        private static object ProbeCli(string exePath, bool enabled)
        {
            if (!enabled)
            {
                return new { surface = "cli", ok = false, skipped = true, detail = "CLI 开关已关闭：入口将拒绝启动（exit 6）。" };
            }

            if (exePath == null)
            {
                return new
                {
                    surface = "cli",
                    ok = false,
                    skipped = false,
                    not_installed = true,
                    detail = "未找到 ExcelAiCli.exe：本安装包未包含 CLI 程序。见下方提示。"
                };
            }

            var watch = Stopwatch.StartNew();
            try
            {
                var output = RunCapture(exePath, "--version", 8000, out var exitCode);
                watch.Stop();
                // 审查 A6：解析 JSON 而不是字符串包含——格式变更（缩进/无空格）不应把正常 CLI 判成失败
                var ok = false;
                string version = null;
                try
                {
                    var parsed = JObject.Parse(output);
                    ok = exitCode == 0 && parsed.Value<bool?>("ok") == true;
                    version = parsed.Value<string>("version");
                }
                catch { }

                return new
                {
                    surface = "cli",
                    ok,
                    skipped = false,
                    exit_code = exitCode,
                    elapsed_ms = watch.ElapsedMilliseconds,
                    exe = exePath,
                    version,
                    detail = ok
                        ? "CLI 握手成功（version=" + (version ?? "?") + "）。"
                        : "CLI 握手失败（exit=" + exitCode + "）：" + Trim(output, 160)
                };
            }
            catch (Exception ex)
            {
                watch.Stop();
                return new
                {
                    surface = "cli",
                    ok = false,
                    skipped = false,
                    elapsed_ms = watch.ElapsedMilliseconds,
                    exe = exePath,
                    detail = "CLI 启动异常：" + ex.GetType().Name + "（" + Trim(ex.Message, 120) + "）"
                };
            }
        }

        private static object ProbeMcp(string exePath, bool enabled)
        {
            if (!enabled)
            {
                return new { surface = "mcp", ok = false, skipped = true, detail = "MCP 开关已关闭：入口将拒绝启动（exit 6）。" };
            }

            if (exePath == null)
            {
                return new
                {
                    surface = "mcp",
                    ok = false,
                    skipped = false,
                    not_installed = true,
                    detail = "未找到 ExcelAiMcpServer.exe：本安装包未包含 MCP 程序。见下方提示。"
                };
            }

            var watch = Stopwatch.StartNew();
            Process process = null;
            Task<string> lineTask = null;
            Task<string> stderrDrain = null;
            try
            {
                var startInfo = new ProcessStartInfo(exePath, "--probe")
                {
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    // 审查 S2'：显式工作目录，避免子进程继承 Excel CWD 造成相对 DLL 加载歧义
                    WorkingDirectory = Path.GetDirectoryName(exePath) ?? Environment.CurrentDirectory
                };

                process = Process.Start(startInfo);
                if (process == null)
                {
                    return new { surface = "mcp", ok = false, skipped = false, exe = exePath, detail = "MCP 进程启动返回空句柄。" };
                }

                // 读空 stderr 防管道缓冲死锁（审查 A4）
                stderrDrain = process.StandardError.ReadToEndAsync();

                // --probe 模式不预热 Excel：initialize 为纯 ASCII（net48 无 StandardInputEncoding，GBK 的 ASCII 子集与 UTF-8 一致）
                process.StandardInput.WriteLine("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{}}");
                process.StandardInput.Flush();

                lineTask = process.StandardOutput.ReadLineAsync();
                var completed = lineTask.Wait(5000); // 审查 S8：握手 5s 足够，避免长阻塞
                watch.Stop();

                if (!completed || lineTask.Result == null)
                {
                    // 审查 R27：进程"启动即崩"与"太慢"必须区分——否则用户被引导去查网络/重试，
                    // 而真实原因是缺依赖（FileNotFound/BadImageFormat）。这里是唯一能拿到 stderr 的地方。
                    var exited = false;
                    var exitCode = -1;
                    try { exited = process.HasExited; if (exited) { exitCode = process.ExitCode; } } catch { }
                    var errText = "";
                    try { if (stderrDrain != null && stderrDrain.Wait(800)) { errText = (stderrDrain.Result ?? "").Trim(); } } catch { }
                    if (errText.Length > 200) { errText = errText.Substring(0, 200); }
                    errText = errText.Replace("\r", " ").Replace("\n", " ");
                    var detail = exited
                        ? ("MCP 进程启动后立即退出（exit=" + exitCode + "），未能握手。" + (errText.Length > 0 ? " 原始错误：" + errText : ""))
                        : "MCP 握手超时（5s 内未返回 initialize 响应）。";
                    return new
                    {
                        surface = "mcp",
                        ok = false,
                        skipped = false,
                        state = exited ? "start_failed" : "timeout",
                        reason_code = exited ? ("EXIT_" + exitCode) : "TIMEOUT",
                        elapsed_ms = watch.ElapsedMilliseconds,
                        exe = exePath,
                        detail = detail
                    };
                }

                var response = lineTask.Result;
                var ok = response.IndexOf("\"protocolVersion\"", StringComparison.Ordinal) >= 0;
                // 只回显 protocolVersion，避免整段 serverInfo（含构建路径）外泄（审查 S6）
                var protocol = "?";
                try { protocol = JObject.Parse(response)["result"]?["protocolVersion"]?.ToString() ?? "?"; } catch { }
                return new
                {
                    surface = "mcp",
                    ok,
                    skipped = false,
                    elapsed_ms = watch.ElapsedMilliseconds,
                    exe = exePath,
                    protocol_version = protocol,
                    detail = ok ? "MCP 握手成功（protocolVersion=" + protocol + "）。" : "MCP 响应异常，未返回 protocolVersion。"
                };
            }
            catch (Exception ex)
            {
                watch.Stop();
                return new
                {
                    surface = "mcp",
                    ok = false,
                    skipped = false,
                    elapsed_ms = watch.ElapsedMilliseconds,
                    exe = exePath,
                    detail = "MCP 启动异常：" + ex.GetType().Name + "（" + Trim(ex.Message, 120) + "）"
                };
            }
            finally
            {
                // 审查 A5：先发 EOF 让 server 主循环自然退出，再温和等待，最后才 Kill
                try { process?.StandardInput.Close(); } catch { }
                try
                {
                    if (process != null && !process.WaitForExit(2000))
                    {
                        process.Kill();
                        process.WaitForExit(2000);
                    }
                }
                catch { }

                // 审查 B3：进程被 Kill 后读写任务会 faulted——必须 observe，避免未观察任务异常
                Observe(lineTask);
                Observe(stderrDrain);

                try { process?.Dispose(); } catch { }
            }
        }

        private static string RunCapture(string exePath, string arguments, int timeoutMs, out int exitCode)
        {
            exitCode = -1;
            var startInfo = new ProcessStartInfo(exePath, arguments)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using (var process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    return "(启动失败：空句柄)";
                }

                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(timeoutMs))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(2000); // 审查 B4：Kill 是异步语义，必须等句柄真正退出
                    }
                    catch { }

                    // 审查 B4：超时分支同样要观察读写任务，避免未观察异常
                    Observe(stdoutTask);
                    Observe(stderrTask);
                    exitCode = -2; // 未取得退出码（与真实退出码 -1 区分）
                    return "(超时 " + timeoutMs + "ms)";
                }

                exitCode = process.ExitCode;
                var stdout = stdoutTask.Wait(2000) ? stdoutTask.Result : string.Empty;
                var stderr = stderrTask.Wait(2000) ? stderrTask.Result : string.Empty;
                Observe(stdoutTask);
                Observe(stderrTask);
                return string.IsNullOrWhiteSpace(stdout) ? stderr : stdout;
            }
        }

        /// <summary>按固定部署位置查找 MCP/CLI 可执行文件；找不到返回 null（不猜、不伪装）。
        /// 审查 S1/A7：**不做 PATH 搜索、不做源码回溯**——用户可写目录 / 相对 PATH 段 / 同名投毒会让
        /// Excel 以可信进程身份执行任意程序。只认安装目录与显式开发覆盖（环境变量）。</summary>
        /// <summary>安装目录可信判定：必须是本产品自己的安装根（LocalAppData\XaiAssistant 或 Program Files），
        /// 拒绝 TEMP/UNC/用户可写目录——否则 shadow-copy 场景下可被同名 exe 冒充。</summary>
        private static bool IsTrustedInstallDir(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir)) { return false; }
            try
            {
                var full = Path.GetFullPath(dir).TrimEnd('\\');
                if (full.StartsWith(@"\\", StringComparison.Ordinal)) { return false; }
                var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var progFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                var progFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                foreach (var root in new[] { localApp, progFiles, progFilesX86 })
                {
                    if (string.IsNullOrWhiteSpace(root)) { continue; }
                    var r = Path.GetFullPath(root).TrimEnd('\\');
                    if (full.StartsWith(r + "\\", StringComparison.OrdinalIgnoreCase)) { return true; }
                }
                return false;
            }
            catch { return false; }
        }

        private static string ResolveConnectorExecutable(string fileName)
        {
            var candidates = new List<string>();
            try
            {
                // 1) 加载项真实所在目录（Assembly.Location）——Office 托管时 BaseDirectory 是 OFFICE16，
                //    不是我们的安装目录，用它永远找不到 exe（实测缺陷）。
                var archOrder = IntPtr.Size == 8 ? new[] { "x64", "x86" } : new[] { "x86", "x64" };
                try
                {
                    var asmDir = Path.GetDirectoryName(typeof(SidebarBridge).Assembly.Location);
                    // 审查 R25：Assembly.Location 可能是 shadow-copy / TEMP / UNC（可写目录），
                    // 直接把那里当"安装目录"等于让可写目录冒充产品目录。只接受受信根。
                    if (IsTrustedInstallDir(asmDir))
                    {
                        candidates.Add(Path.Combine(asmDir, fileName));
                        foreach (var a in archOrder) { candidates.Add(Path.Combine(asmDir, a, fileName)); }
                    }
                }
                catch { }
                candidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName));

                // 2) 安装目录的两种位数子目录
                var localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (!string.IsNullOrWhiteSpace(localApp))
                {
                    candidates.Add(Path.Combine(localApp, "XaiAssistant", "app", fileName));
                    foreach (var a in archOrder)
                    {
                        candidates.Add(Path.Combine(localApp, "XaiAssistant", "app", a, fileName));
                    }

                    // 4) 独立组件目录（M30：mcp-server / CLI 单独解压包也常放这里，双架构都认）
                    foreach (var root in new[] { "mcp", "cli" })
                    {
                        candidates.Add(Path.Combine(localApp, "XaiAssistant", root, fileName));
                        foreach (var a in archOrder)
                        {
                            candidates.Add(Path.Combine(localApp, "XaiAssistant", root, a, fileName));
                        }
                    }
                }

                // 3) 审查 S1'：已移除 JYYJ_CONNECTOR_DEV_DIR 覆盖——环境变量可指向 UNC/任意盘，
                // 等于给"以 Excel 身份执行任意 exe"开了后门。开发场景请把 exe 复制到安装目录。
            }
            catch { }

            foreach (var candidate in candidates)
            {
                if (!File.Exists(candidate))
                {
                    continue;
                }

                // 审查 S1：执行前校验文件身份——同名但非本产品的程序一律拒绝
                if (!LooksLikeOwnExecutable(candidate, fileName))
                {
                    continue;
                }

                return candidate;
            }

            return null;
        }

        /// <summary>文件身份校验：OriginalFilename 匹配，或 ProductName 含本产品标识；版本信息读取失败则拒绝。</summary>
        private static bool LooksLikeOwnExecutable(string path, string fileName)
        {
            try
            {
                var info = FileVersionInfo.GetVersionInfo(path);
                if (string.Equals(info.OriginalFilename, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                var product = info.ProductName ?? string.Empty;
                return product.IndexOf("ExcelAiAssistant", StringComparison.OrdinalIgnoreCase) >= 0
                    || product.IndexOf("JYYJ", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static List<string> BuildConnectorHints(string cliPath, string mcpPath)
        {
            var hints = new List<string>();
            if (cliPath == null || mcpPath == null)
            {
                // 审查 U1'：去掉重复文案与难懂术语，给普通用户可照做的三步
                hints.Add("没找到命令行程序 ExcelAiCli.exe / ExcelAiMcpServer.exe：当前安装包可能不含 CLI/MCP 组件。");
                hints.Add("首选：重新下载完整安装包覆盖安装。");
                hints.Add("高级替代：先安装 Node.js，再在终端运行 npx jyyj-mcp config（本机没有 exe 时会自动拉便携版）。");
            }
            else
            {
                hints.Add("MCP 客户端配置：把下方片段粘到 %APPDATA%\\Claude\\claude_desktop_config.json 的 mcpServers 下。");
            }

            return hints;
        }

        private static string AddinVersionText()
        {
            try
            {
                var assembly = typeof(SidebarBridge).Assembly;
                var info = assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false);
                return info.Length > 0
                    ? ((System.Reflection.AssemblyInformationalVersionAttribute)info[0]).InformationalVersion
                    : assembly.GetName().Version.ToString();
            }
            catch
            {
                return "?";
            }
        }

        private static string Trim(string text, int max)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "(空)";
            }

            var flattened = text.Replace("\r", " ").Replace("\n", " ");
            return flattened.Length <= max ? flattened : flattened.Substring(0, max) + "…";
        }
    }
}