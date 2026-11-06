using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Skills;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 技能包安装/删除协作者（阶段 B 自 SidebarBridge 抽出）。
    /// busy 发布与异常→JSON 兜底经 executeAsync 委托保持单层结构（Bridge 侧透传，不二次包装）；
    /// 技能目录缓存（skillCatalog）仍在 SidebarBridge，经 loadSkills/reloadSkills 委托读写。
    /// </summary>
    internal sealed class SkillPackInstaller
    {
        private readonly string rootFolder;
        private readonly Func<IReadOnlyList<Skill>> loadSkills;
        private readonly Action reloadSkills;
        private readonly Func<Func<Task<string>>, Task<string>> executeAsync;
        private readonly Func<IReadOnlyList<string>> loadToolNames;

        public SkillPackInstaller(string rootFolder, Func<IReadOnlyList<Skill>> loadSkills, Action reloadSkills, Func<Func<Task<string>>, Task<string>> executeAsync, Func<IReadOnlyList<string>> loadToolNames = null)
        {
            this.rootFolder = rootFolder ?? throw new ArgumentNullException(nameof(rootFolder));
            this.loadSkills = loadSkills ?? throw new ArgumentNullException(nameof(loadSkills));
            this.reloadSkills = reloadSkills ?? throw new ArgumentNullException(nameof(reloadSkills));
            this.executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
            this.loadToolNames = loadToolNames;
        }

        public Task<string> ImportSkillPack(string path)
        {
            return executeAsync(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "技能包文件不存在：" + path }));
                    }

                    var tmp = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N"));
                    try
                    {
                        ExtractSkillPack(path, tmp);
                    }
                    catch (Exception ex)
                    {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:55  {ex.GetType().Name}: {ex.Message}");
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "技能包解压失败：" + ex.Message }));
                    }

                    var files = Directory.GetFiles(tmp, "SKILL.md", SearchOption.AllDirectories);
                    if (files.Length == 0)
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "技能包中未找到 SKILL.md：不是有效技能包。" }));
                    }

                    var skillsRoot = Path.Combine(rootFolder, "skills");
                    Directory.CreateDirectory(skillsRoot);
                    var imported = new List<string>();
                    var overwritten = 0;
                    foreach (var skillFile in files)
                    {
                        var skillDir = Path.GetDirectoryName(skillFile);
                        var name = Path.GetFileName(skillDir);
                        var target = Path.Combine(skillsRoot, name);
                        var isNew = !Directory.Exists(target);
                        Directory.CreateDirectory(target);
                        CopyDirectory(skillDir, target);
                        if (!isNew)
                        {
                            overwritten++;
                        }

                        imported.Add(name);
                    }

                    reloadSkills();
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        imported = imported.ToArray(),
                        overwritten,
                        skills_root = skillsRoot
                    }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:95  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "导入失败：" + ex.Message }));
                }
            });
        }

        /// <summary>安全解压技能包。.NET Framework 4.8 的 ZipFile.ExtractToDirectory 不防护 zip-slip
        /// （条目含 ../ 时写入目标目录之外，可覆盖 %APPDATA% 其他文件）。这里手动逐条解压并校验，
        /// 任何逃逸目标目录的条目一律拒绝。</summary>
        private static void ExtractSkillPack(string archivePath, string target)
        {
            Directory.CreateDirectory(target);
            var targetRoot = Path.GetFullPath(target).TrimEnd('\\', '/') + "\\";
            using (var zip = ZipFile.OpenRead(archivePath))
            {
                foreach (var entry in zip.Entries)
                {
                    var entryName = entry.FullName.Replace('/', '\\');
                    if (string.IsNullOrWhiteSpace(entryName))
                    {
                        continue;
                    }

                    if (entryName.Contains("..\\") || entryName.StartsWith("\\", StringComparison.Ordinal) || entryName.StartsWith("\\\\", StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException("技能包包含非法路径条目：" + entry.FullName);
                    }

                    var dest = Path.Combine(target, entryName);
                    if (!Path.GetFullPath(dest).StartsWith(targetRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException("技能包条目越出目标目录：" + entry.FullName);
                    }

                    if (entryName.EndsWith("\\", StringComparison.Ordinal))
                    {
                        Directory.CreateDirectory(dest);
                        continue;
                    }

                    var parent = Path.GetDirectoryName(dest);
                    if (!string.IsNullOrEmpty(parent))
                    {
                        Directory.CreateDirectory(parent);
                    }

                    entry.ExtractToFile(dest, overwrite: true);
                }
            }
        }

        private static void CopyDirectory(string source, string target)
        {
            var sourceFull = Path.GetFullPath(source).TrimEnd('\\', '/') + "\\";
            foreach (var dir in Directory.GetDirectories(source, "*", SearchOption.AllDirectories))
            {
                var relative = dir.Substring(sourceFull.Length);
                Directory.CreateDirectory(Path.Combine(target, relative));
            }

            foreach (var file in Directory.GetFiles(source, "*", SearchOption.AllDirectories))
            {
                var relative = file.Substring(sourceFull.Length);
                var dest = Path.Combine(target, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(dest));
                File.Copy(file, dest, true);
            }
        }

        /// <summary>粘贴 SKILL.md 文本安装（来源=market）。以 front matter 的 name 取名；与内置技能同名则拒绝。</summary>
        public Task<string> InstallSkillFromText(string markdown)
        {
            return executeAsync(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(markdown))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "粘贴内容为空。" }));
                    }

                    if (!markdown.TrimStart().StartsWith("---", StringComparison.Ordinal))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "粘贴内容不是有效 SKILL.md：必须以 --- front matter --- 开头。" }));
                    }

                    var name = SkillMarketText.ParseSkillName(markdown);
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "SKILL.md 的 front matter 中缺少 name 字段，无法确定技能名。" }));
                    }

                    var installed = InstallWrittenSkill(name, markdown, out var overwritten);
                    reloadSkills();
                    if (installed.Rejected)
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "技能「" + name + "」与产品内置技能同名，已拒绝覆盖。" }));
                    }

                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        imported = new[] { installed.Name },
                        overwritten = overwritten ? 1 : 0,
                        rejected = new string[0],
                        conflicted = false
                    }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:204  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "安装失败：" + ex.Message }));
                }
            });
        }

        /// <summary>从 URL 下载技能包安装（来源=market）：支持 .zip 压缩包或 SKILL.md 直链。</summary>
        public Task<string> InstallSkillFromUrl(string url)
        {
            string tmp = null;
            return executeAsync(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(url) || !Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                        (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "URL 无效，仅支持 http/https。" }));
                    }

                    EnsureModernTls();
                    tmp = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N") + ".dl");
                    using (var client = new WebClient { Encoding = Encoding.UTF8 })
                    {
                        client.Proxy = WebRequest.DefaultWebProxy;
                        client.DownloadFile(uri, tmp);
                    }

                    var head = ReadHead(tmp, 512);
                    var isZip = head.Length >= 2 && head[0] == 'P' && head[1] == 'K';
                    if (isZip)
                    {
                        var extract = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N"));
                        ExtractSkillPack(tmp, extract);
                        var files = Directory.GetFiles(extract, "SKILL.md", SearchOption.AllDirectories);
                        if (files.Length == 0)
                        {
                            return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "下载内容不是有效技能包：未找到 SKILL.md。" }));
                        }

                        var skillsRoot = Path.Combine(rootFolder, "skills");
                        Directory.CreateDirectory(skillsRoot);
                        var imported = new List<string>();
                        var rejected = new List<string>();
                        var overwritten = 0;
                        foreach (var file in files)
                        {
                            var skillName = Path.GetFileName(Path.GetDirectoryName(file));
                            if (string.IsNullOrWhiteSpace(skillName))
                            {
                                continue;
                            }

                            var entry = InstallWrittenSkill(skillName, File.ReadAllText(file, Encoding.UTF8), out var ov);
                            if (entry.Rejected)
                            {
                                rejected.Add(entry.Name);
                            }
                            else
                            {
                                imported.Add(entry.Name);
                                if (ov)
                                {
                                    overwritten++;
                                }
                            }
                        }

                        reloadSkills();
                        return Task.FromResult(JsonConvert.SerializeObject(new
                        {
                            ok = imported.Count > 0,
                            imported = imported.ToArray(),
                            overwritten,
                            rejected = rejected.ToArray(),
                            conflicted = rejected.Count > 0
                        }));
                    }
                    else
                    {
                        var text = File.ReadAllText(tmp, Encoding.UTF8);
                        return Task.FromResult(InstallSkillFromText(text).GetAwaiter().GetResult());
                    }
                }
                catch (WebException ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:289  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "下载失败：" + ex.Message }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:293  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "安装失败：" + ex.Message }));
                }
                finally
                {
                    if (tmp != null)
                    {
                        try { File.Delete(tmp); } catch { }
                    }
                }
            });
        }

        /// <summary>把用户输入解析为 GitHub 仓库坐标并安装（AI Skill P1，FR-S1）：支持 owner/repo、
        /// owner/repo/skill（单仓多技能指定安装）、GitHub 链接与 skills.sh 技能页链接；分支先 main
        /// 后 master 回退。市场标记（source/risk_policy=high）与产品契约声明统一在落盘时注入，
        /// 返回值附每个技能的兼容性三档判定（声明 tools vs 本机注册表）。</summary>
        public Task<string> InstallSkillFromRepo(string input)
        {
            string tmp = null;
            return executeAsync(() =>
            {
                try
                {
                    if (!SkillMarketResolver.TryParse(input, out var target))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new
                        {
                            ok = false,
                            error = "无法识别的仓库输入。支持：owner/repo、owner/repo/技能名、GitHub 仓库链接、skills.sh 技能页链接。"
                        }));
                    }

                    EnsureModernTls();
                    tmp = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N") + ".zip");
                    Exception lastError = null;
                    var downloaded = false;
                    foreach (var url in SkillMarketResolver.BuildCodeloadUrls(target.Owner, target.Repo))
                    {
                        try
                        {
                            using (var client = new WebClient { Encoding = Encoding.UTF8 })
                            {
                                client.Proxy = WebRequest.DefaultWebProxy;
                                client.DownloadFile(url, tmp);
                            }

                            var head = ReadHead(tmp, 2);
                            if (head.Length == 2 && head[0] == 'P' && head[1] == 'K')
                            {
                                downloaded = true;
                                break;
                            }

                            lastError = new InvalidOperationException("下载内容不是 zip（分支可能不存在）。");
                        }
                        catch (Exception ex)
                        {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:350  {ex.GetType().Name}: {ex.Message}");
                            lastError = ex;
                        }
                    }

                    List<string> files = null;
                    if (downloaded)
                    {
                        var extract = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N"));
                        try
                        {
                            ExtractSkillPack(tmp, extract);
                        }
                        catch (Exception ex)
                        {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:364  {ex.GetType().Name}: {ex.Message}");
                            return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "仓库 zip 解压失败：" + ex.Message }));
                        }

                        files = Directory.GetFiles(extract, "SKILL.md", SearchOption.AllDirectories).ToList();
                        if (!string.IsNullOrEmpty(target.Skill))
                        {
                            var match = files.FirstOrDefault(file => string.Equals(
                                Path.GetFileName(Path.GetDirectoryName(file)),
                                target.Skill,
                                StringComparison.OrdinalIgnoreCase));
                            if (match == null)
                            {
                                return Task.FromResult(JsonConvert.SerializeObject(new
                                {
                                    ok = false,
                                    error = "仓库中未找到技能目录「" + target.Skill + "」。单仓多技能仓库请用 owner/repo/技能名 形态指定。"
                                }));
                            }

                            files = new List<string> { match };
                        }
                    }
                    else
                    {
                        // 国内镜像兜底（用户真机实证 codeload 直连 TLS 握手被掐断）：jsDelivr 清单 API + 按文件 CDN 两段式，
                        // 镜像形态不支持整仓 zip，按 flat 清单逐文件重建技能目录（保留 references/scripts 附属文件）。
                        var mirrorNote = string.Empty;
                        foreach (var branch in new[] { "main", "master" })
                        {
                            try
                            {
                                string flat;
                                using (var client = new WebClient { Encoding = Encoding.UTF8 })
                                {
                                    client.Proxy = WebRequest.DefaultWebProxy;
                                    flat = client.DownloadString(SkillMirrorPlan.FlatListingUrl(target.Owner, target.Repo, branch));
                                }

                                List<SkillMirrorFile> plan;
                                try
                                {
                                    plan = SkillMirrorPlan.BuildDownloadPlan(flat, target.Skill);
                                }
                                catch (InvalidOperationException planEx)
                                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:409  {planEx.GetType().Name}: {planEx.Message}");
                                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = planEx.Message }));
                                }

                                var mirrorRoot = Path.Combine(Path.GetTempPath(), "xai-skill-" + Guid.NewGuid().ToString("N"));
                                foreach (var item in plan)
                                {
                                    var dest = Path.Combine(mirrorRoot, item.RemotePath.Replace('/', '\\'));
                                    Directory.CreateDirectory(Path.GetDirectoryName(dest));
                                    using (var client = new WebClient { Encoding = Encoding.UTF8 })
                                    {
                                        client.Proxy = WebRequest.DefaultWebProxy;
                                        client.DownloadFile(SkillMirrorPlan.CdnFileUrl(target.Owner, target.Repo, branch, item.RemotePath), dest);
                                    }
                                }

                                files = plan.Where(item => item.RemotePath.EndsWith("/SKILL.md", StringComparison.OrdinalIgnoreCase))
                                    .Select(item => Path.Combine(mirrorRoot, item.RemotePath.Replace('/', '\\')))
                                    .ToList();
                                if (files.Count > 0)
                                {
                                    downloaded = true;
                                    break;
                                }

                                mirrorNote = branch + " 分支清单中未发现 SKILL.md。";
                            }
                            catch (Exception ex)
                            {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:437  {ex.GetType().Name}: {ex.Message}");
                                mirrorNote = branch + "：" + ex.Message;
                            }
                        }

                        if (!downloaded)
                        {
                            var detail = (mirrorNote ?? string.Empty) +
                                         (lastError == null ? string.Empty : (string.IsNullOrEmpty(mirrorNote) ? "" : "；") + "直连：" + lastError.Message);
                            return Task.FromResult(JsonConvert.SerializeObject(new
                            {
                                ok = false,
                                error = "仓库下载失败（已尝试 codeload 直连与 jsDelivr 镜像，main/master 双分支）：" + (string.IsNullOrEmpty(detail) ? "未知错误" : detail)
                            }));
                        }
                    }

                    if (files == null || files.Count == 0)
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "仓库中未找到 SKILL.md。" }));
                    }

                    // 审查 F2：未指定单技能的超大仓库（如 K-Dense-AI 143 技能）全量安装体量过大，护栏拦截并引导指定形态
                    if (string.IsNullOrEmpty(target.Skill) && files.Count > 20)
                    {
                        var example = Path.GetFileName(Path.GetDirectoryName(files[0]));
                        return Task.FromResult(JsonConvert.SerializeObject(new
                        {
                            ok = false,
                            error = "该仓库包含 " + files.Count + " 个技能，全量安装体量过大。请用 owner/repo/技能名 形态指定安装（例如 " + target.Owner + "/" + target.Repo + "/" + example + "）。"
                        }));
                    }

                    var availableTools = loadToolNames == null ? null : loadToolNames();
                    var imported = new List<string>();
                    var rejected = new List<string>();
                    var overwritten = 0;
                    var compat = new List<object>();
                    var skillsRoot = Path.Combine(rootFolder, "skills");
                    Directory.CreateDirectory(skillsRoot);
                    foreach (var file in files)
                    {
                        var skillDir = Path.GetDirectoryName(file);
                        var skillName = Path.GetFileName(skillDir);
                        if (string.IsNullOrWhiteSpace(skillName))
                        {
                            continue;
                        }

                        var markdown = File.ReadAllText(file, Encoding.UTF8);
                        var judge = SkillCompatJudge.Judge(SkillMarketText.ParseDeclaredTools(markdown), availableTools);
                        if (loadSkills().Any(s => s.IsBuiltIn && string.Equals(s.Name, skillName, StringComparison.OrdinalIgnoreCase)))
                        {
                            rejected.Add(skillName);
                            continue;
                        }

                        // 审查 F1：整目录拷贝保留技能的 references/scripts/assets 附属文件（与 ImportSkillPack 同口径），
                        // SKILL.md 在拷贝后重写以注入市场标记与产品契约声明
                        var targetDir = Path.Combine(skillsRoot, skillName);
                        var ov = Directory.Exists(targetDir);
                        Directory.CreateDirectory(targetDir);
                        CopyDirectory(skillDir, targetDir);
                        File.WriteAllText(
                            Path.Combine(targetDir, "SKILL.md"),
                            SkillMarketText.EnsureMarketSource(markdown) + SkillMarketText.ContractAppendix,
                            new UTF8Encoding(false));
                        imported.Add(skillName);
                        if (ov)
                        {
                            overwritten++;
                        }

                        compat.Add(new { name = skillName, tier = judge.Tier, label = judge.Label, missing = judge.Missing });
                    }

                    reloadSkills();
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = imported.Count > 0,
                        imported = imported.ToArray(),
                        overwritten,
                        rejected = rejected.ToArray(),
                        conflicted = rejected.Count > 0,
                        compat = compat.ToArray()
                    }));
                }
                catch (WebException ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:525  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "下载失败：" + ex.Message }));
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:529  {ex.GetType().Name}: {ex.Message}");
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "安装失败：" + ex.Message }));
                                }
                finally
                {
                    if (tmp != null)
                    {
                        try { File.Delete(tmp); } catch { }
                    }
                }
            });
        }

        /// <summary>删除某个用户/市场技能；内置技能拒绝删除。</summary>
        public Task<string> DeleteSkill(string name)
        {
            return executeAsync(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "技能名不能为空。" }));
                    }

                    if (loadSkills().Any(s => s.IsBuiltIn && string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "内置技能不可删除。" }));
                    }

                    var target = Path.Combine(Path.Combine(rootFolder, "skills"), name);
                    if (!Directory.Exists(target))
                    {
                        return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "未找到技能目录：" + target }));
                    }

                    Directory.Delete(target, true);
                    reloadSkills();
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, deleted = name }));
                }
                catch (Exception ex)
                {
                    DiagnosticsLog.Write("[LogCover] SkillPackInstaller.cs:570  " + ex.GetType().Name + ": " + ex.Message);
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = "删除失败：" + ex.Message }));
                }
            });
        }

        /// <summary>把一份技能内容写入用户技能目录；市场标记（source/risk_policy）与产品契约声明在此统一落盘；返回是否覆盖与是否被内置重名拒绝。</summary>
        private SkillInstallResult InstallWrittenSkill(string name, string markdown, out bool overwritten)
        {
            overwritten = false;
            var skillsRoot = Path.Combine(rootFolder, "skills");
            if (loadSkills().Any(s => s.IsBuiltIn && string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
            {
                return new SkillInstallResult(name, true);
            }

            var targetDir = Path.Combine(skillsRoot, name);
            overwritten = Directory.Exists(targetDir);
            Directory.CreateDirectory(targetDir);
            var marked = SkillMarketText.EnsureMarketSource(markdown) + SkillMarketText.ContractAppendix;
            File.WriteAllText(Path.Combine(targetDir, "SKILL.md"), marked, new UTF8Encoding(false));
            return new SkillInstallResult(name, false);
        }

        /// <summary>宿主进程（Excel AppDomain）内 TLS 协议栈可能停留在旧默认，对 GitHub/jsDelivr 握手报
        /// 「未能创建 SSL/TLS 安全通道」——显式启用 TLS1.2/1.3（一次性全局抬升，只升不降）。</summary>
        private static void EnsureModernTls()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | (SecurityProtocolType)12288 /* Tls13 */;
        }

        /// <summary>把技能内容包上市场标记（source: market），供安装落盘统一口径。</summary>
        private static string EnsureMarketSource(string markdown)
        {
            var text = markdown ?? string.Empty;
            var trimmed = text.TrimStart();
            var name = ParseSkillName(text) ?? "market-skill";
            return "---\nsource: market\nname: " + name + "\n---\n\n" + text;
        }

        /// <summary>快速读取 SKILL.md 文本的首个 /^name\\s*:/ 字段（不依赖完整解析）。</summary>
        private static string ParseSkillName(string text)
        {
            var trimmed = text.TrimStart();
            if (!trimmed.StartsWith("---", StringComparison.Ordinal))
            {
                return null;
            }

            var match = Regex.Match(trimmed, @"(?m)^name\s*:\s*([^\r\n]+)");
            return match.Success ? match.Groups[1].Value.Trim().Trim('"', '\'') : null;
        }

        private static byte[] ReadHead(string path, int count)
        {
            using (var fs = File.OpenRead(path))
            {
                var buffer = new byte[Math.Min(count, (int)Math.Min(fs.Length, int.MaxValue))];
                var read = fs.Read(buffer, 0, buffer.Length);
                return read == buffer.Length ? buffer : buffer.Take(read).ToArray();
            }
        }

        private sealed class SkillInstallResult
        {
            public SkillInstallResult(string name, bool rejected)
            {
                Name = name;
                Rejected = rejected;
            }

            public string Name { get; }

            public bool Rejected { get; }
        }
    }
}
