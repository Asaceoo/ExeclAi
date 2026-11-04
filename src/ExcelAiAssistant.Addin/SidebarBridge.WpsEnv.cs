using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Skills;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// SidebarBridge 分域 partial（阶段 A 机械拆分，见 docs/AI交接文档.md）。
    /// 公共 API 即 WebView2 前端契约（hostObjects.excelAiBridge.*），签名不得改动。
    /// </summary>
    public sealed partial class SidebarBridge : IDisposable
    {
        public Task<string> GetVersion()
        {
            return ExecuteAsync(() =>
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version?.ToString() ?? "0.0.0.0";
                var info = assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false);
                var informational = info.Length > 0
                    ? ((System.Reflection.AssemblyInformationalVersionAttribute)info[0]).InformationalVersion
                    : version;
                var inExcel = excelApplication != null;
                // 宿主类型（2026-09-08 用户反馈：关于面板须写清当前宿主是 Excel 还是 WPS）：
                // 判定逻辑与 ExcelToolHost.IsWpsHost 同源——WPS 表格进程名为 et.exe；
                // 环境变量 XAI_FORCE_HOST=WPS 可强制判定（harness/测试用）。
                string hostProcess = "EXCEL";
                try { hostProcess = System.Diagnostics.Process.GetCurrentProcess().ProcessName; }
                catch { }
                var isWpsHost = string.Equals(
                    Environment.GetEnvironmentVariable("XAI_FORCE_HOST")?.Trim(),
                    "WPS", System.StringComparison.OrdinalIgnoreCase)
                    || (string.Equals(hostProcess, "et", System.StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(Environment.GetEnvironmentVariable("XAI_FORCE_HOST")?.Trim(),
                            "Excel", System.StringComparison.OrdinalIgnoreCase));
                var skillFolder = System.IO.Path.Combine(RootFolder, "skills");
                var builtInSkillFolder = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(assembly.Location) ?? string.Empty,
                    "skills");
                var skillCount = CountSkillFiles(skillFolder) + CountSkillFiles(builtInSkillFolder);
                string excelVer = null, excelLabel = null;
                if (inExcel)
                {
                    // dynamic COM 读取包一层 try：避免 null 判断后实例断开的极窄窗口直接抛给 UI
                    try { excelVer = Convert.ToString(excelApplication.Version); }
                    catch { excelVer = null; }
                    excelLabel = ExcelVersionLabel(excelVer);
                }
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    version = version,
                    informational = informational,
                    self_check = new
                    {
                        host_kind = isWpsHost ? "WPS" : "Excel",
                        host_process = hostProcess + ".exe",
                        excel_connected = inExcel,
                        excel_version = excelVer,
                        excel_version_label = excelLabel,
                        webview2_loader = System.IO.File.Exists(System.IO.Path.Combine(
                                System.IO.Path.GetDirectoryName(assembly.Location) ?? string.Empty,
                                "WebView2Loader.dll")),
                        skill_files = skillCount > 0
                    }
                }));
            });
        }

        /// <summary>将 Application.Version（15.0/16.0）映射为友好版式名。16.x 各版式（2016/2019/2021/365）无法仅凭 Version 区分，如实际标注。</summary>
        private static string ExcelVersionLabel(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw;
            switch (raw.Split('.')[0].Trim())
            {
                case "12": return "Excel 2007";
                case "14": return "Excel 2010";
                case "15": return "Excel 2013";
                case "16": return "Excel 2016 或更高（含 2019/2021/365）";
                default: return raw;
            }
        }

        /// <summary>统计目录下 SKILL.md 数量（目录不存在返回 0）。</summary>
        private static int CountSkillFiles(string folder)
        {
            return System.IO.Directory.Exists(folder)
                ? System.IO.Directory.GetFiles(folder, "SKILL.md", System.IO.SearchOption.AllDirectories).Length
                : 0;
        }

        /// <summary>当前进程是否为 WPS 表格（et.exe）。</summary>
        private static bool IsWpsProcess()
        {
            try
            {
                return string.Equals(System.Diagnostics.Process.GetCurrentProcess().ProcessName,
                    "et", System.StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>设置页「WPS 环境」卡片数据源（2026-09-08 用户需求：识别本机 WPS 版本与环境支持）。
        /// 探测口径与 register-auto.ps1 同源：三处常见安装根取版本号最高的 office6\et.exe；
        /// VBA 环境判定优先本进程 VBE 直测（WPS 宿主内最准），否则探测 VBA 7.1 运行库 VBE7.dll 落点；
        /// 个人版/专业版按安装根与版本号推断（11.x 企业分发特征 / 12.x 新版通用分发），明确标注为推断。</summary>
        public Task<string> GetWpsEnvironment()
        {
            return ExecuteAsync(() =>
            {
                string etPath = null, etVersion = null, etArch = null, installRoot = null;
                string localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
                var roots = new[]
                {
                    string.IsNullOrEmpty(localAppData) ? null : System.IO.Path.Combine(localAppData, "Kingsoft", "WPS Office"),
                    @"C:\Program Files (x86)\Kingsoft\WPS Office",
                    @"C:\Program Files\Kingsoft\WPS Office"
                };
                foreach (var root in roots)
                {
                    if (string.IsNullOrEmpty(root) || !System.IO.Directory.Exists(root)) continue;
                    var versionDirs = System.IO.Directory.GetDirectories(root)
                        .OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase);
                    foreach (var vdir in versionDirs)
                    {
                        var cand = System.IO.Path.Combine(vdir, "office6", "et.exe");
                        if (System.IO.File.Exists(cand))
                        {
                            etPath = cand;
                            installRoot = System.IO.Path.Combine(vdir, "office6");
                            break;
                        }
                    }
                    if (etPath != null) break;
                }

                if (etPath != null)
                {
                    try { etVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(etPath).FileVersion; }
                    catch { etVersion = null; }
                    try
                    {
                        using (var fs = System.IO.File.OpenRead(etPath))
                        using (var br = new System.IO.BinaryReader(fs))
                        {
                            fs.Seek(0x3C, System.IO.SeekOrigin.Begin);
                            int peOffset = br.ReadInt32();
                            fs.Seek(peOffset + 4, System.IO.SeekOrigin.Begin);
                            ushort machine = br.ReadUInt16();
                            etArch = machine == 0x8664 ? "64 位" : machine == 0x014C ? "32 位" : "未知(0x" + machine.ToString("X4") + ")";
                        }
                    }
                    catch { etArch = null; }
                }

                // VBA 环境：WPS 宿主内直测 VBE 最准；非 WPS 进程（Excel/设计器预览）探测 VBA 7.1 运行库落点
                bool? vbaAvailable = null;
                string vbaEvidence = null;
                if (excelApplication != null && IsWpsProcess())
                {
                    try { vbaAvailable = ((dynamic)excelApplication).VBE != null; }
                    catch { vbaAvailable = false; }
                    vbaEvidence = vbaAvailable == true
                        ? "当前 WPS 进程 VBE 直测：可访问"
                        : "当前 WPS 进程 VBE 直测：不可访问（个人版未装 VBA 支持库的典型特征）";
                }
                else
                {
                    var vbaCandidates = new[]
                    {
                        System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86), "Microsoft Shared", "VBA", "VBA7.1", "VBE7.dll"),
                        System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles), "Microsoft Shared", "VBA", "VBA7.1", "VBE7.dll"),
                    };
                    foreach (var c in vbaCandidates)
                    {
                        if (System.IO.File.Exists(c))
                        {
                            vbaAvailable = true;
                            vbaEvidence = "检测到 VBA 7.1 运行库：" + c;
                            break;
                        }
                    }
                    if (vbaAvailable == null)
                    {
                        vbaAvailable = false;
                        vbaEvidence = "未检测到 VBA 7.1 运行库（VBE7.dll）；若当前正处于 WPS 中，请从 WPS 内打开本面板以获得 VBE 直测结论";
                    }
                }

                // 加载项注册状态（WPS ET Addins 键 LoadBehavior=3 = 启动加载）
                bool addinRegistered = false;
                try
                {
                    using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Kingsoft\Office\ET\Addins\ExcelAiAssistant.AddIn"))
                    {
                        addinRegistered = k != null && Convert.ToInt32(k.GetValue("LoadBehavior", 0)) == 3;
                    }
                }
                catch { addinRegistered = false; }

                // 版本类型推断（明确标注推断，不当结论）
                string editionGuess = null;
                if (etVersion != null)
                {
                    var major = etVersion.Split('.')[0];
                    if (major == "11") editionGuess = "11.x（专业版/企业分发特征，推断）";
                    else if (major == "12") editionGuess = "12.x（新版通用分发，个人版/专业版均有可能，推断）";
                    else editionGuess = etVersion.Split('.')[0] + ".x（未收录的版本系，推断）";
                }

                var payload = new
                {
                    ok = true,
                    host_kind = IsWpsProcess() ? "WPS" : "Excel",
                    wps_installed = etPath != null,
                    et_path = etPath,
                    install_root = installRoot,
                    wps_version = etVersion,
                    wps_arch = etArch,
                    edition_guess = editionGuess,
                    vba_available = vbaAvailable,
                    vba_evidence = vbaEvidence,
                    addin_registered = addinRegistered
                };
                return Task.FromResult(JsonConvert.SerializeObject(payload));
            });
        }

        /// <summary>列出全部已沉淀的自定义工具/模板（供设置页与 VBA 模板库管理）。
        /// category 可选：preset/ai/user，空则返回全部。每次列出前先幂等 seed 预置模板。</summary>

        /// <summary>用户自定义工具清单（按类别过滤）。</summary>
        public Task<string> ListCustomTools(string category = null)
        {
            return customToolManager.ListCustomTools(category);
        }

        /// <summary>保存用户 VBA 模板。</summary>
        public Task<string> SaveUserTemplate(string name, string description, string code, string tagsText)
        {
            return customToolManager.SaveUserTemplate(name, description, code, tagsText);
        }

        /// <summary>删除用户自定义工具。</summary>
        public Task<string> DeleteCustomTool(string name)
        {
            return customToolManager.DeleteCustomTool(name);
        }

        /// <summary>删除 VBA 模板库模板。</summary>
        public Task<string> DeleteVbaTemplate(string name)
        {
            return customToolManager.DeleteVbaTemplate(name);
        }

        /// <summary>VBA 代码脱域（作用域改写为选区引用）。</summary>
        public Task<string> DeScopeVbaCode(string code)
        {
            return customToolManager.DeScopeVbaCode(code);
        }

        /// <summary>更新自定义工具描述。</summary>
        public Task<string> UpdateCustomToolDescription(string name, string description)
        {
            return customToolManager.UpdateCustomToolDescription(name, description);
        }

    }
}
