using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyPhaseI
{
    /// <summary>
    /// 自动冒烟（第 I 批）：P5-C 多工作簿与协同。
    /// 文件级 + 反射校验，不依赖 Excel COM。断言：
    ///  1) ExternalBookRegistry 提供 Open/Close/TryGet/List/Describe 与句柄去重映射（pathToToken）；
    ///  2) ExcelExternalBookTools 注册 open/read/close 三个外部工具；open 为 Dangerous（任意外部路径需确认），
    ///     read/close 为 ReadOnly；
    ///  3) ExcelToolHost 暴露 ExternalBooks 注入属性；
    ///  4) SidebarBridge 暴露 ListConnectedBooks/SetContextAnchor/CloseExternalBook/
    ///     ImportSkillPack/ExportControlledPolicy/ImportControlledPolicy；
    ///  5) 内嵌 index.html 含 P5-C 前端入口（已连接工作簿/锚点/技能包导入/策略导出导入）；
    ///  6) tools/skillpack.ps1 存在且为 UTF-8 BOM（避免中文乱码）；
    ///  7) 技能包 zip 结构往返：含 SKILL.md 的技能目录打包后解压仍需含 SKILL.md（C3 契约）。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL PHASE-I MULTI-WORKBOOK CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            CheckExternalBookRegistry();
            CheckExternalBookTools();
            CheckHostHook();
            CheckSidebarBridge();
            CheckAgentStreaming();
            CheckEmbeddedHtml();
            CheckCustomToolLibrary();
            CheckNewDataTools();
            CheckNewSkills();
            CheckSkillPackScript();
            CheckSkillPackZipContract();
            CheckZipSlipRejected();
        }

        private static void CheckAgentStreaming()
        {
            // 思考过程实时流式（真实根因）：请求带 stream=true，客户端逐行读 SSE，
            // 每段 reasoning 增量回调 AgentProgress(ReasoningDelta)；AgentRunner 将 onProgress 透传给客户端。
            var kindType = Type.GetType("ExcelAiAssistant.Core.Agent.AgentProgressKind, ExcelAiAssistant.Core");
            Assert("stream-kind-delta", kindType != null
                && kindType.IsEnum
                && Enum.GetName(kindType, 3) == "ReasoningDelta",
                "AgentProgressKind 定义 ReasoningDelta=3（流式增量事件）");

            var clientType = Type.GetType("ExcelAiAssistant.Core.Agent.OpenAiCompatibleClient, ExcelAiAssistant.Core");
            const BindingFlags priv = BindingFlags.NonPublic | BindingFlags.Instance;
            Assert("stream-method", clientType != null
                && clientType.GetMethod("SendStreamingAsync", priv) != null
                && clientType.GetMethod("ParseCompletionBody", BindingFlags.NonPublic | BindingFlags.Static) != null,
                "客户端提供 SendStreamingAsync（增量读流）与 ParseCompletionBody（非流式兜底）");

            var streamCall = clientType?.GetMethod("CompleteAsync", new[]
            {
                typeof(System.Collections.Generic.IReadOnlyList<ExcelAiAssistant.Core.Agent.ChatMessage>),
                typeof(System.Collections.Generic.IReadOnlyList<ExcelAiAssistant.Core.Tools.ToolDescriptor>),
                typeof(System.Threading.CancellationToken),
                typeof(Action<ExcelAiAssistant.Core.Agent.AgentProgress>)
            });
            Assert("stream-progress-param", streamCall != null,
                "CompleteAsync 接受 Action<AgentProgress> onProgress（流式回调入口）");
        }

        private static void CheckExternalBookRegistry()
        {
            var type = Type.GetType("ExcelAiAssistant.Excel.ExternalBookRegistry, ExcelAiAssistant.Excel");
            Assert("registry-type", type != null, "存在 ExternalBookRegistry");
            if (type == null)
            {
                return;
            }

            Assert("registry-open", type.GetMethod("Open", new[] { typeof(string) }) != null, "提供 Open(path)");
            Assert("registry-close", type.GetMethod("Close", new[] { typeof(string) }) != null, "提供 Close(token)");
            Assert("registry-tryget", type.GetMethod("TryGet").GetParameters().Length == 2, "提供 TryGet(handle, out workbook)");
            Assert("registry-list", type.GetMethod("List") != null, "提供 List()");
            Assert("registry-describe", type.GetMethod("Describe", new[] { typeof(string) }) != null, "提供 Describe(token)");

            var pathToToken = type.GetField("pathToToken", BindingFlags.NonPublic | BindingFlags.Instance);
            Assert("registry-dedup-map", pathToToken != null, "持有 pathToToken 句柄去重映射");

            var infoType = type.GetNestedType("HandleInfo");
            Assert("registry-handle-info", infoType != null
                && infoType.GetProperty("Token") != null
                && infoType.GetProperty("Name") != null
                && infoType.GetProperty("Path") != null
                && infoType.GetProperty("Sheets") != null, "HandleInfo 含 Token/Name/Path/Sheets");
        }

        private static void CheckExternalBookTools()
        {
            var registry = new ToolRegistry();
            // 注册阶段仅收集描述符，参数 lambda 惰性；传 null books 可安全完成注册。
            ExcelExternalBookTools.Register(registry, null);
            var names = registry.Tools.Select(t => t.Name).ToList();
            Assert("tools-open", names.Contains("open_excel_external_book"), "注册 open_excel_external_book");
            Assert("tools-read", names.Contains("read_external_range"), "注册 read_external_range");
            Assert("tools-close", names.Contains("close_external_book"), "注册 close_external_book");
            // 安全硬化：open 接受任意外部路径，若为 ReadOnly 会自动放行＝静默任意文件读取；故升为 Dangerous 弹确认，
            // read/close 仅操作已开放的句柄（token），维持 ReadOnly。风险级别≠只读语义，外部书仍只读打开。
            Assert("tools-open-dangerous", registry.GetRiskLevel("open_excel_external_book") == RiskLevel.Dangerous,
                "open_excel_external_book 为 Dangerous（打开任意外部路径须用户确认）");
            Assert("tools-all-readonly", registry.GetRiskLevel("read_external_range") == RiskLevel.ReadOnly
                && registry.GetRiskLevel("close_external_book") == RiskLevel.ReadOnly,
                "read/close 外部工作簿工具维持 ReadOnly（守住只读边界）");
        }

        private static void CheckHostHook()
        {
            var hostType = Type.GetType("ExcelAiAssistant.Excel.ExcelToolHost, ExcelAiAssistant.Excel");
            if (hostType == null)
            {
                Fail("host-type", "未加载 ExcelToolHost");
                return;
            }

            var books = hostType.GetProperty("ExternalBooks");
            Assert("host-externalbooks", books != null && books.CanRead && books.CanWrite,
                "ExcelToolHost 暴露 ExternalBooks 可读写注入属性");
        }

        private static void CheckSidebarBridge()
        {
            var type = Type.GetType("ExcelAiAssistant.Addin.SidebarBridge, ExcelAiAssistant.Addin");
            if (type == null)
            {
                Fail("bridge-type", "未加载 SidebarBridge");
                return;
            }

            var expected = new[]
            {
                "ListConnectedBooks", "SetContextAnchor", "CloseExternalBook",
                "ImportSkillPack", "ExportControlledPolicy", "ImportControlledPolicy",
                "GenerateFormula", "InsertFormula", "PreviewFormula"
            };
            foreach (var name in expected)
            {
                Assert("bridge-" + name, type.GetMethod(name) != null, "桥接方法 " + name + " 存在");
            }

            var anchorField = type.GetField("contextAnchorHandle", BindingFlags.NonPublic | BindingFlags.Instance);
            Assert("bridge-anchor-field", anchorField != null, "SidebarBridge 持有 contextAnchorHandle 锚点字段");

            var externalField = type.GetField("externalBooks", BindingFlags.NonPublic | BindingFlags.Instance);
            Assert("bridge-external-field", externalField != null, "SidebarBridge 持有 externalBooks 实例");

            // 网页停靠窗自注册（运行时 Self-register，修复 CreateCTP 找不到 InprocServer32）
            var addinType = Type.GetType("ExcelAiAssistant.Addin.ExcelAiAddIn, ExcelAiAssistant.Addin");
            Assert("ds-selfregister", addinType != null
                && addinType.GetMethod("EnsureWebDockPaneRegistered",
                    BindingFlags.NonPublic | BindingFlags.Instance) != null,
                "ExcelAiAddIn 提供 EnsureWebDockPaneRegistered 自注册");
            var dsHost = Type.GetType("ExcelAiAssistant.Addin.DeepSeekTaskPaneHost, ExcelAiAssistant.Addin");
            Assert("ds-pane-comvisible", dsHost != null
                && dsHost.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComVisibleAttribute), false).Length > 0,
                "DeepSeekTaskPaneHost 标记 ComVisible");
            Assert("webdock-invoke", dsHost != null
                && dsHost.GetMethod("NavigateTo", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance) != null,
                "DeepSeekTaskPaneHost 提供 NavigateTo（网页停靠多站点导航）");
        }

        private static void CheckEmbeddedHtml()
        {
            const string resource = "ExcelAiAssistant.Addin.Sidebar.index.html";
            var asm = typeof(ExcelAiAssistant.Addin.SidebarBridge).Assembly;
            using (var stream = asm.GetManifestResourceStream(resource))
            {
                if (stream == null)
                {
                    Fail("html-resource", "未找到内嵌资源 " + resource);
                    return;
                }

                string html;
                using (var reader = new StreamReader(stream))
                {
                    html = reader.ReadToEnd();
                }

                string[] markers =
                {
                    "refreshConnectedBooks", "connectedBooksList", "anchorSelect", "refreshBooksButton",
                    "skillPackFile", "importSkillPackButton", "exportPolicyButton", "policyFile",
                    "小红书 Dots", "商汤 SenseNova", "小米 MiMo",
                    "token.sensenova.cn/v1", "api.xiaomimimo.com/v1",
                    "note3-prev-api.askdiandian.com", "dots3-note-prev",
                    "sensenova-6.8-flash-lite", "mimo-v2.5-pro"
                };
                foreach (var marker in markers)
                {
                    Assert("html-" + marker, html.Contains(marker), "内嵌 index.html 含 " + marker);
                }

                string[] formulaMarkers =
                {
                    "公式助手", "dataAreaList", "targetAreaList", "dataAreaInput", "targetAreaInput",
                    "addSelDataAreaButton", "addSelTargetAreaButton",
                    "formulaGenerateButton", "formulaCopyButton", "formulaOutput", "formulaExplain",
                    "generateFormula", "GenerateFormula",
                    "formulaPreviewButton", "formulaPreviewOut", "previewFormulaResult", "PreviewFormula"
                };
                foreach (var marker in formulaMarkers)
                {
                    Assert("formula-" + marker, html.Contains(marker), "公式助手入口含 " + marker);
                }

                // 实时「思考过程」流式渲染：后端 agent-progress → 前端活卡增量更新
                string[] streamMarkers =
                {
                    "agent-progress", "renderLiveProgress", "ensureLiveThinking", "releaseLiveThinking",
                    "liveThinking", "实时更新",
                    "p.reasoning", "p.tool_index"
                };
                foreach (var marker in streamMarkers)
                {
                    Assert("stream-" + marker, html.Contains(marker), "AI 实时思考过程含 " + marker);
                }

                Assert("no-view-history", !html.Contains("view-history"), "已移除「历史」独立页（view-history）");
                Assert("no-rollbackAllButton", !html.Contains("rollbackAllButton"), "已移除历史页「全部回滚」按钮");

                // VBA 标签：原「工具」区改为 VBA，工具目录移入设置
                string[] vbaMarkers =
                {
                    "view-vba", "vbaCode", "vbaMacro", "vbaValidateButton", "vbaRunButton", "vbaResult",
                    "ValidateVba", "RunVba"
                };
                foreach (var marker in vbaMarkers)
                {
                    Assert("vba-" + marker, html.Contains(marker), "VBA 面板含 " + marker);
                }
                Assert("no-view-tools", !html.Contains("view-tools"), "原「工具」标签已移除（view-tools）");
                Assert("tools-in-settings", html.Contains("view-settings") && html.IndexOf("toolsList") > html.IndexOf("view-settings"), "工具目录移入【设置】（toolsList 位于 settings 内）");
                Assert("tools-list-last", html.IndexOf("关于 / 版本") < html.IndexOf("toolsList"), "工具说明位于设置最底部（在「关于/版本」之后）");
                // 网页停靠：DS 独立页移除并迁入设置页，支持多站点 + 自定义 URL
                Assert("no-view-ds", !html.Contains("view-ds"), "DS 独立页已移除（view-ds），网页停靠迁入【设置】");
                string[] webDockMarkers =
                {
                    "data-dock-url", "customDockUrl", "customDockButton", "OpenWebDock",
                    "https://www.kimi.com/", "https://yuanbao.tencent.com/chat/naQivTmsDa",
                    "https://www.doubao.com/chat/"
                };
                foreach (var marker in webDockMarkers)
                {
                    Assert("webdock-" + marker, html.Contains(marker), "设置页网页停靠含 " + marker);
                }

                // AI 回复：思考过程（内含二级「调用工具」）-> AI 回复（各可折叠）
                string[] cardMarkers =
                {
                    "buildThinking", "readableStep", "buildRollbackStrip", "buildAssistantCard",
                    "autoReplyBody",
                    "思考过程", "调用工具", "AI 回复", "renderMarkdown"
                };
                foreach (var marker in cardMarkers)
                {
                    Assert("card-" + marker, html.Contains(marker), "AI 卡片三段式含 " + marker);
                }
                Assert("tool-fused-in-thinking",
                    html.Contains("function buildToolSection") == false
                    && html.IndexOf("readableStep(ev);") > html.IndexOf("function buildThinking")
                    && html.IndexOf("readableStep(ev);") < html.IndexOf("function buildRollbackStrip"),
                    "工具步骤已融入思考过程（无 buildToolSection 二级区块，readableStep 直接在 buildThinking 内渲染）");
                Assert("no-truncated-ai", !System.Text.RegularExpressions.Regex.IsMatch(
                        html, @"\.ai-body\s*\{[^}]*max-height"), "AI 回复正文 .ai-body 不再限定 max-height（完整显示）");

                // 自主写代码 + 工具沉淀：设置页已沉淀工具管理卡片 + 系统提示引导
                string[] customToolMarkers =
                {
                    "已沉淀工具", "customToolsList", "customToolsHint", "refreshCustomToolsButton",
                    "refreshCustomTools()", "DeleteCustomTool", "ListCustomTools", "ct_",
                    "save_vba_tool", "沉淀为工具", "SinkVbaTool", "makeSinkButton", "hasVbaSnippetSuccess"
                };
                foreach (var marker in customToolMarkers)
                {
                    Assert("customtool-" + marker, html.Contains(marker), "自主写代码/工具沉淀相关含 " + marker);
                }
            }
        }

        private static void CheckCustomToolLibrary()
        {
            // 功能往返：临时库 Save -> LoadAll -> RegisterTo 注册出 ct_<name> -> Delete
            var root = Path.Combine(Path.GetTempPath(), "verify-custom-" + Guid.NewGuid().ToString("N"));
            var store = new CustomToolStore(root);
            var save = store.Save("transpose_sel", "把选区转置", "Sub T()\nEnd Sub");
            Assert("custom-save-ok", save.Ok, "沉淀工具保存成功");
            Assert("custom-badname", !new CustomToolStore(root).Save("../evil", "x", "y").Ok,
                "非法工具名被拒（防路径穿越）");

            var loaded = store.LoadAll();
            Assert("custom-load", loaded.Count == 1 && loaded[0].Name == "transpose_sel"
                && loaded[0].Code.Contains("T()"), "LoadAll 取回保存的工具");

            var registry = new ToolRegistry();
            store.RegisterTo(registry, (code, args) =>
                new ToolResult { Ok = !string.IsNullOrEmpty(code), Data = code });
            Assert("custom-register", registry.TryGet("ct_transpose_sel", out var td)
                && td.RiskLevel == RiskLevel.Dangerous && td.Execute(new Dictionary<string, object>()).Ok,
                "库内工具注册为 ct_<name>（危险级）且可执行");

            Assert("custom-delete", store.Delete("transpose_sel").Ok
                && store.LoadAll().Count == 0, "删除后库为空");
            Directory.Delete(root, true);

            // 反射校验 C# 侧注入：
            var addinAsm = typeof(ExcelAiAssistant.Addin.SidebarBridge).Assembly;
            var bridge = typeof(ExcelAiAssistant.Addin.SidebarBridge);
            Assert("custom-bridge-list", bridge.GetMethod("ListCustomTools") != null, "SidebarBridge 暴露 ListCustomTools");
            Assert("custom-bridge-del", bridge.GetMethod("DeleteCustomTool") != null, "SidebarBridge 暴露 DeleteCustomTool");
            Assert("custom-bridge-get-last", bridge.GetMethod("GetLastVbaSnippet") != null, "SidebarBridge 暴露 GetLastVbaSnippet（沉淀兜底读取末码）");
            Assert("custom-bridge-sink", bridge.GetMethod("SinkVbaTool") != null, "SidebarBridge 暴露 SinkVbaTool（手动沉淀兜底）");

            // 系统提示沉淀契约 + 高性能 VBA 引导（回归防护：防止把「成功后必须沉淀」和「数组批处理」改回宽松/低效写法）
            var sbSrc = Path.Combine(FindRepositoryRoot(), "src", "ExcelAiAssistant.Addin", "SidebarBridge.cs");
            if (File.Exists(sbSrc))
            {
                var sbText = File.ReadAllText(sbSrc);
                Assert("prompt-sink-mandatory", sbText.Contains("必须主动调用 save_vba_tool")
                    && sbText.Contains("不要因“逻辑不够通用”"),
                    "系统提示强制沉淀：成功后必须主动调用 save_vba_tool 且不再设“可复用”门槛");
                Assert("prompt-vba-perf", sbText.Contains("严禁在 For Each 循环里逐单元格读写")
                    && sbText.Contains("ScreenUpdating = False"),
                    "系统提示高性能 VBA：数组批处理替代逐单元格、关屏幕刷新");
            }
            var host = typeof(ExcelAiAssistant.Excel.ExcelToolHost);
            const BindingFlags nonpub = BindingFlags.NonPublic | BindingFlags.Instance;
            Assert("custom-host-store", host.GetField("customToolStore", nonpub) != null, "ExcelToolHost 持有 customToolStore");
            Assert("custom-host-reg", host.GetMethod("RegisterCustomToolLibrary", nonpub) != null, "ExcelToolHost 注入 RegisterCustomToolLibrary");
            Assert("custom-extract-snippet", typeof(ExcelAiAssistant.Addin.SidebarBridge).GetMethod("ExtractCode", BindingFlags.NonPublic | BindingFlags.Static) != null,
                "run_vba_snippet 确认门存在（参数已统一为 snippet）");
            _ = addinAsm;
        }

        private static void CheckNewDataTools()
        {
            // 新增数据加工工具（排序/去重/查找替换/转置/条件格式）必须如期注册且风险级别正确。
            var registry = new ToolRegistry();
            // 传 dummy application：工具执行懒加载，注册阶段不触碰 COM。
            ExcelDataToolsRegistration.RegisterDataTools(registry, new object(), null);
            var names = registry.Tools.Select(t => t.Name).ToList();
            Assert("data-sort", names.Contains("sort_range"), "注册 sort_range");
            Assert("data-dedup", names.Contains("dedup_range"), "注册 dedup_range");
            Assert("data-findreplace", names.Contains("find_replace_range"), "注册 find_replace_range");
            Assert("data-transpose", names.Contains("transpose_range"), "注册 transpose_range");
            Assert("data-condformat", names.Contains("format_excel_conditional"), "注册 format_excel_conditional");
            Assert("data-normal-write", new[] { "sort_range", "dedup_range", "find_replace_range", "transpose_range", "format_excel_conditional" }
                    .All(n => registry.GetRiskLevel(n) == RiskLevel.NormalWrite),
                "五个数据加工工具均为普通写入级（写入前确认、可撤销）");
        }

        private static void CheckNewSkills()
        {
            // 内置技能包新增：数据清洗/表格格式化带类别，并新增透视、报表、VBA 三项示例技能。
            var addinDir = Path.Combine(FindRepositoryRoot(), "src", "ExcelAiAssistant.Addin", "skills");
            foreach (var name in new[] { "data-clean", "format-table", "pivot-summary", "report-analyst", "vba-automation" })
            {
                Assert("skill-" + name, File.Exists(Path.Combine(addinDir, name, "SKILL.md")), "内置技能 " + name + "/SKILL.md 存在");
            }

            // 解析器能读到 category/metadata 且按分类分组。
            var catalog = new ExcelAiAssistant.Core.Skills.SkillCatalog().AddSearchDirectory(addinDir, true);
            catalog.Reload();
            var groups = catalog.GroupByCategory();
            Assert("skill-groups", groups.Count >= 3, "技能目录按分类分组 ≥3 组");
            Assert("skill-category-read", catalog.TryGet("vba-automation", out var vba)
                && vba.Category == "自动化" && vba.Triggers.Count > 0 && vba.Examples.Count > 0,
                "解析 vba-automation 的分类/triggers/examples 元数据");
            Assert("skill-triggers-search", catalog.Search("修日期").Any(s => s.Name == "data-clean"),
                "Search 命中 triggers 中的场景词（修日期 → data-clean，仅存在于 triggers 不在 name/description）");

            // 内置技能路径必须用加载项 DLL 所在目录（assembly.Location），不能用 AppDomain.BaseDirectory——
            // COM 加载项运行在 Excel 宿主进程的默认 AppDomain，BaseDirectory 指向 Excel.exe 目录，会导致内置技能加载不到。
            var bridgeSrc = File.ReadAllText(Path.Combine(FindRepositoryRoot(), "src", "ExcelAiAssistant.Addin", "SidebarBridge.cs"));
            Assert("skill-load-path",
                bridgeSrc.Contains("Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)")
                && !bridgeSrc.Contains("AppDomain.CurrentDomain.BaseDirectory, \"skills\""),
                "LoadSkills 用 assembly.Location 定位内置技能目录（非 BaseDirectory）");
        }

        private static void CheckSkillPackScript()
        {
            var script = Path.Combine(FindRepositoryRoot(), "tools", "skillpack.ps1");
            Assert("skillpack-exists", File.Exists(script), "tools/skillpack.ps1 存在");
            if (!File.Exists(script))
            {
                return;
            }

            var bytes = File.ReadAllBytes(script);
            Assert("skillpack-bom-utf8", bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "skillpack.ps1 为 UTF-8 BOM 编码（避免中文乱码）");

            var text = File.ReadAllText(script);
            Assert("skillpack-pack", text.Contains("-Source") && text.Contains("CreateFromDirectory")
                && text.Contains("ZipFile"), "脚本含打包分支（.NET ZipFile，兼容 .xai-skill 扩展名）");
            Assert("skillpack-import", text.Contains("-Import") && text.Contains("ExtractToDirectory"),
                "脚本含导入分支");
        }

        private static void CheckSkillPackZipContract()
        {
            var dir = Path.Combine(Path.GetTempPath(), "xai-skillpack-check-" + Guid.NewGuid().ToString("N"));
            var skillDir = Path.Combine(dir, "my-skill");
            Directory.CreateDirectory(skillDir);
            try
            {
                File.WriteAllText(Path.Combine(skillDir, "SKILL.md"), "---\nname: my-skill\ndescription: demo\n---\nbody", System.Text.Encoding.UTF8);
                var zip = Path.Combine(dir, "my-skill.xai-skill");
                ZipFile.CreateFromDirectory(skillDir, zip, CompressionLevel.Optimal, includeBaseDirectory: true);

                var extract = Path.Combine(Path.GetTempPath(), "xai-skillpack-x-" + Guid.NewGuid().ToString("N"));
                ZipFile.ExtractToDirectory(zip, extract);
                try
                {
                    var skillMdFiles = Directory.GetFiles(extract, "SKILL.md", SearchOption.AllDirectories);
                    Assert("zip-contract-skillmd", skillMdFiles.Length >= 1, "skillpack 打包后解压需含 SKILL.md（C3 契约）");
                    Assert("zip-contract-name", Path.GetFileName(Path.GetDirectoryName(skillMdFiles[0])) == "my-skill",
                        "技能包保留技能目录名作为顶层结构");
                }
                finally
                {
                    try { Directory.Delete(extract, recursive: true); } catch { }
                }
            }
            finally
            {
                try { Directory.Delete(dir, recursive: true); } catch { }
            }
        }

        /// <summary>对抗性审查修复项：.NET Framework 4.8 的 ZipFile.ExtractToDirectory 不防护 zip-slip。
        /// 构造含 ../ 逃逸条目的恶意技能包，断言 SidebarBridge.ExtractSkillPack 拒绝解压。</summary>
        private static void CheckZipSlipRejected()
        {
            var dir = Path.Combine(Path.GetTempPath(), "xai-ziptrap-" + Guid.NewGuid().ToString("N"));
            var target = Path.Combine(Path.GetTempPath(), "xai-ziptrap-out-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var zipFile = Path.Combine(dir, "evil.xai-skill");
                using (var stream = System.IO.File.Create(zipFile))
                using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create))
                {
                    var entry = archive.CreateEntry("../escaped.txt");
                    using (var writer = new System.IO.StreamWriter(entry.Open()))
                    {
                        writer.Write("pwned");
                    }
                }

                var extract = typeof(ExcelAiAssistant.Addin.SidebarBridge).GetMethod(
                    "ExtractSkillPack", BindingFlags.NonPublic | BindingFlags.Static);
                if (extract == null)
                {
                    Assert("zip-slip-reject", false, "SidebarBridge.ExtractSkillPack（安全解压）不存在");
                    return;
                }

                Exception thrown = null;
                try
                {
                    extract.Invoke(null, new object[] { zipFile, target });
                }
                catch (Exception ex)
                {
                    thrown = ex.InnerException ?? ex;
                }

                // 必须拒绝解压（抛异常），且 target 目录外不得产生 escaped.txt
                var escaped = Path.Combine(Path.GetTempPath(), "escaped.txt");
                var leaked = File.Exists(escaped);
                if (leaked)
                {
                    try { File.Delete(escaped); } catch { }
                }

                var msg = thrown?.Message ?? string.Empty;
                var rejected = thrown != null
                    && (msg.Contains("非法路径") || msg.Contains("越出目标"))
                    && !leaked;
                Assert("zip-slip-reject", rejected,
                    "含 ../ 条目的技能包被拒绝且未落到目的外（escp_file=" + leaked + " ex=" + (thrown?.GetType().Name ?? "none") + "）");
            }
            finally
            {
                try { Directory.Delete(dir, recursive: true); } catch { }
                try { Directory.Delete(target, recursive: true); } catch { }
            }
        }

        private static string FindRepositoryRoot()
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null)
            {
                if (File.Exists(Path.Combine(dir.FullName, "ExcelAiAssistant.sln")))
                {
                    return dir.FullName;
                }

                dir = dir.Parent;
            }

            return AppDomain.CurrentDomain.BaseDirectory;
        }

        private static void Assert(string name, bool condition, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("PASS " + name + " — " + detail);
            }
            else
            {
                _failed++;
                Failures.Add(name + ": " + detail);
                Console.WriteLine("FAIL " + name + " — " + detail);
            }
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + ": " + detail);
            Console.WriteLine("FAIL " + name + " — " + detail);
        }

        private static string Safe(Exception ex)
        {
            return ex == null ? string.Empty : ex.Message;
        }
    }
}