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
    [ComVisible(true)]
    public sealed partial class SidebarBridge : IDisposable
    {
        private static readonly string RootFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAiAssistant");

        /// <summary>
        /// 自检历史存储（懒加载）：初始化失败（文件被锁/磁盘满等）时降级为 null 并停止记录，
        /// 绝不能让趋势库的故障拖死整个侧边栏构造。
        /// </summary>
        private SelfCheckHistoryStore selfCheckHistoryStore;
        private bool selfCheckHistoryFailed;

        private SelfCheckHistoryStore SelfCheckHistory
        {
            get
            {
                if (selfCheckHistoryStore != null || selfCheckHistoryFailed)
                {
                    return selfCheckHistoryStore;
                }

                try
                {
                    selfCheckHistoryStore = new SelfCheckHistoryStore(Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "ExcelAiAssistant", "selfcheck-history.db"));
                }
                catch (Exception ex)
                {
                    selfCheckHistoryFailed = true;
                    DiagnosticsLog.Error("自检历史库初始化失败，本会话停用趋势记录：" + ex.Message);
                }

                return selfCheckHistoryStore;
            }
        }

        /// <summary>由加载项在 OnConnection 时注入：以指定标题与 URL 弹出右侧停靠网页窗格（与 JYYJ助手同款 CustomTaskPane）。</summary>
        public static Action<string, string> OpenWebDockPane;

        /// <summary>功能区「公式」组回调注入提示词的静态钩子（本实例构造时指向自身 SendPrompt；Dispose 时清空）。</summary>
        public static Func<string, Task<string>> InjectPrompt;

        /// <summary>功能区「更多组」广播视图切换的静态钩子（本实例构造时指向自身 Publish("switch-view", …)；Dispose 时清空）。
        /// payload 为字符串表示目标视图名（如 view-vba/view-functions），为对象时前端按 {view, sql} 联动预填 SQL 模板。</summary>
        public static Action<object> NavigateView;

        /// <summary>
        /// 功能区「JYYJ快捷」复用本侧边栏会话历史/日志通道的宿主工厂（本实例构造时指向自身 CreateHost，Dispose 时清空）。
        /// 触碰式注入：只增不改，复用同一 operationRepository 与 sessionId，保证 Ribbon 触发的操作落入同一回滚历史。
        /// </summary>
        public static Func<dynamic, ExcelToolHost> ExcelHostFactory;

        private readonly dynamic excelApplication;
        private readonly WebView2 webView;
        private readonly ConfigurationStore configurationStore;
        private readonly OperationRepository operationRepository;
        private readonly string sessionId = Guid.NewGuid().ToString("N");
        private readonly Dictionary<string, TaskCompletionSource<bool>> pendingConfirmations =
            new Dictionary<string, TaskCompletionSource<bool>>();
        private readonly List<ChatMessage> conversation = new List<ChatMessage>();
        private readonly List<ReferenceRegion> references = new List<ReferenceRegion>();
        private readonly SessionStore sessionStore;
        private readonly AgentSettingsStore agentSettingsStore;
        private readonly MemoryStore memoryStore;
        private readonly Dictionary<string, Dictionary<string, object>> retryableToolArgs =
            new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        private readonly WorkflowStore workflowStore;
        private readonly ExternalBookRegistry externalBooks;
        private string contextAnchorHandle;
                        private readonly List<ToolEventRecord> lastToolRecords = new List<ToolEventRecord>();
        private string _lastVbaSnippet;
        private CancellationTokenSource activeRunCts;
        private string activeSessionId;
        private string activeProfileId;
        private SkillCatalog skillCatalog;
        private bool skillCatalogLoaded;
        private readonly SkillPackInstaller skillPackInstaller;
        private readonly CustomToolManager customToolManager;
        private readonly WorkflowManager workflowManager;
        private string cachedWriteMode;
        private int cachedMaxToolCalls = 15;
        private bool cachedWebSearchEnabled = true;
        private string activeSummary;

        public SidebarBridge(dynamic excelApplication, WebView2 webView)
        {
            this.excelApplication = excelApplication ?? throw new ArgumentNullException(nameof(excelApplication));
            this.webView = webView ?? throw new ArgumentNullException(nameof(webView));
            InjectPrompt = prompt => SendPrompt(prompt);
            ExcelHostFactory = _ => CreateHost();
            NavigateView = payload => Publish("switch-view", payload);
            configurationStore = new ConfigurationStore(
                Path.Combine(RootFolder, "models.json"),
                ProtectApiKey,
                UnprotectApiKey);
            Directory.CreateDirectory(RootFolder);
            operationRepository = new OperationRepository(Path.Combine(RootFolder, "operations.db"));
            sessionStore = new SessionStore(Path.Combine(RootFolder, "sessions"));
            memoryStore = new MemoryStore(Path.Combine(RootFolder, "memory.json"));
            workflowStore = new WorkflowStore();
            externalBooks = new ExternalBookRegistry(excelApplication);
            agentSettingsStore = new AgentSettingsStore();
            cachedMaxToolCalls = agentSettingsStore.Load().MaxToolCallsPerTurn;
            cachedWebSearchEnabled = agentSettingsStore.Load().WebSearchEnabled;
            skillPackInstaller = new SkillPackInstaller(
                RootFolder,
                LoadSkills,
                ReloadSkills,
                ExecuteAsync,
                () => CreateRegistry().Tools.Select(tool => tool.Name).ToList());
            customToolManager = new CustomToolManager(ExecuteAsync);
            workflowManager = new WorkflowManager(
                ExecuteAsync,
                CreateRegistry,
                CurrentTurnLabel,
                Publish,
                ConfirmTool,
                () => lastToolRecords,
                TryCaptureSelection,
                workflowStore);
        }

        /// <summary>技能目录缓存失效并重载（供 SkillPackInstaller 经委托调用）。</summary>
        private void ReloadSkills()
        {
            skillCatalogLoaded = false;
            try { LoadSkills(); } catch { }
        }

        private bool TryCaptureSelection(out string address, out string sheet)
        {
            address = string.Empty;
            sheet = string.Empty;
            if (excelApplication == null) return false;
            try
            {
                dynamic selection = excelApplication.Selection;
                address = Convert.ToString(selection?.Address) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(address)) return false;
                try
                {
                    dynamic activeSheet = excelApplication.ActiveSheet;
                    sheet = Convert.ToString(activeSheet?.Name) ?? string.Empty;
                }
                catch (Exception)
                {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.cs:177  silent-catch covered (no exception variable)");
                    sheet = string.Empty;
                }
                return true;
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] SidebarBridge.cs:183  silent-catch covered (no exception variable)");
                return false;
            }
        }


        public string Initialize()
        {
            try
            {
                operationRepository.DeleteOlderThan(DateTime.UtcNow.AddDays(-30));
                operationRepository.KeepLatest(CreateWorkbookHash(), 500);
                var profiles = configurationStore.Load();
                var active = ResolveActiveProfile(profiles);
                var restoredTourn = RestoreSession();
                return JsonConvert.SerializeObject(new
                {
                    ok = true,
                    context = SafeGetCurrentContextObject(),
                    history = GetHistoryRecords(),
                    config = active == null ? null : ToConfigObject(active),
                    profiles = profiles.Select(ToProfileObject).Cast<object>().ToArray(),
                    active_profile_id = active?.Id,
                    write_preference = CurrentWriteMode(),
                    max_tool_calls_per_turn = cachedMaxToolCalls,
                    web_search_enabled = cachedWebSearchEnabled,
                    skills = LoadSkills().Select(skill => new
                    {
                        skill.Name,
                        skill.Description,
                        skill.Version,
                        skill.RiskPolicy,
                        skill.Category,
                        skill.Type,
                        skill.Triggers,
                        skill.Examples,
                        skill.Valid,
                        skill.InvalidReason,
                        skill.IsBuiltIn,
                        skill.Source,
                        tools = skill.Tools
                    }).Cast<object>().ToArray(),
                    references = references.Select(ToReferenceObject).Cast<object>().ToArray(),
                    tools = GetToolCatalogObjects(),
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                });
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write("Initialize FAILED: " + exception);
                return Error(exception);
            }
        }

        public Task<string> LoadProfiles()
        {
            return ExecuteAsync(() =>
            {
                var profiles = configurationStore.Load();
                var active = ResolveActiveProfile(profiles);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    profiles = profiles.Select(ToProfileObject).Cast<object>().ToArray(),
                    active_profile_id = active?.Id
                }));
            });
        }

        public Task<string> SaveProfile(string name, string baseUrl, string apiKey, string modelName, string protocolMode, string temperature, string authScheme = "Bearer", bool isNew = false)
        {
            return ExecuteAsync(() =>
            {
                var profiles = configurationStore.Load().ToList();
                var existing = isNew
                    ? null
                    : profiles.FirstOrDefault(item =>
                        string.Equals(item.Id, activeProfileId, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
                var protocol = Enum.TryParse<AgentProtocolMode>(protocolMode, true, out var parsed)
                    ? parsed
                    : AgentProtocolMode.Auto;
                var temperatureValue = double.TryParse(temperature, out var parsedTemperature)
                    ? parsedTemperature
                    : 0.1;

                ModelProfile profile;
                if (existing != null)
                {
                    profiles.Remove(existing);
                    profile = existing;
                }
                else
                {
                    profile = new ModelProfile();
                }

                profile.Name = name;
                profile.BaseUrl = baseUrl;
                profile.ApiKey = string.IsNullOrWhiteSpace(apiKey) ? profile.ApiKey ?? string.Empty : apiKey;
                profile.ModelName = modelName;
                profile.ProtocolMode = protocol;
                profile.AuthScheme = string.Equals(authScheme, "ApiKey", StringComparison.OrdinalIgnoreCase)
                    ? AuthScheme.ApiKey
                    : AuthScheme.Bearer;
                profile.Temperature = temperatureValue;
                profiles.Insert(0, profile);
                configurationStore.Save(profiles);
                activeProfileId = profile.Id;
                PersistActiveProfileId();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    message = "方案已保存。",
                    id = profile.Id,
                    profiles = profiles.Select(ToProfileObject).Cast<object>().ToArray(),
                    active_profile_id = profile.Id
                }));
            });
        }

        public Task<string> DeleteProfile(string id)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    throw new ArgumentException("方案 ID 不能为空。", nameof(id));
                }

                var profiles = configurationStore.Load().ToList();
                profiles.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                if (profiles.Count == 0)
                {
                    throw new InvalidOperationException("至少保留一个模型方案。");
                }

                configurationStore.Save(profiles);
                if (string.Equals(activeProfileId, id, StringComparison.OrdinalIgnoreCase) ||
                    !profiles.Any(item => string.Equals(item.Id, activeProfileId, StringComparison.OrdinalIgnoreCase)))
                {
                    activeProfileId = profiles[0].Id;
                }

                PersistActiveProfileId();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    profiles = profiles.Select(ToProfileObject).Cast<object>().ToArray(),
                    active_profile_id = activeProfileId
                }));
            });
        }

        public Task<string> SetActiveProfile(string id)
        {
            return ExecuteAsync(() =>
            {
                var profiles = configurationStore.Load();
                if (!profiles.Any(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new ArgumentException("指定的模型方案不存在。");
                }

                activeProfileId = id;
                PersistActiveProfileId();
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, active_profile_id = id }));
            });
        }

        public Task<string> SaveConfig(string baseUrl, string apiKey, string modelName)
        {
            return SaveProfile("默认", baseUrl, apiKey, modelName, "Auto", "0.1");
        }

        public Task<string> TestConnection()
        {
            return ExecuteAsync(async () =>
            {
                var profile = RequireProfile();
                using var client = new OpenAiCompatibleClient(profile);
                using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                var connected = await client.TestConnectionAsync(timeoutSource.Token);
                if (!connected)
                {
                    throw new InvalidOperationException("模型接口没有返回有效响应。");
                }

                return JsonConvert.SerializeObject(new { ok = true, message = "连接成功。" });
            });
        }

        public Task<string> GetContext()
        {
            try
            {
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    context = GetCurrentContextObject()
                }));
            }
            catch (Exception exception)
            {
DiagnosticsLog.Write("[LogCover] SidebarBridge.cs:387  {exception.GetType().Name}: {exception.Message}");
                return Task.FromResult(Error(exception));
            }
        }

        /// <summary>
        /// 看板选区（需求①）：列出当前工作簿内全部看板（sheet + 明细数据源），
        /// 供「看板选区」按钮把数据源插入输入框，让用户与 AI 对话时明确看板作用范围。
        /// </summary>
        public Task<string> GetDashboardSources()
        {
            return ExecuteAsync(() =>
            {
                var service = new ExcelAiAssistant.Excel.ExcelDashboardService(excelApplication);
                var result = service.GetDashboardSources(new Dictionary<string, object>());
                if (!result.Ok)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = false, error = Convert.ToString(result.Data) ?? "看板识别失败。" }));
                }

                // 铺平服务层 {dashboard_count, dashboards:[...]}，前端直接读 c.dashboards 数组
                var jo = Newtonsoft.Json.Linq.JObject.FromObject(result.Data);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    dashboard_count = jo.Value<int?>("dashboard_count") ?? 0,
                    dashboards = jo["dashboards"]
                }));
            });
        }

        /// <summary>导出「当前活动工作表最近一张图表」为 PNG data URI，供 AI 回复里直接显示图片；无图表/失败时给出提示。</summary>
        public Task<string> GetLastChartImage()
        {
            return ExecuteAsync(() =>
            {
                var service = new ExcelAiAssistant.Excel.ExcelChartService(excelApplication);
                var result = service.ExportLatestChartAsDataUri(new Dictionary<string, object>());
                if (result.Ok && result.Data != null)
                {
                    var jo = Newtonsoft.Json.Linq.JObject.FromObject(result.Data);
                    var image = jo.Value<string>("chart_image");
                    var note = image == null ? "模型暂不支持图片（当前工作表没有图表，无法生成图片预览）。" : ((jo.Value<string>("note") ?? "") + (jo.Value<bool>("has_data") ? "" : " 注意：该图表疑似无有效数据系列。"));
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = true,
                        chart_image = image,
                        chart_name = jo.Value<string>("chart_name"),
                        chart_type = jo.Value<int?>("chart_type"),
                        series_count = jo.Value<int?>("series_count"),
                        point_count = jo.Value<int?>("point_count"),
                        has_title = jo.Value<bool?>("has_title"),
                        title = jo.Value<string>("title"),
                        has_data = jo.Value<bool?>("has_data"),
                        width = jo.Value<int?>("width"),
                        note = note
                    }));
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = false,
                    chart_image = (string)null,
                    note = "模型暂不支持图片（图表导出失败）。"
                }));
            });
        }

        public Task<string> GetHistory()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                history = GetHistoryRecords()
            })));
        }

        public Task<string> GetSkills()
        {
            return ExecuteAsync(() =>
            {
                var skills = LoadSkills();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    skills = skills.Select(skill => new
                    {
                        skill.Name,
                        skill.Description,
                        skill.Version,
                        skill.RiskPolicy,
                        skill.Category,
                        skill.Type,
                        skill.Triggers,
                        skill.Examples,
                        skill.Valid,
                        skill.InvalidReason,
                        skill.IsBuiltIn,
                        skill.Source,
                        tools = skill.Tools
                    }).Cast<object>().ToArray()
                }));
            });
        }

        public Task<string> AddReference(string address)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(address))
                {
                    throw new ArgumentException("引用地址不能为空。", nameof(address));
                }

                var (sheet, range) = SplitSheetAndRange(address);
                var (rows, columns) = EstimateDimensions(range);
                references.Add(new ReferenceRegion
                {
                    Id = Guid.NewGuid().ToString("N"),
                    Address = address.Trim(),
                    Sheet = sheet,
                    Range = range,
                    RowCount = rows,
                    ColumnCount = columns
                });
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    references = references.Select(ToReferenceObject).Cast<object>().ToArray()
                }));
            });
        }

        public Task<string> RemoveReference(string id)
        {
            return ExecuteAsync(() =>
            {
                references.RemoveAll(item => item.Id == id);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    references = references.Select(ToReferenceObject).Cast<object>().ToArray()
                }));
            });
        }

        public Task<string> NewSession()
        {
            return CreateSession();
        }

        public Task<string> CreateSession()
        {
            return ExecuteAsync(() =>
            {
                PersistSession();
                var key = ResolveSessionKey();
                activeSessionId = sessionStore.Create(key);
                conversation.Clear();
                activeSummary = null;
                references.Clear();
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                }));
            });
        }

        public Task<string> DeleteSession(string sessionIdArg)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(sessionIdArg))
                {
                    throw new ArgumentException("会话不能为空。", nameof(sessionIdArg));
                }

                var key = ResolveSessionKey();
                sessionStore.Delete(key, sessionIdArg);
                if (string.Equals(activeSessionId, sessionIdArg, StringComparison.OrdinalIgnoreCase))
                {
                    conversation.Clear();
                    references.Clear();
                    activeSummary = null;
                    var sessions = ListSessionInfo();
                    if (sessions.Count > 0)
                    {
                        SwitchToSession(sessions[0].Id);
                    }
                    else
                    {
                        activeSessionId = sessionStore.Create(key);
                    }
                }

                _ = key;
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                }));
            });
        }

        public Task<string> SwitchSession(string sessionIdArg)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(sessionIdArg))
                {
                    throw new ArgumentException("会话不能为空。", nameof(sessionIdArg));
                }

                if (!ListSessionInfo().Any(item => string.Equals(item.Id, sessionIdArg, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new ArgumentException("会话不存在。");
                }

                PersistSession();
                SwitchToSession(sessionIdArg);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                }));
            });
        }

        public Task<string> RenameSession(string sessionIdArg, string title)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(sessionIdArg))
                {
                    throw new ArgumentException("会话不能为空。", nameof(sessionIdArg));
                }

                if (!sessionStore.Rename(ResolveSessionKey(), sessionIdArg, title))
                {
                    throw new ArgumentException("重命名失败：会话不存在或标题为空。");
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                }));
            });
        }

        public Task<string> ArchiveSession(string sessionIdArg, bool archived)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(sessionIdArg))
                {
                    throw new ArgumentException("会话不能为空。", nameof(sessionIdArg));
                }

                if (!sessionStore.SetArchived(ResolveSessionKey(), sessionIdArg, archived))
                {
                    throw new ArgumentException("归档失败：会话不存在。");
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    session = GetActiveSessionObject(),
                    sessions = ListSessionObjects()
                }));
            });
        }

        public Task<string> SearchSessions(string keyword)
        {
            return ExecuteAsync(() =>
            {
                var hits = sessionStore.Search(ResolveSessionKey(), keyword ?? string.Empty);
                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    hits = hits.Select(hit => (object)new
                    {
                        hit.SessionId,
                        hit.Title,
                        hit.UpdatedAt,
                        hit.Archived,
                        hit.TitleMatched,
                        hit.Snippet
                    }).ToArray()
                }));
            });
        }

        public Task<string> ListArchivedSessions()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                sessions = ListSessionObjects(true)
            })));
        }

        public Task<string> ListSessions()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                active_session_id = activeSessionId,
                sessions = ListSessionObjects()
            })));
        }

        public Task<string> GetSession()
        {
            return ExecuteAsync(() => Task.FromResult(JsonConvert.SerializeObject(new
            {
                ok = true,
                active_session_id = activeSessionId,
                session = GetActiveSessionObject(),
                sessions = ListSessionObjects()
            })));
        }

        public bool RespondConfirmation(string callId, bool approved)
        {
            TaskCompletionSource<bool> completion;
            lock (pendingConfirmations)
            {
                if (!pendingConfirmations.TryGetValue(callId ?? string.Empty, out completion))
                {
                    return false;
                }

                pendingConfirmations.Remove(callId);
            }

            completion.TrySetResult(approved);
            return true;
        }

        public Task<string> RollbackOperation(string operationId)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrWhiteSpace(operationId))
                {
                    throw new ArgumentException("操作记录不能为空。", nameof(operationId));
                }

                var registry = CreateRegistry();
                if (!registry.TryGet("undo_excel_operation", out var tool))
                {
                    throw new InvalidOperationException("当前宿主不支持回滚。");
                }

                var result = tool.Execute(new Dictionary<string, object>
                {
                    ["operation_id"] = operationId
                });
                if (!result.Ok)
                {
                    throw new InvalidOperationException(result.Data?.ToString() ?? "回滚失败。");
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = true,
                    message = "已回滚指定操作。",
                    history = GetHistoryRecords()
                }));
            });
        }

        public Task<string> RollbackAllOperations()
        {
            return ExecuteAsync(() =>
            {
                var registry = CreateRegistry();
                if (!registry.TryGet("undo_excel_operation", out var tool))
                {
                    throw new InvalidOperationException("当前宿主不支持回滚。");
                }

                var records = operationRepository.List(new OperationQuery
                {
                    WorkbookHash = CreateWorkbookHash(),
                    Limit = 500
                });
                var plan = UndoPlanner.CreatePlan(records);

                var rolledBack = 0;
                var skipped = 0;
                var failed = new List<string>();
                foreach (var record in plan.Order)
                {
                    var result = tool.Execute(new Dictionary<string, object>
                    {
                        ["operation_id"] = record.Id
                    });
                    if (result.Ok)
                    {
                        rolledBack++;
                        continue;
                    }

                    var message = result.Data?.ToString() ?? string.Empty;
                    if (message.IndexOf("冲突", StringComparison.Ordinal) >= 0 ||
                        message.IndexOf("不可回滚", StringComparison.Ordinal) >= 0)
                    {
                        skipped++;
                    }
                    else
                    {
                        failed.Add(record.Id);
                    }
                }

                return Task.FromResult(JsonConvert.SerializeObject(new
                {
                    ok = failed.Count == 0,
                    rolled_back = rolledBack,
                    skipped = skipped,
                    failed = failed.Count,
                    failed_ids = failed.ToArray(),
                    history = GetHistoryRecords()
                }));
            });
        }

    }
}
