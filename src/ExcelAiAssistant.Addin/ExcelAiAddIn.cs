using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Diagnostics;
using Extensibility;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;

namespace ExcelAiAssistant.Addin
{
    [ComVisible(true)]
    [Guid("8F7A6D91-4C52-45BE-9F55-7D8C2A31E001")]
    [ProgId("ExcelAiAssistant.AddIn")]
    public sealed class ExcelAiAddIn : IDTExtensibility2, Office.ICustomTaskPaneConsumer, Office.IRibbonExtensibility
    {
        private const string TaskPaneTitle = "JYYJ助手";
        private dynamic excelApplication;
        private object ctpFactoryInstance;
        private Office.CustomTaskPane taskPane;
        private Office.CustomTaskPane webDockPane;
        private DeepSeekTaskPaneHost webDockControl;
        private Office.IRibbonUI ribbonUi;
        private SpotlightService spotlight;
        private RibbonCommandService ribbonCommands;
        private Timer startupCleanupTimer;

        public string GetCustomUI(string ribbonId)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream(
                "ExcelAiAssistant.Addin.Ribbon.xml");
            DiagnosticsLog.Write($"GetCustomUI ribbonId={ribbonId} resourceFound={stream != null}");
            if (stream == null) return null;
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            this.ribbonUi = ribbonUi;
        }

        /// <summary>功能区每按钮自定义图标回调。图标以「按钮Id.png」命名嵌为程序集资源。</summary>
        /// <remarks>按 id 用 Lazy&lt;object&gt; 记忆化缓存同一份 IPictureDisp 并全程持有：ExecutionAndPublication
        /// 保证首次并发下每 id 恰创建一次（无重复创建/悬挂），OLE 侧 fOwn=false 不销毁句柄，Office
        /// 在点击/重绘后反复重取都不会失效，从根上避免「点一次图标就消失」。</remarks>
        private static readonly ConcurrentDictionary<string, Lazy<object>> _ribbonIconCache =
            new ConcurrentDictionary<string, Lazy<object>>();

        public object GetImage(Office.IRibbonControl control)
        {
            var id = control?.Id;
            if (string.IsNullOrEmpty(id)) return null;

            var lazy = _ribbonIconCache.GetOrAdd(
                id,
                static iconId => new Lazy<object>(
                    () => LoadRibbonIcon(iconId),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication));

            try
            {
                return lazy.Value;
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"GetImage failed id={id} {exception.GetType().Name}");
                return null;
            }
        }

        private static object LoadRibbonIcon(string id)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(id + ".png"))
            {
                if (stream == null)
                {
                    DiagnosticsLog.Write($"GetImage missing resource id={id}");
                    return null;
                }

                using (var bitmap = new Bitmap(stream))
                {
                    return RibbonImageConverter.ToPictureDisp(bitmap);
                }
            }
        }

        public void OnToggleTaskPane(Office.IRibbonControl control, bool isPressed)
        {
            DiagnosticsLog.Write($"OnToggleTaskPane id={control?.Id} pressed={isPressed}");
            if (isPressed)
                ShowTaskPane();
            else
                HideTaskPane();
        }

        public bool GetTaskPanePressed(Office.IRibbonControl control)
        {
            if (taskPane == null) return false;
            try { return taskPane.Visible; }
            catch (Exception exception) when (IsStalePane(exception))
            {
                DiagnosticsLog.Write($"GetTaskPanePressed stale pane: {exception.GetType().Name}");
                taskPane = null;
                return false;
            }
        }

        public string GetPaneScreentip(Office.IRibbonControl control)
        {
            return "打开或关闭 JYYJ助手侧边栏";
        }

        public void OnToggleSpotlight(Office.IRibbonControl control, bool isPressed)
        {
            SetViewMode(isPressed ? SpotlightMode.Spotlight : SpotlightMode.Off);
        }

        public bool GetSpotlightPressed(Office.IRibbonControl control)
        {
            return spotlight != null && spotlight.Mode == SpotlightMode.Spotlight;
        }

        public string GetSpotlightScreentip(Office.IRibbonControl control)
        {
            return "高亮活动单元格所在整行整列（阅读模式），再次点击关闭";
        }

        public void OnTogglePointer(Office.IRibbonControl control, bool isPressed)
        {
            SetViewMode(isPressed ? SpotlightMode.Pointer : SpotlightMode.Off);
        }

        public bool GetPointerPressed(Office.IRibbonControl control)
        {
            return spotlight != null && spotlight.Mode == SpotlightMode.Pointer;
        }

        public string GetPointerScreentip(Office.IRibbonControl control)
        {
            return "以边框标记活动单元格位置，再次点击关闭";
        }

        // ================= 数据处理组 =================

        public void OnJQClean(Office.IRibbonControl control) => RunQuickCommand("clean_excel_data");

        public void OnJQDedup(Office.IRibbonControl control) => RunQuickCommand("dedup_range");

        public void OnJQSort(Office.IRibbonControl control) => RunQuickCommand("sort_range");

        public void OnJQTranspose(Office.IRibbonControl control) => RunQuickCommand("transpose_range");

        public void OnJQFilter(Office.IRibbonControl control) => RunQuickCommand("filter_excel_data");

        public void OnJQPivot(Office.IRibbonControl control) => RunQuickCommand("create_excel_pivot_summary");

        public void OnJQSuperBoard(Office.IRibbonControl control)
        {
            if (excelApplication == null)
            {
                DiagnosticsLog.Write("OnJQSuperBoard skipped app=null");
                return;
            }
            if (ribbonCommands == null)
            {
                ribbonCommands = new RibbonCommandService(excelApplication);
            }
            ribbonCommands.ExecuteSuperBoard();
        }

        // ================= 文本处理组 =================

        public void OnJQTextSplit(Office.IRibbonControl control) => RunQuickCommand("text_split");

        public void OnJQTextMerge(Office.IRibbonControl control) => RunQuickCommand("text_merge");

        public void OnJQTextClean(Office.IRibbonControl control) => RunQuickCommand("text_clean");

        public void OnJQFillFormula(Office.IRibbonControl control) => RunQuickCommand("fill_formula");

        // ================= 常用组 =================

        public void OnJQInsertText(Office.IRibbonControl control) => RunQuickCommand("insert_text_into_cell");

        public void OnJQDelText(Office.IRibbonControl control) => RunQuickCommand("delete_text_matches");

        public void OnJQRandNum(Office.IRibbonControl control) => RunQuickCommand("fill_random_numbers");

        public void OnJQSequence(Office.IRibbonControl control) => RunQuickCommand("fill_number_sequence");

        public void OnJQInsertBlank(Office.IRibbonControl control) => RunQuickCommand("insert_blank_rows_every_n");

        public void OnJQDelBlank(Office.IRibbonControl control) => RunQuickCommand("delete_blank_rows_range");

        public void OnJQFillDown(Office.IRibbonControl control) => RunQuickCommand("fill_blank_cells_down");

        public void OnJQMergeDup(Office.IRibbonControl control) => RunQuickCommand("merge_duplicate_cells");

        // ================= 更多组 =================

        /// <summary>宏收纳箱：导航到侧边栏 VBA/Python 模板库（现成的代码库视图）。</summary>
        public void OnJQMacroBox(Office.IRibbonControl control) => SwitchView("view-vba");

        /// <summary>函数库：导航到极简函数列表页（仅浏览与复制用法，UDF 注册属独立机制）。</summary>
        public void OnJQFuncLib(Office.IRibbonControl control) => SwitchView("view-functions");

        /// <summary>SQL 查询：导航到 Python 面板并预填 SQLite 查询模板（读取当前工作簿、选区为表）。</summary>
        public void OnJQSqlQuery(Office.IRibbonControl control) => SwitchView("view-python", true);

        /// <summary>更多组统一入口：先确保面板可见，再经 SidebarBridge 向 WebView2 前端广播视图切换。
        /// 全程在 onAction 的 UI 线程同步执行；prefillSql 时联动前端预填 SQL 模板。</summary>
        private void SwitchView(string viewName, bool prefillSql = false)
        {
            try { ShowTaskPane(); }
            catch (Exception exception) { DiagnosticsLog.Write($"SwitchView ShowTaskPane {exception.GetType().Name}"); }

            var navigate = SidebarBridge.NavigateView;
            if (navigate != null)
            {
                try
                {
                    var payload = prefillSql ? (object)new { view = viewName, sql = true } : viewName;
                    navigate(payload);
                }
                catch (Exception exception) { DiagnosticsLog.Write($"SwitchView publish {exception.GetType().Name}"); }
            }

            if (ribbonUi != null) ribbonUi.InvalidateControl("ToggleTaskPane");
        }

        /// <summary>JYYJ快捷 统一入口：按需懒建 RibbonCommandService，全部逻辑在 onAction 的 UI 线程内同步执行。</summary>
        private void RunQuickCommand(string toolName)
        {
            if (excelApplication == null)
            {
                DiagnosticsLog.Write($"RunQuickCommand skipped app=null tool={toolName}");
                return;
            }

            if (ribbonCommands == null)
            {
                ribbonCommands = new RibbonCommandService(excelApplication);
            }

            ribbonCommands.Execute(toolName);
        }

        private void SetViewMode(SpotlightMode mode)
        {
            if (spotlight == null)
            {
                if (excelApplication == null) return;
                spotlight = new SpotlightService(excelApplication);
            }

            spotlight.SetMode(mode);
            // 聚光灯/指针互斥：开启其一刷新另一 toggle 的按下状态
            if (mode == SpotlightMode.Spotlight) ribbonUi?.InvalidateControl("TogglePointer");
            else if (mode == SpotlightMode.Pointer) ribbonUi?.InvalidateControl("ToggleSpotlight");
        }

        public string GetExplainScreentip(Office.IRibbonControl control)
        {
            return "对选中区公式做逐层中文拆解，并提示关联公式的位置与内容";
        }

        public string GetFillScreentip(Office.IRibbonControl control)
        {
            return "对选中区空白格做语义智能填充（先展示预计结果，确认后再回填）";
        }

        public void OnExplainFormula(Office.IRibbonControl control)
        {
            if (!TryGetRectangularSelection(out var addr, out var rows, out var cols, out var sheetPrefix))
            {
                return;
            }

            var prompt =
                $"请解释当前选区 {sheetPrefix}{addr}（行×列 = {rows}×{cols}）里的公式。\n" +
                "步骤：先调用 explain_excel_formula 工具获取结构化拆解。要求：\n" +
                "1) 若选区公式格超过 20 个，只解释激活格并在结果里提示“选区公式过多，仅解释激活格”；\n" +
                "2) 遇到无公式的单元格，明确标注「这是无公式格」，不要编造解释；\n" +
                "3) 对引用的关联公式，说明它位于哪个工作表、哪个单元格（地址需带工作表名），并简述它是什么公式；\n" +
                "4) 用中文分点逐层解释，让用户一眼看懂结构。";
            InjectRibbonPrompt(prompt);
        }

        public void OnAiFillBlanks(Office.IRibbonControl control)
        {
            if (!TryGetRectangularSelection(out var addr, out var rows, out var cols, out var sheetPrefix))
            {
                return;
            }

            var prompt =
                $"请对选区 {sheetPrefix}{addr}（行×列 = {rows}×{cols}）做数据智能填充。\n" +
                "步骤：\n" +
                "1) 先调用 ai_fill_blank_cells 工具读取该区域并展示「预计结果」（含整行邻近上下文与同列分布参考、以及每个空格的建议值与理由）；has_header 请你根据选区内容自行判断；\n" +
                "2) 找不到语义依据的数值空格，一律填「待补」，不要填 0 或均值；\n" +
                "3) 在聊天里列出各空格的建议值，请用户确认；\n" +
                "4) 用户确认后，再调用 apply_fill_blank_cells 把建议值就地回填。只有空白格会被填充，非空格不会覆盖。";
            InjectRibbonPrompt(prompt);
        }

        /// <summary>读取当前选区。仅当为单块矩形选区时返回 true 并给出地址/行列；否则提示「请选择矩形区域」，不注入。</summary>
        private bool TryGetRectangularSelection(out string address, out int rows, out int cols, out string sheetPrefix)
        {
            address = null;
            rows = 0;
            cols = 0;
            sheetPrefix = null;
            if (excelApplication == null)
            {
                return false;
            }

            try
            {
                dynamic selection = excelApplication.Selection;
                int areas = 0;
                try { areas = Convert.ToInt32(selection.Areas.Count); }
                catch (Exception) { areas = 0; }
                string selectionAddr = Convert.ToString(selection.Address);
                if (areas != 1 || string.IsNullOrWhiteSpace(selectionAddr) || selectionAddr.IndexOf(',') >= 0)
                {
                    InjectRibbonPrompt("请先选择单个矩形区域（当前为非矩形/多段选区），再点击该按钮。");
                    return false;
                }

                var sheetName = Convert.ToString(selection.Worksheet.Name);
                sheetPrefix = QuoteIfNeeded(sheetName) + "!";
                address = selectionAddr;
                rows = Convert.ToInt32(selection.Rows.Count);
                cols = Convert.ToInt32(selection.Columns.Count);
                if (rows <= 0 || cols <= 0)
                {
                    InjectRibbonPrompt("当前选区无效，请重新选择矩形区域。");
                    return false;
                }

                return true;
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"TryGetRectangularSelection failed: {exception.GetType().Name}");
                return false;
            }
        }

        private static string QuoteIfNeeded(string name)
        {
            return name.IndexOfAny(new[] { ' ', '-', '(', ')', '!', '\'', '&', '#' }) >= 0
                ? "'" + name + "'"
                : name;
        }

        private static void InjectRibbonPrompt(string prompt)
        {
            var inject = SidebarBridge.InjectPrompt;
            if (inject == null)
            {
                DiagnosticsLog.Write("InjectRibbonPrompt skipped: inject hook not set (panel not opened)");
                return;
            }

            try
            {
                // Ribbon onAction 本就在 UI 线程执行，SendPrompt 内部需访问 WebView2 的
                // CoreWebView2（仅允许 UI 线程）。在此同步调用并阻塞等待，勿用 Task.Run 丢到
                // 线程池线程 —— 否则 Publish 抛「CoreWebView2 can only be accessed from the UI thread」，
                // 导致按钮点击后毫无反应。
                inject(prompt).GetAwaiter().GetResult();
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"InjectRibbonPrompt failed: {exception}");
            }
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            excelApplication = application ?? throw new ArgumentNullException(nameof(application));
            // SQLite Interop 预加载（2026-11-03 WPS ET 0x8007007E 修复）：
            // CodeBase 加载场景下 System.Data.SQLite 托管层探测 SQLite.Interop.dll 走 AppDomain.BaseDirectory
            // （= 宿主 EXE 目录，WPS ET 为 office6，无 x86 子目录）→ LoadLibrary 失败 0x8007007E，
            // WebView2 初始化回调里首次开 SQLite 连接时弹出「无法加载 SQLite.Interop.dll」。
            // 修复 = 会话最早处用绝对路径预加载与本进程位数匹配的 Interop 模块：
            // 已加载模块按文件名再次请求时直接命中，探测路径不再参与；Excel x64 与 WPS x86 双宿主通用。
            try
            {
                var codebaseDir = Path.GetDirectoryName(new Uri(GetType().Assembly.CodeBase).LocalPath);
                var interopSub = Environment.Is64BitProcess ? "x64" : "x86";
                var interopPath = Path.Combine(codebaseDir, interopSub, "SQLite.Interop.dll");
                if (File.Exists(interopPath))
                {
                    var ptr = Win32Native.LoadLibrary(interopPath);
                    DiagnosticsLog.Write(ptr != IntPtr.Zero
                        ? "SQLite.Interop preload ok path=" + interopPath
                        : "SQLite.Interop preload failed gle=" + Marshal.GetLastWin32Error() + " path=" + interopPath);
                }
                else
                {
                    DiagnosticsLog.Write("SQLite.Interop preload skipped (missing) path=" + interopPath);
                }
            }
            catch (Exception preloadException)
            {
                DiagnosticsLog.Write("SQLite.Interop preload exception: " + preloadException.Message);
            }
            // 供侧边栏桥接弹出任意网页停靠窗格（DeepSeek/Kimi/元宝/豆包/自定义 URL）。
            SidebarBridge.OpenWebDockPane = ShowWebDockPane;
            DeepSeekTaskPaneHost.HidePane = HideWebDockPane;
            DeepSeekTaskPaneHost.OpenSystemBrowser = url => System.Diagnostics.Process.Start(url);
            // 把 Core/Excel 层（run_python_*/工具服务/VBA 通道）的日志全部汇入运行日志，覆盖全功能。
            AppLog.SetSink(entry =>
            {
                var level = entry.Level switch
                {
                    CoreLogLevel.Warn => LogLevel.Warn,
                    CoreLogLevel.Error => LogLevel.Error,
                    _ => LogLevel.Info
                };
                var prefix = string.IsNullOrWhiteSpace(entry.Category) ? string.Empty : $"[{entry.Category}] ";
                DiagnosticsLog.Write(level, prefix + entry.Message);
            });
            DiagnosticsLog.Write($"OnConnection connectMode={connectMode}");
            // M2 宿主检测（WPS 兼容版）：仅留痕与暴露 HostKind，Excel 主线行为零改动（需求文档 N1/D9）。
            HostKind = DetectHostKind();
            string hostProcess;
            try { hostProcess = System.Diagnostics.Process.GetCurrentProcess().ProcessName; }
            catch (Exception hostProcessException) {
DiagnosticsLog.Write($"[LogCover] ExcelAiAddIn.cs:411  {hostProcessException.GetType().Name}: {hostProcessException.Message}");
        hostProcess = hostProcessException.GetType().Name; }
            string hostArch;
            try { hostArch = Environment.Is64BitProcess ? "64 位" : "32 位"; }
            catch { hostArch = "未知"; }
            // 会话首行：版本+宿主+进程位数一次留全，调试时按文件首行即可定位运行环境。
            DiagnosticsLog.Write($"Session start version={GetType().Assembly.GetName().Version} host={HostKind} process={hostProcess} bitness={hostArch} .NET={Environment.Version}");
        }

        /// <summary>当前宿主类型：Excel / WPS / Unknown:*。M2 仅为后续宿主感知过滤（M3）提供依据。</summary>
        internal static string HostKind { get; private set; } = "Unknown";

        private static string DetectHostKind()
        {
            try
            {
                var processName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
                if (string.Equals(processName, "et", StringComparison.OrdinalIgnoreCase))
                {
                    return "WPS";
                }
                if (processName.StartsWith("EXCEL", StringComparison.OrdinalIgnoreCase))
                {
                    return "Excel";
                }
                return "Unknown:" + processName;
            }
            catch (Exception exception)
            {
                return "Unknown:" + exception.GetType().Name;
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            if (startupCleanupTimer != null)
            {
                startupCleanupTimer.Stop();
                startupCleanupTimer.Dispose();
                startupCleanupTimer = null;
            }

            RemoveTaskPane();
            RemoveWebDockPane();
            if (spotlight != null)
            {
                spotlight.Dispose();
                spotlight = null;
            }
            SidebarBridge.OpenWebDockPane = null;
            DeepSeekTaskPaneHost.HidePane = null;
            DeepSeekTaskPaneHost.OpenSystemBrowser = null;
            excelApplication = null;
            ctpFactoryInstance = null;
        }

        public void OnStartupComplete(ref Array custom)
        {
            // 视图组 CF 方案自愈：清掉崩溃/异常退出可能遗留的哨兵 CF 规则与定义名。
            // 崩溃恢复自动重开的工作簿在 OnStartupComplete 之后才打开，故再排一个 15s 单次延迟清理兜底。
            try
            {
                SpotlightService.CleanupAllWorkbooks(excelApplication);
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"OnStartupComplete view cleanup: {exception.GetType().Name}");
            }

            // 恢复上次保存的视图模式（聚光灯/指针）：叠加窗方案零工作簿写入，直接恢复即可
            try
            {
                SpotlightMode savedMode = SpotlightService.GetSavedViewMode();
                if (savedMode != SpotlightMode.Off)
                {
                    DiagnosticsLog.Write($"OnStartupComplete restore view mode={savedMode}");
                    SetViewMode(savedMode);
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"OnStartupComplete restore view: {exception.GetType().Name}");
            }

            try
            {
                if (startupCleanupTimer == null)
                {
                    startupCleanupTimer = new Timer { Interval = 15000 };
                    startupCleanupTimer.Tick += (s, e) =>
                    {
                        startupCleanupTimer.Stop();
                        startupCleanupTimer.Dispose();
                        startupCleanupTimer = null;
                        DiagnosticsLog.Write("delayed view cleanup fired");
                        try
                        {
                            // 用户已手动开启视图模式：规则归主实例管理，跳过盲清理，
                            // 避免删掉正在使用的 CF 规则与 JYYJ_SEL/JYYJ_R/JYYJ_C 定义名
                            // （否则主实例缓存失效 → update selection failed 风暴）。
                            if (spotlight != null && spotlight.Mode != SpotlightMode.Off)
                            {
                                DiagnosticsLog.Write($"delayed view cleanup skipped: mode={spotlight.Mode}");
                                return;
                            }

                            SpotlightService.CleanupAllWorkbooks(excelApplication);
                        }
                        catch (Exception exception) { DiagnosticsLog.Write($"delayed view cleanup: {exception.GetType().Name}"); }
                    };
                    startupCleanupTimer.Start();
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"startup cleanup timer: {exception.GetType().Name}");
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            ctpFactoryInstance = CTPFactoryInst;
            DiagnosticsLog.Write("CTPFactoryAvailable");
            ShowTaskPane();
        }

        public void ShowTaskPane()
        {
            if (taskPane != null)
            {
                try
                {
                    // 窗口归属自愈：CTP 挂在创建时的窗口上，用户切换/新开工作簿窗口后
                    // Visible=true 不报错但物理不可见——归属窗口与当前活动窗口不一致时丢弃重建。
                    if (!IsPaneOnActiveWindow(taskPane))
                    {
                        DiagnosticsLog.Write("ShowTaskPane existing pane on other window -> recreate on active window");
                        try { taskPane.Delete(); }
                        catch (Exception deleteEx) { DiagnosticsLog.Write($"ShowTaskPane old pane delete failed: {deleteEx.GetType().Name}"); }
                        taskPane = null;
                    }
                    else
                    {
                        taskPane.Visible = true;
                        // 宽度自愈：被拖到近零宽度或停靠异常时可见性形同丢失。
                        if (taskPane.Width <= 10)
                        {
                            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                            taskPane.Width = 380;
                            DiagnosticsLog.Write("ShowTaskPane width healed to 380");
                        }
                        DiagnosticsLog.Write("ShowTaskPane existing made visible");
                        return;
                    }
                }
                catch (Exception exception) when (IsStalePane(exception))
                {
                    // 底层 CTP 的 RCW 已失效（隐藏后重开、加载项重载等），丢弃引用并重新创建。
                    DiagnosticsLog.Write($"ShowTaskPane existing pane stale: {exception.GetType().Name}");
                }

                taskPane = null;
            }

            if (ctpFactoryInstance == null || excelApplication == null)
            {
                DiagnosticsLog.Write(
                    $"ShowTaskPane skipped factory={ctpFactoryInstance != null} excel={excelApplication != null}");
                return;
            }

            try
            {
                ExcelAiTaskPaneHost.CurrentExcelApplication = excelApplication;
                var factory = (Office.ICTPFactory)ctpFactoryInstance;
                // 显式挂当前活动窗口（旧实现传 Type.Missing，会挂到创建时刻的活动窗口，
                // 切换工作簿窗口后表现为「点了打开但面板在别的窗口」）。
                object targetWindow = null;
                try { targetWindow = excelApplication.ActiveWindow; }
                catch (Exception windowEx) { DiagnosticsLog.Write($"ShowTaskPane active window probe failed: {windowEx.GetType().Name}"); }
                DiagnosticsLog.Write("CreateCTP starting");
                taskPane = factory.CreateCTP("ExcelAiAssistant.TaskPaneHost", TaskPaneTitle, targetWindow ?? (object)Type.Missing);
                DiagnosticsLog.Write($"CreateCTP completed title={taskPane.Title}");
                taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = 380;
                taskPane.Visible = true;
                DiagnosticsLog.Write("Task pane made visible");
            }
            catch (Exception exception)
            {
                taskPane = null;
                ExcelAiTaskPaneHost.CurrentExcelApplication = null;
                DiagnosticsLog.Write($"ShowTaskPane failed: {exception}");
                throw;
            }
        }

        /// <summary>
        /// 判断既有 CTP 是否挂在当前活动窗口上。拿不到窗口句柄（COM 异常/对象失效）时
        /// 保守返回 true，维持既有「仅 Visible=true」的老行为，不误删可用面板。
        /// </summary>
        private bool IsPaneOnActiveWindow(Office.CustomTaskPane pane)
        {
            try
            {
                var paneWindow = pane.Window;
                object activeWindow = null;
                try { activeWindow = excelApplication.ActiveWindow; }
                catch { return true; }
                if (paneWindow == null || activeWindow == null) return true;
                var paneHwnd = (long)((dynamic)paneWindow).Hwnd;
                var activeHwnd = (long)((dynamic)activeWindow).Hwnd;
                var same = paneHwnd == activeHwnd;
                if (!same) DiagnosticsLog.Write($"Pane window mismatch paneHwnd={paneHwnd} activeHwnd={activeHwnd}");
                return same;
            }
            catch (Exception exception) when (IsStalePane(exception))
            {
                return false;
            }
            catch (Exception binderEx)
            {
                // Hwnd 晚绑定不可用等非 COM 失效场景：保守放行老路径，不误删可用面板。
                DiagnosticsLog.Write($"IsPaneOnActiveWindow fallback: {binderEx.GetType().Name} {binderEx.Message}");
                return true;
            }
        }

        private bool IsStalePane(Exception exception)
        {
            return exception is InvalidComObjectException
                || exception is ObjectDisposedException
                || exception is COMException;
        }

        /// <summary>
        /// 弹出右侧网页停靠窗格（与 XAI 侧边栏同款）。单实例承载：已存在则导航到新地址，
        /// 否则按 title/url 新建。由侧边栏桥接（OpenWebDock）转调，支持 DeepSeek / Kimi / 元宝 / 豆包 / 自定义 URL。
        /// </summary>
        public void ShowWebDockPane(string title, string url)
        {
            var pane = webDockPane;
            var control = webDockControl;
            if (pane != null)
            {
                try
                {
                    pane.Visible = true;
                    // 嵌入式互操作下反射拿不到 .Control（navCtrl 常为 null），切站导航改走
                    // LastInstance：由控件构造赋值，实例一定存在于 CTP 内，导航最可靠。
                    // 同时更新 CurrentUrl，避免窗格因 stale 重建时 _initialUrl 仍取旧站而退回首个站点。
                    if (!string.IsNullOrWhiteSpace(url))
                    {
                        DeepSeekTaskPaneHost.CurrentUrl = url;
                        DeepSeekTaskPaneHost.LastInstance?.NavigateTo(url);
                    }

                    DiagnosticsLog.Write("ShowWebDockPane existing made visible nav=" + url);
                    return;
                }
                catch (Exception exception) when (IsStalePane(exception) || exception is InvalidCastException)
                {
                    DiagnosticsLog.Write($"ShowWebDockPane stale pane: {exception.GetType().Name}");
                }

                webDockPane = null;
                webDockControl = null;
            }

            if (ctpFactoryInstance == null || excelApplication == null)
            {
                DiagnosticsLog.Write(
                    $"ShowWebDockPane skipped factory={ctpFactoryInstance != null} excel={excelApplication != null}");
                return;
            }

            try
            {
                EnsureWebDockPaneRegistered();
                DeepSeekTaskPaneHost.CurrentExcelApplication = excelApplication;
                DeepSeekTaskPaneHost.CurrentUrl =
                    string.IsNullOrWhiteSpace(url) ? "https://chat.deepseek.com/" : url;
                var factory = (Office.ICTPFactory)ctpFactoryInstance;
                var paneTitle = string.IsNullOrWhiteSpace(title) ? "网页停靠" : title.Trim();
                var newPane = factory.CreateCTP("ExcelAiAssistant.DeepSeekPaneHost", paneTitle, Type.Missing);
                webDockPane = newPane;
                newPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                newPane.Width = 420;
                // 反射经 COM 拿控件作后备（可能为 null）；导航主要依赖 LastInstance 与构造读取 CurrentUrl。
                DeepSeekTaskPaneHost navCtrl = null;
                try
                {
                    var ctrl = newPane.GetType().InvokeMember(
                        "Control",
                        BindingFlags.GetProperty,
                        null,
                        newPane,
                        null);
                    navCtrl = ctrl as DeepSeekTaskPaneHost;
                }
                catch (Exception cex) { DiagnosticsLog.Write($"ShowWebDockPane control-as: {cex.GetType().Name}"); }
                newPane.Visible = true;
                webDockControl = navCtrl;
                DiagnosticsLog.Write(
                    $"WebDock pane created title={paneTitle} hasControl={navCtrl != null} nav={DeepSeekTaskPaneHost.CurrentUrl}");
            }
            catch (Exception exception)
            {
                webDockPane = null;
                webDockControl = null;
                DiagnosticsLog.Write($"ShowWebDockPane failed: {exception}");
            }
        }

        private const string DeepSeekProgId = "ExcelAiAssistant.DeepSeekPaneHost";
        private const string DeepSeekClsid = "8F7A6D91-4C52-45BE-9F55-7D8C2A31E003";

        /// <summary>
        /// 自注册 DS 停靠窗格的 COM 类（HKCU\Software\Classes，按当前用户、无需管理员）。
        /// 安装脚本偶尔未能持久写入新类，导致 CreateCTP("ExcelAiAssistant.DeepSeekPaneHost")
        /// 报 0x80004005「无法创建指定的 ActiveX 控件」；在运行中的 Excel（真实交互进程）内
        /// 用 .NET Registry 直写一次可自愈，随后 CreateCTP 即可解析到 InprocServer32。
        /// </summary>
        private void EnsureWebDockPaneRegistered()
        {
            try
            {
                {
                var fullType = "ExcelAiAssistant.Addin.DeepSeekTaskPaneHost";
                var assembly = Assembly.GetExecutingAssembly();
                var assemblyFullName = assembly.FullName;
                var location = assembly.Location;
                if (string.IsNullOrEmpty(location))
                {
                    DiagnosticsLog.Write("EnsureDeepSeekPaneRegistered skipped: no assembly location");
                    return;
                }

                var codeBase = new Uri(location).AbsoluteUri;
                var version = assembly.GetName().Version.ToString();
                var cu = Registry.CurrentUser;

                using (var cls = cu.CreateSubKey($@"Software\Classes\CLSID\{{{DeepSeekClsid}}}"))
                {
                    cls.SetValue("", fullType);
                    cls.SetValue("Class", fullType);
                    cls.SetValue("Assembly", assemblyFullName);
                    cls.SetValue("ProgID", DeepSeekProgId);
                }

                using (var prog = cu.CreateSubKey($@"Software\Classes\CLSID\{{{DeepSeekClsid}}}\ProgId"))
                {
                    prog.SetValue("", DeepSeekProgId);
                }

                cu.CreateSubKey($@"Software\Classes\CLSID\{{{DeepSeekClsid}}}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}").Close();

                using (var inproc = cu.CreateSubKey($@"Software\Classes\CLSID\{{{DeepSeekClsid}}}\InprocServer32"))
                {
                    inproc.SetValue("", "mscoree.dll");
                    inproc.SetValue("ThreadingModel", "Both");
                    inproc.SetValue("Class", fullType);
                    inproc.SetValue("Assembly", assemblyFullName);
                    inproc.SetValue("CodeBase", codeBase);
                    inproc.SetValue("RuntimeVersion", "v4.0.30319");
                }

                using (var ver = cu.CreateSubKey($@"Software\Classes\CLSID\{{{DeepSeekClsid}}}\InprocServer32\{version}"))
                {
                    ver.SetValue("Class", fullType);
                    ver.SetValue("Assembly", assemblyFullName);
                    ver.SetValue("CodeBase", codeBase);
                    ver.SetValue("RuntimeVersion", "v4.0.30319");
                }

                using (var progBase = cu.CreateSubKey($@"Software\Classes\{DeepSeekProgId}"))
                {
                    progBase.SetValue("", fullType);
                }

                using (var progCls = cu.CreateSubKey($@"Software\Classes\{DeepSeekProgId}\CLSID"))
                {
                    progCls.SetValue("", "{" + DeepSeekClsid + "}");
                }

                DiagnosticsLog.Write("EnsureDeepSeekPaneRegistered ok");
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"EnsureDeepSeekPaneRegistered failed: {exception.Message}");
            }
        }

        private void HideWebDockPane()
        {
            if (webDockPane == null)
            {
                return;
            }

            try { webDockPane.Visible = false; }
            catch (Exception exception) when (IsStalePane(exception))
            {
                DiagnosticsLog.Write($"HideWebDockPane stale pane: {exception.GetType().Name}");
                webDockPane = null;
                webDockControl = null;
            }
        }

        private void RemoveWebDockPane()
        {
            if (webDockPane == null)
            {
                return;
            }

            try
            {
                var control = ((dynamic)webDockPane).Control;
                webDockPane.Delete();
                control?.Dispose();
            }
            catch (ObjectDisposedException)
            {
            }
            catch (Exception exception) when (IsStalePane(exception))
            {
                DiagnosticsLog.Write($"RemoveWebDockPane stale pane: {exception.GetType().Name}");
            }
            finally
            {
                webDockPane = null;
                webDockControl = null;
            }
        }

        private void HideTaskPane()
        {
            if (taskPane == null) return;
            try { taskPane.Visible = false; }
            catch (Exception exception) when (IsStalePane(exception))
            {
                DiagnosticsLog.Write($"HideTaskPane stale pane: {exception.GetType().Name}");
                taskPane = null;
            }
        }

        private void RemoveTaskPane()
        {
            if (taskPane == null)
            {
                return;
            }

            try
            {
                var control = ((dynamic)taskPane).Control;
                taskPane.Delete();
                control?.Dispose();
            }
            catch (ObjectDisposedException)
            {
            }
            catch (Exception exception) when (IsStalePane(exception))
            {
                DiagnosticsLog.Write($"RemoveTaskPane stale pane: {exception.GetType().Name}");
            }
            finally
            {
                taskPane = null;
            }
        }
    }

    /// <summary>把 32bpp ARGB 位图转成 OLE IPictureDisp（供功能区 getImage 回调消费）。
    /// 走 PICTYPE_ICON + GetHicon 透传 alpha，避免 GetHbitmap 丢透明导致黑色底。</summary>
    internal static class RibbonImageConverter
    {
        private static readonly Guid IPictureDispIID = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
        private const int PICTYPE_ICON = 3;

        [DllImport("oleaut32.dll")]
        private static extern int OleCreatePictureIndirect(
            ref PICTDESC pPictDesc,
            ref Guid refiid,
            [MarshalAs(UnmanagedType.Bool)] bool fOwn,
            [MarshalAs(UnmanagedType.Interface)] out object ppvRet);

        [StructLayout(LayoutKind.Sequential)]
        private struct PICTDESC
        {
            public int cbSizeOfstruct;
            public int picType;
            public IntPtr hIcon;
            public IntPtr hPal;
            public IntPtr reserved;
            public int unused;

            public static PICTDESC ForIcon(IntPtr hicon)
            {
                return new PICTDESC
                {
                    cbSizeOfstruct = 16,
                    picType = PICTYPE_ICON,
                    hIcon = hicon
                };
            }
        }

        public static object ToPictureDisp(Bitmap bitmap)
        {
            // fOwn=false：OLE 不接管 hicon 生命周期；此 hicon 缓存在 _ribbonIconCache 全程持有，
            // 进程存活期内不销毁（24 个图标，占用可忽略），因此 Office 重绘、重复取图都不会失效。
            // 切勿改为 fOwn=true——Office 释放该 picture 时会销毁 hicon，下一次重取被缓存复用即得到悬空图标。
            var desc = PICTDESC.ForIcon(bitmap.GetHicon());
            var iid = IPictureDispIID; // ref 参数需局部变量，不能直接传静态只读字段
            var hr = OleCreatePictureIndirect(ref desc, ref iid, false, out object pictureDisp);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            return pictureDisp;
        }

    }
    /// <summary>Win32 原生库加载（SQLite Interop 预加载专用）。</summary>
    internal static class Win32Native
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr LoadLibrary(string fileName);
    }

}
