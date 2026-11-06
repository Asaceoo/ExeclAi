using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Tools;

using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Search;

namespace ExcelAiAssistant.Excel
{
    public sealed class ExcelToolHost
    {
        private readonly dynamic application;
        private readonly ExcelContextService contextService;
        private readonly OperationRepository operationRepository;
        private readonly UndoService undoService;
        private readonly string sessionId;

        /// <summary>
        /// 受控执行策略（脚本/进程类工具）。默认关闭以便安全；可在设置中开启并配置白名单。
        /// </summary>
        public ScriptOptions ScriptOptions { get; set; } = new ScriptOptions();

        private readonly CustomToolStore customToolStore = new CustomToolStore(CustomToolStore.DefaultRoot());

        /// <summary>单次写入最大格数：超限直接拒绝（防 COM 大数组封送 + 撤销快照 4 份拷贝导致的 OutOfMemoryException）。</summary>
        private const long MaxSingleWriteCells = 400_000;

        /// <summary>
        /// 写入输出偏好：auto（跟随模型判断）/ formula（公式优先）/ value（直接写值）。
        /// 仅当工具的 value_mode 为 auto（或未指定）时生效，显式指定的 value_mode 仍优先。
        /// </summary>
        public string WriteModePreference { get; set; } = "auto";

        /// <summary>
        /// run_workflow 工具执行体（由加载项桥接注入）。未注入时不注册该工具，避免流程能力缺失时暴露空壳。
        /// </summary>
        public Func<Dictionary<string, object>, ToolResult> RunWorkflowTool { get; set; }

        /// <summary>
        /// 跨工作簿只读引用注册表（P5-C1）。由 SidebarBridge 持有并跨轮注入，保证外部只读句柄多轮复用。
        /// </summary>
        public ExternalBookRegistry ExternalBooks { get; set; }

        public ExcelToolHost(dynamic application)
            : this((object)application, null, null, null)
        {
        }

        public ExcelToolHost(
            dynamic application,
            OperationRepository operationRepository,
            string workbookHash,
            string sessionId,
            string conversationTag = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            contextService = new ExcelContextService(application);
            this.operationRepository = operationRepository;
            this.sessionId = sessionId ?? Guid.NewGuid().ToString("N");

            // 每次新建宿主时从独立 JSON 加载受控执行配置，保证设置页改动即时且重启后仍生效。
            ScriptOptions.Apply(new ControlledExecutionStore().Load());

            if (operationRepository != null)
            {
                undoService = new UndoService(operationRepository, RestoreSnapshot);
                undoService.TurnLabel = conversationTag;
            }
        }

        /// <summary>读取当前受控执行配置（含持久化状态）。</summary>
        public ControlledExecutionConfig GetControlledExecutionConfig()
        {
            return ScriptOptions.ToConfig();
        }

        /// <summary>保存受控执行配置到独立 JSON 并立即应用到当前宿主。</summary>
        public void SaveControlledExecution(ControlledExecutionConfig config)
        {
            if (config == null)
            {
                return;
            }

            new ControlledExecutionStore().Save(config);
            ScriptOptions.Apply(config);
        }
        public ToolRegistry CreateRegistry()
        {
            // M4 宿主感知过滤（WPS 兼容版）：WPS 宿主且无 VBA 工程访问时，剔除 VBA 执行类工具。
            // Excel 宿主 filterVba 恒为 false，路径与主线零差异（需求文档 D9 严格最小范围）。
            var filterVba = IsWpsHost() && !VbaProjectAccessible();
            if (filterVba)
            {
                hostFilterInfo = "WPS-no-VBA: 移除 run_vba_snippet/validate_vba_snippet 与沉淀工具库执行体（需 WPS VBA 环境）";
            }

            var registry = new ToolRegistry();
            // D11 约定：看板工具不进宿主 CreateRegistry（WPS full 巡检基线 144 不含 dashboard 块）。
            // v3.7.1.49 曾在此按宿主补传 dashboard——但 IsWpsHost 按宿主进程名判定，out-of-proc
            // 巡检 harness 进程内判定失效（144→166 亮红），已撤销；功能区超级看板走
            // ExecuteSuperBoard 的表单专用 registry（加载项进程内宿主判定可靠）。
            ExcelToolHostRegistration.RegisterM3Tools(registry, args => CleanData(args), args => FilterData(args), new ExcelStructureService((object)application, RecordStructuralUndo), new ExcelPivotChartService((object)application, RecordStructuralUndo, new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange)));
            ExcelAdvancedToolRegistration.RegisterAdvancedTools(
                registry,
                new ExcelReadService(application),
                new ExcelEditService(application),
                new ExcelScratchpadService(application),
                args => WriteCellsBatch(args),
                RecordStructuralUndo);
            registry.Register(new ToolDescriptor(
                "get_current_excel_context",
                "获取当前工作簿、工作表和选区信息。",
                RiskLevel.ReadOnly,
                _ => Execute(() => new ToolResult { Ok = true, Data = ToSnakeCaseData(contextService.GetCurrentContext()) })));
            registry.Register(new ToolDescriptor(
                "preview_excel_range",
                "预览指定或当前选区的前几行和维度。",
                RiskLevel.ReadOnly,
                args => Execute(() => PreviewRange(args)),
                AdvancedToolSchemas.PreviewRange()));
            registry.Register(new ToolDescriptor(
                "write_excel_cell",
                "向指定单元格写入值、文本或公式。",
                RiskLevel.NormalWrite,
                args => Execute(() => WriteCell(args)),
                AdvancedToolSchemas.WriteCell()));
            registry.Register(new ToolDescriptor(
                "write_excel_range",
                "按区域批量写入二维数组数据。",
                RiskLevel.NormalWrite,
                args => Execute(() => WriteRange(args)),
                AdvancedToolSchemas.WriteRange()));
            registry.Register(new ToolDescriptor(
                "clear_excel_range",
                "清空指定区域的内容（保留格式与批注）。",
                RiskLevel.NormalWrite,
                args => Execute(() => ClearRange(args)),
                AdvancedToolSchemas.ClearRange()));
            registry.Register(new ToolDescriptor(
                "append_excel_rows",
                "在锚点列末行下方追加二维数据（锚点列已有数据时自动向下续写，无需手算起始行）。",
                RiskLevel.NormalWrite,
                args => Execute(() => AppendRowsToSheet(args)),
                AdvancedToolSchemas.AppendRows()));
            registry.Register(new ToolDescriptor(
                "format_excel_range",
                "设置数字格式、字体、填充颜色或自动列宽。",
                RiskLevel.NormalWrite,
                args => Execute(() => FormatRange(args)),
                AdvancedToolSchemas.FormatRange()));
            registry.Register(new ToolDescriptor(
                "undo_excel_operation",
                "按操作记录回滚指定区域写入。",
                RiskLevel.NormalWrite,
                args => Execute(() => UndoOperation(args)),
                AdvancedToolSchemas.UndoOperation()));

            ExcelSecondPhaseRegistration.RegisterSecondPhaseTools(
                registry,
                (object)application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange),
                ScriptOptions);

            ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(
                registry,
                (object)application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange),
                RecordStructuralUndo);

            // 联网搜索工具（v3.8.1.83，只读）：无独立搜索 API——执行时按探测结果分流，
            // 提供方带服务端搜索时由客户端在请求层注入参数（模型自行引用），本工具负责
            // 无服务端搜索提供方（DeepSeek 官方等）的 cn.bing.com 抓取兜底（标注来源类型）。
            registry.Register(new ToolDescriptor(
                "web_search",
                "联网搜索（只读）：当分析需要行业基准/外部参考数据/事实背景（如行业平均良率、汇率、大宗商品价、政策事件），或用户要求搜索时调用。query 只传搜索关键词（禁止把工作簿原始数据当搜索词外发）。结果标注来源类型；引用外部结果时须向用户注明为公开网页信息，不得当作权威统计口径。",
                RiskLevel.ReadOnly,
                args => Execute(() => WebSearch(args)),
                SearchSchemas.WebSearch()));

            ExcelJyyjtoolRegistration.RegisterJyyjtoolTools(
                registry,
                (object)application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange),
                RecordStructuralUndo);

            // P0-2 L2 通用结构操作层 + P1-3 外部数据源（M24 优化路线图）
            var elementOps = new ExcelElementOps((object)application);
            var draftService = new ExcelDraftService((object)application);
            ExcelToolHostRegistration.RegisterElementOps(registry, elementOps, draftService);
            ExcelToolHostRegistration.RegisterExternalData(registry, elementOps, RecordStructuralUndo);
            // FIX（审查 E1）：set_element 改挂宿主 WriteCell 链——获得 undo 快照/操作审计/COM 忙重试，
            // 与 write_excel_cell 同语义（参数集 sheet/address/value 完全一致）。
            if (registry.TryGet("set_element", out var setElementTool))
            {
                registry.Replace(new ToolDescriptor("set_element", setElementTool.Description, RiskLevel.NormalWrite,
                    args => Execute(() => WriteCell(args)), setElementTool.Parameters));
            }

                        ExcelDataToolsRegistration.RegisterDataTools(
                registry,
                application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange));

            ExcelUtilityToolsRegistration.RegisterUtilityTools(
                registry,
                application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange));

            ExcelExplainToolsRegistration.RegisterExplainTools(registry, application);
            ExcelVerifyRegistration.RegisterVerifyTools(registry, (object)application, RecordStructuralUndo);
            // 截图：AI 视觉自检地基（只读，不改单元格）
            ExcelScreenshotToolsRegistration.RegisterScreenshotTools(registry, application);
            ExcelAiFillToolsRegistration.RegisterAiFillTools(
                registry,
                application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange));
            var recordUndo = new Action<string, object, string, object, object>(RecordStructuralUndo);
            ExcelFlagToolsRegistration.RegisterFlagTools(
                registry,
                application,
                new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange),
                recordUndo);

            if (RunWorkflowTool != null)
            {
                registry.Register(new ToolDescriptor(
                    "run_workflow",
                    "执行预定义流程模板，按步骤顺序执行：参数覆盖、危险确认、单步失败可重试。",
                    RiskLevel.NormalWrite,
                    RunWorkflowTool,
                    WorkflowToolSchemas.RunWorkflow()));
            }

            if (ExternalBooks != null)
            {
                ExcelExternalBookTools.Register(registry, ExternalBooks);
            }

            RegisterCustomToolLibrary(registry, registerStoredLibrary: !filterVba);
            ExcelWebToolsRegistration.RegisterWebTools(registry);
            ExcelEnhanceToolsRegistration.RegisterEnhanceTools(
                registry,
                new ExcelEnhanceService(
                    (object)application,
                    new Func<Dictionary<string, object>, object[,], string, string, ToolResult>(WriteGenericRange),
                    RecordStructuralUndo));

            if (filterVba)
            {
                registry.Remove("run_vba_snippet");
                registry.Remove("validate_vba_snippet");
            }

            // M17：写路径工具保护态集中守卫——非 ReadOnly 且未豁免的工具统一包装
            // SHEET_PROTECTED 前置探测（默认全量覆盖，豁免清单见 SheetProtectionGuard）。
            SheetProtectionGuard.WrapWriteTools(registry, (object)application);

            // AI 自主闭环工具（必须在保护包装之后注册）：
            // apply_correction 是组合工具，顶层无统一 sheet 参数；其 attempts 里的每个子工具都已各自被
            // SheetProtectionGuard 包装，保护态判定按子工具自己的 sheet 参数生效——外层再包会误判活动表。
            ExcelSelfCheckToolsRegistration.RegisterSelfCheckTools(registry, application);

            // 本机 AI 客户端接入（detect / connect / test）：不碰工作簿，注册在保护态守卫之后，
            // 避免 SheetProtectionGuard 对无 sheet 参数的工具误判活动表保护态。
            ExcelAiClientToolsRegistration.RegisterAiClientTools(registry);

            return registry;
        }

        /// <summary>宿主感知过滤结果说明；空串表示未触发过滤（Excel 宿主恒为空）。</summary>
        public string HostFilterInfo => hostFilterInfo;

        private string hostFilterInfo = string.Empty;

        /// <summary>当前进程是否为 WPS 表格（et.exe）。环境变量 XAI_FORCE_HOST=WPS 可强制判定（harness/测试用）。</summary>
        private static bool IsWpsHost()
        {
            try
            {
                var forced = Environment.GetEnvironmentVariable("XAI_FORCE_HOST");
                if (!string.IsNullOrWhiteSpace(forced))
                {
                    return string.Equals(forced.Trim(), "WPS", StringComparison.OrdinalIgnoreCase);
                }

                var processName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
                return string.Equals(processName, "et", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>探测宿主是否开放 VBA 工程对象模型访问（WPS 个人版默认无 VBA 环境）。</summary>
        private bool VbaProjectAccessible()
        {
            try
            {
                return ((dynamic)application).VBE != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 注册「沉淀新工具」入口 + 把已沉淀的自定义工具库全量挂到 registry。
        /// save_vba_tool 只负责写库文件；真正执行沉淀代码走 RunCustomStoredTool，与
        /// run_vba_snippet 同链（受控执行开启后才可运行，且执行前仍需确认）。
        /// registerStoredLibrary=false（WPS 无 VBA 环境）时跳过沉淀库注册——沉淀工具全部由 VBA 驱动。
        /// </summary>
        private void RegisterCustomToolLibrary(ToolRegistry registry, bool registerStoredLibrary = true)
        {
            this.customToolStore.SeedPresets(); // 三源之一：内置预置模板（幂等，不覆盖用户改动）

            registry.Register(new ToolDescriptor(
                "save_vba_tool",
                "把已编写并验证有效的 VBA 代码沉淀为可在今后会话直接复用的自定义工具。传入唯一工具名 name（字母开头，字母/数字/下划线，≤40 字符）、用途说明 description（AI 凭此判断何时调用）、完整的 VBA 代码 code。每次你用 run_vba_snippet 成功执行自写代码后都必须调用本工具沉淀，不要因代码“不够通用”而跳过——本工具会先征求用户确认，是否保存由用户决定。",
                RiskLevel.NormalWrite,
                args => Execute(() => this.customToolStore.Save(Required(args, "name"), Required(args, "description"), Required(args, "code"))),
                CustomLibrarySchemas.SaveVbaTool()));

            if (registerStoredLibrary)
            {
                this.customToolStore.RegisterTo(registry, RunCustomStoredTool);
            }
        }

        private ToolResult RunCustomStoredTool(string storedCode, Dictionary<string, object> args)
        {
            var script = new ScriptRunnerService(application, ScriptOptions);
            var packed = new Dictionary<string, object>(args) { ["snippet"] = storedCode };
            return ExcelToolHost.Execute(() => script.RunVbaSnippet(packed));
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            // Y6 修复：缺参必须走 INVALID_ARGS 人话通道（原返回空串会让 save_custom_tool 静默存空值）。
            if (args != null && args.TryGetValue(key, out var value) && value != null)
            {
                var text = Convert.ToString(value);
                if (text.Length > 0) { return text; }
            }

            throw new ArgumentException($"缺少必填参数：{key}");
        }

        /// <summary>
        /// 供结构级工具记录可撤销操作（工作表/行列插入等）。宿主未启用历史时可安全忽略。
        /// </summary>
        public void RecordStructuralUndo(string tool, object args, string undoKind, object affected, object payload)
        {
            if (undoService == null || operationRepository == null)
            {
                return;
            }

            dynamic workbook = null;
            string workbookPath = null;
            try
            {
                workbook = application.ActiveWorkbook;
                workbookPath = Convert.ToString(workbook?.FullName ?? string.Empty);
            }
            finally
            {
                ExcelContextService.Release((object)workbook);
            }

            undoService.CaptureStructural(
                sessionId,
                CreateWorkbookHash(workbookPath),
                workbookPath,
                tool,
                args is Dictionary<string, object> map ? map : new Dictionary<string, object>(),
                undoKind,
                affected,
                payload);
        }

        /// <summary>
        /// 供第二期写类工具复用的通用区域写入回调，包含 undo 快照捕获。
        /// </summary>
        public ToolResult WriteGenericRange(Dictionary<string, object> args, object[,] values, string valueMode, string tool)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                var address = GetOptionalString(args, "address");
                if (string.IsNullOrWhiteSpace(address))
                {
                    return Invalid("INVALID_ADDRESS", "缺少 address。");
                }

                var rows = values.GetLength(0);
                var columns = values.GetLength(1);
                var resized = ExcelRangeHelper.Resize(address, rows, columns);
                if (string.IsNullOrEmpty(resized))
                {
                    return Invalid("INVALID_ADDRESS", $"地址无效：{address}。");
                }

                range = sheet.Range[resized];
                return WriteRangeCore(workbook, sheet, range, values, valueMode, args, tool);
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                // run_python_excel write_ranges 的底层通道：入口阶段 COM 异常同样归因（v3.7.1.62 审查补口）。
                return ComWriteError(valueMode, ce, sheet, range);
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>清空指定区域内容（保留格式与批注）。</summary>
        public ToolResult ClearRange(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                var totalCells = (long)rows * columns;
                if (totalCells > MaxSingleWriteCells)
                {
                    return Invalid("INVALID_RANGE",
                        $"清空区域 {rows} 行 × {columns} 列 = {totalCells} 个单元，超过单次上限 {MaxSingleWriteCells} 格。请给出有限区间（如 A1:A1000），不要使用整列/整行地址。");
                }

                // 统一走 WriteRangeCore（空矩阵清值 + 撤销快照 + operation_id）。
                // 直调 ClearContents 既无格数护栏也无快照，整列清空后无法回滚。
                var cleared = WriteRangeCore(workbook, sheet, range, new object[rows, columns], "value", args, "clear_excel_range");
                if (!cleared.Ok)
                {
                    return cleared;
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheet.Name}!{range.Address}" },
                    Data = new
                    {
                        cleared = true,
                        address = Convert.ToString(range.Address),
                        operation_id = ExcelToolHostHelpers.GetOperationId(cleared)
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                return ComWriteError("clear", ce, sheet, range);
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>在锚点列的末行下方追加二维数据；锚点列上已有数据时自动向下续写。</summary>
        public ToolResult AppendRowsToSheet(Dictionary<string, object> args)
        {
            var startCell = GetString(args, "start_cell");
            // Y2-3 修复：values 缺失原走 args["value"] 抛 KeyNotFoundException → 被收敛为 INVALID_SHEET 误导 AI 查表名。
            // 双兜底皆缺时抛 ArgumentException 走 INVALID_ARGS 人话通道。
            object input;
            if (args.TryGetValue("values", out var rawValues))
            {
                input = rawValues;
            }
            else if (args.TryGetValue("value", out var fallbackValue))
            {
                input = fallbackValue;
            }
            else
            {
                throw new ArgumentException("缺少必填参数：values（或单值参数 value）");
            }
            var rejected = RejectObjectArrayInput(input, "write_excel_cells_batch");
            if (rejected != null)
            {
                return rejected;
            }

            var normalized = NormalizeValues(input);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic anchor = null;
            dynamic range = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                anchor = sheet.Range[startCell];
                var column = Convert.ToInt32(anchor.Column);
                var startRow = Convert.ToInt32(anchor.Row);
                // 从表末向上找锚点列最低已用行，续写在其之后；锚点列空则从锚点行开始。
                dynamic lastCell = null;
                try
                {
                    lastCell = sheet.Cells[Convert.ToInt32(sheet.Rows.Count), column].End(-4162); // xlUp
                    var used = Convert.ToInt32(lastCell.Row);
                    startRow = used >= startRow ? used + 1 : startRow;
                }
                catch
                {
                    // 无法探测末行时从锚点行开始
                }

                var rows = normalized.GetLength(0);
                var columns = normalized.GetLength(1);
                var resized = ExcelRangeHelper.Resize(ExcelRangeHelper.CellAddress(startRow, column), rows, columns);
                if (string.IsNullOrEmpty(resized))
                {
                    return Invalid("INVALID_ADDRESS", $"起始单元格无效：{startCell}。");
                }

                range = sheet.Range[resized];
                return WriteRangeCore(workbook, sheet, range, normalized, "auto", args, "append_excel_rows");
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                return ComWriteError("auto", ce, sheet, range);
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)anchor);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>
        /// 公式助手：把生成的公式写入目标选区右上角单元格（首格）并记录撤销快照。
        /// </summary>
        public ToolResult InsertFormulaToRange(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            var formula = GetString(args, "formula");
            if (string.IsNullOrWhiteSpace(formula))
            {
                return Invalid("INVALID_VALUE", "缺少公式内容。");
            }

            var normalized = formula.TrimStart();
            if (!normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = "=" + normalized;
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic anchor = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[address];
                anchor = range.Cells.Item[1, 1];
                var single = new object[1, 1] { { normalized } };
                return WriteRangeCore(workbook, sheet, anchor, single, "formula", args, "insert_formula_to_range");
            }
            finally
            {
                ExcelContextService.Release((object)anchor);
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        public ToolResult CleanData(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            // operations 与 cells 同为「对象数组」语义：模型经 JSON 传来的是 JArray of JObject，
            // NormalizeValues 会把 JObject 当 IEnumerable 展开（枚举出 JProperty）导致假性失败，
            // 因此用同一归一化路径（NormalizeCellEntries），仅对非 JSON 形态回退旧矩阵路径。
            List<Dictionary<string, object>> operations;
            if (args.TryGetValue("operations", out var rawOperations))
            {
                operations = NormalizeCellEntries(rawOperations);
                if (operations == null)
                {
                    var legacy = NormalizeValues(rawOperations);
                    operations = new List<Dictionary<string, object>>();
                    for (var index = 0; index < legacy.GetLength(0); index++)
                    {
                        if (legacy[index, 0] is Dictionary<string, object> entry)
                        {
                            operations.Add(entry);
                        }
                        else
                        {
                            throw new ArgumentException("operations 必须是对象数组。");
                        }
                    }
                }
            }
            else
            {
                throw new ArgumentException("缺少必填参数：operations");
            }
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                if (rows * (long)columns > MaxSingleWriteCells)
                {
                    return Invalid("TOO_MANY_CELLS", $"数据区域过大（{rows} 行 × {columns} 列），请缩小范围或分批处理。");
                }

                var current = FlattenValues(range.Value2, rows, columns);
                var removedRows = 0;
                var changedCells = 0;
                foreach (var operation in operations)
                {
                    var name = GetOptionalString(operation, "tool") ?? GetOptionalString(operation, "name");
                    switch (name)
                    {
                        case "remove_blank_rows":
                            var removed = DataTransformService.RemoveBlankRows(current);
                            removedRows += removed.ChangedCount;
                            current = removed.Values;
                            break;
                        case "fix_dates":
                            var dateResult = DataTransformService.FixDates(current, ExcelToolHostHelpers.GetInt(operation, "column"), GetOptionalString(operation, "date_format"));
                            changedCells += dateResult.ChangedCount;
                            current = dateResult.Values;
                            break;
                        case "fill_blanks":
                            var filled = DataTransformService.FillBlanks(current, operation.TryGetValue("column", out var column) ? Convert.ToInt32(column) : (int?)null, operation.TryGetValue("value", out var value) ? value : string.Empty);
                            changedCells += filled.ChangedCount;
                            current = filled.Values;
                            break;
                        case "trim":
                            var trimmed = DataTransformService.TrimWhitespace(current);
                            changedCells += trimmed.ChangedCount;
                            current = trimmed.Values;
                            break;
                        case "remove_non_printable":
                            var noControl = DataTransformService.RemoveNonPrintable(current);
                            changedCells += noControl.ChangedCount;
                            current = noControl.Values;
                            break;
                        default:
                            return Invalid("UNSUPPORTED_OPERATION", $"不支持的数据清洗操作：{name}");
                    }
                }

                // 工具分类审查批次修复：全空行清洗后 0 行输出无合法区域可写（Excel 区域最小 1 格），
                // 直接返回人话结果，不写任何内容（此前会经 Resize 钳行写 1 行空值）。
                if (current.GetLength(0) == 0)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            operation_id = (string)null,
                            removed_rows = removedRows,
                            changed_cells = changedCells,
                            output_rows = 0,
                            output_columns = current.GetLength(1),
                            note = "清洗后区域已无内容（如全部为空行），未写入任何内容。"
                        }
                    };
                }

                var resizedAddress = ExcelRangeHelper.Resize(address, current.GetLength(0), current.GetLength(1));
                if (string.IsNullOrEmpty(resizedAddress))
                {
                    return Invalid("INVALID_RANGE", $"清洗结果形状（{current.GetLength(0)} 行 × {current.GetLength(1)} 列）无法映射到合法区域，请缩小范围后重试。");
                }
                var target = sheet.Range[resizedAddress];
                var result = WriteRangeCore(workbook, sheet, target, current, "auto", args, "clean_excel_data");
                ExcelContextService.Release((object)target);
                if (result.Ok)
                {
                    result.Data = new { operation_id = ExcelToolHostHelpers.GetOperationId(result), removed_rows = removedRows, changed_cells = changedCells, output_rows = current.GetLength(0), output_columns = current.GetLength(1) };
                }
                return result;
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        public ToolResult FilterData(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            var columnIndex = ExcelToolHostHelpers.GetInt(args, "column");
            var @operator = GetString(args, "operator");
            var value = GetOptionalString(args, "value");
            var targetAddress = GetOptionalString(args, "target_address");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                if (rows * (long)cols > MaxSingleWriteCells)
                {
                    return Invalid("TOO_MANY_CELLS", $"筛选数据过大（{rows} 行 × {cols} 列），请缩小范围或分批处理。");
                }

                if (columnIndex < 1 || columnIndex > cols)
                {
                    return Invalid("INVALID_COLUMN", $"筛选列号（column={columnIndex}）越界，应在 1..{cols} 之间。");
                }

                var opError = ValidateFilterOperator(@operator, value);
                if (opError != null)
                {
                    return Invalid("INVALID_OPERATOR", opError);
                }

                var source = FlattenValues(range.Value2, rows, cols);
                var filtered = DataTransformService.ApplyFilter(source, columnIndex, @operator, value);
                // 工具分类审查批次修复（数据破坏链）：0 匹配时 Resize 曾被钳成 1 行地址且 Refit 清空语义
                // 会清空源区域首行——0 匹配直接返回人话结果，不写任何内容。
                if (filtered.Values.GetLength(0) == 0)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            matched_rows = 0,
                            output_address = (string)null,
                            note = "筛选条件未匹配到任何行，未写入任何内容。请检查筛选条件（列号/运算符/值）后重试。"
                        }
                    };
                }
                var baseAddress = string.IsNullOrWhiteSpace(targetAddress) ? address : targetAddress;
                var output = ExcelRangeHelper.Resize(baseAddress, filtered.Values.GetLength(0), filtered.Values.GetLength(1));
                var target = sheet.Range[output];
                var result = WriteRangeCore(workbook, sheet, target, filtered.Values, "auto", args, "filter_excel_data");
                ExcelContextService.Release((object)target);
                if (result.Ok)
                {
                    result.Data = new { operation_id = ExcelToolHostHelpers.GetOperationId(result), matched_rows = filtered.MatchedRows, output_address = output };
                }
                return result;
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        public ToolResult FormatRange(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            var skipped = new List<string>();
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[GetString(args, "address")];

                if (args.TryGetValue("number_format", out var fmt) && !string.IsNullOrEmpty(Convert.ToString(fmt)))
                {
                    try { range.NumberFormat = Convert.ToString(fmt); }
                    catch { skipped.Add("number_format"); }
                }

                if (args.TryGetValue("bold", out var bold))
                {
                    if (TryParseBool(bold, out var b)) range.Font.Bold = b;
                    else skipped.Add("bold");
                }

                if (args.TryGetValue("italic", out var italic))
                {
                    if (TryParseBool(italic, out var it)) range.Font.Italic = it;
                    else skipped.Add("italic");
                }

                if (args.TryGetValue("font_color", out var fc) && TryParseColor(fc, out var fci)) range.Font.Color = fci;
                else if (args.ContainsKey("font_color")) skipped.Add("font_color");

                if (args.TryGetValue("back_color", out var bc) && TryParseColor(bc, out var bci)) range.Interior.Color = bci;
                else if (args.ContainsKey("back_color")) skipped.Add("back_color");

                if (args.TryGetValue("auto_fit_columns", out var af) && TryParseBool(af, out var afb) && afb)
                {
                    try { range.EntireColumn.AutoFit(); }
                    catch { skipped.Add("auto_fit_columns"); }
                }

                var data = skipped.Count == 0
                    ? (object)new { formatted = true }
                    : new { formatted = true, skipped = string.Join(",", skipped) };
                return new ToolResult { Ok = true, AffectedRanges = { $"{sheet.Name}!{range.Address}" }, Data = data };
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>校验筛选算子与 value 的匹配：不认识的算子、数值算子配非数值 value、equals/not_equals/contains 配空 value 均返回错误文案；
        /// 合法返回 null。避免“数值缺失→全被过滤出道空结果”或“contains 空值→全表匹配”等静默语义。</summary>
        private static string ValidateFilterOperator(string op, string value)
        {
            switch (op)
            {
                case "equals":
                case "not_equals":
                case "contains":
                    if (string.IsNullOrEmpty(value))
                    {
                        return $"筛选运算符 {op} 需要非空的 value 参数。";
                    }

                    return null;
                case "empty":
                case "not_empty":
                    return null;
                case "gt":
                case "gte":
                case "lt":
                case "lte":
                    return double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _)
                        ? null
                        : $"筛选运算符 {op} 需要数值型的 value 参数（当前：\"{value}\"）。";
                default:
                    return $"不支持的筛选运算符：{op}（应为 equals/not_equals/contains/empty/not_empty/gt/gte/lt/lte）。";
            }
        }

        /// <summary>宽容布尔解析：接受 bool / 1|0 / 字符串 true|false|yes|no|on|off|是|否|是|否（不区分大小写）。解析失败返回 false。</summary>
        private static bool TryParseBool(object value, out bool result)
        {
            if (value is bool b) { result = b; return true; }
            if (value is int n) { result = n != 0; return true; }
            if (value is long l) { result = l != 0; return true; }

            var s = value?.ToString()?.Trim().ToLowerInvariant();
            switch (s)
            {
                case "1":
                case "true":
                case "yes":
                case "y":
                case "on":
                case "是":
                case "对":
                    result = true;
                    return true;
                case "0":
                case "false":
                case "no":
                case "n":
                case "off":
                case "否":
                case "不对":
                    result = false;
                    return true;
                default:
                    result = false;
                    return false;
            }
        }

        /// <summary>宽容颜色解析：接受 Excel 颜色 int、'#RRGGBB'/'RRGGBB' 十六进制、常见中英文颜色名。
        /// 统一转成 Excel 的 BGR 整型（R 在最低字节）。无法识别返回 false。</summary>
        private static bool TryParseColor(object value, out int color)
        {
            color = 0;
            if (value == null) return false;
            if (value is int ci) { color = ci; return true; }
            if (value is long cl) { color = (int)cl; return true; }

            var s = value.ToString().Trim();
            if (s.Length == 0) return false;

            if (s[0] == '#') s = s.Substring(1);
            if (s.Length == 6 && s.All(Uri.IsHexDigit))
            {
                var rgb = Convert.ToInt32(s, 16); // #RRGGBB
                // Excel 用 BGR 布局：R 在最低字节（可为负 int，交给 Excel 按 Color 解释）
                color = (int)(((rgb >> 16) & 0xFF) | (((rgb >> 8) & 0xFF) << 8) | ((rgb & 0xFF) << 16));
                if (color < 0) unchecked { color = (int)((uint)color); }
                return true;
            }

            // 常见颜色名：赋值给中间 / 直接用 BGR 值
            var named = s.ToLowerInvariant();
            switch (named)
            {
                case "black":
                case "黑色":
                case "黑":
                    color = 0;
                    return true;
                case "white":
                case "白色":
                case "白":
                    color = 16777215;
                    return true;
                case "red":
                case "红色":
                case "红":
                    color = 255;
                    return true;
                case "green":
                case "绿色":
                case "绿":
                    color = 65280;
                    return true;
                case "blue":
                case "蓝色":
                case "蓝":
                    color = 16711680;
                    return true;
                case "yellow":
                case "黄色":
                case "黄":
                    color = 65535;
                    return true;
                case "dark blue":
                case "deep blue":
                case "深蓝":
                case "深蓝色":
                    color = 7884319; // #1F4E78
                    return true;
                case "light blue":
                case "浅蓝":
                case "浅蓝色":
                    color = 16053492; // #ADD8E6
                    return true;
                case "orange":
                case "橙色":
                case "橙":
                    color = 33023;
                    return true;
                case "gray":
                case "grey":
                case "灰色":
                case "灰":
                    color = 8421504;
                    return true;
                case "purple":
                case "紫色":
                case "紫":
                    color = 8388736;
                    return true;
                case "cyan":
                case "青色":
                case "青":
                    color = 16776960;
                    return true;
                case "magenta":
                case "品红":
                    color = 16711935;
                    return true;
                case "brown":
                case "棕色":
                case "棕":
                    color = 2763429;
                    return true;
                default:
                    return false;
            }
        }
        public static string CreateWorkbookHash(string workbookFullName)
        {
            var source = workbookFullName ?? string.Empty;
            using var sha256 = SHA256.Create();
            var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(source));
            var builder = new StringBuilder(hash.Length * 2);
            foreach (var value in hash)
            {
                builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        /// <summary>统一工具执行入口：带 COM 忙自动补偿。本程序集内所有工具注册共用。</summary>
        internal static ToolResult Execute(Func<ToolResult> action)
        {
            var attempt = Compensable(action);
            if (attempt != null)
            {
                return attempt;
            }

            // COM 忙：短暂等待后自动补偿重试一次，仍失败才报错。
            Thread.Sleep(200);
            return Compensable(action) ??
                   new ToolResult
                   {
                       Ok = false,
                       ErrorCode = "EXCEL_COM_ERROR",
                       Data = "Excel 正忙，操作未完成，请稍后重试。"
                   };
        }

        /// <summary>
        /// web_search 工具执行（v3.8.1.83）：只读，与 Excel COM 无关；
        /// 运行时读取设置开关（关闭时明确拒绝并指路），开启走 cn.bing.com 抓取兜底。
        /// </summary>
        private static ToolResult WebSearch(Dictionary<string, object> args)
        {
            var settings = new ExcelAiAssistant.Core.Configuration.AgentSettingsStore().Load();
            if (!settings.WebSearchEnabled)
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "WEB_SEARCH_DISABLED",
                    Data = "联网搜索已在设置页关闭。可在「设置 → 联网搜索」重新开启。"
                };
            }

            return ExcelAiAssistant.Core.Search.WebSearchService.Search(args);
        }

        private static ToolResult Compensable(Func<ToolResult> action)
        {
            try
            {
                return action();
            }
            catch (OutOfMemoryException)
            {
                // 大数组写入经 COM 封送 / 撤销快照放大时可能触发 .NET OOM。若任其逃逸会一路冒泡到
                // AgentRunner，把整轮对话打崩（特征：SendPrompt RunAsync 异常: System.OutOfMemoryException）。
                // 此处把 OOM 收敛为工具级友好错误，让对话继续；WriteRangeCore 的 MaxSingleWriteCells 确定性
                // 护栏已先行拦截绝大多数超大数组，这里是边界兜底。
                ExcelTrace.Error("tool_host", "Compensable", new OutOfMemoryException("OOM 边界兜底（工具级收敛，具体上下文见异常消息）"));
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "EXCEL_MEMORY_ERROR",
                    Data = "Excel/COM 内存不足（OutOfMemoryException）：数据量过大，请缩小 address 或分批写入/处理后再试。"
                };
            }
            catch (ArgumentException argEx)
            {
                // M16 全量审查：Required()/参数校验抛出的 ArgumentException 原被兜成 EXCEL_COM_ERROR，
                // 错误码与真实原因（缺参/参数非法）严重不符。收敛为结构化 INVALID_ARGS。
                ExcelTrace.Warn("tool_host", "INVALID_ARGS | " + argEx.Message);
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "INVALID_ARGS",
                    Data = argEx.Message
                };
            }
            catch (KeyNotFoundException keyEx)
            {
                // M16 全量审查：表名不存在等（如 JYYJ Resolve 的 fail-closed 分支）收敛为 INVALID_SHEET。
                ExcelTrace.Warn("tool_host", "INVALID_SHEET | " + keyEx.Message);
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "INVALID_SHEET",
                    Data = keyEx.Message
                };
            }
            catch (Exception ex)
            {
                if (IsComBusy(ex))
                {
                    return null;
                }

                ExcelTrace.Error("tool_host", "Compensable", ex);
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "EXCEL_COM_ERROR",
                    Data = ex.Message
                };
            }
        }

        /// <summary>
        /// 把强类型模型（PascalCase 属性）投影为 snake_case 的 JSON 对象（M16 全量审查）：
        /// 与其它工具返回的 snake_case 匿名对象保持同一契约，AI 消费端不再拿到 PascalCase 键。
        /// </summary>
        private static readonly Newtonsoft.Json.JsonSerializerSettings SnakeCaseSettings =
            new Newtonsoft.Json.JsonSerializerSettings
            {
                ContractResolver = new Newtonsoft.Json.Serialization.DefaultContractResolver
                {
                    NamingStrategy = new Newtonsoft.Json.Serialization.SnakeCaseNamingStrategy()
                }
            };

        internal static object ToSnakeCaseData(object model)
        {
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(model, SnakeCaseSettings);
            return Newtonsoft.Json.JsonConvert.DeserializeObject(json);
        }

        /// <summary>判断是否为 COM 忙/调用被拒绝类瞬时错误（RPC_E_SERVERCALL_RETRYLATER、RPC_E_CALL_REJECTED、RPC_E_SERVERCALL_REJECTED）。</summary>
        internal static bool IsComBusy(Exception ex)
        {
            if (!(ex is COMException com))
            {
                return false;
            }

            var hr = unchecked((int)com.HResult);
            return hr == -2147417846 || hr == -2147418111 || hr == -2147417851;
        }

        private ToolResult PreviewRange(Dictionary<string, object> args)
        {
            var requested = GetOptionalString(args, "address");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic values = null;

            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = string.IsNullOrWhiteSpace(requested)
                    ? application.Selection
                    : sheet.Range[requested];
                var rowCount = Math.Min(Convert.ToInt32(range.Rows.Count), 20);
                var columnCount = Convert.ToInt32(range.Columns.Count);
                var truncated = rowCount < Convert.ToInt32(range.Rows.Count);

                if (rowCount > 0 && columnCount > 0)
                {
                    var firstCell = range.Cells[1, 1];
                    var lastCell = range.Cells[rowCount, columnCount];
                    values = sheet.Range[firstCell, lastCell].Value2;
                    ExcelContextService.Release((object)firstCell);
                    ExcelContextService.Release((object)lastCell);
                }

                var sample = FlattenValues(values, rowCount, columnCount);
                return new ToolResult
                {
                    Ok = true,
                    Truncated = truncated,
                    Data = new
                    {
                        address = Convert.ToString(range.Address),
                        rows = Convert.ToInt32(range.Rows.Count),
                        columns = columnCount,
                        sampleRows = rowCount,
                        sample = sample
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)values);
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>
        /// 防公式注入：value 写入前把以 = + - @ 或 Tab/CR 开头的字符串加前导单引号，强制按文本落盘。
        /// 返回被转义的格数，供结果如实上报（写公式请走 value_mode=formula）。
        /// </summary>
        private static int EscapeFormulaLikeCells(object[,] values)
        {
            var escaped = 0;
            for (var r = values.GetLowerBound(0); r <= values.GetUpperBound(0); r++)
            {
                for (var c = values.GetLowerBound(1); c <= values.GetUpperBound(1); c++)
                {
                    if (!(values[r, c] is string text) || text.Length == 0)
                    {
                        continue;
                    }

                    var first = text[0];
                    if (first == '=' || first == '+' || first == '-' || first == '@' || first == '\t' || first == '\r')
                    {
                        values[r, c] = "'" + text;
                        escaped++;
                    }
                }
            }

            return escaped;
        }

        /// <summary>
        /// values 参数期望二维数组。模型若传对象数组（[{A:1,B:2}]），NormalizeValues 会把 JObject
        /// 当行枚举成 KeyValuePair 静默写入错误内容——必须在归一化前显式拒绝并指向正确的批量工具。
        /// </summary>
        private static ToolResult RejectObjectArrayInput(object input, string toolHint)
        {
            // 扫描全部行（不只首行）：[null, {"A":1}] 这类混合数组首元素是 JValue，
            // 放过会让第二行 JObject 被 NormalizeValues 枚举成 KeyValuePair 写错内容。
            if (input is Newtonsoft.Json.Linq.JArray array)
            {
                foreach (var item in array)
                {
                    if (item is Newtonsoft.Json.Linq.JObject || item is Newtonsoft.Json.Linq.JProperty)
                    {
                        return Invalid("INVALID_ARGS",
                            "values 必须是二维数组（如 [[\"A\",\"B\"],[1,2]]）；对象数组（[{A:1,B:2}]）请改用 " + toolHint + "。");
                    }
                }
            }

            // 进程内对象数组（List<Dictionary<string,object>>）同样不是二维矩阵。
            if (input is System.Collections.IEnumerable sequence && !(input is string))
            {
                foreach (var item in sequence)
                {
                    if (item is IDictionary<string, object>)
                    {
                        return Invalid("INVALID_ARGS",
                            "values 必须是二维数组（如 [[\"A\",\"B\"],[1,2]]）；对象数组请改用 " + toolHint + "。");
                    }
                }
            }

            return null;
        }

        /// <summary>cells 参数的「对象数组」语义归一化：JArray of JObject / Dictionary 列表 → List&lt;Dictionary&gt;；形态不支持返回 null。</summary>
        private static List<Dictionary<string, object>> NormalizeCellEntries(object raw)
        {
            if (!(raw is System.Collections.IEnumerable sequence) || raw is string)
            {
                return null;
            }

            var result = new List<Dictionary<string, object>>();
            foreach (var item in sequence)
            {
                if (item is Dictionary<string, object> dict)
                {
                    result.Add(dict);
                    continue;
                }

                if (item is Newtonsoft.Json.Linq.JObject jo)
                {
                    result.Add(jo.ToObject<Dictionary<string, object>>());
                    continue;
                }

                return null;
            }

            return result.Count > 0 ? result : null;
        }

        private ToolResult WriteCellsBatch(Dictionary<string, object> args)
        {
            object rawCells;
            if (!args.TryGetValue("cells", out rawCells))
            {
                throw new ArgumentException("缺少必填参数：cells");
            }

            // cells 归一化：模型经 JSON 传来的是 JArray of JObject；NormalizeValues 会把 JObject 当
            // IEnumerable<object>（其枚举元素是 JProperty）展开，导致「cells 必须是对象数组」的假性失败。
            // 先按「对象数组」语义逐元素转 Dictionary，仅对非 JSON 形态回退旧 NormalizeValues 路径。
            var cells = NormalizeCellEntries(rawCells);
            if (cells == null)
            {
                var legacy = NormalizeValues(rawCells);
                cells = new List<Dictionary<string, object>>();
                for (var index = 0; index < legacy.GetLength(0); index++)
                {
                    if (legacy[index, 0] is Dictionary<string, object> entry)
                    {
                        cells.Add(entry);
                    }
                    else
                    {
                        throw new ArgumentException("cells 必须是对象数组。");
                    }
                }
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic cellRange = null;
            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                var writes = new List<object>();
                var operationIds = new List<string>();
                for (var index = 0; index < cells.Count; index++)
                {
                    var cell = cells[index];
                    var address = GetString(cell, "address");
                    var valueMode = GetOptionalString(cell, "value_mode") ?? "value";
                    object value;
                    if (!cell.TryGetValue("value", out value) || value == null)
                    {
                        throw new ArgumentException($"cells[{index}] 缺少 value。");
                    }

                    var single = new object[1, 1];
                    single[0, 0] = UnwrapScalar(value);
                    cellRange = sheet.Range[address];
                    var result = WriteRangeCore(
                        workbook,
                        sheet,
                        cellRange,
                        single,
                        valueMode,
                        new Dictionary<string, object>(cell, StringComparer.OrdinalIgnoreCase),
                        "write_excel_cells_batch");
                    ExcelContextService.Release((object)cellRange);
                    cellRange = null;
                    if (!result.Ok)
                    {
                        // 批量写非原子：中途失败时如实回报"已写哪几格 + 对应 operation_id"，
                        // 用户/AI 可据此逐条回滚；旧实现直接 return 丢掉已写明细，表已半改却无人知情。
                        return new ToolResult
                        {
                            Ok = false,
                            ErrorCode = result.ErrorCode ?? "BATCH_WRITE_PARTIAL",
                            Data = new
                            {
                                sheet = Convert.ToString(sheet.Name),
                                failed_address = address,
                                written_cells = writes.Count,
                                operation_ids = operationIds.ToArray(),
                                writes,
                                note = "批量写入中途失败：writes 中的单元格已写入（可用 operation_ids 逐条回滚），失败地址见 failed_address。"
                            }
                        };
                    }

                    var operationId = ExcelToolHostHelpers.GetOperationId(result);
                    if (operationId != null)
                    {
                        operationIds.Add(Convert.ToString(operationId));
                    }

                    writes.Add(new { address, written = true });
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        written_cells = writes.Count,
                        operation_ids = operationIds.ToArray(),
                        writes
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)cellRange);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult WriteCell(Dictionary<string, object> args)
        {
            // FIX（M28/R3）：单格写入只接受标量——矩阵值原先会被 RefitToRegion 静默裁到 1×1，
            // 或者区域地址先报 INVALID_ADDRESS 掩盖「改用 write_excel_range」的可执行指引。
            if (args.TryGetValue("value", out var rawValue) &&
                rawValue is System.Collections.IEnumerable && !(rawValue is string))
            {
                return Invalid("INVALID_ARGS",
                    "write_excel_cell/set_element 只接受标量值；写二维矩阵请用 write_excel_range（values 参数）。");
            }

            var address = GetString(args, "address");
            if (!ExcelAddressParser.TryParseCell(address, out _, out _))
            {
                return Invalid("INVALID_ADDRESS", $"单元格地址无效：{address}");
            }

            return WriteRange(new Dictionary<string, object>(args, StringComparer.OrdinalIgnoreCase)
            {
                ["address"] = address
            });
        }

        internal ToolResult WriteRangeCore(
            dynamic workbook,
            dynamic sheet,
            dynamic range,
            object[,] values,
            string valueMode,
            Dictionary<string, object> args,
            string tool)
        {
            var rows = values.GetLength(0);
            var columns = values.GetLength(1);

            // 单次写入格数护栏：超大数组经 COM Value2/撤销快照(前后各读 values+formulas 共 4 份)会暴涨内存，
            // 实测可触发 System.OutOfMemoryException（COM 封送 SAFEARRAY 拷贝 + 快照放大）。确定性拦截，让模型拆小分批写。
            var totalCells = rows * (long)columns;
            if (totalCells > MaxSingleWriteCells)
            {
                return Invalid("INVALID_RANGE",
                    $"单次写入 {rows} 行 × {columns} 列 = {totalCells} 个单元，超过单次上限 {MaxSingleWriteCells} 格，容易触发 Excel/内存不足(OOM)。请缩小 address 或分批写入（每批不超过 {MaxSingleWriteCells} 格）。");
            }

            var targetRows = Convert.ToInt32(range.Rows.Count);
            var targetColumns = Convert.ToInt32(range.Columns.Count);

            // 目标区域也要过格数护栏：整列地址（A:A=1,048,576 行）会让输入矩阵被 Refit 放大到百万格，
            // 随后全列快照 (Value2+Formula) 直接 OOM——旧实现只校验输入矩阵，整列地址可绕过。
            var targetCells = (long)targetRows * targetColumns;
            if (targetCells > MaxSingleWriteCells)
            {
                return Invalid("INVALID_RANGE",
                    $"目标区域 {targetRows} 行 × {targetColumns} 列 = {targetCells} 个单元，超过单次上限 {MaxSingleWriteCells} 格。请给出有限区间（如 A1:A1000），不要使用整列/整行地址。");
            }

            var shapeRefit = rows != targetRows || columns != targetColumns;
            var refitSourceCells = shapeRefit ? rows * columns : 0;
            if (shapeRefit)
            {
                // 目标区域是用户/模型指定的地址。AI 对区域形状的估计常与数据形状有裁剪级偏差
                // （额外/缺少一行一列），逐格硬失败会把整轮写入打回、只能靠 AI 换工具兜底——
                // 这正是 1.19.0 起「write_excel_range 写入失败」的根因。因此算法改为按行优先
                // 把提供的值拟合进目标区域：值多则截断到区域、值少则区域剩余格清空（区域替换语义，
                // 与形状恰好一致时的行为保持一致），不再返回 SHAPE_MISMATCH。
                values = RefitToRegion(values, targetRows, targetColumns);
                rows = targetRows;
                columns = targetColumns;
            }

            var resolvedMode = ResolveValueMode(valueMode);
            int? textLockedCells = null;

            var beforeValues = FlattenValues(range.Value2, rows, columns);
            var beforeFormulas = FlattenValues(range.Formula, rows, columns);

            var escapedCells = 0;
            if (string.Equals(resolvedMode, "formula", StringComparison.OrdinalIgnoreCase))
            {
                try { range.Formula = values; }
                catch (OutOfMemoryException) { return OomWriteError(rows, columns); }
                catch (System.Runtime.InteropServices.COMException ce) { return ComWriteError(resolvedMode, ce); }
            }
            else
            {
                // v3.8.1.76 写入文本保护（审查批次真机实证）：COM Value2 写入会按"手工输入"语义
                // 重解析字符串——18 位身份证→5.787E+17（精度丢失）、"007"→7（前导零丢失）。
                // 对身份证/手机号/前导零编码等形态先逐格锁 NumberFormat="@" 再整块写值（Core 纯
                // 逻辑 TextPreserveAdvisor 判定，宁窄勿宽；日期串不拦，走 coerced_cells 上报）。
                var textLockCells = Core.Diagnostics.TextPreserveAdvisor.FindTextLockCells(values);
                textLockedCells = textLockCells.Count;
                if (textLockCells.Count > 0)
                {
                    foreach (var rc in textLockCells)
                    {
                        dynamic cellRng = null;
                        try
                        {
                            cellRng = range.Cells.Item[rc.Key, rc.Value];
                            cellRng.NumberFormat = "@";
                        }
                        catch (Exception xaiLock)
                        {
                            AppLog.Warn("host", "ExcelToolHost.cs:WriteRangeCore 文本锁格式设置失败(该格继续原样写入): " + xaiLock.Message);
                        }
                        finally
                        {
                            ExcelContextService.Release((object)cellRng);
                        }
                    }
                }

                // 防公式注入（用户拍板 A 方案）：仅显式 value 模式按"字面值"写入——以 = + - @ 或 Tab/CR
                // 开头的字符串会被 Excel 按手工输入解析为公式（WEBSERVICE/DDE 可外连）。
                // auto/formula 保留旧语义（= 开头按公式），要写公式请用 value_mode=formula 或 auto。
                if (string.Equals(resolvedMode, "value", StringComparison.OrdinalIgnoreCase))
                {
                    escapedCells = EscapeFormulaLikeCells(values);
                }

                try { range.Value2 = values; }
                catch (OutOfMemoryException) { return OomWriteError(rows, columns); }
                catch (System.Runtime.InteropServices.COMException ce) { return ComWriteError(resolvedMode, ce); }
            }

            var afterValues = FlattenValues(range.Value2, rows, columns);
            var afterFormulas = FlattenValues(range.Formula, rows, columns);

            // v3.8.1.76 重解析可见性：输入字符串被 Excel 存为数值/日期（如日期串→序列号）时上报明细，
            // 把"写入/回读不符"从静默漂移变成 AI 可复核的显式信息（长字符串才上报防噪声）。
            List<object> coercedCells = null;
            var coercedCount = 0;
            if (!string.Equals(resolvedMode, "formula", StringComparison.OrdinalIgnoreCase))
            {
                coercedCells = CollectCoercedCells(values, afterValues);
                coercedCount = coercedCells.Count;
            }

            var sheetName = Convert.ToString(sheet.Name);
            var rangeAddress = Convert.ToString(range.Address);
            string operationId = null;

            if (undoService != null)
            {
                var record = undoService.Capture(
                    sessionId,
                    CreateWorkbookHash(Convert.ToString(workbook.FullName)),
                    Convert.ToString(workbook.FullName),
                    tool,
                    args,
                    new CellSnapshot
                    {
                        Sheet = sheetName,
                        Address = rangeAddress,
                        BeforeValues = beforeValues,
                        AfterValues = afterValues,
                        BeforeFormulas = beforeFormulas,
                        AfterFormulas = afterFormulas
                    });
                operationId = record.Id;
            }

            return new ToolResult
            {
                Ok = true,
                AffectedRanges = { $"{sheetName}!{rangeAddress}" },
                Data = new
                {
                    operation_id = operationId,
                    written_cells = rows * columns,
                    value_mode = valueMode.ToLowerInvariant(),
                    text_locked_cells = textLockedCells,
                    escaped_cells = escapedCells,
                    escaped_note = escapedCells > 0
                        ? "有单元格以 = + - @ 等公式前缀开头，已按字面文本转义写入（防公式注入）。如需写入公式，请显式传 value_mode=formula。"
                        : null,
                    coerced_cells = coercedCells,
                    coerced_count = coercedCount,
                    coerced_note = coercedCount > 0
                        ? "以下输入字符串被 Excel 重解析存储（日期串→日期序列号/长数字→数值）。如需按文本保留，请先对该列设置文本格式再写入；长数字/前导零编码已由写入层自动锁文本。"
                        : null,
                    shape_refit = shapeRefit,
                    // M16 全量审查：形状拟合时如实上报源数据格数（目标格数=written_cells），
                    // AI 可据此判断截断/清空的规模，而不是只看到一个 shape_refit=true。
                    refit_source_cells = refitSourceCells > 0 ? (int?)refitSourceCells : null,
                    before_values = beforeValues,
                    after_values = afterValues,
                    before_formulas = beforeFormulas,
                    after_formulas = afterFormulas
                }
            };
        }

        private static ToolResult OomWriteError(int rows, int columns)
        {
            return Invalid(
                "WRITE_OOM",
                $"写入 {rows} 行 × {columns} 列时 Excel/COM 内存不足（OutOfMemoryException）。数据量过大，请缩小 address 或分批写入后再试。");
        }

        private static List<object> CollectCoercedCells(object[,] input, object[,] after)
        {
            var list = new List<object>();
            if (input == null || after == null)
                return list;
            var rows = Math.Min(input.GetLength(0), after.GetLength(0));
            var cols = Math.Min(input.GetLength(1), after.GetLength(1));
            for (var r = 1; r <= rows; r++)
            {
                for (var c = 1; c <= cols; c++)
                {
                    var inp = input[r - 1, c - 1];
                    if (!Core.Diagnostics.TextPreserveAdvisor.ShouldReportCoercion(inp))
                        continue;
                    var got = after[r - 1, c - 1];
                    if (got is double || got is DateTime)
                    {
                        list.Add(new
                        {
                            row = r,
                            column = c,
                            input = Convert.ToString(inp),
                            stored_as = got is DateTime ? "date" : "number",
                            stored_value = Convert.ToString(got)
                        });
                        if (list.Count >= 20)
                            return list;
                    }
                }
            }
            return list;
        }

        private static ToolResult ComWriteError(string mode, System.Runtime.InteropServices.COMException ce, dynamic sheet = null, dynamic range = null)
        {
            // v3.7.1.59：0x800A03EC 家族不再回裸 HRESULT——探测保护/只读/隐藏/合并/数组公式五类上下文，
            // 由 ComWriteErrorAdvisor（Core 纯逻辑）给出按优先级排列的可行动建议；探测失败不阻断归因。
            // 工具分类审查批次修复（接线断点）：advice 曾被计算后丢弃、返回消息硬编码笼统建议——
            // 现把归因建议并入返回消息，AI/用户直接拿到「保护态→verify_protection」等可行动指引。
            var ctx = ProbeWriteContext(sheet, range);
            var advice = Core.Diagnostics.ComWriteErrorAdvisor.Advise(ce.HResult, ctx);
            return Invalid(
                "WRITE_FAILED",
                $"写入（value_mode={mode}）失败：{ce.Message}。" + (advice ?? "请改用 value 直写或拆小区域后重试。"));
        }

        /// <summary>写入失败上下文探测（M17 探测纪律：零 FinalRelease、逐项 try-try-catch 失败即跳过）。</summary>
        private static Core.Diagnostics.ComWriteErrorAdvisor.WriteContext ProbeWriteContext(dynamic sheet, dynamic range)
        {
            var ctx = new Core.Diagnostics.ComWriteErrorAdvisor.WriteContext();
            if (sheet != null)
            {
                try { ctx.SheetProtected = Convert.ToBoolean(sheet.ProtectContents); } catch (Exception) { }
                try { ctx.SheetVisibleState = Convert.ToInt32(sheet.Visible); } catch (Exception) { }
                try { ctx.SheetName = Convert.ToString(sheet.Name); } catch (Exception) { }
                try { ctx.WorkbookReadOnly = Convert.ToBoolean(sheet.Parent.ReadOnly); } catch (Exception) { }
            }

            if (range != null)
            {
                try { ctx.Address = Convert.ToString(range.Address); } catch (Exception) { }
                try { ctx.TargetHasMergedCells = Convert.ToBoolean(range.MergeCells); } catch (Exception) { }
                try { ctx.TargetHasArrayFormula = Convert.ToBoolean(range.HasArray); } catch (Exception) { }
            }

            return ctx;
        }

        private ToolResult WriteRange(Dictionary<string, object> args)
        {
            var address = GetString(args, "address");
            var valueMode = GetOptionalString(args, "value_mode") ?? "auto";
            var input = args.TryGetValue("values", out var rawValues) ? rawValues : args["value"];
            var rejected = RejectObjectArrayInput(input, "write_excel_cells_batch");
            if (rejected != null)
            {
                return rejected;
            }

            var normalized = NormalizeValues(input);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;

            try
            {
                ResolveTarget(ref workbook, ref sheet, GetOptionalString(args, "sheet"));
                range = sheet.Range[address];
                // 统一走 WriteRangeCore：内含形状自适应重排 + 撤销快照捕获。
                return WriteRangeCore(workbook, sheet, range, normalized, valueMode, args, "write_excel_range");
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                // 入口阶段（ResolveTarget / Range 解析）抛出的 0x800A03EC 同样走归因（v3.7.1.62 审查补口）。
                return ComWriteError(valueMode, ce, sheet, range);
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult UndoOperation(Dictionary<string, object> args)
        {
            if (operationRepository == null || undoService == null)
            {
                return Invalid("UNDO_UNAVAILABLE", "当前宿主未启用操作历史。");
            }

            var requestedId = GetOptionalString(args, "operation_id");
            var record = string.IsNullOrWhiteSpace(requestedId)
                ? operationRepository.List(new OperationQuery { Limit = 1 }).FirstOrDefault()
                : operationRepository.Get(requestedId);

            if (record == null)
            {
                return Invalid("OPERATION_NOT_FOUND", "没有找到可回滚操作。");
            }

            if (!record.CanUndo)
            {
                return Invalid("UNDO_FAILED", "该操作已不可回滚。");
            }

            if (!WorkbookMatches(record))
            {
                return Invalid("UNDO_WORKBOOK_MISMATCH", "该操作记录属于其他工作簿，无法在当前工作簿回滚。");
            }

            if (string.Equals(record.UndoKind, "sheet-insert", StringComparison.OrdinalIgnoreCase))
            {
                return UndoSheetInsert(record);
            }

            if (string.Equals(record.UndoKind, "sheet-rename", StringComparison.OrdinalIgnoreCase))
            {
                return UndoSheetRename(record);
            }

            if (string.Equals(record.UndoKind, "dimension-insert", StringComparison.OrdinalIgnoreCase))
            {
                return UndoDimensionInsert(record);
            }

            if (string.Equals(record.UndoKind, "chart-insert", StringComparison.OrdinalIgnoreCase))
            {
                return UndoChartInsert(record);
            }

            if (string.Equals(record.UndoKind, "sparkline-insert", StringComparison.OrdinalIgnoreCase))
            {
                return UndoSparklineInsert(record);
            }

            var snapshot = Newtonsoft.Json.JsonConvert.DeserializeObject<CellSnapshot>(record.PayloadJson);
            if (snapshot == null)
            {
                return Invalid("OPERATION_NOT_FOUND", "操作快照无效。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = workbook.Worksheets.Item[snapshot.Sheet];
                range = sheet.Range[snapshot.Address];
                var currentValues = FlattenValues(
                    range.Value2,
                    snapshot.AfterValues?.GetLength(0) ?? 0,
                    snapshot.AfterValues?.GetLength(1) ?? 0);
                var result = undoService.Undo(record.Id, currentValues);
                return new ToolResult
                {
                    Ok = result.Restored,
                    ErrorCode = result.Restored ? null : "UNDO_FAILED",
                    Data = new
                    {
                        operation_id = record.Id,
                        conflict_detected = result.ConflictDetected,
                        message = result.Message
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private bool WorkbookMatches(OperationRecord record)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                var path = Convert.ToString(workbook?.FullName ?? string.Empty);
                return string.IsNullOrEmpty(record.WorkbookHash) ||
                       string.Equals(record.WorkbookHash, CreateWorkbookHash(path), StringComparison.Ordinal);
            }
            finally
            {
                ExcelContextService.Release((object)workbook);
            }
        }

        /// <summary>撤销工作表重命名（sheet-rename）：把当前名改回旧名；旧名已被占用时拒绝，避免覆盖他表。</summary>
        private ToolResult UndoSheetRename(OperationRecord record)
        {
            var payload = Newtonsoft.Json.Linq.JObject.Parse(record.PayloadJson ?? "{}");
            var oldName = payload.Value<string>("old_name");
            var newName = payload.Value<string>("new_name");
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
            {
                return Invalid("UNDO_FAILED", "操作快照无效。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic probe = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");

                var oldNameTaken = false;
                try { probe = workbook.Worksheets.Item[oldName]; oldNameTaken = probe != null; }
                catch (System.Runtime.InteropServices.COMException) { oldNameTaken = false; }
                finally { ExcelContextService.Release((object)probe); }

                if (oldNameTaken)
                {
                    return Invalid("UNDO_CONFLICT", "原表名「" + oldName + "」已被其他工作表占用，无法回滚重命名。");
                }

                sheet = workbook.Worksheets.Item[newName];
                sheet.Name = oldName;
                return new ToolResult
                {
                    Ok = true,
                    Data = new { restored_name = oldName, previous_name = newName }
                };
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                return Invalid("UNDO_FAILED", "回滚重命名失败：" + ce.Message);
            }
            finally
            {
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult UndoSheetInsert(OperationRecord record)
        {
            var payload = Newtonsoft.Json.JsonConvert
                .DeserializeObject<SheetInsertUndoPayload>(record.PayloadJson);
            if (payload == null || payload.Sheets == null || payload.Sheets.Length == 0)
            {
                return Invalid("UNDO_FAILED", "操作快照无效。");
            }

            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var removed = new List<string>();
                foreach (var name in payload.Sheets)
                {
                    dynamic sheet = null;
                    try
                    {
                        sheet = workbook.Worksheets.Item[name];
                        var sheetName = Convert.ToString(sheet.Name);
                        sheet.Delete();
                        removed.Add(sheetName);
                    }
                    catch
                    {
                        // 工作表已不存在时忽略。
                    }
                    finally
                    {
                        ExcelContextService.Release((object)sheet);
                    }
                }

                if (removed.Count > 0)
                {
                    operationRepository.MarkNotUndoable(record.Id);
                }

                return new ToolResult
                {
                    Ok = removed.Count > 0,
                    ErrorCode = removed.Count > 0 ? null : "UNDO_FAILED",
                    Data = new
                    {
                        operation_id = record.Id,
                        undo_kind = "sheet-insert",
                        removed_sheets = removed.ToArray(),
                        message = removed.Count > 0 ? $"已删除工作表：{string.Join("、", removed)}" : "没有需要回滚的工作表。"
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult UndoDimensionInsert(OperationRecord record)
        {
            var payload = Newtonsoft.Json.JsonConvert
                .DeserializeObject<DimensionInsertUndoPayload>(record.PayloadJson);
            if (payload == null || payload.Count <= 0 || string.IsNullOrWhiteSpace(payload.Address))
            {
                return Invalid("UNDO_FAILED", "操作快照无效。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic whole = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = workbook.Worksheets.Item[payload.Sheet];
                range = sheet.Range[payload.Address];
                if (string.Equals(payload.Dimension, "column", StringComparison.OrdinalIgnoreCase))
                {
                    whole = range.Resize[1, payload.Count].EntireColumn;
                }
                else
                {
                    whole = range.Resize[payload.Count, 1].EntireRow;
                }

                whole.Delete();
                operationRepository.MarkNotUndoable(record.Id);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        operation_id = record.Id,
                        undo_kind = "dimension-insert",
                        sheet = payload.Sheet,
                        dimension = payload.Dimension,
                        count = payload.Count,
                        message = $"已删除插入的 {payload.Count} 个{payload.Dimension}。"
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)whole);
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult UndoChartInsert(OperationRecord record)
        {
            var payload = Newtonsoft.Json.JsonConvert.DeserializeObject<ChartInsertPayload>(record.PayloadJson);
            if (payload == null || string.IsNullOrWhiteSpace(payload.Sheet) || string.IsNullOrWhiteSpace(payload.ChartName))
            {
                return Invalid("OPERATION_NOT_FOUND", "操作快照无效。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic chartObject = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = workbook.Worksheets.Item[payload.Sheet];
                try
                {
                    chartObject = sheet.ChartObjects(payload.ChartName);
                }
                catch
                {
                    return Invalid("CHART_NOT_FOUND", $"图表“{payload.ChartName}”已不存在，可能已被手动删除。");
                }

                chartObject.Delete();
                operationRepository.MarkNotUndoable(record.Id);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        operation_id = record.Id,
                        undo_kind = "chart-insert",
                        sheet = payload.Sheet,
                        chart_name = payload.ChartName,
                        message = $"已删除图表“{payload.ChartName}”。"
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)chartObject);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private ToolResult UndoSparklineInsert(OperationRecord record)
        {
            var payload = Newtonsoft.Json.JsonConvert.DeserializeObject<SparklineInsertPayload>(record.PayloadJson);
            if (payload == null || string.IsNullOrWhiteSpace(payload.Sheet) || string.IsNullOrWhiteSpace(payload.DestRange))
            {
                return Invalid("OPERATION_NOT_FOUND", "操作快照无效。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = workbook.Worksheets.Item[payload.Sheet];
                range = sheet.Range[payload.DestRange];
                range.SparklineGroups.Clear();
                operationRepository.MarkNotUndoable(record.Id);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        operation_id = record.Id,
                        undo_kind = "sparkline-insert",
                        sheet = payload.Sheet,
                        dest_range = payload.DestRange,
                        message = $"已清除 {payload.DestRange} 的迷你图。"
                    }
                };
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private bool RestoreSnapshot(CellSnapshot snapshot)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = workbook.Worksheets.Item[snapshot.Sheet];
                range = sheet.Range[snapshot.Address];
                range.Value2 = snapshot.BeforeValues;
                return true;
            }
            catch
            {
                // 回滚失败（如受影响项是图表/形状、地址已失效）不应把 COM 底层异常抛出：
                // 返回 false，由 UndoService 归为“恢复快照失败”，保持可回滚语义清晰。
                return false;
            }
            finally
            {
                ExcelContextService.Release((object)range);
                ExcelContextService.Release((object)sheet);
                ExcelContextService.Release((object)workbook);
            }
        }

        private void ResolveTarget(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static object[,] NormalizeValues(object input)
        {
            switch (input)
            {
                case object[,] array:
                    return array;
                case IEnumerable<object> sequence:
                {
                    var items = sequence.Cast<object>().ToList();
                    // 行元素兼容两种形态：object[]（实现 IEnumerable<object>）与 int[]/double[] 等
                    // 非泛型枚举（仅实现 IEnumerable<int>，协变不适用值类型）。后者若被当标量
                    // 或直接强转 IEnumerable<object> 会抛 InvalidCastException
                    // （"Int32[] → IEnumerable<object>"，全量审查实测），这里统一走非泛型枚举兜底。
                    static List<object> ToRow(object item)
                    {
                        if (item is string) return null;
                        if (item is IEnumerable<object> seq) return seq.ToList();
                        if (item is System.Collections.IEnumerable en) return en.Cast<object>().ToList();
                        return null;
                    }

                    var rows = items.Select(ToRow).ToList();
                    var scalarOnly = rows.All(row => row == null);
                    if (scalarOnly)
                    {
                        var result = new object[items.Count, 1];
                        for (var index = 0; index < items.Count; index++)
                        {
                            result[index, 0] = UnwrapScalar(items[index]);
                        }

                        return result;
                    }

                    var width = rows.Max(row => row?.Count ?? 1);
                    var matrix = new object[items.Count, width];
                    for (var rowIndex = 0; rowIndex < items.Count; rowIndex++)
                    {
                        var row = rows[rowIndex];
                        for (var columnIndex = 0; columnIndex < width; columnIndex++)
                        {
                            matrix[rowIndex, columnIndex] =
                                row != null && columnIndex < row.Count ? UnwrapScalar(row[columnIndex]) : null;
                        }
                    }

                    return matrix;
                }
                default:
                    if (input is string text)
                    {
                        // AI 偶尔会把要写入的二维数据字符串化成一个 JSON 数组字面量再传给工具，
                        // 若当标量写入会整段落进单格（written_cells=1、区域只见一格文本）。仅在
                        // 字符串恰为合法二维 JSON 数组时解析展开；否则仍当普通文本单元格值写入。
                        var parsed = TryParseJsonMatrix(text);
                        if (parsed != null)
                        {
                            return parsed;
                        }
                    }

                    var single = new object[1, 1];
                    single[0, 0] = UnwrapScalar(input);
                    return single;
            }
        }

        /// <summary>把「JSON 数组字面量字符串」解析为二维矩阵；非二维数组或解析失败返回 null。</summary>
        private static object[,] TryParseJsonMatrix(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var trimmed = text.TrimStart();
            if (trimmed.Length == 0 || trimmed[0] != '[')
            {
                return null;
            }

            try
            {
                var arr = Newtonsoft.Json.Linq.JArray.Parse(text);
                if (arr.Count == 0)
                {
                    return null;
                }

                var rows = new List<object[]>();
                foreach (var item in arr)
                {
                    if (item is Newtonsoft.Json.Linq.JArray row)
                    {
                        if (row.Count == 0)
                        {
                            return null;
                        }

                        var values = new object[row.Count];
                        var hasObject = false;
                        for (var i = 0; i < row.Count; i++)
                        {
                            if (row[i] is Newtonsoft.Json.Linq.JObject)
                            {
                                hasObject = true;
                                break;
                            }

                            values[i] = row[i] is Newtonsoft.Json.Linq.JValue jv ? jv.Value : null;
                        }

                        if (hasObject)
                        {
                            return null; // 含对象成员，属非二维标量表，不擅自推断
                        }

                        rows.Add(values);
                    }
                    else if (item is Newtonsoft.Json.Linq.JValue scalar)
                    {
                        rows.Add(new object[] { scalar.Value }); // 顶层级联的单列
                    }
                    else
                    {
                        return null;
                    }
                }

                var width = rows.Max(r => r.Length);
                var matrix = new object[rows.Count, width];
                for (var r = 0; r < rows.Count; r++)
                {
                    for (var c = 0; c < width; c++)
                    {
                        matrix[r, c] = c < rows[r].Length ? rows[r][c] : null;
                    }
                }

                return matrix;
            }
            catch
            {
                return null;
            }
        }

        private static object UnwrapScalar(object value)
        {
            if (value is Newtonsoft.Json.Linq.JValue jValue)
            {
                return jValue.Value;
            }

            if (value is IEnumerable<object> list && !(value is string))
            {
                var first = list.FirstOrDefault();
                return UnwrapScalar(first);
            }

            return value;
        }

        /// <summary>把源矩阵按行优先拟合进目标区域：值多于区域则截断，值少于区域则其余格保持空（区域替换语义）。</summary>
        private static object[,] RefitToRegion(object[,] source, int rows, int columns)
        {
            var result = new object[rows, columns];
            var sourceRows = source.GetLength(0);
            var sourceColumns = source.GetLength(1);
            var total = rows * columns;
            var index = 0;
            for (var r = 0; r < sourceRows && index < total; r++)
            {
                for (var c = 0; c < sourceColumns && index < total; c++, index++)
                {
                    result[index / columns, index % columns] = source[r, c];
                }
            }

            return result; // 未填充格保持 null，写入时对应目标格被清空
        }

        private static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source == null || rows == 0 || columns == 0)
            {
                return result;
            }

            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }

                return result;
            }

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            return GetOptionalString(args, key) ??
                   throw new ArgumentException($"缺少必填参数：{key}");
        }

        private static string GetOptionalString(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            return value is Newtonsoft.Json.Linq.JValue jValue
                ? Convert.ToString(jValue.Value)
                : Convert.ToString(value);
        }

        /// <summary>当工具显式要求 auto（或未指定）时，按用户写入偏好解析为 formula/value；否则尊重显式值。</summary>
        private string ResolveValueMode(string valueMode)
        {
            var mode = string.IsNullOrWhiteSpace(valueMode) ? "auto" : valueMode.Trim().ToLowerInvariant();
            if (mode != "auto")
            {
                return mode;
            }

            if (string.Equals(WriteModePreference, "formula", StringComparison.OrdinalIgnoreCase))
            {
                return "formula";
            }

            if (string.Equals(WriteModePreference, "value", StringComparison.OrdinalIgnoreCase))
            {
                return "value";
            }

            return "auto";
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult
            {
                Ok = false,
                ErrorCode = code,
                Data = message
            };
        }
    }
}











