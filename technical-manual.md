[JYYJ助手技术手册-3.8.1.83.md](https://github.com/user-attachments/files/32271227/JYYJ.-3.8.1.83.md)
# JYYJ助手 技术手册

> Excel COM Add-in · .NET Framework 4.8 · WebView2 · OpenAI 兼容协议 · 版本 3.8.1.83

## 一、总体架构

JYYJ助手 是一款原生 **Excel COM 加载项**（.NET Framework VSTO 风格的托管 COM Add-in），采用「侧边栏 WebView2 前台 + 托管 COM 桥 + 工具服务层 + Agent 编排」的分层结构。

### 运行环境与 Excel 版本兼容

本加载项采用经典 **COM Add-in（IDTExtensibility2）+ 侧边栏任务窗格** 模型（非 VSTO 定制任务窗格 HKCU 直写），对 Excel 版本为**前向兼容**：只要 Excel 支持 COM 加载项即可运行。

- **支持 Excel 版本**：2013(15.0) / 2016(16.0) / 2019 / 2021 / Microsoft 365(16.x)，**32 位 / 64 位均可**。编译期引用 `office.dll 15.0.0.0`（Office 2013 互操作），运行时经 COM 自动绑定向后兼容，无需按新版 Excel 重编。
- **按 Excel 位数注册**：安装器/脚本用 `Registry64/Registry32` 分别写 `.NET COM 类 + Software\Microsoft\Office\Excel\Addins<ProgId>`，`LoadBehavior=3`（启动加载）。32/64 位 Excel 各加载对应架构程序集。
- **运行时依赖**：.NET Framework 4.8、WebView2 Runtime（WebView2 需运行在 Excel 进程内；Win10/11 一般自带，缺失时侧边栏空白，见运行日志/健康监控）。`GetVersion` 自检会上报 `excel_version`（`Application.Version`，15.0/16.0 等）供侧边栏「关于 / 版本」展示与排障。
- **受 Excel 版本影响的边界**：主要是**内置函数可用范围**（函数库按首版可用版本标徽章，XLOOKUP/动态数组/LAMBDA 等需 2021/365）。加载项自身能力在上述版本一致。
- **不引入超出运行环境的平台依赖**：全部写操作走自检兜底、可回滚、运行日志留痕（编码器红线），与 Excel 版本无关。

```text
┌────────────────────────────────────────────────────────────┐
│ Excel (x64 / x86)                                          │
│  ┌──────────────────────────────────────────────────────┐  │
│  │ ExcelAiAssistant.Addin  (COM Add-in 宿主)             │  │
│  │  · Ribbon 选项卡 / 侧边栏宿主 (TaskPaneHost)           │  │
│  │  · SidebarBridge  ←→ WebView2 页面 (index.html)        │  │
│  │  · DiagnosticsLog / SmokeRunner (运行日志与自检)        │  │
│  └───────────────┬──────────────────────────────────────┘  │
│                  │ 进程内托管调用                            │
│  ┌───────────────▼──────────────────────────────────────┐  │
│  │ ExcelAiAssistant.Core (Agent 引擎 / 配置 / 安全 / 技能) │  │
│  │  AgentRunner · ToolCatalog · OpenAiCompatibleClient   │  │
│  │  SafetyPolicy · UndoService · SkillCatalog · Workflow │  │
│  └───────────────┬──────────────────────────────────────┘  │
│                  │ 读取/写入/分析                          │
│  ┌───────────────▼──────────────────────────────────────┐  │
│  │ ExcelAiAssistant.Excel (140 代码注册工具服务层)              │  │
│  │  读取/编辑/数据/高级/图表/透视/导入/报告/解释/标记/     │  │
│  │  工具/验证/网页/工作簿/脚本 等 *Service + 注册表        │  │
│  └──────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────┘
   外部：模型服务(OpenAI 兼容 HTTP) · 本机文件 · 网页(WebFetch)
```

### 进程模型

- 一切工作都在 **Excel 进程内** 完成，直接持有并操作 `Application / Workbook / Worksheet / Range` COM 对象，无独立后台进程。
- 前台 UI 由 **WebView2** 承载，JS 与 C# 通过 `SidebarBridge` 桥接（postMessage / native 互操作）。
- 同一时间只加载与当前 Excel 位数匹配的程序集（由安装注册决定）。

---

## 二、目录结构与工程

```text
D:\execl1.0-wps
├─ src\
│  ├─ ExcelAiAssistant.Addin\      # COM 加载项宿主 + 侧边栏 UI
│  │   ├─ ExcelAiAddIn.cs / ExcelAiTaskPaneHost.cs / DeepSeekTaskPaneHost.cs
│  │   ├─ SidebarBridge.cs         # JS↔C# 桥
│  │   ├─ Sidebar\index.html       # 侧边栏单页 UI
│  │   ├─ DiagnosticsLog.cs / DiagnosticSmokeRunner.cs
│  │   └─ NativeConfirmDialog.cs
│  ├─ ExcelAiAssistant.Core\      # 引擎（无 UI 依赖，可单测）
│  │   ├─ Agent\        # AgentRunner / 协议解析 / 工具目录 / 会话
│  │   ├─ Configuration\ # AgentSettings / Configuration / Memory / WritePreference
│  │   ├─ Operations\    # OperationRecord / UndoPlanner / UndoService
│  │   ├─ Safety\?Security\ # SafetyPolicy / ControlledExecution / PathGuard
│  │   ├─ Skills\        # SkillCatalog
│  │   ├─ Tools\         # ToolRegistry / VbaPresetCatalog / VbaReusabilizer / BlankFill / Formula / Outlier / CustomToolStore
│  │   └─ Models\ Workflow\  # ModelProfile / ToolCall / WorkflowExecutor
│  ├─ ExcelAiAssistant.Excel\    # 全部 Excel 工具服务 + schema + 注册（140 代码注册工具（另含 ct_ 用户沉淀动态注册） / 18 个 verify_* 自检）
│  └─ ExcelAiAssistant.Tests\    # xUnit 单元测试（979 个）
├─ tools\
│  ├─ installer\         # JYYJ助手-Setup-wps.iss (Inno Setup 6) + register-auto.ps1 等注册脚本
│  ├─ release\           # release.py（bump/pack 唯一发版入口）、gate.py（一键门禁）、
│  │                     # check_wps_full.py（WPS full 基线校验）+ expected_wps_full.tsv
│  ├─ RealMachineVerify\ # Excel 真机验证 harness（23 用例组 119 断言）
│  └─ WpsToolVerify\     # WPS 真机验证 harness（full 基线 131 组）
└─ dist\                 # 构建产物与安装器（WPS 版产物一律带 wps 后缀）
```

依赖要点（见各 `.csproj`）：`Newtonsoft.Json`、`System.Data.SQLite`（含按架构的 `SQLite.Interop.dll`）、`WebView2Loader.dll`；目标框架为 .NET Framework 4.8。

---

## 三、加载项宿主（ExcelAiAssistant.Addin）

| 组件 | 职责 |
|------|------|
| `ExcelAiAddIn` | COM 加载项主入口，实现 `IDTExtensibility2`，挂接功能区与侧边栏生命周期 |
| `ExcelAiTaskPaneHost` | 侧边栏任务窗格宿主，承载 WebView2 |
| `DeepSeekTaskPaneHost` | 独立停靠浏览器窗格宿主（承载拒绝 iframe 的站点） |
| `SidebarBridge` | JS↔C# 双向桥：工具调用、配置读写、日志、确认信号；管理单轮工具调用上限（1–1000） |
| `NativeConfirmDialog` | 危险操作的原生确认对话框 |
| `DiagnosticsLog / DiagnosticSmokeRunner` | 运行日志登记与启动自检 |

侧边栏顶层视图：`聊天 / 公式 / VBA / 流程 / 设置`，所有交互以单页 `index.html` 承载。

---

## 四、Agent 引擎（ExcelAiAssistant.Core.Agent）

Agent 采用 **OpenAI 兼容的 Tool-Calling 协议** 编排多轮任务，核心类：

| 类 | 职责 |
|----|------|
| `AgentRunner` | 主循环：组装系统提示 → 调用模型 → 解析工具调用 → 分发执行 → 回填结果 → 直到完成 |
| `OpenAiCompatibleClient` | HTTP 客户端，对接任意 OpenAI 兼容接口 /{chat/completions} |
| `ToolCatalogBuilder / ToolSchemaBuilder` | 把工具注册表的 JSON Schema 汇总成模型可见的工具清单 |
| `AgentProtocolParser` | 解析模型返回的消息与工具调用（支持流式 / 思考步骤） |
| `AgentExperience` | 对话体验层（步骤折叠、工具事件视图） |
| `AgentToolEventView` | 工具调用事件的可视化数据（供侧边栏「思考过程」渲染） |
| `SessionStore` | 会话上下文持久化 |
| `ConfirmationDiffBuilder` | 生成危险操作前后差异，驱动确认面板 |
| `ModelProfile / ToolCall` | 模型配置与工具调用消息模型 |

### 执行循环

```text
for (limit = 配置的单轮上限) {
  steps = AgentRunner.RunOneTurn(ctx)        // 调用模型，得文本 + 工具调用
  if (!steps.ToolCalls) break                 // 无工具调用 → 结束
  foreach (call in steps.ToolCalls) {
    if (call.Risk == High && !UserConfirmed) { 等待确认; }
    result = ToolHost.Dispatch(call.Name, call.Args)   // 见下节
    ctx.AppendToolResult(call.Id, result)
  }
}
```

---

## 五、工具层（ExcelAiAssistant.Excel）

所有 Excel 能力封装为按风险分级、带 JSON Schema 的 **工具集**，由 `ExcelToolHost` 统一调度，`ToolRegistry` 建目录。各 `XxxToolsRegistration` 负责把对应 `XxxSchemas` 与实现注册进目录。

| 注册入口 | 能力域 | 实现服务 |
|----------|--------|----------|
| `ExcelDataToolsRegistration` | 读取 / 写入 / 范围 | `ExcelReadService / ExcelEditService / WorkbookService` |
| `ExcelAiFillToolsRegistration` | 智能填充空格 | `ExcelAiFillService + BlankFillAnalyzer` |
| `ExcelAdvancedToolRegistration` | 去重 / 排序 / 筛选 / 合并分列 | `ExcelAdvancedEditService` |
| `ExcelEnhanceToolsRegistration` | 格式 / 条件格式 / 表头 | `ExcelEnhanceService` |
| `ExcelExplainToolsRegistration` | 公式解释 / 引用解析 | `ExcelExplainService + FormulaAnalyzer` |
| `ExcelFlagToolsRegistration` | 异常值标记高亮 | `ExcelFlagCellsService + OutlierDetector` |
| `ExcelStructure / Pivot / Chart` | 结构整理 / 透视表 / 图表 | `ExcelStructureService / PivotTableService / ExcelChartService / ExcelPivotChartService` |
| `ExcelUtilityToolsRegistration` | 工具集（文本、统计等） | `ExcelUtilityToolsService + JyyjtoolAlgorithms` |
| `ExcelVerifyRegistration` | 自检（18 个 verify_*：写入/格式/条件格式/排序筛选/公式/图表/清洗/透视/标记/结构/迷你图/数据验证/合并/超链接/表格/图片/冻结窗格/保护态） | `ExcelVerifyService + ExcelVerifyJudgments` |
| `ExcelProfileRegistration` | 列质量画像（只读：空值/唯一值/错误值分布，Power Query 同口径） | `ColumnProfileEngine` |
| `ExcelWebToolsRegistration` | 网页取数 | `WebFetchService` |
| `ExcelExternalBookTools` | 多工作簿只读协同 | `ExternalBookRegistry` |
| 数据分析 / 报告 | 统计建模 / 文本报告 | `ExcelAnalysisService / ExcelReportService + ReportTextBuilder` |
| 脚本 / 高进度 | 脚本、外部进程、数据转换 | `ScriptRunnerService / PythonWorkerHost / DataTransformService / ExcelScratchpadService / ExcelImportService` |
| 看板渲染 / 截图 | 六类看板 + 超级看板 + 三维度截图 | `ExcelDashboardService`（部分方法经 dashboard 专用注册块，仅 Excel 宿主） |

### 5.1 截图链路（screenshot_dashboard）

只读工具，三个参数维度按优先级分流：`chart_name`（图表级）> `address`（任意选区级）> `dashboard_sheet`（看板全景级）；全部不传时缺省截**当前活动选区**（「选中哪儿截哪儿」）。

- **图表级**：`Chart.Export` 直接导出 PNG；存在性预检，未找到回 `CHART_NOT_FOUND` 人话错误。
- **选区级**：`CopyPicture(xlBitmap)` → 独立 STA 线程剪贴板直取 Bitmap → PNG（`address_clipboard` 路线）；失败降级「EMF 复制 + 临时 ChartObject 粘贴 + Export」（`address_chart_paste`）。
- **看板级**：截图区域 = `UsedRange ∪ 全部浮动对象`（图表/切片器在 UsedRange 之外，须扩展包围盒换算行列）；剪贴板直取失败降级同上（`chart_paste`）。
- **健壮性**（v3.7.1.65）：`CopyPictureWithRetry` 3 次退避重试（400/800ms，防前次剪贴板直取刚完成即 CopyPicture 的瞬态 `0x800A03EC` 竞争）；看板级首拷异常改走降级路线而非外抛；空白 fail-closed 防御扩展到**全部路线**（非白占比 <0.5% 判 `SCREENSHOT_BLANK` 并删除产物，绝不假成功）。
- **AI 读图闭环**：返回值含 `screenshot_path / route / non_blank_ratio / image_data_uri`（宽 ≤720px PNG base64 缩略图）；SKILL.md 契约要求 AI 读图自审（排版/遮挡/空白）→修复→重截复检后再交付。

> **关键约束**
>
> - 危险级别工具在 schema 中标注 `risk:high`，执行前必须经确认门。
> - 结果不“就地覆盖”时，必须显式把源区域地址设为写回地址；相对逻辑坐标叠加选区绝对原点（`range.Row/range.Column`）。
> - 公式引用解析需覆盖 `$A$1`、`C:C`、`'表'!B2`、`[Book]Sheet!A1` 等变体，R1C1 样式标记为去参数哨兵。
> - 工具扩展沿用“分期”演进（旧产物标记为 三期 = 导入/分析/可视化基础能力，四期 = 数据分析与报告生成）。

---


### 5.2 工具注册全景（178 个，按注册分块）

> 下表由注册代码直接生成（v3.8.1.83 口径），工具名/风险级/职责与代码一一对应，供排查「AI 能做什么」与工具数对账使用。看板工具块（render_* 等）按 D11 口径仅 Excel 宿主注册；WPS full 巡检对账基数见 `tools/release/check_wps_full.py` 的 EXPECTED_TOOL_COUNT。

#### 宿主注册（ExcelToolHost.CreateRegistry）：上下文/写入/图表入口/外部兜底

| 工具 | 风险 | 职责 |
|------|------|------|
| `append_excel_rows` | 写入 | 在锚点列末行下方追加二维数据（锚点列已有数据时自动向下续写，无需手算起始行）。 |
| `clear_excel_range` | 写入 | 清空指定区域的内容（保留格式与批注）。 |
| `format_excel_range` | 写入 | 设置数字格式、字体、填充颜色或自动列宽。 |
| `get_current_excel_context` | 只读 | 获取当前工作簿、工作表和选区信息。 |
| `preview_excel_range` | 只读 | 预览指定或当前选区的前几行和维度。 |
| `run_workflow` | 写入 | 执行预定义流程模板，按步骤顺序执行：参数覆盖、危险确认、单步失败可重试。 |
| `save_vba_tool` | 写入 | 把已编写并验证有效的 VBA 代码沉淀为可在今后会话直接复用的自定义工具。传入唯一工具名 name（字母开头，字母/数字/下划线，≤40 字符）、用途说明 description（AI 凭此判断何时调用）、完整的 VBA 代码 code。每次你用 run_vba_snippet 成功执行自写代码后都必须调用本工具沉淀，不要因代码“不够通用”而跳过——本工具会先征求用户确认，是否保存由用户决定。 |
| `undo_excel_operation` | 写入 | 按操作记录回滚指定区域写入。 |
| `web_search` | 只读 | 联网搜索（只读）：当分析需要行业基准/外部参考数据/事实背景（如行业平均良率、汇率、大宗商品价、政策事件），或用户要求搜索时调用。query 只传搜索关键词（禁止把工作簿原始数据当搜索词外发）。结果标注来源类型；引用外部结果时须向用户注明为公开网页信息，不得当作权威统计口径。 |
| `write_excel_cell` | 写入 | 向指定单元格写入值、文本或公式。 |
| `write_excel_range` | 写入 | 按区域批量写入二维数组数据。 |

#### 基础读写与结构（ExcelToolHostRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `add_excel_worksheet` | 结构 | 新增工作表，可指定名称和插入位置。 |
| `build_pivot_sheet` | 结构 | 分步生成透视分析 sheet（超级看板向导 Step2）：按勾选行维度与值字段构建可见、可拖拽的原生数据透视表，失败降级静态快照。与看板主链隔离。 |
| `check_chart_version` | 只读 | 图表版本徽章探测（只读）：检测当前 Excel 是否支持版本敏感图表（树状图/旭日图/漏斗图需 2021+，直方图/排列图/瀑布图需 2016+），不支持时给替代建议。 |
| `clean_excel_data` | 写入 | 清洗指定区域：去空行、修日期、填充空白。 |
| `create_chart_advanced` | 结构 | 技巧型图表族单入口（Excel 2013+）：type=bullet 子弹图/waffle 华夫图/slider 滑珠图/funnel 漏斗图/butterfly 蝴蝶图/gantt 甘特图；2016+ 原生图表：waterfall 瀑布图/treemap 树状图/sunburst 旭日图/pareto 排列图。 |
| `create_combo_chart` | 结构 | 创建双轴组合图（柱形主轴 + 指定系列折线次轴）。适用于量纲不同的双指标趋势对比（如销售额柱形 + 销量折线）。 |
| `create_dashboard_kpi_brief` | 结构 | 创建 KPI 速览看板（kpi_brief 模板）：隐藏 SUMIFS 辅助区 + 4-8 张 KPI 卡片（数值+环比位+IFERROR 兜底）。源数据只读；重生成幂等覆盖。 |
| `create_doughnut` | 结构 | 创建环形图（可选分离型，孔径可调 10-90）。适用于构成占比展示（≤5 类效果最佳）。 |
| `create_excel_chart` | 结构 | 根据区域创建基础图表；柱/条/折线/雷达在系列≤3 且每系列≤12 点时自动打开数据标签（饼/环显示百分比），data_labels=on/off 可强制覆盖，返回值 data_labels 字段回显实际状态。 |
| `create_excel_pivot_summary` | 写入 | 按指定标签列和数值列生成汇总透视结果。 |
| `create_excel_table` | 结构 | 把区域转为智能表（ListObject）：结构化引用、排序筛选、透选取数的基石。WPS 表格同样支持。 |
| `create_super_table` | 结构 | 超级数据表（看板数据底座）：把明细源规范化为一张可见 sheet——明细长表（类型规整：月份锁文本、文本数字数值化）+ 派生汇总区（KPI/月度/品类/地区，同表 SUMIF 公式），看板渲染与多切片器共享同一数据源。 |
| `create_truncated_column` | 结构 | 特大值截断柱：某值远超其余（>均值 N 倍）时截断显示并在柱顶标注真实值，避免其余分类被压扁。 |
| `delete_dashboard` | 高危 | 一键删除看板：删除看板 sheet 与隐藏辅助表并清理联动切片器。删除前回读断言看板结构，普通工作表会被拒绝（防误删）。 |
| `delete_dashboard_chart` | 结构 | 删单图：删除看板上指定图表对象（chart_name 逐字取自 list_dashboard_charts），强制数值守恒断言。切片器/下拉清理走 delete_dashboard。 |
| `delete_excel_worksheet` | 高危 | 删除指定工作表。 |
| `detect_dashboard_data_readiness` | 只读 | 数据门槛探测（生成看板前必跑）：合计行/合并单元格/文本型数字/空列硬拦截，ready=false 时给可读理由与清洗建议。只读不改数据。 |
| `edit_dashboard_theme` | 结构 | 看板换主题：只重刷卡片底色/字色等特征配色，不动数值、公式与图表（支持销售概览/KPI 摘要看板）。 |
| `filter_excel_data` | 写入 | 按列条件筛选数据，并把匹配结果写入目标区域。 |
| `freeze_panes` | 结构 | 冻结/解冻窗口窗格：锚点单元格上方所有行与左侧所有列被冻结（滚动表头/首列不动的标准做法）。 |
| `get_dashboard_sources` | 只读 | 看板数据源识别（只读）：列出当前工作簿内全部看板（看板 sheet + 隐藏辅助表 + 明细数据源引用）。AI 理解看板作用范围、回答「这个看板的数据来自哪里」时调用。 |
| `insert_image` | 结构 | 在指定锚点单元格左上角插入本地图片（随工作簿保存，不链接源文件）。WPS 表格同样支持。 |
| `list_dashboard_charts` | 只读 | 看板图表清单（只读）：枚举看板上全部图表（名称/类型/锚点/系列数）与当前主题。删图/换主题前必调。 |
| `manage_hyperlinks` | 写入 | 超链接管理：list 枚举整表超链接（可按 address 过滤）/add 新建（外部网址 link_url 或内部位置 link_subaddress 二选一）/update 修改/remove 删除（支持区域）/remove_all 清空整表。写动作前置工作表保护检查。 |
| `rename_excel_worksheet` | 结构 | 重命名工作表。 |
| `render_dashboard` | 结构 | 渲染 sales_overview 看板（视频同款）：地区下拉联动 + 4 KPI 卡 + 双轴趋势 + 品类环形，一次调用完整生成。 |
| `render_dim_cross` | 结构 | 渲染多维交叉看板：行×列维度交叉矩阵 + 双下拉联动（行/列筛选，(全部) 支持全量视图）+ 矩阵条形图。 |
| `render_pnl_brief` | 结构 | 财务损益看板：收入/成本/费用 KPI 卡 + 科目明细数据条 + 净利润。数据三列：科目/金额/类型（revenue/cost）。 |
| `render_rank_dynamic` | 结构 | 排名动态看板：时间下拉联动 + 榜首/Top N 合计/全量合计 KPI 卡 + Top N 排行条形图。数据三列：类别/时间/数值。 |
| `render_super_dashboard` | 结构 | 渲染超级看板（三层结构一键生成）：自动建超级数据表底座 + 汇总透视层 + 固定四宫格看板（5 张形状 KPI 卡随指标视图切换、品类/地区下拉联动、透视分析可选）。支持 extra_kpis 派生指标（如利润占比=销售总额/销售量）。生成后须调 verify_dashboard_deep 核验与 screenshot_dashboard 截图。 |
| `render_target_track` | 结构 | 目标达成看板：总体达成率/达标项数/最佳指标 KPI 卡 + 逐指标完成率数据条。数据三列：指标/实际/目标。 |
| `screenshot_dashboard` | 只读 | 截图工具（只读）：把看板整体、指定图表或任意选区（address）导出 PNG，返回文件路径与像素统计，供用户直接查看视觉效果。用户要求截图/预览/看看效果时必须调用本工具；生成或修改看板后也应调用并把路径告知用户。AI 无法直接看图，视觉判断以像素统计+用户反馈为准。 |
| `set_linkage` | 结构 | 为已有看板设置联动筛选：native 模式创建切片器（多选/直观），cell 模式创建下拉（兼容 Excel 2013）。须先 render_dashboard 生成看板。 |
| `suggest_chart_mapping` | 只读 | 图表选型建议（只读启发式）：分析选区各列画像（数值比/去重基数/日期特征/负值），给出图表或看板工具推荐清单，每条附理由与数据要求。生成图表/看板前先调用，把建议复述给用户并确认数据源/行/列映射。 |
| `verify_dashboard` | 只读 | 看板自检：看板 sheet 与隐藏辅助区是否存在、辅助区 KPI 重算值是否全部为有效数值（守恒对账）。fail-closed。 |
| `verify_dashboard_deep` | 只读 | 看板深度核验（只读）：六项检查各带期望/实际/依据——看板与辅助表存在、辅助表 KPI 重算全数值、看板卡片显示值守恒（显示值==辅助表值）、联动控件、数值守恒对账（全量视图 KPI==源列直和）、图表渲染像素核验。最终回复必须逐项转述。 |

#### JYYJ 实用工具（ExcelJyyjtoolRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `amount_to_uppercase` | 写入 | 把数值金额转换为中文大写金额（壹贰叁…元角分）。 |
| `bankers_round_values` | 写入 | 按银行家舍入（四舍六入五取偶）保留指定小数位。 |
| `compare_two_sheets` | 只读 | 逐格对比两张表，输出新增/删除/变更差异。 |
| `convert_1d_to_2d` | 写入 | 一维数据按每行固定列数折叠为二维表。 |
| `convert_2d_to_1d` | 写入 | 宽表转长表：保留前 N 列标识，其余列拆成多行（逆透视）。 |
| `convert_chinese_variant` | 写入 | 简繁中文之间互相转换。 |
| `convert_text_to_pinyin` | 写入 | 把中文文本转成拼音全拼或首字母。 |
| `copy_range_as_image` | 写入 | 把区域复制为图片（静态快照）并粘贴到锚点。 |
| `copy_visible_range_only` | 写入 | 仅复制筛选后可见的行到目标区域，跳过隐藏行。 |
| `create_pivot_table` | 写入 | 对明细数据做轻量透视（按行字段分组并对数值列求和/计数/平均）。 |
| `create_sheet_by_value` | 结构 | 以指定单元格的值作为新工作表明称（可复制某表内容）。 |
| `cross_reference_columns` | 写入 | 跨区键值匹配回填：以源区域键列去查找区域匹配，命中则把查找区域值列的值回填到源区域目标列（免 VLOOKUP 公式）。 |
| `delete_blank_rows_range` | 结构 | 删除区域内整行为空或首列为空的行。 |
| `delete_text_matches` | 写入 | 按正则表达式删除匹配文本，其余保留。 |
| `extract_text_characters` | 写入 | 提取单元格文本中的指定类型字符（中文/英文/数字/字母数字/特殊符号）。 |
| `fill_blank_cells_down` | 写入 | 向下填充空单元格：沿用上方同列非空值（或无公式）。 |
| `fill_number_sequence` | 写入 | 向下或向右填充公差为 step 的等差序列。 |
| `fill_random_numbers` | 写入 | 向指定区域填充随机数，可设值域与小数位。 |
| `fuzzy_match_columns` | 只读 | 两列做相似度模糊匹配，返回最佳配对与阈值判定。 |
| `generate_permutation_list` | 写入 | 生成 n 个元素的排列（或两两组合）并写入单元格。 |
| `insert_blank_rows_every_n` | 结构 | 每隔 N 行插入若干空行（工资条空格/分段）。 |
| `insert_text_into_cell` | 写入 | 按字符位置或某个分隔符第 N 次出现的位置，在文本中插入子串。 |
| `make_payslip_rows` | 写入 | 把表头与数据区生成为工资条：每行数据上方重复表头。 |
| `merge_duplicate_cells` | 结构 | 合并一列内相邻的相同单元格（垂直合并同类项）。 |
| `pick_numbers_to_sum` | 只读 | 凑数：从候选数字中找一组合尽量接近目标值（支持精确组合与贪心近似）。 |
| `protect_sheet_watermark` | 结构 | 锁定/解锁工作表，或在指定单元格写入水印文本。 |
| `smart_paste_delimiter` | 写入 | 智能粘贴：把一列含分隔符的文本拆成多列（自动识别分隔符）。 |
| `split_sheet_by_field` | 结构 | 按字段取值把一张表拆分成多个新工作表（各含子集数据）。 |
| `wrap_formulas_with_iferror` | 写入 | 给区域内的公式加上 IFERROR 包裹，避免错误值残留。 |

#### 第三阶段：导入导出/分析/清洗/校验（ExcelThirdPhaseRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `add_insight_card` | 写入 | 看板洞察卡标准工具：把分析结论生成为圆角矩形洞察卡（标题加粗+正文白字，五色主题可选），落在指定工作表的锚点单元格处，卡高按文字量自动估算（正文超 8 行截断省略）。写入后自动读回文本校验（verdict=ok）。卡片文字必须来自工具真实分析结论（禁止编造数值）；replace_existing=true 可清空该表旧洞察卡重写（防反复调用堆积）；卡片是展示形状，随删表移除、不参与单元格快照回滚。适合配合情感分析/假设检验/聚类画像的结论做看板可视化。 |
| `analyze_cluster` | 只读 | 聚类分群标准工具（只读）：K-Means++ 初始化 + Lloyd 迭代 + Z-score 标准化，k 在 2~6 内按平均轮廓系数自动选定（可显式指定 k），返回各簇画像（样本数/特征均值）与逐行簇分派。完整观测行参与聚类（缺失行剔除并计数）；分派明细超 1000 行时省略并提示。DBSCAN/层次聚类等走 python-advanced-analysis 技能。 |
| `analyze_column_correlation` | 只读 | 计算数值列两两之间的相关系数（皮尔逊或秩相关）。成对完整观测（两列同一行均为数值），返回附 basis 计算依据（方法/口径/CORREL 公式示例/成对样本数 n）。 |
| `analyze_column_profile` | 只读 | 列质量画像（Power Query 式单面板）：一次给出该列的质量（空值率/错误率/唯一值/类型占比）、分布（Top N 高频值）与剖析（数值列 min/max/均值/中位数/标准差；文本列长度与空白统计），并附质量提示（高空值/含错误/混合类型/带空格/恒定值）。数据清洗与选型前建议先对目标列跑一次。返回附 basis 计算依据（来源列区间/明细行区间/质量与统计口径）。 |
| `analyze_hypothesis_test` | 只读 | 假设检验标准工具（只读）：四类常用检验——two_sample_t（Welch 双样本 t）/paired_t（配对 t）/anova（单因素方差分析 F）/chi_square（卡方独立性）。返回统计量、自由度、精确 p 值、显著性判断（可调 α）、效应量（Cohen's d/η²/Cramér's V）、分组摘要与前提提醒（正态性/方差齐性/期望频数），附 basis（T.TEST/F.DIST.RT/CHISQ.TEST 公式示例）。统计口径纯 C# 可复算；非参检验与复杂设计走 python-advanced-analysis 技能。 |
| `analyze_range_statistics` | 只读 | 对数据区域做数值/文本列的统计摘要（均值、中位数、标准差等）。返回附 basis 计算依据（来源区域/明细行区间/统计口径/公式示例）。 |
| `analyze_regression` | 只读 | 多元线性回归标准工具（只读）：普通最小二乘（OLS）拟合因变量与一个或多个特征列，返回系数表（估计/标准误/t/p）、R²/调整 R²/F 整体检验、残差标准误/Durbin-Watson/VIF/Cook's D 诊断与逐行拟合值/残差明细（超 1000 行省略）。完整观测行参与回归（缺失行剔除并计数）；完全共线或常数列 fail-closed 并提示。统计口径与 statsmodels 对齐、纯 C# 可复算；非线性/逻辑回归/正则化走 python-advanced-analysis 技能。 |
| `analyze_text_variants` | 只读 | 文本列变体聚类建议标准工具（只读，探索式清洗）：对文本列（如客户名/供应商名/品类）输出疑似变体组——空白/大小写/全半角归一化差异与编辑距离≤阈值的错字/后缀差异各成一组，每组给代表值、变体成员（含出现次数与匹配类型）与合并方向选项（取最长/取最短/取出现最多）。本工具只建议不合并：用户确认后合并动作须用 write_excel_cells_batch 将变体单元格改写为代表值（改值语义，走确认门可回滚；不得用 dedup_range——其语义为删除重复行）。唯一值超 1500 或需进阶模糊匹配走 python-advanced-analysis 技能。 |
| `analyze_trend_forecast` | 只读 | 轻量趋势预测（只读）：对时间/序列标签列+数值列做线性回归或移动平均外推，返回未来 N 期预测值、趋势方向、斜率与拟合优度 r2，附 basis 计算依据（样本区间/SLOPE/INTERCEPT/FORECAST.LINEAR 公式示例）。有效数值点不足 6 个时拒绝预测（INSUFFICIENT_POINTS）；季节性 Holt-Winters/ARIMA 属高阶预测请走 python-advanced-analysis 技能（statsmodels）。 |
| `analyze_value_distribution` | 只读 | 统计某列取值分布与出现次数最多的前 N 个值。返回附 basis 计算依据（计数区间/口径/COUNTIF 公式示例）。 |
| `create_excel_sparkline` | 写入 | 在目标单元格插入迷你图（折线/柱状/胜负）。 |
| `diagnose_chart` | 只读 | 图表读回诊断（只读）：无期望全面体检——类型/系列/点数/空系列/数据源/标题/渲染像素核验/数据守恒对账，输出根因归类（root_cause）、诊断依据（findings）、修复建议（suggestions）与渲染截图（chart_image 供 AI 读图核验），诊断全程记录。生成图表后自检或排查图表问题时调用。 |
| `export_excel_chart_image` | 写入 | 把图表导出为 PNG 图片文件。 |
| `generate_analysis_report` | 写入 | 对数据区域生成模板化文字报告，可输出到工作表、Markdown 或 HTML。 |
| `import_delimited_file` | 写入 | 把分隔文本文件导入到工作表的指定区域。 |
| `import_excel_file_sheet` | 结构 | 导入另一个 Excel 文件的工作表（作为新表或取值复制）。 |
| `insert_recommended_chart` | 结构 | 按数据形状一键插入推荐图表（自动选型/坐标列/中文标题）。 |
| `merge_delimited_files` | 写入 | 把多个分隔文本文件按表头合并输出到一个区域。 |
| `preview_external_file` | 高危 | 预览外部文本文件前若干行，自动识别分隔符与编码。 |
| `read_business_glossary` | 只读 | 读取当前活动工作簿同目录的《_口径.md》业务口径规则文件（Markdown，每行一条「- 术语: 定义」，随业务文件走）。生成看板/分析结论/统计摘要前建议先调用本工具；回复中涉及口径的关键数字必须逐条引用条目原文；口径与用户当次指令冲突时用户指令优先，但必须向用户明示差异，不得静默择一。找不到文件时返回创建模板。隐私红线：文件仅本地读取，绝不外发。 |
| `recommend_excel_chart` | 只读 | 根据数据形状推荐合适的图表类型与坐标列。 |
| `repair_chart` | 结构 | 图表读回修复：按诊断结论修复——空系列/点数错乱重绑源数据、117~123 原生新图临时表重建、补标题；修复后自动复检并返回前后对比与截图（AI 读图）。与 diagnose_chart 配套使用。 |
| `verify_chart` | 只读 | 图表自检：校验图表是否真正生成、数据系列/点数是否有效、数据源区域（来自 series 公式）与预期类型/标题/数据源/系列数是否相符，并对渲染图做像素级核验“数据是否真的显示出来”，返回渲染图片供核对。生成任何图表后都应调用一次。 |

#### 第二阶段：Python 通道/透视/图片（ExcelSecondPhaseRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `apply_page_setup` | 写入 | 设置工作表的打印页面布局。 |
| `convert_formula_references` | 写入 | 切换公式中的绝对/相对引用。 |
| `copy_visible_rows_to_sheet` | 结构 | 把筛选后的可见行复制到目标区域。 |
| `drill_excel_range` | 只读 | 下钻单元格：读取其显示值、公式与引用来源预览。 |
| `export_selection_to_markdown` | 只读 | 把选区导出为 Markdown 表格文本。 |
| `export_sheets_to_pdf` | 写入 | 把整本或指定工作表导出为 PDF。 |
| `import_markdown_to_excel` | 写入 | 把 Markdown 表格文本导入为 Excel 区域。 |
| `install_python_packages` | 高危 | 常用 Python 库一键安装（pandas/matplotlib/openpyxl/xlwings 等，白名单+国内镜像兜底）。必须先把 python_setup_guide 返回的库目录与用途说明展示给用户、等用户明确选择要装哪些后再调用。 |
| `list_external_links` | 只读 | 列出当前工作簿的外部链接与查询连接。 |
| `list_python_packages` | 只读 | 列出当前 Python 环境已安装的库（名称/版本/安装位置）与 Python 解释器路径、版本、位数。用户问装了哪些库/库在哪时调用。 |
| `merge_or_split_cells` | 结构 | 合并/取消合并区域，或按相同内容合并、拆分并填充。 |
| `merge_workbooks` | 结构 | 把指定工作簿的各工作表合并进当前工作簿。 |
| `python_setup_guide` | 只读 | Python 环境检测与小白级安装教程：当前 Python 是否安装/位数是否匹配 Excel、逐步安装引导、常用库目录（每个库做什么的详细说明与安装命令）。用户环境缺 Python/缺库或询问如何装库时必调。 |
| `replace_external_links` | 高危 | 替换工作簿中的外部链接。 |
| `reset_cell_value_type` | 写入 | 把区域单元格重置为文本、数字或日期类型。 |
| `run_cli_command` | 高危 | 在受控白名单下运行 CLI 命令。 |
| `run_exe_command` | 高危 | 在受控白名单下运行可执行文件。 |
| `run_python_code` | 高危 | 在受控 Python（常驻 worker）下执行纯计算/数据处理代码，结果以 stdout 文本返回。数据通道：可传 input_data（JSON 对象/数组/标量）注入命名空间变量 input_data 直接使用；需要结构化结果时调用 xai_result(obj)（支持 dict/list/标量，numpy/pandas 自动转换），执行器以 result 字段返回 JSON（示例：xai_result({'got': input_data['x']}) → result={\"got\":...}；结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：persist=true 保留命名空间跨调用（变量/自定义函数可复用；注意 run_python_code 与 run_python_excel 共享同一 worker 的 persist 命名空间，两者变量互通）；xai_state 为常驻 dict（无论 persist 与否都保留）；state_clear=true 一并清空命名空间与 xai_state；返回的 state_reset=true 表示 worker 刚重建/超时重启/被清空，旧状态已丢失，勿假定仍存在。状态仅存内存、进程退出即清、不落盘。约束：① 默认 30 秒超时（可传 timeout_ms 放宽，上限 300 秒），严禁死循环、阻塞等待 input()、访问网络或超大迭代（如 range(1e9)），否则必然超时并重置 worker；② 只做纯 Python 计算，超大输出会被截断；③ 若要把结果写回 Excel 单元格，请改用 run_vba_snippet 用数组一次性写回，不要在 python 里拼大字符串再回填；④ 序列/字典下标访问前先做 len()/in 边界检查（IndexError/KeyError 高频源，行数据列数不足时先截断再处理）。 |
| `run_python_excel` | 高危 | 在受控 Python（常驻 worker + xlwings）中直接读写当前 Excel 活动工作簿。已自动前置 import xlwings as xw；用 xw.apps.active 连当前实例、xw.books.active 定位活动簿（示例：xw.books.active.sheets[0].range(\"A1\").value = 42）。数据通道：可传 input_data（JSON 对象/数组/标量）注入 input_data 变量；需要结构化结果时调用 xai_result(obj)，执行器以 result 字段返回 JSON（结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：persist=true 保留命名空间跨调用（与 run_python_code 共享同一 worker 的 persist 命名空间，变量互通）；xai_state 为常驻 dict（无论 persist 与否都保留）；state_clear=true 一并清空命名空间与 xai_state；返回的 state_reset=true 表示 worker 刚重建/超时重启/被清空，旧状态已丢失，勿假定仍存在。状态仅存内存、进程退出即清、不落盘。安全契约：本工具可绕过操作记录直写 Excel，风险与 run_vba_snippet 同级（Dangerous + 确认门 + 受控执行总开关）；首次使用自动探测 Python 版本/位数（要求 ≥3.9 且与 Excel 位数一致，不符给出安装指引）并自动 pip 安装缺失依赖 xlwings/pywin32（镜像兜底 → --user 兜底 → 装后 import 自检，失败分级提示）；AI 声明 write_ranges 时执行前快照该区域、执行失败自动回滚；若只需纯计算请用 run_python_code（低风险）。约束：默认 30 秒超时（可传 timeout_ms 放宽，上限 300 秒），严禁死循环/阻塞/访问网络；代码必须能在数秒内自行终止。另：openpyxl/xlw 磁盘引擎只处理未打开的文件，要改当前已打开的工作簿必须走 xw.books.active 直写（对已打开文件 to_excel/save 会报 PermissionError——Excel 进程持锁）。 |
| `run_vba_snippet` | 高危 | 当现有工具无法完成用户需求时，编写 VBA 代码片段用本工具执行以达成目的（万能兜底工具）。代码必须直接操作当前 ActiveWorkbook/ActiveSheet/Selection，逻辑自包含、可独立运行，且必须能在数秒内自行终止（本工具无外部超时中断，死循环会一直运行直到手动中断），并在执行前先向用户展示代码确认。参数 snippet 为完整 VBA 代码（注意：VBE 代码模块源码在中文系统按 ANSI/GBK 存储，snippet 与 args 内请使用 ASCII/中文等系统代码页可表示字符，emoji 等非 BMP 字符注入后会被替换为 ? 而损坏）。数据通道：可传 args（JSON 对象/数组/标量，支持嵌套）注入模块级 gXaiArgs，代码内用 XaiArg(\"key\") 读取（返回字符串/数字/布尔/数组，缺键返回 Null）；需要把计算结果结构化交回时，调用 XaiReturn 值（支持标量/数组/字典），执行器以 result 字段返回 JSON（示例：XaiReturn Array(v, 42) → result=[\"A\",42]；结果超 result_limit_kb 时 result 为截断字符串预览且 result_truncated=true）。持久状态：用 XaiStateSet/XaiStateGet 在多次调用间保留键值状态，state_clear=true 显式清空；返回的 state_reset=true 表示状态模块刚重建、旧状态已丢失（工作簿首次调用或升级后），勿假定旧状态仍存在。状态仅存内存、Excel 关闭即清、不落盘。铁律（违者会把整段代码判定非法驳回）：① 大批量读写一律用数组/Range.Value2 整体赋值（先 arr=Range.Value2，处理数组后再一次写回），严禁在 For/While 循环里用 Cells()/Range().Value 逐格读写；② 严禁对多行多列矩形 Range 直接调用 .AutoFit（如 ws.Range(\"G12:K18\").AutoFit 运行时会抛 1004），改按行/列调用 .Rows(n:m).AutoFit/.Columns(x:y).AutoFit 或设 RowHeight/ColumnWidth=0；③ 开头关闭 Application.ScreenUpdating 与 Calculation=Manual，结束前恢复；④ 成功执行后必须再调用 save_vba_tool 沉淀。 |
| `split_workbook` | 结构 | 把当前工作簿按工作表拆分为独立文件。 |
| `trace_formula_references` | 只读 | 定位公式直接引用（前导项）的单元格与值。 |
| `transform_excel_formulas` | 写入 | 对区域公式批量应用 IFERROR/IFNA/ROUND 等转换。 |
| `validate_vba_snippet` | 写入 | 校验 VBA 代码片段语法/结构。 |

#### 校验族（ExcelVerifyRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `verify_cell_format` | 只读 | 单元格格式自检：读取目标区域的数字格式/对齐/字体（名称/字号/颜色/加粗/斜体）/边框/填充/锁定属性，与 AI 操作时传入的预期值（expected_*）逐项比对，返回 verdict/issues/selfcheck。设置任何单元格格式后都应调用一次。 |
| `verify_conditional_format` | 只读 | 条件格式自检：读取目标区域的条件格式规则（类型/公式/运算符/填充色/字体色），与预期值（expected_rule/expected_formula1/expected_operator 等）比对，返回 verdict/issues/selfcheck。设置任何条件格式后都应调用一次。 |
| `verify_data_clean` | 只读 | 数据清洗自检：对区域身体行（跳过表头）逐检查项扫描——no_blank_rows(无整行空白)/no_blanks(指定列无空白)/no_duplicates(无重复记录)/dates_valid(日期列文本可解析)/no_text_numbers(无文本型数字)，任一发现问题即报 issue，返回 verdict/issues/selfcheck。执行 clean_excel_data/dedup_range 等清洗后都应调用一次。 |
| `verify_data_validation` | 只读 | 数据验证自检：读取目标区域的验证规则（类型/来源/运算符/边界），与预期（expected_type/expected_source/expected_operator/expected_value1/value2）比对，返回 verdict/issues/selfcheck。执行 data_validation 后都应调用一次。 |
| `verify_formula` | 只读 | 公式自检：检查目标单元格是否含公式、是否返回错误值（#REF!/#VALUE! 等）、公式内容与预期是否一致、计算值与预期是否相符（数值按容差比较），返回 verdict/issues/selfcheck。写入任何公式后都应调用一次。 |
| `verify_freeze_panes` | 只读 | 冻结窗格自检：读取目标工作表视图的冻结状态与分割行列号，与 expected_frozen/expected_cell 比对，返回 verdict/issues/selfcheck。执行 freeze_panes 后都应调用一次。 |
| `verify_hyperlink` | 只读 | 超链接自检：枚举工作表超链接（锚点/网址/内部位置/显示文本），按 expected_count 核对整表总数、按 address 定位目标单元格/区域并按 expected_url/expected_subaddress/expected_text 核对内容，或按 expected_present=false 断言删除生效，返回 verdict/issues/selfcheck。执行 manage_hyperlinks 的写动作后都应调用一次。 |
| `verify_image` | 只读 | 图片自检：枚举工作表中的图片（不含图表/文本框），按 expected_count 核对数量、按 anchor 定位锚点单元格内的图片并按 expected_width/expected_height 核对尺寸（磅，容差 2），返回 verdict/issues/selfcheck。执行 insert_image 后都应调用一次。 |
| `verify_marker` | 只读 | 标记自检：扫描目标区域单元格的非默认填充色（被标记），按 expected_count 核对标记总数、按 expected_fill_colors 核对实际标记色集合、按 expected_marked 核对期望被标记地址是否命中，返回 verdict/issues/selfcheck。执行 flag_cells 等标记操作后都应调用一次。 |
| `verify_merge` | 只读 | 合并状态自检：读取目标区域 MergeCells 实际状态（true 整区已合并/false 无合并/null 混合），与 expected_merged 比对，返回 verdict/issues/selfcheck。执行 merge_or_split_cells/merge_duplicate_cells 后都应调用一次。 |
| `verify_pivot` | 只读 | 透视表自检：读取透视表汇总输出区域（首行为表头 Row Labels/Total），按 expected_rows 中的标签逐行抽查汇总值，并可按 expected_count 核对数据行数，返回 verdict/issues/selfcheck。执行 create_excel_pivot_summary 后都应调用一次。 |
| `verify_protection` | 只读 | 保护状态自检（只读探测）：回报工作簿结构保护/工作表保护/单元格锁定与筛选排序许可，支持 expected_protected 与 expected_writable 期望比对；写入被拒（0x800A03EC）时的首选排查工具。 |
| `verify_sort_filter` | 只读 | 排序筛选自检：基于目标区域数据矩阵的实际值验证——检查排序列（expected_sort_column/direction）从表头行起的身体行是否按预期方向单调，以及每个数据行是否都满足筛选谓词（expected_filter_field/criteria1/operator）。与内存 sort_range / filter_excel_data 工具语义一致（不依赖原生 Sort/AutoFilter 对象）。返回 verdict/issues/selfcheck。执行排序或筛选后都应调用一次。 |
| `verify_sparkline` | 只读 | 迷你图自检：读取目标区域的 SparklineGroups（类型/落点/数据源），与预期（expected_type/expected_data_source）比对，返回 verdict/issues/selfcheck。执行 create_excel_sparkline 后都应调用一次。 |
| `verify_structure` | 只读 | 结构自检：读取工作簿工作表名单与可选区域行列数，与期望比对（expected_sheet_exists/expected_sheet_absent/expected_row_count/expected_column_count），返回 verdict/issues/selfcheck。新增/重命名/删除工作表等结构变更后都应调用一次。 |
| `verify_table` | 只读 | 智能表自检：按名称或地址定位 ListObject，核对存在性/表头/行列数（expected_headers/expected_row_count/expected_column_count），返回 verdict/issues/selfcheck。执行 create_excel_table 及拆表/建表类操作后都应调用一次。 |
| `verify_write_result` | 只读 | 写入结果自检：读取目标区域实际值，与预期值矩阵（expected_values）逐格比对（数值按容差、文本忽略大小写与空白），并可按 expected_count 核对非空单元格数，返回 verdict/issues/selfcheck。执行 write_excel_cell/write_excel_cells_batch/write_excel_range 等写入后都应调用一次。 |

#### 高级工具（ExcelAdvancedToolRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `append_calculation_scratchpad` | 写入 | 向 AI 计算草稿追加输入、公式与期望输出并回填结果。 |
| `clear_calculation_scratchpad` | 结构 | 清空 AI 计算草稿内容。 |
| `copy_paste_excel_range` | 结构 | 在工作簿内复制并粘贴区域。 |
| `create_toc_sheet` | 写入 | 生成包含各工作表链接的工作簿目录页。 |
| `delete_excel_dimension` | 高危 | 删除行或列，涉及数据破坏请谨慎使用。 |
| `find_excel_cells` | 只读 | 在指定区域查找匹配的值或公式单元格。 |
| `get_current_selection_values_and_formulas` | 只读 | 读取当前选区的数值、公式与数字格式。 |
| `get_excel_range_values_and_formulas` | 只读 | 读取区域的数值、公式与数字格式；超过阈值仅返回摘要。 |
| `get_excel_sheet_list` | 只读 | 列出工作簿中所有工作表的名称、可见性与 UsedRange 概览。 |
| `insert_excel_dimension` | 结构 | 插入行或列。 |
| `manage_excel_filter` | 写入 | 按列条件过滤（隐藏不匹配行）或清除该过滤；基于行隐藏模拟，不产生原生 AutoFilter。 |
| `manage_excel_named_range` | 写入 | 查询、创建、更新或删除命名区域。 |
| `preview_excel_sheet` | 只读 | 预览工作表的 UsedRange 维度、前几行、尾部几行与列类型推断。 |
| `select_excel_range` | 只读 | 选中并激活指定区域。 |
| `write_excel_cells_batch` | 写入 | 向多个不连续单元格批量写入值或公式。 |

#### 实用工具（ExcelUtilityToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `date_normalize` | 写入 | 把一列日期统一为标准格式（自动识别或指定输入格式，支持 Excel 序列日期），无法解析的可替换为指定文本；写入结果并记录撤销。 |
| `fill_formula` | 写入 | 向区域内批量填充公式模板（用 {row} 表示当前行号自动替换），整块可撤销；适合按行批量套用公式。 |
| `merge_sheets` | 写入 | 把多个工作表的数据纵向合并到目标工作表（可选带表头，只取第一个源表的表头），适合月度/分部门台账汇总。 |
| `number_pad` | 写入 | 把一列编号补零/加前后缀格式化为定长文本（如 1 → NO-000001-A），适合规范台账编号；写入结果并记录撤销。 |
| `text_clean` | 写入 | 对区域内文本做清洗：去首尾空格、合并连续空白、去控制字符、统一大小写、删除指定字符、查找替换；写入结果并记录撤销。 |
| `text_merge` | 写入 | 把多列文本按分隔符合并到一列（如姓+名拼全名、多列地址拼完整地址），可跳过空单元格；写入结果并记录撤销。 |
| `text_split` | 写入 | 把一列文本按分隔符拆分为多列（如姓名拆姓/名、地址拆省市区），可保留或清空源列；写入结果并记录撤销。 |

#### 数据工具（ExcelDataToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `dedup_range` | 写入 | 对区域内数据去重：按整行或指定列判重，保留首次/末次出现，其余重复行删除并向上压缩，尾部补空；整块可撤销。 |
| `find_replace_range` | 写入 | 在区域内批量查找并替换文本（可区分大小写、可整格匹配），替换单元格数可控并记录撤销。 |
| `format_excel_conditional` | 写入 | 对区域应用条件格式：大于/小于/介于/重复值高亮或数据条，可用 rule=clear 一键清除规则。 |
| `sort_range` | 写入 | 对区域内数据按一列或多列排序（数值/文本智能比较，升/降序），表头保持不动；写入结果并记录撤销。 |
| `transpose_range` | 写入 | 把区域行列互换（转置），可选择写入当前选区左上角或指定目标位置。 |

#### 增强工具（ExcelEnhanceToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `cell_comment` | 写入 | 管理单元格批注：get 读取、set 写入、clear 清除一个区域的批注。 |
| `data_validation` | 写入 | 设置数据有效性：下拉列表（list，source 为单元格范围或逗号列表）或整数区间校验（interval）。 |
| `rate_change_analysis` | 写入 | 读取 label+value 两列，逐行计算绝对变化与百分比变化（环比），可选将结果写回；适合时间/周期序列分析。只读分析，写回需显式 output_address。 |

#### 智能填充（ExcelAiFillToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `ai_fill_blank_cells` | 只读 | 只读分析一个区域（含表头+数据）里的空白单元格：返回空缺清单、按列分组的表头与非空样本、以及可参考的其他数据样本，不写入任何内容。供模型做语义推理后逐格填出建议值。 |
| `apply_fill_blank_cells` | 写入 | 把确认后的建议值就地回填到指定区域的空白单元格（非空白格一律不动，会被拒绝），整块写回并记录可撤销。需开启受控执行总开关，否则拒绝。 |

#### 外部工作簿（ExcelExternalBookTools）

| 工具 | 风险 | 职责 |
|------|------|------|
| `close_external_book` | 只读 | 显式关闭外部只读句柄并释放资源。 |
| `open_excel_external_book` | 高危 | 以只读句柄打开指定路径的工作簿，返回句柄与工作表清单，供后续 read_external_range 引用。只读，不会改动该工作簿。 |
| `read_external_range` | 只读 | 读取外部只读句柄中指定工作簿、工作表、区域的值，作为上下文引用。 |

#### 公式解释（ExcelExplainToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `explain_excel_formula` | 只读 | 只读解释一个公式：真机拆解其函数/参数/引用单元格的实际值，返回结构化信息供转述。无法真机求值的引用（跨表/UDF/数组等）会显式标记为降级，不会编造解释。 |

#### 标记工具（ExcelFlagToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `flag_cells` | 写入 | 对区域内数据做标记式清洗识别：异常值(IQR/Z-score)、格式异常(文本型数字/坏日期)、空值；在源区用异色高亮，并生成「异常标记诊断」表（含行/列/类型/原因/建议），诊断表可整体删除并恢复。只标记不修改数据。 |

#### 网页抓取（ExcelWebToolsRegistration）

| 工具 | 风险 | 职责 |
|------|------|------|
| `fetch_url` | 只读 | 只读抓取一个 http/https 网页或文本 URL，返回其内容文本（自动转纯文本，默认脱敏隐藏邮箱/手机/身份证）供分析；不写工作簿、不执行任何代码、不访问本地文件。 |
## 六、安全与撤销

安全由 `ExcelAiAssistant.Core.Safety` / `Security` 与 `Operations` 三部分协同：

| 模块 | 机制 |
|------|------|
| `SafetyPolicy` | 统一风险判定：危险工具强制确认门 |
| `PreRunVbaGate` | VBA 执行前静态护栏，拦截“多行多列矩形直接 AutoFit”等致 1004 的写法 |
| `ControlledExecution (Config/Policy/Store)` | 脚本 / CLI / EXE 白名单制，默认关闭，黑名单直接拒绝 |
| `PathGuard` | 文件路径安全校验，防越权访问 |
| `ConfirmationDiffBuilder` | 对整行删除等高风险操作生成变更差异供用户确认 |
| `UndoPlanner / UndoService / OperationRepository` | 操作记录与结构化撤销，支持可逆回滚 |

> **实现注意（RCW 生命周期）**：COM 的 .NET RCW 按底层对象标识缓存；同一 Worksheet 无论从哪条路径解析均为同一 RCW。工具对解析出的 sheet 在 finally 调 `FinalReleaseComObject` 会连带销毁调用方持有的同名 RCW，导致后续访问抛「COM 对象与其基础 RCW 分开」。因此生产工具必须**无状态**：每次操作前重新解析工作表、用完即放，绝不跨调用持有 Worksheet RCW。

---

## 七、对话、偏好与配置

| 类 | 职责 |
|----|------|
| `ConfigurationStore` | 本地配置读写（模型、偏好等） |
| `AgentSettingsStore` | Agent 级设置（模型方案、API Key、Base URL） |
| `MemoryStore` | 长期记忆：跨会话偏好与备注，注入每次系统提示 |
| `WritePreferenceStore` | 写入偏好持久化：auto / formula / value |

配置与数据目录：

```text
%APPDATA%\ExcelAiAssistant\      # 配置、记忆、会话、自定义工具 (custom-tools)
%APPDATA%\ExcelAiAssistant\skills\<技能名>\SKILL.md   # AI 技能包
%LocalAppData%\XaiAssistant\      # 安装的应用文件 app\x64 / app\x86 + 注册脚本
```

### 写入偏好语义

影响助手写入计算结果的默认方式：`公式优先`（写入 `=item` 引用，可联动追溯）、`直接写值`（数值/文本）、`自动`（跟随模型判断）。

### 智能填充参考上下文

采用“整行邻近上下文（同行各列值）+ 同列分布”双料；数值空格无语义依据时填 `待补`；默认参考范围 `same_sheet`，跨表需显式请求并提供来源。

### 写入文本保护与重解析上报（v3.8.1.76）

`write_excel_range`/`append_excel_rows` 等 Value2 写入通道的文本保护层（Core 纯逻辑 `TextPreserveAdvisor`，真机自检实证驱动：18 位身份证曾被 COM 写入重解析为 5.787E+17 精度丢失、"007" 前导零丢失）：

- **锁文本形态（宁窄勿宽）**：纯数字字符串长度 ≥ 11（手机号/身份证/长编码，≥16 位 double 必丢精度）与前导 0 的纯数字字符串（编码语义），写入前逐格预设 `NumberFormat="@"` 再整块写值；返回值 `text_locked_cells` 回显锁定格数。日期串不拦（Excel 原生日期化与手工输入一致）。
- **重解析可见性**：写入后回读比对，输入为长字符串（≥8 字符）而实际被存为数值/日期的格子经 `coerced_cells`（行/列/输入值/存储形态，上限 20 格）+ `coerced_count`/`coerced_note` 上报，把「写入/回读不符」从静默漂移变成 AI 可复核的显式信息；短数字串数值化属无歧义等价不上报防噪声。
- 真机闭环：MacroChainProbe M14（身份证 18 位锁文本原样回读 / "007" 保留 / "1987-07-04"→31962 上报 coerced），15/15 全过；Core 单测 16 例锁口径。

### 条件格式自检降级（v3.8.1.76）

`verify_conditional_format` 在 WPS 宿主下 FormatCondition 的 COM dynamic 分派可能不暴露 `Interior/Font`（RuntimeBinderException「未包含 Interior 的定义」，dynamic 丢运行时成员同族），原实现一票否决整个自检。现改为读取降级：规则类型/公式/运算符校验照常执行，颜色读不到时同步跳过颜色断言（expected 置空防误判"颜色不符"），并在 `issues` 明示"请人工目视核对颜色"、`color_check_skipped=true` 回显、AppLog.Warn 留痕。

### 快捷宏与数据分析报告链路

快捷宏由前端 `MACRO_CATALOG`（32 条，group 分组）驱动，精选行上限由前端 `MACRO_FEATURED_MAX` 单点把关（C# 侧 Get/SaveCustomMacros 仅透传）；老用户升级迁移语义：持久化精选与旧默认完全一致视为从未自定义，自动升级为新默认，自定义过则保留（`reconcileFeaturedNames` 做"精确→去 emoji 宽松→丢弃告警"三段对账）。

「👥 数据专家团」宏（v3.8.1.83 新增，catalog 不进精选）的编排契约：多角色协同叙事——🧭组长（口径/计划/轻确认直行）→📊数据工程师（体检+清洗建议清单不擅改）→📈分析师（基线对比先行+标准工具组合，需要外部基准时调用 web_search：只传关键词、结果标注网页抓取来源并与本地数据区分、WEB_SEARCH_DISABLED 时明示）→🎨可视化设计师（suggest 映射确认后建图+逐图自审）→🧭组长合并（generate_analysis_report 落盘+数字对账自检回显+六段式按贡献角色署名）。与洞察分析宏的分工：洞察=单角色全景直出；专家团=多角色分工叙事（工具与纪律完全同源，差异在过程呈现与角色署名）。

**联网搜索链路**（v3.8.1.83，主 A 兜 B，无独立搜索 API）：①主路线 A——`SearchProviderDetector` 按模型方案 baseUrl 前缀自动探测提供方（zhipu/moonshot/qwen/openrouter/qianfan 命中服务端搜索；deepseek 等归入兜底），`SearchRequestInjector` 在 `OpenAiCompatibleClient.CompleteAttemptAsync` 组装请求时按提供方注入（智谱 tools 追加 web_search 声明 / Kimi tools 追加 builtin_function $web_search / 通义顶层 enable_search / OpenRouter 模型名 :online / 千帆顶层 web_search.enable），注入仅 native 协议、Fallback 降级重试自动跳过；②兜底 B——`WebSearchService`（Core.Search）抓取 cn.bing.com（浏览器 UA/Accept-Language/TLS12，b_algo 结构解析经真实网络实测），结果标注 source=bing_scrape「网页抓取」来源，网络失败/零结果 fail-closed（SEARCH_NETWORK_FAILED/SEARCH_EMPTY）；③产品 `web_search` 只读工具注册于 `ExcelToolHost.CreateRegistry`（工具数 151→152），执行时读 AgentSettings.WebSearchEnabled 总开关（关闭返回 WEB_SEARCH_DISABLED 并指路设置页）；④AgentRunner 对 `$` 前缀内置函数调用（Kimi $web_search 回传形态）拦截走 bing 抓取兜底，避免 UNKNOWN_TOOL 打断回合；⑤隐私红线：query 仅关键词出网，工作簿原始数据不外发。

「💡 洞察分析」宏（v3.8.1.82 新增，catalog 不进精选）的编排契约：一键全景洞察报告，与数据分析报告（快速摸底）、深度分析（单方法深挖）三宏分工。关键编排要素：①轻确认直行——开场 2~3 行复述体检结论+分析计划+基线判断后直接继续，仅写数据操作与方法分歧停等；②基线探测对比先行——prompt 层扫描同期/目标/预算列对，有基线全部指标先对比后结论，无基线在数据概览段显式声明（对应需求稿 docs/需求-报告洞察增强.md Q4，工具层 detect 基线探测按 P0-1 独立批次）；③无分析目的不追问、报告头部明示（Q3 拍板）；④洞察引擎按数据形态组合 analyze_* 族与检验/聚类/回归标准工具，分工与抽样纪律同深度分析宏；⑤报告自检回显——generate_analysis_report 落盘后抽查关键数字对账/空段/基线声明一致性（Q6 的 prompt 层先行版，工具化自检按 P0-2 排期）；⑥六段式输出中"核心洞察"段强制对比先行四要素：对比数字→业务解读→来源依据→置信度。

「📊 数据分析报告」宏（精选第一位）的编排契约：

1. **视角识别（第 0 步）**：显式指定 > 列名推断 > 通用资深分析师兜底；14 个职能视角包（生产/PMC物控/质量/销售/电商运营/财务/采购供应链/市场营销/研发项目/物流仓储/客户服务/风控合规/人效HR/CEO）内联于宏 prompt，六段式骨架不变，仅切换第②段指标框架、第③④段典型分析与第⑤段建议方向；CEO 包为唯一特殊输出格式（决策备忘录式）；**自定义视角**：用户指定内置清单之外的视角时按通用框架现场生成指标框架并声明（v1 不持久化）。
2. **工具链**：`read_business_glossary`（口径命中逐条引用原文，口径条目优先于包内默认定义）→ `analyze_column_profile` / `analyze_range_statistics`（含 IQR 离群点明细，报告第④段逐条引用）/ `analyze_column_correlation`（threshold 参数化，strong 标注）/ `analyze_value_distribution` / `analyze_trend_forecast`（序列形态时轻量预测）→ `generate_analysis_report（output=html）` → `recommend_excel_chart` 只读推荐。
3. **硬性要求**：全部数字基于工具真实返回、禁止编造；不自动建图；不默认写工作表网格。

「🔬 深度分析」宏（高阶路由，📋 一键报表组）：视角识别同上 → 数据摸底（画像/统计/轻量预测基线）→ 按数据形态+视角优先范式出 **Top-3 方法候选**并复述"方法/检验什么/结论写到哪"复述确认 → 执行分工（描述统计、**常用四类假设检验（analyze_hypothesis_test）、KMeans 聚类分群（analyze_cluster）、多元线性回归（analyze_regression）**走内置 analyze_*，其余高阶（非参检验/逻辑与非线性回归/正则化/时序/文本）走 run_python_excel 按 python-advanced-analysis 契约）→ 结果落位（轻量结论写右侧空白列，成体系产物落「深度分析」工作表，全部经确认门可回滚）→ 五段式汇报（结论/方法依据/人话/局限风险含相关≠因果/后续建议）。

`analyze_hypothesis_test`（v3.8.1.70 新增只读工具，工具数 146→147）：假设检验标准工具——四类：two_sample_t（Welch 双样本 t）/paired_t（配对 t）/anova（单因素 F）/chi_square（独立性列联表）。返回统计量/自由度/精确 p 值/显著性判断（可调 α）/效应量（Cohen's d、η²、Cramér's V）/分组摘要/前提提醒（正态性、方差齐性、期望频数<5 提示改 Python 精确检验）。p 值由 Core 纯逻辑 `HypothesisTester` 计算（lgamma Lanczos + 不完全 gamma/beta 特殊函数，t/F/卡方分布），10 例单测对统计表值断言（t=2,df=10→p≈0.0734；chi2=3.84,df=1→p≈0.050；F=1→p=0.5）并经 scipy 交叉验证。basis 附 T.TEST/F.DIST.RT/CHISQ.TEST 公式示例。

`analyze_cluster`（v3.8.1.71 新增只读工具，工具数 147→148）：聚类分群标准工具——Z-score 标准化（可关）+ K-Means++ 初始化 + Lloyd 迭代（5 次重启取最优惯性），k 在 2~6 内按**平均轮廓系数最大**自动选定（可显式指定 k，k_trace 回显各 k 得分）。返回各簇画像（样本数/特征原始量纲均值）、逐行簇分派（≤1000 行全量，超出省略并提示）、完整观测行数与剔除计数，basis 说明标准化口径/定 k 依据/球形簇边界。实现为 Core 纯逻辑 `ClusterAnalyzer`（KMeans++ D² 抽样、确定性种子），6 例单测覆盖三分离簇自动定 k、显式 k、标准化恢复小量纲维度簇、低样本 fail-closed、轮廓范围与确定性、1 维数据。

`analyze_regression`（v3.8.1.72 新增只读工具，工具数 148→149）：多元线性回归标准工具——普通最小二乘（OLS）正规方程求解（部分主元高斯消元求逆），输出系数表（估计/标准误/双侧 t/p）、R²/调整 R²/F 整体检验、残差标准误/Durbin-Watson/VIF/Cook's D 诊断与逐行拟合值/残差明细（≤1000 行，超出省略）。因变量列显式 `target_column` 优先（缺省取区域最后一列），特征列显式 `columns` 优先（缺省自动选除因变量外数值占比≥60% 的列）；完整观测行参与回归（缺失行剔除并计数）；完全共线/常数列 fail-closed（`REGRESSION_FAILED` 并提示剔除冗余特征），残差自由度 <2 拒算。全部公式（含 VIF 闭式 `((X'X)^-1)_jj·S_j` 与 Cook's D `h_ii/(1-h_ii)²`）先在 Python 沙箱与 statsmodels/scipy 全参数对齐验证（误差 ≤1e-12）再移植 C#，实现为 Core 纯逻辑 `RegressionAnalyzer`，5 例单测（statsmodels 锁定参考值全指标对照/精确数据系数还原/完全共线抛错/样本不足抛错/诊断明细行对齐）+ 探针 M10 用例（已知数据断言系数 3.0825/1.6925/1.7925、R²=0.9998）。basis 附 LINEST/TREND/RSQ 公式示例与「R² 高≠因果、VIF>10 共线、Cook's D 大影响大」边界提醒。

`analyze_text_variants`（v3.8.1.73 新增只读工具，工具数 149→150，FR-K2 探索式清洗落地）：文本列变体聚类建议——对名称类文本列（客户/供应商/品类）输出疑似变体组：入口先 Trim 折叠首尾空白变体，再按「归一化（去空白/全角转半角/大小写折叠）相等=exact、编辑距离≤阈值（默认 2，可调 1~4）且共享首字符=fuzzy」贪心聚类（按频次降序，代表值=簇内出现最多原值）。每组返回代表值、成员（原值+出现次数+匹配类型）与合并方向选项（取最长/取最短/取出现最多）。**只建议不合并**：用户确认方向后，合并动作由 AI 用 `write_excel_cells_batch` 将变体单元格改写为代表值（改值语义，走确认门可回滚；不得用 dedup_range——其语义为删除重复行）。实现为 Core 纯逻辑 `TextVariantClusterer`（编辑距离 DP 长度差>4 剪枝、串长 >64 只做 exact、唯一值 >1500 报 `TOO_MANY_UNIQUE_VALUES` 引导 python-advanced-analysis 走 rapidfuzz），算法先经 Python 沙箱预验证（3 组真值召回 100%/零跨组混入/零噪声误合；LCP≥1 首字符约束专为挡「腾讯科技/小米科技」类短中文串距离误合，沙箱实证），7 例单测（归一化 exact/中文错字 fuzzy/短串误合防护/合并方向选项/AC-K2 缩比召回/唯一值上限/空值计数）+ 探针 M11 用例（2 组变体真机建议断言）。指纹算法仅对拉丁字符有效，中文不做承诺（需求 v0.2 审查口径）。

`add_insight_card`（v3.8.1.74 新增写工具，工具数 150→151，高阶蓝图批7 洞察卡进看板）：把 AI 分析结论生成为圆角矩形洞察卡（Shapes.AddShape msoShapeRoundedRectangle），落指定工作表锚点单元格处。参数 sheet/cell/title/body/color（blue/green/orange/red/gray）/width_pt/replace_existing；卡高按文字量自动估算（Core 纯逻辑 `InsightCardLayout`：CJK 全角/ASCII 半角字宽换行、正文超 8 行截断省略），标题 11pt 加粗+正文 10pt 白字。写入后内嵌 TextFrame2 读回校验（verdict=ok 才交付，否则删半成品报 `INSIGHT_CARD_FAILED`）；`replace_existing=true` 幂等清空该表 `xai_insight_*` 旧卡重写。AutoVerifyPolicy 豁免登记（形状不在单元格快照回滚体系，读回自检内嵌返回值）。AutoVerifyPolicy 覆盖审计、6 例布局单测（字宽/卡高增长/截断封顶/手动换行/主题色映射/窄卡兜底）+ 探针 M12 用例（读回 ok+形状真机存在性复核）。配套 Python 侧：snownlp 中文情感分析库入 PythonLibCatalog（24→25 库）与缺模块自愈白名单，python-advanced-analysis v2.3.0 固化文本情感契约（sentiments 0~1 打分、≥0.6 积极/≤0.4 消极业务阈值约定、批量结果写右侧空白列、结论配占比+典型样本、可配 add_insight_card 进看板）；深度分析宏 prompt 同步情感路由与洞察卡落位。

**大表抽样与超时分级契约（v3.8.1.75，python-advanced-analysis v2.4.0，蓝图 P1-2）**：数据超 5 万行时，建模类分析（聚类/回归/网格搜索/蒙特卡洛）不得直接全量拟合——先 `df.sample(n=50000, random_state=42)` 随机抽样或按关键维度 `groupby(...).apply(分层 sample)`，**抽样参数（n/seed/分层列）必须写进结论三件套的「计算依据」**并告知用户"结论基于 X 万行中的 n 行样本"；全量可完成的统计（均值/占比/分布/情感打分）不抽样。超时分级：按场景显式传 `timeout_ms`（上限 300 秒，Clamp 已固化）——统计检验/单次拟合 60s、聚类/降维/时序建模 120s、网格搜索/蒙特卡洛/大批量文本 300s 顶格；默认 30 秒只够轻量计算，大表建模不传必超时并重置 worker（旧状态丢失）。深度分析宏 prompt 同步抽样与分级提醒。真机用例：探针 M13（6 万行块写 BIG_TMP → run_python_excel 抽样 n=2000/seed=42/timeout 120s → xai_result 断言 total_rows=60000/sample_n=2000，链路与分级超时全链验证）。

**数据分析能力测试卷探针（v3.8.1.75 后验收批次，XaiProbe `exam` 模式）**：自拟测试题真机作答覆盖 18 项能力全矩阵——样例卷 `tools/fixtures/analysis-exam.xlsx`（14 sheet，每题确定性数据+scipy/statsmodels 预计算锁定真值）。21 题：基础 8（E1 描述统计锁定均值/中位数、E2 分布频次、E3 相关系数 0.9985/0.0015 与 strong 标记、E4 画像 number=9/text=4/blank=2、E5 口径文件条目、E6 趋势第 13 期=220、E7 IQR 离群 count=3、E8 变体 3 组建议）+ 高阶 10（E9a-d 四类检验 t=-9.3531/-5.1657、F=75、chi2=16.6667 全显著；E10 完美回归系数 2/3 与 R²=1；E11 四簇自动 k=4 高轮廓；E12 statsmodels Holt 预测区间断言；E13 sklearn PCA PC1>0.95；E14 IsolationForest 5 极端点全检出；E15 snownlp 正负分类；E16 statsmodels 比例 z 检验 z=-1.963 与功效不足提示；E17 scipy linprog 最优 36+蒙特卡洛 π；E18 TOPSIS 读簿排名 甲>丙>乙）。收敛史：轮 1=16/21 → 轮 3=21/21，5 处 FAIL 全部定性为测试卷侧错误（画像 Schema 断言旧字段/口径 Data 匿名集合未 JSON 序列化/paired 2 列表缺显式列位/statsmodels 0.15 Holt 须显式 initialization_method='estimated'/effectsize API 名为 proportion_effectsize 单数），**产品工具零缺陷**。真机守护补充：`XaiProbe exam <exam.xlsx> [pythonPath]`；口径题夹具由探针运行时写入临时目录 `_口径.md`。

`analyze_trend_forecast`（v3.8.1.69 新增只读工具，工具数 145→146）：轻量趋势预测。输入 address + time_column/value_column + periods(1~12) + method(linear/moving_average) + window；输出未来 N 期预测值、趋势方向、slope/r2 与 basis（SLOPE/INTERCEPT/FORECAST.LINEAR 公式示例）。有效数值点 <6 报 `INSUFFICIENT_POINTS` fail-closed（预测即外推，拒绝误导性输出）；季节性/复杂序列明确引导走 Python 技能，与 statsmodels 分工口径不越界。实现为 Core 纯逻辑 `TrendForecaster` + 服务层薄壳（参数解析/ReadMatrix/basis 组装），线性最小二乘与移动平均基准外推均有纯单测（含截距 90 语义、低样本抛错、标签零填充位宽保留）。

`generate_analysis_report` 参数化（同批）：新增可选 `sections`（章节白名单：概览/趋势/异常/重点列摘要/相关性/建议，缺省全部）与 `correlation_threshold`（|r| 判定阈值，默认 0.7，范围 0.05~0.95）；`analyze_column_correlation` 同步新增 `threshold` 参数并输出 `strong` 标注与 `strong_pair_count`。`ReportTextBuilder.Build` 以可选参数透传，九处章节产出统一走白名单守卫，三种输出形态（worksheet/markdown/html）一致生效。

路由同源：`report-analyst` 技能（v1.2.0，tools 含 analyze_trend_forecast）与宏同口径；`python-advanced-analysis` 技能 v2.0 路由表扩至十范式（检验/回归归因/分群/时序/降维结构/异常进阶/文本情感/实验功效/优化模拟/综合评价）并固化统计口径零容忍纪律（检验前查前提/分类变量编码/聚类前标准化/回归报告 R²与残差/评价先统一指标方向）。

真机守护：MacroChainProbe（XaiProbe `macrochain` 模式）14 用例——口径读取（缺文件返模板）、列画像/统计摘要（含 basis 与离群点）、相关性（threshold/strong）、分布、趋势预测（12 点已知序列断言第 13 期=220、趋势=上升、basis 公式齐）、低样本 fail-closed（INSUFFICIENT_POINTS）、假设检验（M8）、聚类（M9 自动定 k=3/轮廓 0.976）、回归（M10 系数还原/R²/诊断面）、文本变体聚类（M11 变体组建议/改值语义提醒）、洞察卡（M12 读回自检 ok/形状真机存在）、大表抽样（M13 6 万行/sample n=2000/seed=42/timeout 120s 分级）、HTML 报告落盘六章节。探针断言口径：匿名对象数组须 JSON 序列化后判读（直接 ToString 输出类型名造成假 FAIL）；样例源簿 `tools/fixtures/macrochain-source.xlsx`（含「01 比较类」「02 趋势类」两表，勿被其他步骤覆盖）。

---

## 八、VBA 生成、执行与沉淀

| 能力 | 实现 |
|------|------|
| 预置模板库 | `VbaPresetCatalog`：覆盖人事、财务等多领域；模板“输入含表头 + 循环从第 2 行起”避免表头被首个数据覆盖；跨领域结果写新表避免覆盖用户数据；频数表补溢出桶 |
| 自动沉淀 | `VbaReusabilizer + VbaAutoSinkPolicy`：成功执行的宏自动进入代码库；模板命名 ≤20 字、凝练、突出重点 |
| 自定义工具 | `CustomToolStore`：用户保存的自定义工具模板 |
| 执行门控 | `run_vba_snippet` 执行前经 `PreRunVbaGate` 静态护栏 |

「载入VBA面板」用于把本轮成功执行的（但未自动沉淀的）代码复制到 VBA 面板手动处理：可修改、再校验、执行，或另存为模板。校验为普通操作直接返回结果；执行为高危操作需在弹出的确认面板审阅后点「执行代码」。

---

## 九、技能包与流程编排

### AI 技能包

`SkillCatalog` 扫描 `%APPDATA%\ExcelAiAssistant\skills\` 下的 `SKILL.md`，聊天中以 `/技能名` 强制加载调用。技能包导入需作 zip-slip 防护：手动逐条解压校验，条目路径不得含 `../` 或绝对路径，目标须在安全范围内，越界抛异常。

### 技能市场与 AI Skill（v3.8.1.77）

外部技能安装/管理/调用的完整闭环（需求稿 `docs/需求-AISkill技能市场.md`，grill-me 八问拍板：安装与执行分离、兼容性诚实降级、市场来源强制高风险策略、产品契约优先）：

- **安装器四入口**：本地 zip（`ImportSkillPack`）、URL 直链 zip/SKILL.md（`InstallSkillFromUrl`）、粘贴 SKILL.md（`InstallSkillFromText`）、**仓库坐标**（`InstallSkillFromRepo`，新增）——`SkillMarketResolver` 统一解析 `owner/repo`、`owner/repo/技能名`（单仓多技能仓库如 K-Dense-AI/scientific-agent-skills 指定安装，避免 143 技能全量落盘）、GitHub 链接与 skills.sh 技能页链接为 GitHub 仓库坐标，codeload zip 分支先 main 后 master 回退，直连失败自动切 jsDelivr 镜像（flat 清单+按文件 CDN 重建目录，`SkillMirrorPlan`），且宿主内下载显式启用 TLS1.2/1.3（真机实证国内 codeload TLS 握手被掐的场景）；
- **市场来源标记**：市场安装统一经 `SkillMarketText.EnsureMarketSource` 注入 `source: market` 并**强制覆盖 `risk_policy: high`**（第三方自声明低风险不可信），正文尾部追加产品契约优先声明（中文字体探测链/确认门回滚/统计口径/禁编造）；
- **兼容性三档**（`SkillCompatJudge`）：技能 front matter 声明的 tools 与本机注册表 diff——全命中=可用、部分命中=部分可用（列缺失工具）、全未命中=知识参考；纯知识技能（未声明 tools）恒为可用。设置页技能列表逐技能显示档位徽章；
- **技能清单注入上下文**（`SkillListInjector`）：系统提示注入已装技能清单（`- /名称（来源）：描述`，上限 30 行，描述截 60 字符），仅在未显式加载技能时注入——模型可主动建议用户 `/技能名` 加载，技能全文经渐进披露只在确认后进入上下文；
- **来源标注**：`/技能名` 加载后返回值带 `used_skill_source`（builtin/local/market），聊天面板"技能已加载"徽章显示来源；
- **预置推荐**：设置页「AI 技能包」卡内置 10 个推荐技能（excel-analysis / data-visualization / seaborn / statsmodels / matplotlib-viz / grilling + v3.8.1.80 扩充的 statistical-analysis（Anthropic 官方）/ pandas-pro / analytics-data-analysis / data-analysis-jupyter，两轮调研实证，见 `调研-数据分析与可视化skill调研.md`），一键按仓库坐标安装；推荐条目支持**点击展开能力说明与使用说明**（`SKILL_MARKET_RECOMMENDED` 每项带 `detail`/`usage` 字段，前端渐进披露渲染）；v3.8.1.81 起 5 个第一梯队技能升级为**产品内置技能**（`src/ExcelAiAssistant.Addin/skills/{statistical-analysis,pandas-pro,analytics-data-analysis,data-analysis-jupyter,data-visualization}`，front matter 含 version/type/category/tools/triggers/examples 四件套登记，tools 全部为实测注册工具名，正文为中文方法论文档），用户技能目录同名市场副本移除、推荐区对应条目渲染「已内置」禁用态（前端按 `allSkillsCache` 内置名集合判定）；
- **验证**：Core 纯逻辑单测 29 例（解析/三档/注入/标记），前端标签配平 35/35·div 372/372·script node --check，gate 全绿（单测 1171/Excel 真机 119/WPS full 131 组/双泄漏）；v3.8.1.80 批次 5 个第一梯队技能经镜像通道真实下载落盘核验（name 字段与目录一致、市场标记+契约声明注入同口径）。

### 流程编排

`WorkflowExecutor / WorkflowStore / WorkflowTemplate / FlowScopeNormalizer` 实现可复用多步骤模板：`@workflow:名称` 触发、把本次成功工具步骤另存为流程、逐条调用并支持危险确认与单步失败重试。流程范围经 `FlowScopeNormalizer` 归一化，骨架跨工作簿/跨工作表保持一致。

---

## 十、构建与打包

```text
Directory.Build.props         → 版本号单一来源（<Version> / <InformationalVersion>，csproj 继承）
tools/release/release.py      → bump --ver X.Y.Z.W：props + README + 四份手册（文件名/版本戳/引用）一键同步
                                pack：双架构全量重建 → 组装 dist → zip → 版本戳断言 → ISCC
tools/release/gate.py --pack  → 一键门禁直通发版（见 §十二 发布门禁）
tools/installer/*.iss         → Inno Setup 6（/DAppVersion 注入版本，版本号不落 iss）
```

产物命名（WPS 版一律带 `wps` 后缀）：

```text
dist\JYYJ助手-<ver>-wps-x64\ / -wps-x86\   便携解包目录
dist\JYYJ助手-<ver>-wps-x64.zip / -x86.zip 便携包
dist\installer\JYYJ助手-<ver>-Setup-wps.exe 安装器
dist\JYYJ助手-<ver>-wps-安装说明.txt        安装说明
```

版本维护要点：当前版本统一由 `tools/release/release.py bump --ver` 自动同步——`Directory.Build.props`（Version/InformationalVersion）、`README.md`、四份手册（JYYJ助手技术手册/用户手册-<版本>.md/.html，**文件名、版本戳、以及 CLAUDE.md / AI交接文档 / README 中的手册引用**一并改名改写）；改动后跑门禁回归。禁止手工 sed 三处版本号（历史红线：M15 增量编译旧戳事故）。

### 双架构依赖布局

```text
app\x64\   → ExcelAiAssistant.*.dll + x64 的 SQLite.Interop.dll + WebView2Loader.dll
app\x86\   → ExcelAiAssistant.*.dll + x86 的 SQLite.Interop.dll + WebView2Loader.dll
```

---

## 十一、安装、COM 注册与卸载

安装器 `JYYJ助手-3.8.1.83-Setup-wps.exe`（Inno Setup 6，约 4.7MB）与便携包 `-wps-x64.zip`/`-wps-x86.zip`（约 3.4MB），自动探测 Excel/WPS 位数并写对应注册表视图：

**安装器形态（v3.7.1.5，WPS 版）**：

- **轻量版** `-Setup-wps.exe`（约 4.7MB）：缺 .NET 4.8 / WebView2 时提示并指引下载；离线完整版形态仅 Excel 主版本提供（`-Setup-离线完整版.exe`，约 316MB），WPS 版暂无；WebView2 缺失检测逻辑同源（per-machine 键在 32 位注册表视图，须 HKLM/HKCU 双视图探测，见踩坑手册 §6.7）；
- **宿主检测与强杀（v3.7.1.5 重构，M17.2/M17.5）**：检测 Excel / WPS 宿主进程（`tasklist` 全量输出按名搜索——`/FI` 过滤器实测会静默失灵）→ 普通权限 `taskkill /F /T` 强杀全家族进程名（EXCEL.EXE / et.exe / wps.exe / wpp.exe / wpsoffice.exe）≤3 轮 → 仍存活则 `ShellExec('runas')` 弹 UAC 以管理员再强杀。**注意 `PrivilegesRequired=lowest` 下 Inno 永远以普通权限运行，用户右键"以管理员身份运行"也会被降权**，`ShellExec('runas')` 是唯一有效提权路径；击杀输出持久化到 `%LocalAppData%\XaiAssistant\logs\hostkill.log`（原 `{tmp}` 安装结束即删、失败无据可查，M17.5 改持久化）；`/SILENT` 静默安装自动强杀宿主（批量部署 + 自动化验证前提）；
- `register-auto.ps1` 按 `CLSID\{GUID}` 带花括号写键，并在注册前/卸载时自愈清理历史无括号死键（干净机器必报「运行错误」的产品级根因，见踩坑手册 §6.8）；
- 随包整合 `安装说明.txt`（小白版）与 `diagnose.ps1` + `一键诊断.cmd`（桌面生成诊断报告），开始菜单提供「一键诊断」入口。

### 注册原理

```text
探测 Excel 位数：读 HKLM App Paths\excel.exe → 取 PE 头 machine 值
  0x8664 → x64  → 写 Registry64 视图
  0x014C → x86  → 写 Registry32 视图
用 .NET Registry API (显式 RegistryView) 写：
  HKCU\Software\Classes\CLSID\{GUID}        (Assembly/Class/ProgID/InprocServer32/CodeBase…)
  HKCU\Software\Classes\{ProgId}
  HKCU\Software\Classes\CLSID\{GUID}\Implemented Categories\{catId}
  HKCU\Software\Microsoft\Office\Excel\Addins\{ProgId}   LoadBehavior = 3
```

注册编排脚本：`register-auto.ps1`（UTF-8 带 BOM），由 64 位 PowerShell（`{sysnative}`）静默执行，显式 `RegistryView` 对进程位数免疫，保证 32/64 位 Excel 各从对应视图加载。

### 前置检测（ISS [Code]）

- 检测 Excel / WPS 宿主进程并多轮强杀（含 UAC 提权轮，详见上文「宿主检测与强杀」）；交互模式失败时弹归因弹窗（UAC 未确认 / 安全软件拦截 / WPS 后台组件拉起 / 服务或自动化拉起需重启四类指引）。
- 检测 .NET Framework 4.x（`NDP\v4\Full`）与 WebView2 Runtime（`EdgeUpdate\Clients\{GUID}`），缺失仅友好提示、不阻断。

### 卸载

`unregister-all.ps1` 遍历 `Registry64 / Registry32` 两个视图，清理 `CLSID`、`ProgId`、`Excel\Addins` 全部注册项，再删文件。数据目录（custom-tools / agent-settings.json / 记忆 / 技能）不被卸载删除，需手动清理。

> **兼容性注意**：COM 注册须用与 Excel PE 架构匹配的 `RegistryView`；安装必须保留用户数据、仅覆盖应用目录；PowerShell 脚本必须 UTF-8 带 BOM 以防 PowerShell 5.1 中文解析错误。

---

## 十二、可观测性与质量保证

### 运行日志 · 一键自检

`DiagnosticsLog` 实时登记助手本轮的写入数据 / 生成的 VBA 代码 / 创建的图表（登录与否均登记）；侧边栏「运行日志 · 自检」的**一键自检**用 AI 研判最近会话，核验 AI 生成的数据、代码与图表是否正确；**全功能巡检**按钮对全部核心子系统逐项体检（详见下文），结果同样写入运行日志。

**v2.19.4.0 运行日志全功能覆盖**：新增 `Core\Diagnostics\AppLog.cs` 跨程序集日志门面（`Action<LogEntry> Sink` + Info/Warn/Error 三级 + 可选 category），`ExcelAiAddIn` 启动时 `AppLog.SetSink` 注入 `DiagnosticsLog`，打通 Core/Excel 层无法直取 Addin 层日志的架构隔离——此后全部写操作/工具执行统一汇入运行日志，按级别过滤即可排查任意功能错误。链路点：
- `AgentRunner.ExecuteToolAsync` 逐工具记耗时与**三路径**日志（成功→`Info("tool")`、失败→`Info` 带错误码、异常→`Error` 带 SafeDetail），成功摘要截 200 字符；子项自动自检同样记 `"verify"` 耗时。
- `PythonWorkerHost.DrainStderr` 用 `BeginErrorReadLine` 捕获原被静默丢弃的 Python stderr 并记 `Warn("python")`；新增 `NoCr()` 折行（换行折叠 + Trim，>300 截断加省略号）防多行 traceback 撑爆单行日志。

### 自动自检体系（执行层硬保证）

写入类工具成功后，服务端在 `AgentRunner` 执行层**自动**调用配套 `verify_*` 工具回读核对，不依赖模型自觉：

- `AutoVerifyPolicy.TryPlanAll`：按主工具映射出全部自检计划（批量回写逐格出计划，≤5 格封顶），并从主工具实参自动推导预期参数（写入值即期望值）。
- `AutoVerifyPolicy.EnrichFromResult`：用主工具执行结果富化计划（追加行的目标地址取自 `AffectedRanges`、透视的输出地址/行数/守恒对账参数取自结果）。
- `AutoVerifyMonitor`：元自检——每回合统计「可自检写操作数 vs 实际执行的自检计划数」，出现"有写操作但 0 条自检执行"即记 `SELF_CHECK_LINK_SUSPECT` 告警进运行日志（自检链路失效不再静默），摘要同时随 `AgentRunResult.AutoVerifySummary` 返回会话层。
- 映射覆盖 **26 类主操作**（写入/公式/排序/清洗/去重/追加/批量/格式/条件格式/透视/标记/图表/工作表增删改/图片/冻结窗格等），另有 77 个写能力工具在 `AutoVerifyPolicy.DocumentedWriteToolExemptions` 中**带理由豁免**；`MappedPrimaryTools` 与映射行为由一致性单测锁死，覆盖审计在 Excel 真机 harness（用例组 17）与全功能巡检两处执行。
- 自动自检在 `verify_write_result` 上默认开启 `type_aware`（文本型数字判不通过）。

### 自检工具族（verify_*）

| 自检方向 | 验证器 |
|----------|--------|
| 图表 | `verify_chart`：类型 / 系列 / 点数 / data_ranges / 渲染像素核验 / selfcheck |
| 单元格格式 | `verify_cell_format` |
| 条件格式 | `verify_conditional_format` |
| 排序 / 筛选 | `verify_sort_filter` |
| 公式 | `verify_formula`（错误值 #DIV/0! 等判 error） |
| 写入结果 | `verify_write_result`（逐格比对 + `type_aware` 文本型数字检测） |
| 数据清洗 | `verify_data_clean`（no_blank_rows/no_blanks/no_duplicates/dates_valid/no_text_numbers/no_text_matches/all_blank） |
| 透视 | `verify_pivot`（行值抽查 + 明细/汇总**数量守恒对账**） |
| 标记 | `verify_marker` |
| 结构 | `verify_structure`（工作表存在/缺席断言 + 区域行列数） |
| 迷你图 | `verify_sparkline` |
| 数据验证 | `verify_data_validation` |
| 合并单元格 | `verify_merge` |
| 超链接 | `verify_hyperlink` |
| 表格 | `verify_table` |
| 图片 | `verify_image`（Shapes 锚点/数量/尺寸，容差 2 磅；尺寸断言不配 anchor 时 fail-closed） |
| 冻结窗格 | `verify_freeze_panes`（SplitRow/SplitColumn + 锚点行列 -1 判定） |
| 保护态 | `verify_protection`（只读探测，配合 SheetProtectionGuard 负向用例） |

共 **18 个** verify_* 验证器，统一产出 `verdict + issues + selfcheck` 三要素；verdict≠ok 时模型必须先修正再向用户报告（系统提示词强制）。

### 语义守恒护栏

- **排序行守恒**：`sort_range` 写入前比对排序前后身体行多重集，不一致返回 `CONSERVATION_VIOLATED` 并拒绝写入。
- **透视数量守恒**：明细列合计 == 汇总列合计，口径与 `PivotTableService.SumByRow` 严格对齐。
- **类型感知回读**：期望数值 vs 实际文本型数字判不通过。

### 踩坑失败模式库

`PitfallLibrary`（Core.Diagnostics）把《踩坑记录与规避手册》提炼为「日志签名 → 症状 → 手册级建议」规则表（15+ 条，含 HRESULT 0x800A9C68、RCW、WebView2 跨线程、CONFIRM_TIMEOUT 等）。一键自检 AI 研判对运行日志逐条签名比对，命中即并入问题清单与研判提示词。新坑在手册补充章节后同步追加规则即可被自动识别。

### 全功能健康巡检

`FeatureHealthSweep`（Addin）：在**新建临时工作簿**（不触碰用户数据）中对全部核心子系统逐项体检——写入/公式/排序/清洗/追加/批量/透视守恒/格式/条件格式/标记/图表/结构操作各自跑「主工具+自动自检闭环」，另含回滚链路端到端（写入→覆写→undo→断言恢复）、自检覆盖审计、踩坑库、自检历史库、配置与技能加载。逐项结果写入运行日志并回传前端。VBA/Python 数据通道由一键自检覆盖，两者互补。

### 自检历史与趋势

`SelfCheckHistoryStore`（SQLite，`%APPDATA%\ExcelAiAssistant\selfcheck-history.db`）：回合元监控摘要与每次研判结论落库（上限 500 条自动裁剪）；`Stats(20)` 计算最近通过率与高频问题 Top3，随一键自检结果返回前端展示；侧边栏在每轮对话结束弹「本轮自检」摘要。

### 发布门禁（gate.py）

`python tools\release\gate.py --pack` 一条命令直通发版，fail-fast 任一环失败即中止：

1. Release 构建 + 全量 xUnit 单测（979 个，含 Sidebar 内联 JS 静态审计、AutoVerify 一致性锁、缺陷原型审计 M16DefectPatternAudit）；
2. Excel 真机回归（`tools\RealMachineVerify`：独立 Excel 实例 + 临时工作簿，23 用例组 119 断言，含负路径）；
3. **宿主进程泄漏断言（M17.5 新增）**：真机步前后对 Excel/WPS 宿主进程快照对比，harness 中断泄漏当场亮红（堵死"遗留僵尸进程毒害下一次安装"通路）；
4. WPS full 真机 harness（`tools\WpsToolVerify`）+ `check_wps_full.py` 基线校验（OK/预期 FAIL/工具数 139（剔除 ct_ 动态注册）逐组锁定，`EXPECTED_TOOL_COUNT` 常量为有意设置的变更卡点）；
5. `release.py pack`：双架构全量重建（`--no-incremental`，M15 红线）→ 组装 dist → zip → 版本戳断言（新戳在位 ≥3 / 旧戳零残留 / 守卫符号在位）→ ISCC 打安装器（Error 32 暂态锁定自动退避重试）。

### 测试与验证

`ExcelAiAssistant.Tests` 含 **979 个 xUnit 单测**（安全管线、撤销、协议解析、技能目录、公式分析、异常检测、写偏好、VBA 门控、自动自检策略/元监控/粘合层集成、JS 静态审计、M16 缺陷原型审计等）；`tools\RealMachineVerify`（Excel）与 `tools\WpsToolVerify`（WPS）是真机 harness，遵循**无状态**原则：每次操作前重新解析工作表、跨表断言显式按 `Workbooks[...].Worksheets["name"]` 定位、不依赖 `ActiveSheet`、用 `range.Cells.Item[1,1]` 作写回锚点；out-of-proc 收尾不调 Close/Quit（动态 COM 退出路径会硬故障，见踩坑 §5.12），按 PID 强杀兜底；门禁对真机步执行宿主进程泄漏断言（M17.5）。
