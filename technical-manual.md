
[Uploading technical-manual.md…]()
# JYYJ助手 技术手册

> Excel COM Add-in · .NET Framework 4.8 · WebView2 · OpenAI 兼容协议 · 版本 2.19.4.0

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
│  │ ExcelAiAssistant.Excel (80+ 工具服务层)              │  │
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
d:\execl1.0
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
│  ├─ ExcelAiAssistant.Excel\    # 全部 Excel 工具服务 + schema + 注册
│  └─ ExcelAiAssistant.Tests\    # NUnit 单元测试
├─ tools\
│  ├─ installer\        # *.iss (Inno Setup) + register-*.ps1 / unregister-*.ps1
│  ├─ package-xai.ps1   # 构建 x64/x86 产物并打包
│  ├─ install-xai.ps1   # 开发者本机安装注册
│  └─ Verify* / Probe*  # 真机/协议验证工具
└─ dist\                # 构建产物与安装器
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
| `ExcelVerifyRegistration` | 自检（格式/条件格式/排序筛选/公式/图表） | `ExcelVerifyService + ExcelVerifyJudgments` |
| `ExcelWebToolsRegistration` | 网页取数 | `WebFetchService` |
| `ExcelExternalBookTools` | 多工作簿只读协同 | `ExternalBookRegistry` |
| 数据分析 / 报告 | 统计建模 / 文本报告 | `ExcelAnalysisService / ExcelReportService + ReportTextBuilder` |
| 脚本 / 高进度 | 脚本、外部进程、数据转换 | `ScriptRunnerService / PythonWorkerHost / DataTransformService / ExcelScratchpadService / ExcelImportService` |

> **关键约束**
>
> - 危险级别工具在 schema 中标注 `risk:high`，执行前必须经确认门。
> - 结果不“就地覆盖”时，必须显式把源区域地址设为写回地址；相对逻辑坐标叠加选区绝对原点（`range.Row/range.Column`）。
> - 公式引用解析需覆盖 `$A$1`、`C:C`、`'表'!B2`、`[Book]Sheet!A1` 等变体，R1C1 样式标记为去参数哨兵。
> - 工具扩展沿用“分期”演进（旧产物标记为 三期 = 导入/分析/可视化基础能力，四期 = 数据分析与报告生成）。

---

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

### 流程编排

`WorkflowExecutor / WorkflowStore / WorkflowTemplate / FlowScopeNormalizer` 实现可复用多步骤模板：`@workflow:名称` 触发、把本次成功工具步骤另存为流程、逐条调用并支持危险确认与单步失败重试。流程范围经 `FlowScopeNormalizer` 归一化，骨架跨工作簿/跨工作表保持一致。

---

## 十、构建与打包

```text
src/ExcelAiAssistant.Addin.csproj → 版本 (Version / InformationalVersion) 2.19.4.0
tools/package-xai.ps1  →  分别构建 x64 与 x86 的完整 app + 各架构原生依赖
                          →  dist\JYYJ助手-<ver>-x64\app 与 -x86\app
tools/installer/*.iss  →  Inno Setup 6 打包合并安装器
                          →  dist\installer\JYYJ助手-<ver>-Setup.exe
```

版本维护要点：加载项当前版本需在 `ExcelAiAssistant.Addin.csproj` 的 `Version/InformationalVersion`、ISS 的 `#define AppVersion`、及安装器元数据三处**同步一致**。

### 双架构依赖布局

```text
app\x64\   → ExcelAiAssistant.*.dll + x64 的 SQLite.Interop.dll + WebView2Loader.dll
app\x86\   → ExcelAiAssistant.*.dll + x86 的 SQLite.Interop.dll + WebView2Loader.dll
```

---

## 十一、安装、COM 注册与卸载

合并安装器 `JYYJ助手-2.19.4.0-Setup.exe`（Inno Setup 6）同时装 x64 与 x86 两套产物，自动探测 Excel 位数并写对应注册表视图：

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

- 检测 Excel 是否运行（`FindWindow("XLMAIN")`），占用则提示关闭（重试/取消，阻断）。
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
- 映射覆盖 **16 类主操作**（写入/公式/排序/清洗/去重/追加/批量/格式/条件格式/透视/标记/图表/工作表增删改），另有 55+ 个写能力工具在 `AutoVerifyPolicy.DocumentedWriteToolExemptions` 中**带理由豁免**；`MappedPrimaryTools` 与映射行为由一致性单测锁死。
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
| 数据清洗 | `verify_data_clean`（no_blank_rows/no_blanks/no_duplicates/dates_valid/no_text_numbers） |
| 透视 | `verify_pivot`（行值抽查 + 明细/汇总**数量守恒对账**） |
| 标记 | `verify_marker` |
| 结构 | `verify_structure`（工作表存在/缺席断言 + 区域行列数） |

统一产出 `verdict + issues + selfcheck` 三要素；verdict≠ok 时模型必须先修正再向用户报告（系统提示词强制）。

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

### 发布门禁

`tools\package-xai.ps1` 打包前强制执行双关：①全量单元测试（含 Sidebar 内联 JS 括号/引号静态审计，拦截踩坑 §4.1）；②真机自检回归（`tools\RealMachineVerify`：独立 Excel 实例 + 临时工作簿，正向/负向/覆盖审计 21 用例组 80 断言）。任一失败中止打包；`-SkipReleaseGate` 可显式跳过但会打印警告。`package-xai.ps1 -BumpVersion` 从 csproj 读当前版本按 `-BumpPart`（默认 Patch）递增，并自动同步三处版本号（csproj `Version/InformationalVersion`、`README.md`、复制生成 `<新版本>-Setup.iss`），再跑双关打包。

### 测试与验证

`ExcelAiAssistant.Tests` 含 800+ xUnit 单测（安全管线、撤销、协议解析、技能目录、公式分析、异常检测、写偏好、VBA 门控、自动自检策略/元监控/粘合层集成、JS 静态审计等）；`tools\Verify*` 与 `tools\RealMachineVerify` 是真机 harness，遵循**无状态**原则：每次操作前重新解析工作表、跨表断言显式按 `Workbooks[...].Worksheets["name"]` 定位、不依赖 `ActiveSheet`、用 `range.Cells.Item[1,1]` 作写回锚点；out-of-proc 收尾不调 Close/Quit（动态 COM 退出路径会硬故障，见踩坑 §5.12），按 PID 强杀兜底。
