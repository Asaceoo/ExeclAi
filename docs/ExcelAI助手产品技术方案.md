# Excel AI 助手完整产品与技术方案

## 1. 结论与定位

### 1.1 一句话定位

面向 Windows Excel 的本地智能办公 Agent。用户在 Excel 侧边栏用自然语言下达指令，系统感知当前工作簿、工作表和选区，调用约 25 个受控工具完成读取、分析、写入、格式化、透视表、图表和数据清洗，并通过操作历史支持回滚。

### 1.2 首版边界

- 产品形态：VSTO Excel 加载项，宿主提供 `Custom Task Pane`，侧边栏 UI 使用 WebView2。
- 操作对象：Excel 当前已打开的工作簿；只要 Excel 能打开并编辑该文件，即可处理 `.xlsx`、`.xlsm`、`.xls` 等格式。
- 平台范围：Windows 10/11；Excel 2016、2019、2021 和 Microsoft 365；x64 与 x86 均需测试。
- 不承诺：后台批量处理未打开文件、WPS 表格兼容、Mac/Web Excel、多语言界面、在线授权订阅。
- 后续扩展：增加“导入外部文件到当前工作簿”能力，再评估批量文件处理。

### 1.3 目标用户

| 用户 | 核心诉求 | 首版优先级 |
| --- | --- | --- |
| 财务、审计、运营人员 | 清洗底稿、汇总数据、复核公式、生成分析结果 | P0 |
| 普通办公用户 | 理解表格、生成公式、做简单图表 | P0 |
| 高级办公用户 | 批量整理、目录页、Markdown 转换、复杂自动化 | P1 |
| 开发/数据分析人员 | 运行脚本、扩展 Skill | P2，首版默认关闭 |

## 2. MVP 功能范围

### 2.1 核心体验

1. **常驻侧边栏聊天**：从 Ribbon 打开任务窗格，保留会话状态；Excel 切换工作簿时刷新上下文。
2. **自动选区引用**：每次提问携带工作簿名、工作表名、UsedRange 摘要和当前选区地址；大区域只发送采样和统计摘要。
3. **手动多区域引用**：用户可添加多个引用区域，每个引用显示 `[Workbook]Sheet!Address`、行数、列数、预览和 token 估算。
4. **Agent 循环**：模型可连续请求工具，宿主执行后把结果回传模型；单轮最多 15 次工具调用，总超时默认 180 秒。
5. **工具卡片**：聊天流展示“正在思考”“已使用工具”“参数摘要”“执行结果”“修改 diff”“确认按钮”。
6. **强制指令**：输入 `@tool_name` 强制使用指定工具；输入 `/skill-name` 强制加载指定 Skill。
7. **AI 计算草稿**：复杂计算写入 `_AI_CALC_SHEET`，由 Excel 公式引擎精确求值后回读结果；草稿可作为审计轨迹保留。
8. **历史会话**：SQLite 保存会话、消息、工具事件、模型配置元数据和操作记录。

### 2.2 验收场景

以下 5 条全部通过即认为 MVP 达标：

1. 自动识别当前选区并回答结构、字段类型、空值、数值范围或文本摘要问题。
2. 在用户指定区域生成销售明细和汇总公式，公式必须引用正确且随源数据更新。
3. 清洗数据：删除全空行、修复常见日期格式、拆分姓名/编号等文本列、按关键字批量填充。
4. 创建数据透视表或图表，图表类型和放置位置符合用户要求。
5. 高危操作前确认；普通写入产生新旧值 diff；修改后能从历史记录回滚。

### 2.3 明确不做

首版不提供：

- 云端账号体系和订阅授权；
- Word/PPT Agent；
- 财务账套数据库查询；
- 自动联网搜索；
- 无确认运行 VBA、Python、PowerShell 或外部 EXE；
- 后台扫描并修改未打开的文件夹内文件。

## 3. 架构设计

### 3.1 总体架构

```text
Excel 进程
├─ VSTO Runtime / ThisAddIn
│  ├─ Ribbon 按钮：打开助手 / 配置 / 历史
│  ├─ Custom Task Pane 宿主
│  │  └─ WebView2 控件
│  │     ├─ 聊天界面
│  │     ├─ 引用区域管理
│  │     ├─ 工具确认卡片
│  │     ├─ 设置页面
│  │     └─ 历史回滚视图
│  ├─ ExcelApplicationService（COM 封装）
│  │  ├─ ContextService
│  │  ├─ RangeIOService
│  │  ├─ SheetService
│  │  ├─ FormatService
│  │  ├─ PivotChartService
│  │  ├─ DataCleanService
│  │  └─ ScratchpadService
│  └─ SafetyPipeline
│     ├─ 参数校验
│     ├─ 区域大小限制
│     ├─ 风险分级
│     ├─ 快照采集
│     ├─ 确认门禁
│     └─ UndoRepository
├─ LocalAgentCore
│  ├─ SessionManager
│  ├─ PromptBuilder
│  ├─ ModelClientFactory
│  ├─ FunctionCallingAdapter
│  ├─ JsonProtocolFallbackAdapter
│  ├─ ToolRegistry
│  ├─ SkillLoader
│  └─ EventPublisher
├─ Storage
│  ├─ SQLite：sessions/messages/tool_events/operations
│  ├─ JSON：配置方案与安全策略
│  ├─ DPAPI：API Key 密文
│  └─ Files：undo_backups/
└─ OptionalRunnerHost（默认关闭）
   ├─ PowerShell Runner
   ├─ Python Runner
   ├─ VBA Runner
   └─ Exe Allowlist Runner
```

### 3.2 技术栈

| 层 | 技术 | 说明 |
| --- | --- | --- |
| 宿主 | C#/.NET Framework 4.8 VSTO | 兼容 Excel 2016 及以上；如后续放弃 2016 可迁移 .NET 8 + COM Add-in |
| UI | WPF/VSTO 宿主控件 + WebView2 | WebView2 使用 Evergreen Runtime |
| 前端 | TypeScript + React + Vite | 聊天流、diff、设置和历史复用组件 |
| 状态 | Zustand 或 Redux Toolkit | 会话、待确认操作、上下文引用 |
| 本地存储 | SQLite + Microsoft.Data.Sqlite 或 System.Data.SQLite | 注意 x86/x64 原生依赖 |
| 模型协议 | HttpClient + OpenAI Chat Completions 兼容协议 | 支持原生 tools 与 JSON fallback |
| 序列化 | System.Text.Json | 工具参数统一校验 |
| 日志 | Serilog 文件滚动日志 | 不记录单元格明文，只记录地址、行数和摘要指标 |
| 安装 | Inno Setup 或 WiX | 注册 VSTO，检测 .NET/WebView2/VSTO Runtime |

### 3.3 关键流程

#### 提问流程

1. WebView2 发送用户输入和手动引用列表。
2. 宿主读取当前上下文：
   - active workbook name/path；
   - active sheet name；
   - current selection address/type/count；
   - UsedRange dimensions；
   - header row sample；
   - selection 前 N 行采样；
   - 数值列 min/max/avg/null count。
3. PromptBuilder 组装系统提示词、Skill、上下文摘要、历史压缩消息和用户问题。
4. ModelClient 调用模型。
5. 若返回 tool call，ToolRegistry 校验参数并进入 SafetyPipeline。
6. 低风险立即执行；中风险前端展示 diff 并等待确认；高风险显示影响范围并要求显式确认。
7. 执行前创建快照；执行后写 SQLite 操作记录并推送结果。
8. 结果回传模型继续推理，直至最终回复或达到上限。

#### 回滚流程

1. 用户在历史面板选择一条 operation。
2. 系统校验目标区域当前值是否仍匹配快照中的 after digest。
3. 匹配则恢复 before snapshot；不匹配则提示冲突，允许查看差异后强制覆盖。
4. 回滚本身也生成新的 operation，原记录不物理删除。
5. 完整备份型操作还原临时副本时，提示保存当前文件后再恢复。

## 4. Excel 能力与工具协议

### 4.1 工具清单

第一期实现 25 个进阶工具。

| 分组 | 工具 | 风险 | 说明 |
| --- | --- | --- | --- |
| 上下文 | `get_current_excel_context` | L0 | 当前工作簿、工作表、选区、文件路径、保护状态 |
| 上下文 | `get_excel_sheet_list` | L0 | 工作表名称、可见性、UsedRange 概览 |
| 读取 | `preview_excel_sheet` | L0 | UsedRange 维度、前后若干行、列类型推断 |
| 读取 | `preview_excel_range` | L0 | 指定或选区的前 N 行、行列数、类型摘要 |
| 读取 | `get_excel_range_values_and_formulas` | L0 | 读取值、公式、数字格式；超过阈值只返回摘要 |
| 读取 | `get_current_selection_values_and_formulas` | L0 | 读取当前选区 |
| 导航 | `select_excel_range` | L0 | 选中和激活区域 |
| 名称 | `manage_excel_named_range` | L1/L2 | 查询 L0；创建/更新 L1/L2 |
| 写入 | `write_excel_cell` | L1 | 单格写入值或公式 |
| 写入 | `write_excel_cells_batch` | L1 | 不连续单元格批量写入 |
| 写入 | `write_excel_range` | L1 | 二维数组按区域写入 |
| 计算 | `append_calculation_scratchpad` | L1 | 向 AI 计算草稿追加输入、公式、期望输出 |
| 计算 | `clear_calculation_scratchpad` | L2 | 清空 AI 计算草稿 |
| 数据 | `manage_excel_filter` | L1/L2 | 设置筛选 L1；清除已有筛选 L2 |
| 数据 | `find_excel_cells` | L0 | 查找值、公式或格式匹配 |
| 格式 | `format_excel_range` | L1 | 字体、填充、边框、数字格式、对齐、列宽行高 |
| 结构 | `create_excel_sheet` | L1 | 新建工作表 |
| 结构 | `rename_excel_sheet` | L2 | 重命名工作表 |
| 结构 | `delete_excel_sheet` | L3 | 删除工作表，强制确认 |
| 结构 | `insert_excel_dimension` | L2 | 插入行列 |
| 结构 | `delete_excel_dimension` | L3 | 删除行列，强制确认 |
| 复制 | `copy_paste_excel_range` | L1/L2 | 同簿复制粘贴，目标非空升级风险 |
| 分析 | `create_excel_pivot_table` | L1 | 创建透视表 |
| 图表 | `create_excel_chart` | L1 | 创建常用图表 |
| 目录 | `create_toc_sheet` | L1 | 生成本簿目录页 |

第二期扩展：

| 工具 | 说明 |
| --- | --- |
| `export_selection_to_markdown` / `import_markdown_to_excel` | Markdown 表格双向转换 |
| `merge_or_split_cells` | 相同内容合并、拆分填充 |
| `transform_excel_formulas` | IFERROR、ROUND、错误处理批量转换 |
| `convert_formula_references` | 绝对/相对引用切换 |
| `trace_formula_references` / `drill_excel_range` | 公式引用追踪 |
| `copy_visible_rows_to_sheet` | 筛选可见行复制 |
| `apply_page_setup` | 页面打印布局 |
| `reset_cell_value_type` | 文本/数字/日期类型重置 |
| `export_sheets_to_pdf` | 工作表导出 PDF |
| `merge_workbooks` / `split_workbook` | 合并拆分工作簿 |
| `list_external_links` / `replace_external_links` | 外部链接检查替换 |
| `validate_vba_snippet` / `run_vba_snippet` | VBA 校验与受控执行 |
| `run_python_code` / `run_cli_command` / `run_exe_command` | 本地脚本与进程受控执行 |

### 4.2 JSON Schema 约定

所有工具入参使用统一 envelope：

```json
{
  "tool": "write_excel_range",
  "args": {
    "sheet": "Sheet1",
    "address": "A1:C3",
    "values": [["日期", "金额"], ["2026-01-01", 100]],
    "value_mode": "auto"
  },
  "reason": "写入用户要求的销售明细",
  "call_id": "model-generated-id"
}
```

公共约束：

- 未传 `sheet` 时默认 active worksheet，但响应中必须显式说明实际使用的 sheet。
- 地址仅接受 A1 形式，禁止 R1C1；跨簿地址由宿主解析并限制为当前工作簿。
- 所有二维数组必须矩形，最大 10 万 cell。
- `value_mode` 支持 `auto/raw_text/formula`；`formula` 必须以 `=` 开始。
- 工具结果统一包含 `ok`、`data`、`warning`、`affected_ranges`、`truncated`、`error_code`。

### 4.3 Function Calling 双模式

#### 原生模式

请求 OpenAI Chat Completions `tools` 字段；模型返回 `tool_calls` 后执行。适用于 DeepSeek、Qwen/DashScope、OpenAI、Kimi 等兼容端点。

#### Fallback JSON 协议

当端点不支持 function calling、返回非法 tool call，或用户开启兼容模式时，要求模型输出单个 JSON 对象：

```json
{
  "action": "tool",
  "tool": "preview_excel_range",
  "args": {},
  "reason": ""
}
```

或最终回复：

```json
{
  "action": "final",
  "content": ""
}
```

Fallback 解析规则：

1. 优先整体 JSON parse；
2. 失败后提取第一个平衡的 `{...}`；
3. 再失败则视为 final content；
4. 连续两次非法输出终止本轮并让用户重述。

### 4.4 系统提示词骨架

```text
你是 Windows Excel 侧边栏里的专业办公 Agent。

能力：
1. 只通过提供的工具操作当前 Excel 工作簿。
2. 先理解结构和选区，再决定是否需要读取更多数据。
3. 大区域不得整表读取；先 preview，再按需读取。
4. 写公式时使用 A1 引用，绝对/相对引用要正确。
5. 任何修改必须说明影响范围。
6. 无法确定目标位置时先询问用户。
7. 不建议绕过权限、破解密码或伪造财务数据。
8. 输出中文，保持简洁、可核对。

当前环境：
- Excel 版本：<excel_version>
- 工作簿：<workbook>
- 工作表：<worksheet>
- 选区：<address>
- UsedRange：<rows>x<cols>
- 保护状态：<protection>

可用工具：
<tools_json_schema>

已加载技能：
<skills>

引用区域：
<references>

执行策略：
1. 读取类工具可直接调用。
2. 中风险修改必须先生成 diff，等待用户确认。
3. L3 工具必须强制确认。
4. 连续失败两次应停止并向用户解释原因。
```

## 5. Skill 系统

### 5.1 存放路径

```text
%APPDATA%\ExcelAiAssistant\skills\
  sales-data-analysis\
    SKILL.md
    assets\
      prompts\...
      samples\...
  finance-cleanup\
    SKILL.md
```

安装目录不放用户数据；Skill 支持内置目录和用户目录合并加载，同名用户 Skill 覆盖内置项。

### 5.2 SKILL.md 格式

```markdown
---
name: sales-data-analysis
description: 分析销售明细，识别字段，生成汇总公式和图表。用于销售数据统计、趋势分析和汇总表场景。
version: 1.0.0
tools:
  - get_current_selection_values_and_formulas
  - write_excel_range
  - create_excel_chart
risk_policy: normal
---

# 销售数据分析

## 触发条件

用户要求分析销售明细、计算销售额、按人员/产品/月份汇总。

## 步骤

1. 预览当前选区和表头。
2. 识别日期、客户、产品、数量、单价、金额字段。
3. 缺少金额列时可写入 `数量 * 单价`。
4. 按需求生成分组公式或透视表。
5. 需要可视化时创建柱状图或折线图。
```

### 5.3 加载机制

1. 启动时索引所有 `SKILL.md`。
2. 校验 front matter；缺少 `name/description/tools` 的 Skill 标记为 invalid，不入自动匹配池。
3. PromptBuilder 默认只注入 name + description，不注入全文。
4. 模型可调用 `load_skill(name)` 元工具获取正文；该调用不计入 25 个业务工具。
5. `/skill-name` 强制加载并注入全文。
6. Skill 声明的 `tools` 是白名单；未声明工具的 Skill 不能调用对应工具。
7. 单轮最多加载 3 个 Skill，Prompt 总长度超限时按相关性截断。

## 6. 安全、撤销与数据隐私

### 6.1 风险分级

| 等级 | 行为 | 策略 |
| --- | --- | --- |
| L0 | 读取、预览、查找、选中 | 自动执行，记录审计事件 |
| L1 | 普通写入、新建表、格式化、创建图表 | 显示 diff/影响范围，默认需一次确认；可在设置中对单格写入关闭确认 |
| L2 | 重命名、插行列、清筛选、复制覆盖、清草稿 | 显示详细 diff，强制确认 |
| L3 | 删表、删行列、运行脚本/EXE、替换链接、合并拆分工作簿 | 默认不可用；高危开关显式开启后仍逐次确认 |

### 6.2 写前快照

普通区域操作记录：

```json
{
  "operation_id": "uuid",
  "session_id": "uuid",
  "tool": "write_excel_range",
  "args_hash": "sha256",
  "workbook_path": "...",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "before": [["...", "..."]],
  "after": [["...", "..."]],
  "before_formulas": [],
  "after_formulas": [],
  "number_formats": [],
  "created_at": "2026-08-22T16:00:00+08:00"
}
```

限制：

- 快照最多 500 cells；超过时改为完整文件备份。
- 完整备份保存在 `%APPDATA%\ExcelAiAssistant\undo_backups\{workbook-hash}\`，默认保留 7 天。
- 每个工作簿最多 500 条操作记录，超过后滚动淘汰最旧记录及其孤儿备份。
- 删除工作表前导出该表到隐藏临时文件；恢复时插入到原 index。
- 删除行列前记录整行/整列值、公式、格式、高度宽度、隐藏状态。

### 6.3 API Key 与日志

- API Key 使用 Windows DPAPI 以 CurrentUser 范围加密后落盘。
- 日志默认 Info；调试日志需手动开启。
- 日志不保存 API Key、完整单元格矩阵和文件全文。
- 上传给模型的采样数据必须在聊天中显示“本次发送范围/token 估算”。
- 每次模型请求前展示实际引用范围；用户可选择移除某个引用。

### 6.4 数据发送控制

| 数据 | 默认行为 |
| --- | --- |
| 当前选区 <= 300 cells | 发送值/公式结构化样本 |
| 当前选区 > 300 cells | 只发送地址、维度、表头、类型和统计摘要；用户点击“读取完整选区”后才读取 |
| UsedRange > 10000 cells | 只发送结构概览 |
| 手动引用 | 每个引用单独显示范围和估算 token，用户可移除 |
| 敏感格式 | 检测身份证/手机号/银行卡正则并提示用户确认后再发送 |

## 7. 数据模型

### 7.1 SQLite 表

```sql
CREATE TABLE model_profiles (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  base_url TEXT NOT NULL,
  api_key_cipher BLOB NOT NULL,
  model_name TEXT NOT NULL,
  temperature REAL DEFAULT 0.1,
  max_tokens INTEGER,
  protocol_mode TEXT CHECK(protocol_mode IN ('native','fallback','auto')) DEFAULT 'auto',
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE sessions (
  id TEXT PRIMARY KEY,
  title TEXT NOT NULL,
  workbook_hash TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE messages (
  id TEXT PRIMARY KEY,
  session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
  role TEXT NOT NULL,
  content TEXT NOT NULL,
  context_json TEXT,
  token_in INTEGER,
  token_out INTEGER,
  created_at TEXT NOT NULL
);

CREATE TABLE tool_events (
  id TEXT PRIMARY KEY,
  session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
  message_id TEXT,
  call_id TEXT NOT NULL,
  tool TEXT NOT NULL,
  args_json TEXT NOT NULL,
  result_json TEXT NOT NULL,
  risk_level INTEGER NOT NULL,
  status TEXT NOT NULL,
  started_at TEXT NOT NULL,
  finished_at TEXT NOT NULL
);

CREATE TABLE operations (
  id TEXT PRIMARY KEY,
  session_id TEXT NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
  workbook_hash TEXT NOT NULL,
  workbook_path TEXT,
  tool TEXT NOT NULL,
  args_json TEXT NOT NULL,
  affected_json TEXT NOT NULL,
  undo_kind TEXT NOT NULL CHECK(undo_kind IN ('cells','full_file','sheet_file','rows_columns')),
  payload_path TEXT,
  payload_json TEXT,
  before_digest TEXT,
  after_digest TEXT,
  can_undo INTEGER DEFAULT 1,
  created_at TEXT NOT NULL
);

CREATE TABLE skills (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  directory TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  description TEXT NOT NULL,
  version TEXT,
  enabled INTEGER DEFAULT 1,
  valid INTEGER DEFAULT 1,
  updated_at TEXT NOT NULL
);
```

### 7.2 配置文件

`%APPDATA%\ExcelAiAssistant\config.json`：

```json
{
  "general": {
    "language": "zh-CN",
    "theme": "system",
    "startupPane": false
  },
  "agent": {
    "maxToolCallsPerTurn": 15,
    "turnTimeoutSeconds": 180,
    "previewRows": 20,
    "sampleCellLimit": 300,
    "largeRangeCellLimit": 10000
  },
  "safety": {
    "confirmNormalWrites": true,
    "enableHighRiskTools": false,
    "scriptTimeoutSeconds": 30,
    "scriptOutputLimitKB": 64
  },
  "undo": {
    "retentionDays": 30,
    "maxOperationsPerWorkbook": 500,
    "backupRetentionDays": 7,
    "cellSnapshotLimit": 500
  }
}
```

## 8. UI 设计

### 8.1 信息架构

侧边栏五个主页面：

1. 聊天：消息流、输入框、引用条、工具卡片。
2. 引用：当前选区和手动添加区域。
3. 历史：会话列表、操作记录、筛选、回滚。
4. 技能：启用/禁用、路径、描述、工具白名单。
5. 设置：模型方案、Agent 参数、安全和撤销。

Ribbon 提供：

- “智能助手”：显示/隐藏 Task Pane；
- “新会话”；
- “引用当前选区”；
- “AI 配置”；
- “历史与回滚”；
- “帮助”。

### 8.2 聊天卡片

每条 AI 请求下方展示：

```text
[上下文]
工作簿：Book1.xlsx
工作表：Sheet1
选区：A1:E25
发送：结构 + 前 5 行样本，约 1,120 tokens

[思考过程]（可折叠）

[工具]
@preview_excel_range
参数：A1:E25, rows=5
结果：5 rows x 5 cols

[待确认修改]
目标：Sheet1!F2:F25
动作：写入公式 =IF(E2="","",D2*E2)
旧值：24 个空单元格
新值：24 个公式
[确认执行] [取消]
```

### 8.3 交互细节

- Enter 发送，Shift+Enter 换行。
- `@` 弹出工具补全，Tab 选择。
- `/` 弹出 Skill 补全，Tab 选择。
- 执行中显示取消按钮；取消后不再发起下一次模型请求，但已经开始的工具操作尽量完成并记录。
- Excel 编辑单元格时禁用写入按钮，避免 COM `0x800AC472` 冲突。
- 文件只读或工作表保护时显示明确原因，并提供“跳转到解除保护”的提示，不代输密码。

## 9. 非功能需求

| 指标 | 目标 |
| --- | --- |
| 冷启动加载插件 | Excel 打开后额外延迟 < 1.5 秒 |
| Task Pane 首帧 | < 800 ms |
| 读取 10000 cells 并摘要 | < 700 ms，不含模型耗时 |
| 写入 10000 cells | < 2 秒 |
| 撤销面板加载最近 50 条 | < 400 ms |
| 内存增量 | 空闲 < 180 MB |
| 崩溃率 | 会话崩溃率 < 0.5% |
| 安装成功率 | 干净 Win10/Win11 + Excel 2016/365 成功率 > 99% |

兼容性要求：

- x64/x86 Excel；
- Excel 2016/2019/2021/365；
- 高 DPI 100%、125%、150%、200%；
- 中文/英文 Windows；
- OneDrive 同步文件、受保护视图、只读推荐模式；
- 工作簿保护和工作表保护；
- 包含公式、条件格式、表格对象、数据验证、图片、宏的 `.xlsm`。

## 10. 测试计划

### 10.1 单元测试

- 地址解析与跨表地址拒绝；
- UsedRange 边界；
- 二维数组矩形校验；
- 类型转换和日期序列号；
- 公式本地化符号处理；
- Function Calling 与 Fallback parser；
- Skill front matter 校验；
- 风险分级决策表；
- 快照/回滚幂等性；
- DPAPI 加解密；
- SQLite migration。

### 10.2 集成测试

使用自动化 Excel 实例验证：

1. 新建工作簿生成销售明细；
2. 写入 SUM/SUMIFS/COUNTIFS/XLOOKUP；
3. 删除含公式和条件格式的行列后回滚；
4. 删除工作表后恢复；
5. `.xlsm` 保持宏不丢失；
6. 受保护视图下的降级提示；
7. OneDrive 文件同步冲突提示；
8. x86/x64 SQLite/COM 注册；
9. WebView2 与宿主桥接断线重连；
10. 模型返回非法 JSON/tool args。

### 10.3 UAT 场景

| 编号 | 场景 | 通过标准 |
| --- | --- | --- |
| U1 | 选中销售明细提问“这份数据有什么特征？” | 正确识别字段、行数、缺失值和主要分布 |
| U2 | “在 H2:I10 生成产品和销售总额” | 表头、公式、引用范围正确 |
| U3 | “删除全空行，把 A 列文本日期改成日期” | diff 正确、执行成功、可回滚 |
| U4 | “基于 A1:F200 做月度销售折线图” | 图表类型、数据源和放置位置正确 |
| U5 | “删除 Sheet2” | L3 确认后删除，可从备份恢复 |
| U6 | “运行这段 VBA 清理样式” | 高危开关关闭时拒绝；开启后逐次确认 |
| U7 | 切换 DeepSeek/Qwen/OpenAI 方案 | 测试连接成功；工具循环正常 |
| U8 | 断网/错误 Key/限流 | 有清晰错误和重试建议，不留半成品写入 |

## 11. 里程碑

假设 2 名 C#、1 名前端、0.5 名产品/测试，全职推进。

| 阶段 | 时间 | 交付 |
| --- | --- | --- |
| M0 技术验证 | 第 1 周 | VSTO Task Pane、WebView2、COM 读写、x86/x64 PoC |
| M1 核心闭环 | 第 2-3 周 | 上下文、预览、读取、三种写入、模型双模式、基础聊天 |
| M2 安全与撤销 | 第 4-5 周 | 风险分级、确认卡、快照、SQLite 历史、回滚 |
| M3 进阶工具 | 第 6-7 周 | 25 个工具、计算草稿、目录页、透视表、图表 |
| M4 Skill 与设置 | 第 8 周 | SkillLoader、白名单、多模型配置、连接测试 |
| M5 稳定打包 | 第 9 周 | 性能、异常、安装器、卸载、日志、UAT |
| M6 内测发布 | 第 10 周 | 内测包、用户手册、已知问题清单 |

若只有 1 名开发者并行承担前后端，工期建议放宽至 14-18 周。

## 12. 成本估算区间

人力成本按中国一线城市外包/交付团队粗略区间估算：

| 项目 | 区间 |
| --- | --- |
| MVP 10 周，2.5-3.5 人投入 | 人民币 18万-38万 |
| 1 人自研 14-18 周 | 机会成本约 8万-25万，视薪资水平 |
| UI 定制设计 | 1万-5万 |
| 代码签名证书 EV/OV，年付 | 约 $150-$600 |
| 测试机与 Office 许可 | 1万-3万一次性或已有设备摊销 |
| 模型费用 | 由用户自带 Key 承担，产品方无固定成本 |

如果外购成熟 VSTO 商业框架或委托资深团队，价格会上浮；如果采用更激进的 .NET 8 + Excel COM Add-in 方案并放弃 Excel 2016，可减少部分老技术债但需要重新验证企业机器兼容性。

## 13. 主要风险

| 风险 | 影响 | 缓解 |
| --- | --- | --- |
| COM 修改破坏 Excel 原生撤销栈 | 用户误以为 Ctrl+Z 可撤回 AI 修改 | UI 明确提示使用助手历史回滚；高危操作备份 |
| Excel 忙碌/编辑态导致 RPC 抱歉 | 写入失败 | 检测 `0x800AC472`，等待重试，编辑态禁用按钮 |
| x86/x64 原生依赖错乱 | 安装后无法启动 | SQLite/WebView2/VC++ runtime 分架构测试 |
| 模型工具参数幻觉 | 写错区域 | JSON Schema 校验、目标存在性检查、diff 确认 |
| 大区域导致 token 超限 | 调用失败或费用暴涨 | 强制采样和统计摘要，token 预估 |
| OpenAI 兼容端点行为不一致 | Function Calling 不可用 | native/fallback 双模式 |
| VBA/脚本滥用 | 数据破坏或安全风险 | 默认关闭、白名单、超时、输出截断、逐次确认 |
| `.xls` 老格式兼容差 | 保存/格式丢失 | 只保证当前打开可读写；重大变更前完整备份 |
| 企业安全软件拦截 VSTO | 安装失败 | 提供诊断工具和签名版本 |
| 用户把敏感表整份发给云端 | 隐私事故 | 默认采样、敏感正则提醒、发送范围可视化 |

## 14. 发布验收清单

发布前必须满足：

1. 5 个 MVP 验收场景全部通过。
2. Excel 2016/2019/2021/365 x64/x86 至少各完成一轮冒烟。
3. 卸载后不残留 Excel 加载项注册和计划任务；用户数据可选保留。
4. 所有 L2/L3 操作有确认记录和回滚路径。
5. API Key 不出现在明文配置和日志。
6. 崩溃日志能定位未处理异常。
7. 安装包在无 WebView2 Runtime 机器上自动安装或给出下载指引。
8. README 覆盖安装、配置模型、首次使用、回滚、卸载和故障排查。

## 15. 下一步

建议立即开始三件事：

1. 建 .NET Framework 4.8 VSTO 工程，打通 Excel COM、Task Pane、WebView2 和 x86/x64 调试。
2. 定义 25 个工具的 JSON Schema 和 SafetyPipeline 接口，先写契约测试。
3. 实现 `get_current_excel_context -> preview_excel_range -> write_excel_range` 最小 Agent 循环，用真实销售表跑通第 1、2 个验收场景。

## 16. 已锁定决策记录

| 决策点 | 结论 |
| --- | --- |
| 产品形态 | Windows 本机 VSTO Excel 加载项，不做纯 Office.js 首版 |
| UI | VSTO Custom Task Pane + WebView2 前端 |
| Excel 范围 | Excel 2016/2019/2021/365，x64 和 x86；不做 WPS |
| 文件范围 | 操作当前已打开工作簿；Excel 可打开的 `.xlsx/.xlsm/.xls` 均可处理 |
| Agent 协议 | 优先原生 Function Calling，失败降级自定义 JSON 协议 |
| 工具范围 | 第一期 25 个进阶工具；本地脚本/EXE 类放二期且默认关闭 |
| 高危能力 | 默认关闭；显式开启后每次执行前仍强制确认 |
| 撤销策略 | 普通操作记录区域新旧值；大变更和删除类自动完整备份 |
| 保留期限 | 操作记录默认 30 天/每簿 500 条；完整备份 7 天 |
| Skill | 本地 `SKILL.md` + front matter + 工具白名单 |
| 模型配置 | 多方案命名保存，支持 Base URL/API Key/模型名/连接测试 |
| 数据隐私 | 用户可见并控制发送范围；默认采样和摘要 |
| 存储 | `%APPDATA%\ExcelAiAssistant`，SQLite + DPAPI 加密 Key |
| 语言 | 第一版中文 only，但文案集中资源文件 |
| 授权 | MVP 免费内测，不做授权系统 |
| 安装 | Inno Setup/WiX 正式安装包，后续购买代码签名证书 |
