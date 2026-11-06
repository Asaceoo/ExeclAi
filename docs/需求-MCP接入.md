# 需求-MCP 接入（把助手工具链开放给外部 AI）

> 版本：v0.1（2026-09-11，规划态立项——**仅方案，未实施**；触发：用户问"看板功能做个 MCP 给 AI 链接可以不"）
> 状态：待评审选型（M-A / M-B 二选一）；建议先走 grill-me 拷问会
> 上游：`src/ExcelAiAssistant.Excel/ExcelToolHostRegistration.cs`（147 工具注册表）、`docs/需求-原生数据看板.md`（v0.9）

## 1. 背景与可行性结论

**可行。** 全部业务能力已收敛在 `ToolRegistry`（工具名 + Schema + Execute 委托），与 UI 解耦——Md1Probe/RealMachineVerify 已证明同一注册表可在无 Add-in UI 的独立进程中驱动 Excel（out-of-proc COM）。MCP server 只需做一层薄封装：**MCP 协议 ↔ ToolRegistry 调用**。

## 2. 两种形态（二选一或分期）

| 形态 | 说明 | 优点 | 代价/风险 |
|---|---|---|---|
| **M-A 独立 MCP server**（推荐先做） | 独立进程：启动/附着 Excel COM → 构建 ToolRegistry → 以 stdio 暴露 MCP tools。适合 Claude Desktop / 任何 MCP 客户端，**无需安装插件** | 复用现有 147 工具与自检/回滚/确认门代码；交付快（约 1~2 天） | COM 会话生命周期管理（忙等待 0x8001010A、PID 强杀清理，踩坑手册 §5.2 模式已有） |
| **M-B Add-in 内置 MCP 端点** | 已安装的助手在 WebView2/本地端口暴露 MCP（SSE），外部 AI 连上来复用运行中实例 | 复用用户正在用的 Excel 会话与确认 UI | 认证/会话管理/端口安全（仅 loopback）约 3~5 天；WebView2 侧改动面大 |

**建议**：P1 做 M-A（`tools/McpServer/`，与 SlicerProbe 同族的控制台宿主），P2 评估 M-B。

## 3. 工具面（P1 暴露清单）

- **P1（看板/图表子集，~15 个）**：`detect_dashboard_data_readiness`、`suggest_chart_mapping`、`render_dashboard`、`render_rank_dynamic`、`render_dim_cross`、`render_pnl_brief`、`render_target_track`、`create_dashboard_kpi_brief`、`create_chart_advanced`、`create_combo_chart`、`create_doughnut`、`set_linkage`、`verify_dashboard`、`verify_chart`、`delete_dashboard`；
- **P2**：其余工具族分批暴露（读写/清洗/公式/超链接/VBA 等），按 RiskLevel 配置暴露白名单；
- `Dangerous` 级（delete_dashboard 等）MCP 侧默认**二次确认**：返回 `requires_confirmation` 载荷，客户端回传 confirm 才执行（映射现有确认门语义）。

## 4. MCP 协议映射

- `tools/list` ← ToolRegistry 枚举（name/description/inputSchema=M3ToolSchemas 直转 JSON Schema）；
- `tools/call` ← descriptor.Execute(args)；返回 `ToolResult` → `content[].text`（Ok=false 时 `isError=true` + ErrorCode）；
- 传输：P1 用 stdio（Claude Desktop 原生支持）；Excel 实例策略 = server 启动时附着已有实例或新建（配置项，默认新建+会话结束强杀，SpawnedPids 同款红线：绝不动用户实例）。

## 5. 红线与验收

- 不绕过任何安全设施：保护态守卫、确认门、覆盖审计语义原样保留；
- 用户文件只读红线：MCP 模式下打开用户文件必须走副本或显式只读参数（UD 套件同款）；
- AC-M1：Claude Desktop 配置后能列出新工具并成功渲染一份样例看板且自检通过；
- AC-M2：Dangerous 工具无确认不执行；AC-M3：Excel 忙（0x8001010A）时优雅重试/报错不崩溃；
- AC-M4：会话结束无 Excel 进程残留。

## 6. 成本与期次

| 期次 | 内容 | 估算 |
|---|---|---|
| P1 | M-A server + 看板子集 15 工具 + stdio + 确认门映射 | 1~2 天 |
| P2 | 白名单配置化 + 其余工具族 + 文档 | 1~2 天 |
| P3 | M-B Add-in 端点（SSE + loopback + 认证） | 3~5 天 |

## N2/AC-M1 口径声明（M27 第四期收口）

- **AC-M1**：P1 交付口径已从「看板子集」正式改为**全工具面**——MCP server 暴露完整注册表（当前 188：167 代码注册 + 4 智能体桥接 + 浮动 ct_ 沉淀），读写分级由 RiskLevel 驱动，只读模式（缺省）自动剔除全部非 ReadOnly 工具。
- **N2（确认口径）**：MCP/CLI 侧**只有 Dangerous 级走确认码门**（P0-3 一次性码，人工从文件读码）。StructuralChange 级（建表/删形状/导入等）依设计**不进门**，安全兜底为三层：①执行层 verify_* 自动自检；②OperationRepository 审计落库（operations-mcp.db / operations-cli.db）；③undo_excel_operation 可回滚。与加载项侧 SafetyPipeline 的差异是刻意的——加载项面向交互式用户（面板按钮即确认），MCP 面向自主 Agent（逐次人工确认会阻断自动化，故用「可回滚」换「免打断」）。需要更强确认的场景由调用方在客户端策略层叠加。
- **确认码泄露防线（D1，M27 收口）**：确认码只存在于 pending 文件内容中；文件名（`<tool>-pending.txt`）与响应文案均不含码——AI 无法从任何可回显通道自答。回归锁定见 `McpSecuritySemanticsTests`。
