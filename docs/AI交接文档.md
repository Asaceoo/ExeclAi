﻿# JYYJ助手 · AI 交接文档

> 本文档写给**接手本项目的 AI 助手**（ZCode/Claude/Cursor 等），目标是让它在 10 分钟内建立完整项目认知，并按既定 SOP 安全地更新版本与功能。
> 最后更新：2026-11-03（v3.9.1.0）——M21-M23 三轮对抗审查收敛 33 项修复（MCP server 两轮 20 项 + CLI 第三轮 13 项）、新增 MCP server 与 CLI 双入口（173 工具面）、README 重写（3 概念图+用法+20 实例+免责声明）、双手册免责声明章、真机 119 断言全过、1321 单测、GitHub 推送通道恢复（SSH over 443；M26（v3.9.1.0）L2元素层/草稿模式/外部数据源/AI智能体桥接/MCP懒初始化 + 三视角审查两轮收敛 代理）。前次：2026-10-02（v3.8.3.1）。
---

## 0. 你必须遵守的工作流（用户全局指令）

进入审查阶段后必须执行：①三视角对抗性审查（工具实现者/审查者/AI 使用方）→ ②多轮迭代修复直到连续一轮无新问题 → ③工具闭环拿真实报错（不许凭推测判断通过）→ ④清单兜底（性能/边界/残留/兼容）→ ⑤真机验证（真实 Excel 实例，不是只跑单测）。

任务清单保持干净：只记录当前有效路径，不携带废弃方案。

---

## 1. 项目一句话

**JYYJ助手**：内嵌 Excel 的 AI 办公助手（COM Add-in + WebView2 侧边栏）。用户用自然语言操作表格，AI 通过 167 个工具执行写入/清洗/分析/图表/超链接/VBA/流程编排；**所有写入操作被执行层自动自检兜底、保护态集中守卫（SHEET_PROTECTED）、可回滚、运行日志全量留痕**。

- 当前版本：**v3.9.1.0**（版本史与各批次详情见本文档顶部「最后更新」行；门禁基线：**1321 单测** + md1 62/62 + ud 31/31 + udchart 48/48 + Excel 真机 119/119 + WPS full 131 组 + 167 工具卡点）
- 技术栈：.NET Framework 4.8 · C#（LangVersion latest, **TreatWarningsAsErrors=true**）· VSTO 风格 COM · WebView2 · System.Data.SQLite · Newtonsoft.Json · xUnit
- 软件地址：https://github.com/Asaceoo/ExeclAi
- 支持 Excel：**2013 / 2016 / 2019 / 2021 / Microsoft 365**（32/64 位），.NET Framework 4.8，Windows 10/11，需 WebView2 Runtime

---

## 2. 目录与关键文件地图

```
D:\execl1.0\
├─ ExcelAiAssistant.sln
├─ Directory.Build.props            # LangVersion=latest + 警告即错误（新代码必须零警告）
├─ README.md                        # 版本号在这里（发版要同步改）
├─ src\
│  ├─ ExcelAiAssistant.Addin\       # COM 宿主层
│  │  ├─ ExcelAiAddIn.cs            # 加载项入口（OnConnection 等）
│  │  ├─ SidebarBridge.cs           # ★ 5000+ 行：前端桥、会话、系统提示词(4270+行附近)、
│  │  │                             #   一键自检(AI研判)、全功能巡检入口、自检历史接入
│  │  ├─ FeatureHealthSweep.cs      # 全功能健康巡检（22 项子系统体检，临时工作簿）
│  │  ├─ DiagnosticSmokeRunner.cs   # 旧版一键自检 smoke（VBA/Python 通道探针）
│  │  ├─ DiagnosticsLog.cs          # 运行日志（进程内环形缓冲 + 文件）
│  │  ├─ Diagnostics\AppLog.cs      # ★ 跨程序集日志门面（Core/Excel 层统一汇入，Addin 启动注入）
│  │  └─ Sidebar\index.html         # ★ 前端单页（内联 JS/CSS；注意 JS 静态审计单测）
│  ├─ ExcelAiAssistant.Core\        # Agent 引擎（无 COM 依赖，纯逻辑可单测）
│  │  ├─ Agent\AgentRunner.cs       # ★ 回合编排：模型调用/工具执行/确认门/自动自检包装
│  │  ├─ Agent\AutoVerifyPolicy.cs  # ★ 主工具→verify 映射（TryPlanAll/EnrichFromResult/
│  │  │                             #   MappedPrimaryTools/DocumentedWriteToolExemptions）
│  │  ├─ Agent\AutoVerifyMonitor.cs # 元自检：自检覆盖率统计与链路失效告警
│  │  ├─ Agent\ToolSummaryBuilder.cs# 工具结果的人可读摘要（前端卡片）
│  │  ├─ Agent\OpenAiCompatibleClient.cs # 模型客户端（实现 IChatCompletionClient 接口）
│  │  ├─ Diagnostics\PitfallLibrary.cs   # 踩坑失败模式库（签名→建议，AI 研判消费）
│  │  ├─ Diagnostics\SelfCheckHistoryStore.cs # 自检历史 SQLite（趋势，上限500条）
│  │  ├─ Tools\ToolRegistry.cs      # 工具注册表 + NormalizeValues（注意 JSON 形态陷阱）
│  │  └─ Configuration\…            # AgentSettings/WritePreference/Memory 存储等
│  ├─ ExcelAiAssistant.Excel\       # 工具服务层（全部 Excel 能力）
│  │  ├─ ExcelToolHost.cs           # ★ CreateRegistry()=全部工具组装根；WriteRangeCore(快照+撤销)；
│  │  │                             #   NormalizeCellEntries(JSON对象数组归一化)；UndoOperation
│  │  ├─ ExcelVerifyService.cs / ExcelVerifyJudgments.cs / VerifySchemas.cs
│  │  │                             # ★ verify_* 自检工具（判定器纯静态可单测）
│  │  ├─ ExcelDataToolsService.cs   # 排序（含行守恒硬护栏 WriteBack）/清洗/去重/筛选
│  │  └─ 各 *Registration.cs        # 分域工具注册
│  └─ ExcelAiAssistant.Tests\       # 979 个 xUnit 单测（含粘合层集成测试/JS静态审计）
├─ tools\
│  ├─ package-xai.ps1               # ★ 打包脚本（内置发布门禁：单测+真机，-SkipReleaseGate 逃逸）
│  ├─ RealMachineVerify\            # ★ 真机 harness（22 用例组 119 断言，独立 Excel 实例）
│  └─ installer\*.iss + register-auto.ps1   # Inno Setup 与 COM 注册编排
├─ dist\                            # 发布产物（当前版本 3.9.1.1，WPS 版）
│  ├─ installer\JYYJ助手-3.9.1.1-Setup-wps.exe           # WPS 版轻量安装包（ISCC 编译，4.70MB）

│  ├─ JYYJ助手-3.9.1.1-wps-安装说明.txt        # WPS 版安装说明（含 VBA 支持库三条安装路径+网址，随安装器/便携包分发）
│  ├─ JYYJ助手-3.9.1.1-wps-x64.zip / -x86.zip      # 便携包（3.44/3.42MB）
│  ├─ JYYJ助手-3.9.1.1-wps-x64\ / -x86\            # 便携解包目录（含一键安装.cmd）
│  ├─ JYYJ助手-<旧版>-单文件便携版.cab          # 可选，makecab 单独产出（package-xai 不生成）
│  └─ archive\ + _旧版本待删除\                 # 旧版本归档/待删
└─ docs\                            # 用户手册(HTML)/技术手册(MD)/踩坑手册(MD)/本文件
```

---

## 3. 架构与数据流（一轮对话发生了什么）

```text
用户输入 → SidebarBridge.SendPrompt
  ├─ CreateSystemPrompt()          # 系统提示词：含自动自检协作规则、踩坑规避、技能注入
  ├─ CreateHost/CreateRegistry()   # ExcelToolHost（挂 OperationRepository 撤销历史）
  ├─ new AgentRunner(client, registry, ConfirmTool, options{Trace→DiagnosticsLog})
  └─ AgentRunner.RunAsync：        # 循环直到 action=final 或达工具上限(默认15)
       模型 CompleteAsync → 工具调用
       → 确认门（NormalWrite 需用户确认；Dangerous 默认拒绝）
       → tool.Execute(args)        # Excel COM 操作
       → RunAutoVerify()           # ★ 执行层自动自检（见第 4 节）
       → 观察结果回模型（tool 消息 / auto_verify 观察消息）
  回合结束 → EmitAutoVerifySummary（元监控摘要→result+运行日志）
  → tool-error 事件、session-updated、selfcheck-summary(前端 toast)
  → 自检摘要落 SelfCheckHistoryStore
前端(WebView2 index.html) ⇄ Publish(事件)：busy / agent-progress / tool-error /
  confirmation-request / selfcheck-summary / session-updated
```

前端调用 C# 的方式：`window.chrome.webview.hostObjects.excelAiBridge.<公开方法名>()`——**桥接层 public Task<string> 方法即前端 API**。

---

## 4. 自检与质量体系（本项目最核心资产，勿破坏）

分六层，每层失效都有下一层兜底：

| 层 | 机制 | 关键文件 |
|---|---|---|
| ① 执行层自动自检 | 16 类主写操作成功后**服务端自动**执行 verify_* 回读核对；预期参数从实参/执行结果自动推导；批量回写逐格(≤5格)；模型收到 `auto_verify` 观察消息，verdict≠ok 必须修正后才能报完成 | AutoVerifyPolicy / AgentRunner.RunAutoVerify |
| ② verify_* 工具族（18个） | chart / cell_format / conditional_format / sort_filter / formula / write_result(含 type_aware 文本型数字检测) / data_clean / pivot(含明细汇总数量守恒) / marker / structure / data_validation / sparkline / merge / hyperlink / table / protection / image / freeze_panes（以源码注册为准）。统一 verdict+issues+selfcheck | ExcelVerifyService / ExcelVerifyJudgments |
| ③ 语义守恒护栏 | 排序行多重集守恒（写前拦截 CONSERVATION_VIOLATED）；透视守恒对账（口径与 SumByRow 严格一致）；类型感知 | ExcelDataToolsService.WriteBack / ExcelVerifyJudgments |
| ④ 元自检 | AutoVerifyMonitor：回合内"有可自检写操作但 0 条自检执行"→ `SELF_CHECK_LINK_SUSPECT` 告警进运行日志；PitfallLibrary 有对应规则可被 AI 研判识别 | AutoVerifyMonitor / AgentRunner |
| ⑤ 失败模式库与研判 | PitfallLibrary（15+条"日志签名→手册级建议"）+ 一键自检 AI 研判 + **全功能巡检**（22 项子系统体检，临时工作簿，逐项写运行日志） | PitfallLibrary / SidebarBridge / FeatureHealthSweep |
| ⑥ 验证与门禁 | 979 xUnit 单测（含粘合层集成测试、映射↔switch 一致性锁、JS 静态审计、M16 缺陷原型源码级审计 11 条）→ 真机 harness 119 断言 → WPS full 基线校验 → 打包门禁 | Tests / tools\RealMachineVerify / tools\release\gate.py |

**覆盖审计**：所有写能力工具（NormalWrite/StructuralChange/Dangerous）必须在 `MappedPrimaryTools 44 类）或 `DocumentedWriteToolExemptions 95 条带理由豁免）里，`ct_` 前缀=用户自定义沉淀工具自动豁免。判定入口 `AutoVerifyPolicy.IsExemptedFromAutoVerify`，harness 用例17 与全功能巡检两处执行——**新工具不表态就发不出去**。

---

## 4A. 错误码与恢复路径（AI 使用方视角，M17.3 补）

### SHEET_PROTECTED（保护态集中守卫，出现频率最高）

**含义**：目标工作表处于保护状态（审阅→保护工作表），写工具被宿主层守卫前置拦截，**尚未触碰任何单元格**。错误消息自带表名与工具名。

**恢复闭环（按序执行）**：
1. 调 `verify_protection`（sheet=<错误消息中的表名>）确认保护状态与锁定格分布；
2. **不要绕过守卫重试**——同参数重试必然再被拦，徒耗回合；
3. 向用户确认后调 `protect_sheet_watermark`（action=unprotect，无密码或用户提供密码时）；
4. 重 `verify_protection` 确认 `expected_writable=true`；
5. 重放原写入；完成后如用户需要，可再以 action=protect 恢复保护。

**注意**：`add/rename/delete_excel_worksheet` 三个结构类工具按**工作簿结构保护**（ProtectStructure）判定，解除保护用 protect_sheet_watermark（action=unprotect，sheet 传目标表）。

### 守卫豁免的四类工具（遇到失败不要按 SHEET_PROTECTED 归因）

以下 26+1 类工具豁免于保护态守卫（豁免理由内置于 `SheetProtectionGuard.Exemptions`，ct_ 前缀用户沉淀工具按前缀规则豁免）——它们失败时请按各自错误码归因：

| 类别 | 代表工具 | 为什么豁免 | 失败时怎么办 |
|---|---|---|---|
| 脚本/进程类 | run_vba_snippet / run_python_excel / run_cli_command / undo_excel_operation | 执行体语义自洽，保护态冲突由脚本/恢复自身报错 | 读脚本自身错误输出；undo 看快照错误码 |
| 工作簿/文件级 | split_workbook / import_* / export_*_pdf / open_excel_external_book | 不写目标工作表单元格，保护态与它们无关 | 看文件路径/格式错误码 |
| 视图级 | freeze_panes | 窗口视图设置，保护态下原生可用 | 极少失败；看窗口状态 |
| 产出为新表 | split_sheet_by_field / create_sheet_by_value / create_toc_sheet | 写入目标是新建表，源表只读 | 看新建表/参数错误码 |

### 自动自检（AutoVerify）两档口径

- **44 类主写工具**：成功后服务端自动 verify 回读，verdict≠ok 时 AI 必须修正后才能报完成；
- **其余写能力工具**：在带理由豁免清单（`AutoVerifyPolicy.DocumentedWriteToolExemptions`）——多为确定性变换或自有校验语义，**AI 应自行用对应 verify_* 工具核对**；
- **断言型 verify**（cell_format/sort/conditional_format 等）：无期望值 = `no_expected` 失败；**探测型 verify**（protection/freeze/image/hyperlink）：无期望值 = ok（观察用途）。

### 其它高频错误码速查

| 错误码 | 含义 | 恢复路径 |
|---|---|---|
| INVALID_SHEET | 表名不存在（不静默回退活动表） | 先 `get_excel_sheet_list` 核对表名精确拼写 |
| CONSERVATION_VIOLATED | 排序/写回行数守恒不成立（写前拦截，未落盘） | 核对 region/参数后重试，勿盲目重发同参 |
| NOT_ALLOWLISTED / CLI_DENIED | 脚本不在白名单 | 提示用户在设置中加白，勿尝试路径变换绕过 |
| BLOCKED_PRIVATE_ADDR | fetch_url 目标为内网/回环地址 | 属安全拦截，不提供绕过路径 |
| INVALID_ARGS / 解析失败 | 守卫解析不到目标表时**不越权拦截**（放行回工具自身错误路径），最终由工具返回错误码 | 表名/参数错误：先 get_excel_sheet_list 核对，再修正参数重试 |

---

## 5. 全部用户可见功能（清单随版本演进）

- **对话执行**：自然语言操作表格；先选区再提问；模糊需求 AI 主动确认；单轮工具上限可调（1-1000）
- **数据处理**：读写、清洗（去空行/修日期/填空白）、排序（行守恒）、去重、筛选、转置、透视汇总（守恒对账）、异常标记、文本/日期/数字变换等 60+ 工具
- **分析可视化**：图表生成+渲染像素核验、迷你图、统计分析、分析报告
- **公式**：生成、逐层解释、引用转换
- **函数库面板**：侧边栏「函数库」页，内置 **319** 个 Excel 内置函数（财务/逻辑/文本/日期/查找/数学·三角/统计/其他），支持按类浏览、关键字搜索、常用榜、**按首版可用版本筛选（全部/2007/2010/2013/2016/2019/2021/365，列表与详情显示版本徽章）**，点击插入公式骨架（v2.18.11.0 扩充 / v2.19.0.0 补版本标注与版本筛选 / **v2.19.1.0 修复 FX_LIB 加载 TDZ 崩溃** / **v2.19.4.0 新增 7 个 365 函数 GROUPBY/PIVOTBY/PERCENTOF/TRIMRANGE/REGEXTEST/REGEXEXTRACT/REGEXREPLACE 并补版本标注审计**）
- **VBA**：生成、静态校验（拦截 2D AutoFit 等 1004 写法）、执行（确认门+探针）、自动沉淀为 ct_ 工具
- **Python**：独立常驻 worker 执行数据分析（run_python_excel 等）
- **流程编排**：多步流程模板，步骤级自检，失败单步重试
- **技能包**：内置 12 个（工资条/考勤/花名册等），支持 zip/SKILL.md 在线安装，`/技能名` 调用
- **多工作簿**：跨簿只读引用；网页停靠窗格（预设站点+自定义 URL）
- **回滚**：逐操作回滚按钮 + 冲突保护；快照存 SQLite
- **自检与监控**：每轮自动自检 + 「本轮自检」toast + 一键自检(AI 研判) + 全功能巡检 + 功能健康监控条 + 自检趋势（最近20次通过率/高频问题）
- **运行日志全功能覆盖（v2.19.4.0 增强）**：新增 `AppLog` 跨程序集日志门面，Addin 启动注入 `DiagnosticsLog` 接收器，使 Core/Excel 层全部写操作/工具执行/Python worker 统一汇入运行日志；工具执行记录耗时与成功/失败/异常三路径（含子项自检）；Python worker stderr 落盘（日志折叠多行不撑爆），定位任意功能错误不再靠猜
- **设置页**：模型方案（DPAPI 加密）、写入偏好（auto/formula/value）、单轮上限、受控执行白名单、长期记忆、运行日志（级别过滤/导出/清空/自动刷新）、软件介绍（能做什么/新手三步/七大面板/支持版本与环境/安全隐私）、帮助中心（FAQ+找不到报错喂数据），关于/版本展示**版本号 + Excel 版式名**（15.0→2013、16.0→"2016 或更高"，2016/2019/2021/365 仅 Version 无法区分、如实标注）、工具目录（v2.19.3.0 增强：七大面板修正+友好版式名）
- **快捷功能区参数弹窗**：18 个快捷工具带可配置弹窗，每个配置项带说明文字与真实选项（转置支持跨表写回、文本拆分可指定结果起始列、文本清洗支持查找替换、透视可指定输出位置等）；窗口可自由缩放；选项经双向契约审计防"假选项"（v2.18.8.0 新增/修复）
- **M13+ 新增能力**：create_excel_table+verify_table（注意 WPS ListObjects.Add 会向下扩一行，WPS 下 verify_table 不断言行数）、insert_image+verify_image、freeze_panes+verify_freeze_panes（统一 Split 路径双宿主通用）、超链接工具族+verify_hyperlink、函数库 Excel/WPS 双视图、设置页「WPS 环境·指南」卡片
- **自检族扩展**：verify_* 自检族已扩至 **18 个**（源码注册为准），MappedPrimaryTools 44 类 / DocumentedWriteToolExemptions 95 条；发版走 `release.py bump` + `gate.py --pack`（见 §6.3，本仓库勿用 package-xai.ps1）

---

## 6. 开发与验证 SOP（改代码必读）

### 6.1 验证金字塔（从快到慢，全过才算完）

```powershell
cd D:\execl1.0-wps
dotnet build ExcelAiAssistant.sln                 # 0 警告 0 错误（警告即错误）
dotnet test ExcelAiAssistant.sln                  # 979 个单测
cd tools\RealMachineVerify && dotnet build && dotnet run
#   ↑ 真机 harness：独立 Excel 实例 + 临时工作簿，119 断言（含回滚端到端、负向用例、覆盖审计、自动自检链路）
#   退出码 0 = 通过；EXCEL.EXE 由 harness 按 PID 自杀，不残留
```

### 6.2 加一个新工具的标准动作

1. `XxxService` 实现 + `XxxSchemas` schema（参数名、类型、中文描述——schema 是模型的唯一使用说明）
2. 对应 `*Registration.cs` 注册：名称 / 描述 / **RiskLevel** / 执行委托 / schema
3. **覆盖审计表态**（缺了这步，门禁直接拦）：
   - 有自动自检语义 → `AutoVerifyPolicy.TryPlan` 加 case + `MappedPrimaryTools` 加名字 + `AutoVerifyPolicyConsistencyTests` 模板实参加分支
   - 无确定性期望 → `DocumentedWriteToolExemptions` 加条目并写明理由
4. 若入参是「对象数组」（JArray of JObject）：**必须**走 `NormalizeCellEntries` 归一化，禁止直接 `NormalizeValues`（JObject 会被当行展开，见踩坑 §7）
5. 单测（判定逻辑进纯静态类）；若影响自检链路 → 同步更新 harness 用例
6. 全 6.1 验证 + 更新 `docs/` 手册 + 踩坑沉淀（新坑进手册，日志可识别的进 PitfallLibrary）

### 6.3 发版 SOP

```powershell
# 1) 一键门禁直通发版（本仓库唯一发版路径；package-xai.ps1 是 Excel 主线遗留，勿用）
python tools\release\release.py bump --ver <新版本>   # props+README+四份手册（改名/戳/引用）一次同步
python tools\release\gate.py --pack                   # 单测→Excel 真机→泄漏断言→WPS full 基线→戳断言→ISCC
# 2) 仅打包不升版：python tools\release\release.py pack [--from <旧版本>]
# 3) 验证：dist\installer\JYYJ助手-<新版本>-Setup-wps.exe 生成、gate 日志 GATE PASS
# 4) 更新 docs 手册正文（md 改后用 tools/release/gen_manual_html.py 重生成 HTML）+ README 功能清单（neat-freak 同步）
```

### 6.4 dist 产物形态约定（WPS 版产物一律带 `wps` 后缀）

| 形态 | 路径 |
|---|---|
| 安装器（轻量） | `dist\installer\JYYJ助手-<版本>-Setup-wps.exe`（ISCC 编译，约 4.7MB；release.py pack 产出） |
| 便携包 | `dist\JYYJ助手-<版本>-wps-x64.zip / -wps-x86.zip`（release.py pack 产出，进程内压缩） |
| 便携解包目录 | `dist\JYYJ助手-<版本>-wps-x64\ / -wps-x86\`（app/ 为运行文件） |
| 安装说明 | `dist\JYYJ助手-<版本>-wps-安装说明.txt`（新版从旧版替换版本号继承） |
| 用户支持 | `diagnose.ps1`/`一键诊断.cmd`（一键诊断：Excel 版本/依赖/CLSID 注册实况/运行日志一报告收齐；开始菜单"JYYJ助手 一键诊断"入口）。ProgID/CLSID 清单与 register-auto.ps1 同步维护，见踩坑手册 §6.8 |

> 离线完整版（`-Setup-离线完整版.exe`，约 316MB）与单文件便携 CAB 仅 Excel 主版本提供，WPS 版暂无。

---

## 7. 高危陷阱浓缩（全文见 docs/踩坑记录与规避手册.md，代码改动前先对照）

1. **COM RCW 无状态铁律（§1.1）**：RCW 按对象标识缓存，`FinalReleaseComObject` 会连带销毁他人持有的同名 RCW。工具必须"每次重新解析、用完即放"，绝不跨调用缓存 Worksheet/Range。
2. **JSON「对象数组」入参（§7 新增教训，已踩 3 次）**：`write_excel_cells_batch` 的 cells、`clean_excel_data` 的 operations 是 JArray of JObject，直接 `NormalizeValues` 会把 JObject 当行展开（枚举出 JProperty）→ 假性"必须是对象数组"。必须走 `NormalizeCellEntries` 归一化。
3. **out-of-proc 动态 COM 收尾（§5.12）**：独立 Excel 进程里 `workbook.Close()/app.Quit()` 走动态绑定会在退出路径硬故障（连 corrupted-state handler 都拦不住）。harness 收尾只按 PID 强杀（启动前快照 PID）。
4. **日志自动刷新方向（本版修复）**：运行日志"新→旧"排列时滚动必须钉顶部；滚到底部=看最老。
5. **带 `$` 地址（本版修复）**：issue 地址偏移前必须去掉全部 `$`，否则所有问题标签错位到左上角。
6. **verify 无期望即误报**：verify_pivot/verify_marker 等在无 expected_* 时判 `no_expected`（非 ok）——自动自检必须先从主工具结果富化出期望，否则逼模型进假失败循环。
7. **dynamic 传染（§2）**：dynamic 形参使下游全变动态分派，元组元素名/匿名类成员不可访问；跨程序集读匿名结果用 JsonConvert+JObject。
8. **WebView2 单线程（§3.1）**：依赖 WebView2/UI 的操作不能 Task.Run；Ribbon onAction 已在 UI 线程。
9. **十六进制 HRESULT 进测试需转十进制**；`"\bisect"` 这类 `\b` 合法转义会静默变退格符。
10. **前端内联 JS**：一个语法错误废掉整个侧边栏——改动 index.html 后必跑单测（内置静态审计），不要 C# 裸拼复杂 JS。

---

## 8. 已知边界与未完成事项（按价值排序）

1. **变换类工具自动映射增量**：豁免清单里 20+ 个确定性变换（transpose/date_normalize/text_clean 等）期望值可推导，可建"变换类自检框架"批量转正。
2. **undo/clear 转正映射**：undo 后自动 verify（expected=快照 before_values）、clear 后 expected=空矩阵——各十几行代码。
3. **筛选子集守恒**：filter 输出行⊆源行 + 行数对账（当前靠模型按输出区核对）。
4. **VBA/Python 变化面观察消息**：脚本执行后自动区域 diff 附加给模型（只告知不判定）。
5. **导出类文件核验**：verify_file（存在/大小/魔数）映射到 export 系列工具。
6. **门禁进 CI**：单测关可挂 GitHub Actions（真机关需 Excel，留本机）；打包后自动打 tag。
7. **docs/ 手册的版本同步与 HTML 重生成**：四份手册（JYYJ助手技术手册/用户手册-<版本>.md/.html）的**文件名、版本戳、指针引用**由 `release.py bump` 的 `sync_manuals` 自动同步（M17.6 起，含零残留断言；本仓库不走 package-xai.ps1——那是 Excel 主线的脚本）；HTML 正文由 `tools/release/gen_manual_html.py <md> <旧html取CSS> <输出>` 从 md 重生成（复用现有风格）；PDF 需人工重生成（bump 只提醒）。改 md 正文后先重生成 HTML 再发版。

---

## 9. 快速自证（接手后先跑这个）

```powershell
cd D:\execl1.0-wps
dotnet build ExcelAiAssistant.sln                 # 应 0 警告 0 错误
dotnet test ExcelAiAssistant.sln --no-build       # 应 979 全过
cd tools\RealMachineVerify && dotnet build && dotnet run
#                                                  # 应 119 断言全过，EXCEL.EXE 无残留
```

三条全绿 = 环境健康、质量体系完整，可以开始接需求。任何一条红 = 先修环境/回归，再谈新功能。
- 手册全面更新批次已收口（2026-09-15，纯文档批次合规跳过 bump，版本仍 3.9.1.1）：用户手册新增功能全景章（178 工具×14 域，从注册代码直出）+32 宏总目录+19 内置技能清单+4.2 功能区一键操作（聚光灯/19 一键命令/宏收纳箱/函数库/SQL查询此前零记载），修 3.9/联网搜索结构错位与设置中心过时口径（上限 10000）；技术手册新增 5.2 工具注册全景（178 按注册分块）；双手册 HTML 重生成+校验全过；提取脚本沉淀在 _dev_tmp/extract_tools_desc.py（描述含 \" 须 ((?:[^"]|\\.)*) 正则、name 与 RiskLevel 间可有注释行）；178=Excel 宿主注册全集，WPS full 基线 152 不变
- v3.8.1.84 已发版收口（2026-09-16，WPS 副本）：「打开面板」窗口归属自愈批次——用户报点打开面板无反应，真机日志三次 ShowTaskPane existing made visible 且 WebView2 正常实锤「代码成功、物理不可见」：CTP 挂在创建时刻的活动窗口（CreateCTP 第三参 Type.Missing），切换工作簿窗口后 Visible=true 不报错但面板物理在旧窗口。修复：ShowTaskPane 加窗口归属 Hwnd 比对（IsPaneOnActiveWindow），不一致删旧 pane 重建并显式挂 ActiveWindow；宽度自愈 Width≤10 重置右侧停靠+380；句柄拿不到保守回退不误删（泛化 catch 规避 RuntimeBinderException CS0246）。gate 全绿（单测 1200/Excel 119/WPS full 131 组/双泄漏）、双手册（技术手册第三章自愈小节+用户手册 FAQ#11）、四形态 pack（残留 83=0）、静默安装 x64/x86 落位 3.9.1.1 一次通过——上批受保护僵尸 EXCEL 已随用户手动装 83 解除
- TaskPane 窗口归属自愈缺陷模式沉淀（v3.8.1.84 实证修复）：「工具日志说 visible、用户看不见」类 UI 问题先查 CTP 窗口归属——CreateCTP 第三参 Type.Missing 挂创建时刻的 ActiveWindow，切换工作簿窗口后 Visible=true 成功但物理在旧窗口；修法是 Hwnd 比对+不一致删除重建+显式挂 ActiveWindow+宽度自愈（Width≤10 视同不可见）；排查口诀：日志见 ShowTaskPane existing made visible 而用户看不见即此形态；Addin 工程无 Microsoft.CSharp 引用，dynamic 派生异常（RuntimeBinderException）不可直接 catch 类型名（CS0246），用泛化 catch+GetType().Name 记日志