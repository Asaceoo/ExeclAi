# 设计-WPS 兼容版需求文档

> 版本：v1.0（grilling 会话产出，2026-09-08）
> 状态：待 M0 POC 验证后进入开发
> 关联文档：`docs/功能扩展调研报告.md`（其中"WPS 兼容=不做"的结论由本文档修正）、`docs/JYYJ助手技术手册-3.9.1.3.md`、`docs/踩坑记录与规避手册.md`

---

## 1. 背景与动机

本项目（JYYJ助手）是内嵌 Excel 的 AI 办公助手：COM 加载项（`IDTExtensibility2`）+ Ribbon（`IRibbonExtensibility`）+ CustomTaskPane 侧边栏（WebView2），130 工具经 `dynamic` 晚绑定操作 Excel COM 对象模型。

此前调研将 WPS 兼容定性为"架构硬约束、不做"。经技术核查，该结论**需要修正**：

1. WPS 表格原生支持 COM 加载项机制（`IDTExtensibility2`），注册键为 `HKCU\Software\Kingsoft\Office\ET\Addins`，个人版经 `AddinsWL` 白名单加载；
2. 本项目 Excel 执行层全部为 `dynamic` 晚绑定，无 PIA 强依赖，天然可对接 WPS 的 ET 对象模型（与 Excel COM 高度兼容）；
3. WPS 支持 customUI Ribbon XML，回调机制一致。

因此"全部功能跑在 WPS 上"不是重写，而是**双宿主适配**：真正的风险集中在侧边栏宿主（CustomTaskPane）、注册路径、以及少数依赖 Excel 独有能力的工具。

## 2. 目标与非目标

### 2.1 目标

- G1：同一解决方案、同一套工具层/服务层代码，同时宿主于 Excel 与 WPS 表格（双宿主）。
- G2：核心功能（读写、清洗、格式化、筛选、透视、图表等高频写工具）在 WPS 下 100% 真机验证通过。
- G3：硬缺口工具（VBA、Power Query 类等）在 WPS 下**显式标注、不静默失败**（三档标注，见 §5）。
- G4：WPS 真机 harness 纳入自动打包门禁，与现有 Excel 门禁同等强度。
- G5：双安装包发版（Excel 包 / WPS 包），版本号 base 同步递增、单一来源不变。

### 2.2 非目标（严格最小范围）

- N1：不顺带重构 Excel 侧。宿主抽象（HostKind）只在 WPS 路径生效，Excel 现有行为零回归、零重构。
- N2：不做 WPS 文字 / WPS 演示适配（本项目是纯表格工具）。
- N3：不做 WPS JS 加载项（wpsjs）重写。
- N4：VBA 工具不承诺在未装 VBA 环境的 WPS 上可用，只承诺"明确告知 + 降级路径"。
- N5：**原生数据看板（超级看板）与高级图表能力仅在 Excel 宿主提供，WPS 侧不做**（2026-09-10 决策，见 D11）。

## 3. 决策记录（grilling 落定）

| # | 决策点 | 结论 | 理由 |
|---|--------|------|------|
| D1 | 功能覆盖 | 核心优先 + 三档标注 | 130 工具不可能字面 100% 等价；硬缺口显式化优于静默失败 |
| D2 | WPS 版本范围 | 个人版 + 专业版/教育版双版本都要 | 用户基数与稳定性兼顾，验收矩阵双环境 |
| D3 | 侧边栏形态 | POC 定夺 + 浮动窗兜底 | CustomTaskPane 在 WPS 支持无官方保证，不赌 |
| D4 | 代码形态 | 同解决方案双宿主（HostKind 抽象） | 一套工具层代码，避免双维护漂移 |
| D5 | 验收门禁 | WPS harness 全过，并入打包门禁 | 无门禁的兼容承诺是空话 |
| D6 | 发版形态 | 双包双后缀（`-wps` 后缀） | 用户不会装错；隔离发版风险 |
| D7 | 版本号规则 | base 同步递增，单一来源（Addin.csproj） | 同源码同版本，功能差异由三档清单注明 |
| D8 | POC 时机 | POC = M0 硬门 | 最大不确定项前置暴露，未过即止损 |
| D9 | 降级交互 | 宿主感知过滤 | WPS 下工具清单自动过滤不可用项，用户问到时明确告知原因 |
| D10 | 范围边界 | 严格最小范围 | 只做兼容必需改动，回归面最小 |
| D11 | 看板/高级图表宿主范围 | **仅 Excel 宿主实现，WPS 不做**（2026-09-10） | 原生数据看板（DashboardSpec 渲染：KPI 卡片/SUMIFS 联动/双轴组合图/环形图/切片器）与高级图表工具族均为 Excel 专属；WPS 侧不实现、不编排、不验证，节约双宿主门禁成本（`verify_dashboard` 等新自检仅挂 Excel 真机 harness）。WPS 下该能力按 §5.1 ❌ 档处理：宿主感知过滤，不出现在 WPS 工具清单，用户问及时明确告知原因与替代建议（用现有基础图表 + 透视工具拼装） |

## 4. 用户与版本矩阵

| 宿主 | 版本 | 加载机制 | 验收环境 |
|------|------|----------|----------|
| WPS 个人版 | ≥12.x（本机实测 12.1.0.26899，32 位） | `ET\Addins` 注册键 + `AddinsWL` 白名单 | 本机真机（POC + harness） |
| WPS 专业版/教育版 | v11.8.2+ | `ET\Addins` 注册键 | 按需协调真实环境冒烟 |
| Excel | 2013/2016/2019/2021/365（现状不变） | 现有 Microsoft Office Addins 键 | 现有 harness，零回归验证 |

POC 环境事实（2026-09-08 探测）：
- 本机已装 WPS Office 12.1.0.26899，`C:\Program Files (x86)\Kingsoft\WPS Office\12.1.0.26899\office6\et.exe`（位于 x86 目录，**疑似 32 位进程，M0 需确认 PE 位数**）。
- COM 自动化 ProgID：`KET.Application`（WPS 表格）；Excel 仍为 `Excel.Application`。

## 5. 功能覆盖与三档标注

### 5.1 三档定义

| 档位 | 定义 | 用户侧行为 |
|------|------|-----------|
| ✅ 支持 | harness 断言覆盖且通过 | 正常出现在工具清单 |
| 🔶 降级 | 可用但功能子集/行为差异（如图表部分类型） | 正常出现，描述中注明差异 |
| ❌ 不可用 | 宿主无对应能力（如未装 VBA 环境的 VBA 工具、Power Query 类） | 宿主感知过滤：不出现在 WPS 工具清单；用户问及时明确告知原因与替代建议 |

### 5.2 已知差异点（审计输入清单，M3 逐项落档）

| 能力域 | 差异预期 | 初判档位 |
|--------|----------|----------|
| 读写/选区/工作表管理 | ET 对象模型高度兼容 | ✅ |
| 数据清洗/筛选/格式化 | 预计兼容，个别 NumberFormat 行为差异待测 | ✅/🔶 |
| 图表 | 部分图表类型/属性差异 | 🔶 待测 |
| **看板/高级图表（D11，2026-09-10）** | **组合图（双轴）、环形图、原生看板渲染（DashboardSpec）、切片器联动：仅 Excel 提供** | ❌（WPS 不做，宿主感知过滤） |
| 透视汇总 | 属性级差异待测 | 🔶 待测 |
| 条件格式 | API 差异待测 | 🔶 待测 |
| VBA 执行 | WPS 需另装 VBA 环境；个人版默认无宏 | ❌（未装环境时） |
| Power Query 类 | WPS 无对应能力 | ❌ |
| 动态数组函数（XLOOKUP/FILTER 等） | WPS 函数库差异 | 🔶 按函数标注 |
| 跨工作簿只读引用 / SQLite 历史 / 回滚 | 纯本地逻辑，不依赖宿主差异 | ✅ |
| 模型配置 / AI 工具编排 | 与宿主无关 | ✅ |

> 覆盖审计纪律沿用：WPS 差异清单复制自 `AutoVerifyPolicy.MappedPrimaryTools`(16 类)，**新工具不表态就发不出去**；`verify_*` 自检族在 WPS 下同样 fail-closed。

## 6. 架构方案

### 6.1 双宿主拓扑

```
ExcelAiAssistant.Addin (宿主层)
 ├─ ExcelAiAddIn : IDTExtensibility2, ICustomTaskPaneConsumer, IRibbonExtensibility
 │   └─ HostKind 检测：OnConnection 时按宿主 ProgID/进程判定 Excel | WPS
 ├─ WPS 路径：注册键/白名单由安装器写入；侧边栏按 M0 结论走任务窗格或浮动窗
 └─ Excel 路径：现有逻辑零改动
ExcelAiAssistant.Excel (执行层，不动架构)
 └─ 仅在差异点注入 HostKind 分流（dynamic 晚绑定不变）
```

### 6.2 注册层（安装器改动）

| 注册项 | 键 | 值 |
|--------|-----|-----|
| WPS 加载项 | `HKCU\Software\Kingsoft\Office\ET\Addins\<ProgID>` | FriendlyName / Description / LoadBehavior=3 |
| WPS 白名单（个人版） | `HKCU\SOFTWARE\kingsoft\Office\ET\AddinsWL` | ProgID 字符串值 |
| COM 类 | RegAsm 双位数注册（现状不变） | 按宿主 PE 位数匹配（复用 `RegistryView` 经验） |

### 6.3 侧边栏两形态（M0 分支决定）

- **形态 A（原生任务窗格）**：WPS 支持 `ICustomTaskPaneConsumer` 时沿用现有 CustomTaskPane，体验与 Excel 一致。
- **形态 B（浮动窗兜底）**：不可用时退化为独立 WinForms/WebView2 浮动窗（可贴边吸附）。复用 `SpotlightService` 原生 `Hwnd` 经验；Excel 侧**不启用**该路径（N1 约束）。

### 6.4 Ribbon 适配

- customUI XML 与回调机制 WPS 一致，主路径免改。
- 已知坑：WPS 对透明背景 Bitmap 填浅灰底 → 统一走 IPictureDisp（现有 `getImage` 记忆化缓存已满足）；WPS 无深色主题图标自动反色 → 接受（记录为已知差异）。

## 7. 里程碑

| 里程碑 | 内容 | 退出条件 | 性质 |
|--------|------|----------|------|
| **M0 POC（硬门）** | 手工注册现有 v2.19.5 产物到 WPS 两键 + AddinsWL（不改源码，可回滚）→ 启动 et.exe 验证 OnConnection / Ribbon / CustomTaskPane / WebView2 | 产出《加载可行性报告》：加载成功与否 + 侧边栏形态分支（A/B）确定；**不通过则项目止损或转独立进程方案（KET.Application 外部连接）重新立项** | 前置硬门 |
| M1 注册层 | 安装器/注册代码双键注册、白名单、位数适配 | 本机个人版注册后自动加载 | 可与 M2 并行 |
| M2 宿主 UI | HostKind 检测、Ribbon 微调、浮动窗兜底（若 M0 判 B） | WPS 下 Ribbon 全按钮可点、侧边栏可用 | |
| M3 执行层审计 | 130 工具三档分类落档、差异点分流、宿主感知过滤上线 | 差异清单评审通过；WPS 下清单过滤正确 | |
| M4 验证门禁 | 新建 WPS harness（`KET.Application` 连接，复用 RealMachineVerify 框架 80 断言）；verify_* WPS 适配；并入 `package-xai.ps1` | harness 断言 100% 过；打包门禁含 WPS 关 | |
| M5 发版 | 双包产出、README/交接文档更新、SOP 落地 | 见 §8 验收标准 | |

## 8. 验收标准

1. **M0 报告通过**：加载成功 + 侧边栏形态确定（本报告是 M1-M5 的输入）。
2. **WPS harness 全过**：核心类工具断言 100% 通过；硬缺口工具验证"标注与实际行为一致"（❌ 档确实被过滤、🔶 档差异与文档一致）。
3. **自检族不降级**：WPS 下可用的 `verify_*` 正常工作，判定无期望值时 fail-closed。
4. **双版本矩阵**：个人版（本机 12.1.0.26899）+ 专业版（真实环境协调）各过一遍冒烟清单。
5. **Excel 零回归**：现有 872 单测 + 真机 80 断言全过，Excel 包行为与当前版本无差异。
6. **门禁闭环**：`package-xai.ps1` 双关变三关（单测 + Excel 真机 + WPS harness），逃逸开关同样覆盖。

## 9. 发版与版本号

- **双包双后缀**：`JYYJ助手-<base>-Setup.exe`（Excel）与 `JYYJ助手-<base>-wps-Setup.exe`（WPS/双宿主）。
- **base 同步递增**：同一源码一个版本号，单一来源仍为 `src/ExcelAiAssistant.Addin.csproj`（Version/InformationalVersion）；`-wps` 后缀由打包脚本按渠道附加，**不是独立版本线**。
- 发版 SOP 三处同步（csproj / README / iss）不变；README 与交接文档增补 WPS 包形态、三档功能清单（含 ❌/🔶 明细）。
- 产物禁止无版本号（用户强制规范，两包均适用）。

## 10. 风险与缓解

| 风险 | 等级 | 缓解 |
|------|------|------|
| CustomTaskPane 在 WPS 不可用 | 中高 | M0 前置暴露；形态 B 浮动窗兜底已定 |
| 个人版白名单加载在个别 WPS 版本受限 | 中 | M0 本机实测；专业版为兜底验收环境；失败即止损（D8） |
| ET 对象模型 API 差异导致工具行为偏差 | 中 | 三档标注 + harness 断言 + verify_* fail-closed |
| WebView2 在 WPS 进程内宿主行为异常 | 低中 | M0 验证；浮动窗为独立进程级兜底 |
| 双包版本漂移 | 低 | base 同步 + 打包脚本统一加后缀 |
| 32/64 位错配（WPS 位数未知时） | 低 | 安装器按 et.exe PE 位数探测注册（复用现有 RegistryView 方案） |

## 11. 开放问题（M0 回答）

1. et.exe 实际 PE 位数（32/64）？
2. CustomTaskPane（`ICustomTaskPaneConsumer`）在 WPS 12.1 是否触发？
3. WebView2 侧边栏在 WPS 进程内能否正常初始化？
4. Ribbon 图标在 WPS 实际渲染效果（透明底/灰底核查）？
5. `KET.Application` GetActiveObject/CreateObject 是否可用（harness 连接方式依据）？

---

*本文档为 grilling 会话共识产物：D1-D10 十项决策已与需求方逐条确认。M0 未通过前，M1-M5 不启动。*
