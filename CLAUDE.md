﻿﻿﻿# JYYJ助手 · AI 行为准则

> 本文档告诉**接手本仓库的 AI 助手**"在这里该怎么做事"：红线、命令、文档去哪读、验证门禁、发版约定。完整项目现状/架构/功能清单见 `docs/AI交接文档.md`（它是你 10 分钟建认知的入口）；技术细节见 `docs/JYYJ助手技术手册-3.8.5.0.md`；经验教训见 `docs/踩坑记录与规避手册.md`。

<!-- neat-freak: initialized at 2026-09-06 -->

## 1. 项目定性

**JYYJ助手**：内嵌 Excel 的 AI 办公助手——COM Add-in + WebView2 侧边栏，用户用自然语言操作表格，AI 经 167 个工具执行（口径 = `tools/release/check_wps_full.py` 的 EXPECTED_TOOL_COUNT，2026-11-03 校准：156+11 L2元素层/草稿模式/外部数据源）写入/清洗/分析/图表/超链接/VBA/流程编排；**所有写操作被执行层自动自检兜底、可回滚、运行日志全量留痕**。

- 当前版本：**v3.8.3.1**（源码即最新；2026-10-02 M20 手册按源码校准 156 工具/20 verify_*/1291 单测并上传 GitHub）。2026-09-19 工具分类审查十一批收官：侧边栏七大板块全部扫完（末批：聊天附件白名单制+2MB/10MB 护栏），累计修复 20 真缺陷，单测 1291。批次详录见 docs/版本史.md。**版本号书写约定：本文件与 docs/AI交接文档.md 头部中历史版本号一律带 v 前缀**（bump 只替换裸版本号，v 前缀免疫）。
## 2. 你在这里做什么

- 维护 Excel 加载项功能：新增/修复工具、技能、前端面板、流程模板、安装器。
- 改代码前先看对应能力的**现成真机 harness**（`tools\Verify*`、`tools\RealMachineVerify`），小步改、每步验证。
- 本项目禁止"盲改"。任何写能力变更必须过 §6 门禁。任何新踩的坑必须沉淀（§8）。

## 3. 工作流（用户全局指令）

- 交互顺序：`思考 → 工具调用 → 结果 → 思考 → … → 结论`，不堆叠思考/工具。
- 收到的踩坑`必须写进 docs/踩坑记录与规避手册.md` 对应分类，而非只留会话记忆。
- 每一轮结束后收敛上下文：不留非必要的历史返工噪声，只保留正向流程。

## 4. 红线（违反即返工）

- **改文档/记忆层用神经网络语言，改代码不改文档**：任何行为准则之外，不得删除用户数据文件。
- **Install/COM 注册**：32 位与 64 位一起装，用 `RegistryView` 匹配 Excel PE 架构。
- **版本号单一来源（M17.1 红线）**：全仓库唯一版本定义点 = 根 `Directory.Build.props`；升版**只允许** `python tools/release/release.py bump --ver x.y.z.w`（同步 props+README），打包**只允许** `python tools/release/release.py pack`（自动构建/组装/zip/版本戳断言/ISCC `/DAppVersion` 注入模板 iss）。**禁止手工 sed csproj/README/iss**——csproj 已不定义版本（继承 props），iss 模板 `JYYJ助手-Setup-wps.iss` 不含硬编码版本。
- **PowerShell 5.1 脚本必须 UTF-8 带 BOM** 存储，否则 GBK 解析报错。
- **确认门禁**：破坏性操作（整行删除等）必须显式 user 确认；面板"允许执行"按钮即确认。
- **编码器红线**：不得引入超出程序运行环境的目标平台依赖；所有写操作走自检。新增工具必入覆盖审计（见 §6）。

## 5. 命令速查

| 操作 | 命令 |
|------|------|
| 构建（全解） | `dotnet build ExcelAiAssistant.sln -c Release`（警告即错误） |
| 单测 | `dotnet test ExcelAiAssistant.sln -c Release`（当前 1291 用例全过） |
| 真机验证 | `tools\RealMachineVerify`（独立 Excel 实例，22 用例组 119 断言） |
| 打包（Excel 主线遗留） | `tools\package-xai.ps1`（仅 Excel 主副本 D:\execl1.0 发版流；**本仓库发版只走 release.py**，见下行） |
| 安装器 | `tools\installer\JYYJ助手-Setup-wps.iss`（模板化，版本号由 ISCC `/DAppVersion=` 注入，勿手改版本） |
| WPS full 基线校验 | `python tools\release\check_wps_full.py --log <full日志>`（结果分布 vs 基线，新 FAIL 即亮红；有意变更走 `--update` 后人工逐行评审基线 tsv） |
| 升版/打包 | `python tools\release\release.py bump --ver x.y.z.w` → `python tools\release\release.py pack`（详见 §10） |
| **一键门禁** | `python tools\release\gate.py`（构建→单测→Excel 真机→WPS full→基线校验，真机步前快照宿主进程、步后断言无新增泄漏，`--pack` 直通发版；fail-fast 任一环失败即中止） |
| 版本控制 | 本目录**是 git 仓库**（2026-09-19 c2f6103 首次入库）：origin = gitclone 镜像 `github.com/Asaceoo/ExeclAi`（直连被镜像拦截），**推送走 SSH** `git@github.com:Asaceoo/ExeclAi.git`（HTTPS 无写凭据）；本地分支 `master`、远端分支 `main`，推远端用 `git push git@github.com:Asaceoo/ExeclAi.git HEAD:main`；每批收口 git commit（feat/fix/chore+批号）。与 Excel 主副本 `D:\execl1.0` 的同源修复靠人工双副本同步 |

## 6. 验证与门禁（写工具必过）

- **覆盖审计**：所有写能力工具（NormalWrite/StructuralChange/Dangerous）必须在 `MappedPrimaryTools`(44 类)或 `DocumentedWriteToolExemptions`(95 条带理由豁免)里；`ct_`前缀用户自定义工具自动豁免。判定入口 `AutoVerifyPolicy.IsExemptedFromAutoVerify`，两个执行点：harness 用例17、全功能巡检。**新工具不表态就发不出去**。
- **保护态守卫（M17）**：所有写能力工具默认受 `SheetProtectionGuard` 前置探测（目标表 ProtectContents / 结构类 ProtectStructure），豁免必须带理由（`SheetProtectionGuard.Exemptions`，ct_ 前缀自动豁免）；守卫是加法前置——探测异常一律放行，不得在守卫内拦截或改写工具自身错误路径。**探测类 COM 代码零 FinalRelease**（RCW 身份共享，见兼容差异清单 M17 条目）。
- **verify_* 自检族（18 个）**：write_result / cell_format / conditional_format / sort_filter / formula / chart / data_clean / pivot / marker / structure / data_validation / sparkline / merge / hyperlink / table / protection / image / freeze_panes（以源码注册为准）；判定无期望值时 fail-closed，不得静默放行。
- **流程编排**：内置模板 verify 工具必须带 `{selection_sheet}`；自检结果与主步骤独立命名空间，失败标红且 ok 必须为主结果+自检均通过。

## 7. 编码约定（核心）

- **无状态 COM**：每次操作前重新解析工作表，跨表断言显式 `Workbooks[...].Worksheets["name"]`，不依赖 `ActiveSheet`，用 `Cells.Item[1,1]` 作写回锚点。释放用 `Release((object)x)` 静态绑定。
- **写入格式化**：写 ISO 日期文本先 `NumberFormat="@"`；Range 无状态非法（RCW 分离异常常见）。
- **集合 1 基**：Excel COM 集合从 1 开始；新增表用 `Add(Missing.Value, count)` 追加并断言新表用 `count+1`。
- **Python 交互**：`sheets.active`（不要 `sheets[0]`）；`sel_range` 注入格式（`'表名'!$A$1:$B$6` 单引号+绝对引用+`!`）不得改。
- **WebView2**：`ExecuteScriptAsync` 返回对象→JSON 对象；返回字符串会被再套引号（不要 stringify 后再解析）。
- **Ribbon 图标**：`getImage` 记忆化缓存同一 IPictureDisp + `fOwn=false`，勿每回调新建。

## 8. 文档指针表

| 想知道 | 去读 |
|--------|------|
| 项目整体/接手 10 分钟 | `docs/AI交接文档.md` |
| 用户怎么用 | `docs/JYYJ助手用户手册-3.8.5.0.md`、同名 `.html`（`docs/readme.md` 仅是迁移指针） |
| 技术细节（注册/COM/侧边栏/打包） | `docs/JYYJ助手技术手册-3.8.5.0.md` |
| 踩过的坑（血泪） | `docs/踩坑记录与规避手册.md`（docs 层权威，勿再另建 project_memory.md——已废弃） |
| Ribbon 命令/技能编写 | Excel 主副本 `D:\execl1.0\.trae\documents\ribbon-command-authoring-guide.md`（本 WPS 副本内无此文件） |
| 新增功能设计 | `docs/设计-*.md`、`docs/功能扩展调研报告.md` |
| MCP 接入（已落地 P1） | `docs/需求-MCP接入.md`（M21 落地 M-A server，tools/McpServer） |
| 优化路线图 | `docs/需求-优化路线图.md`（M24 竞品对标产出，P0-P3 分期待拍板） |

## 9. 记忆与踩坑

- 每修一个 bug / 过一道坑，**追加**到 `docs/踩坑记录与规避手册.md` 对应分类（文档层权威，防跨会话上下文丢失；各平台 agent 记忆仅留精简指针，稳定知识"毕业"升入 docs）。
- docs 层为权威；agent 记忆层需"毕业"——稳定知识升入 docs，记忆保留精简指针，防膨胀（单文件 ≤25KB/200 行）。
- 规模体检：本项目无 CLAUDE/MEMORY 之外建议引入更多长文档，能用指针就用指针，控制根级文档厚度。

## 10. 发版 SOP（M17.1 自动化版）

1. `python tools/release/release.py bump --ver x.y.z.w`（唯一改版本入口：Directory.Build.props 单一来源 + README 自动同步）
2. `python tools/release/gate.py --pack` 一条命令完成：单测 → 真机 harness → WPS full 基线校验 → 戳断言 → ISCC（也可分步跑，见命令速查）
3. `python tools/release/release.py pack`（自动：双架构 --no-incremental 重建 → 组装 dist 便携目录（说明文件从上一版继承+版本替换）→ 进程内 zip → 版本戳断言（新戳在位/旧戳零残留/守卫符号）→ ISCC 编译安装器）
4. 产出 dist 四形态 → 更新 `docs/AI交接文档.md` 的当前版本与功能清单，归档旧版本产物