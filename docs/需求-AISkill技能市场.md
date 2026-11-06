# 需求：AI Skill 技能市场（安装外部技能 + 聊天调用）

> 状态：**grill-me 已拍板冻结（2026-09-15）**——八问全部按 AI 倾向答案冻结（超时默认采纳，非逐条显式确认，有异议可回滚重议）；本稿为范围冻结稿，实施按 P1 一批收口。
> 依据：用户 2026-09-15 指令"能不能做个 AiSkill 功能，能安装 https://www.skills.sh/ 等技能网站的技能，能在聊天中调用 skill 的功能"。
> 事实底座：skills.sh 为 Vercel 开放 Agent Skills 目录（技能本体在 GitHub 仓库，SKILL.md + references 格式，`npx skills add owner/repo` 安装，150 万+ 安装量，发布方含 anthropics/microsoft/feishu 官方）；用户本机即该生态真实用户（grill-me/neat-freak 等由此安装）。

## 0. 既有拍板修订声明

修订 2026-09-13 WorkBuddy 借鉴拍板中「SkillHub UGC 市场：明确不学」一项——改为**开放安装、执行层守门**（Q1A）：技能是 prompt 层知识文本，安装放开；技能指导触发的**写操作仍走产品既有确认门/回滚/auto-verify，门禁哲学一寸不让**。其余借鉴拍板（技能发现缺口待补等）不变且由本需求落地。

## 1. 目标 / 非目标

**目标**：在聊天侧边栏中安装 skills.sh 等技能站点的外部技能（SKILL.md 格式）、管理已装技能、聊天中调用并让 AI 主动感知技能存在。

**非目标**（本期不做）：技能内容审核背书（不做内容警察，风险控制在执行层）、P2 前的内容预览审计 UI、自动更新检查、skills.sh 排行/热度 API 接入、技能市场内付费/账号体系。

## 2. 拍板结论（八问冻结）

| # | 议题 | 拍板 |
|---|---|---|
| Q1 | 开放 vs 门禁 | **A：安装与执行分离**——装任意技能（来源标注+用户责任），写操作走既有确认门/回滚，门禁不放松 |
| Q2 | 兼容性 | **A：安装时解析 front matter 与本机注册表 diff，输出「可用/部分可用/知识参考」三档标注**，诚实降级不装糊涂 |
| Q3 | 安装通道 | **A：产品内置 GitHub 直拉，零外部依赖**——三入口：`owner/repo`、skills.sh 页面链接自动解析到 repo、本地 zip；复用 SkillPackImporter 的 zip-slip 校验管线；不依赖 node/npx |
| Q4 | 安全 | **A：三层**——①安装走确认门（展示名称/来源/描述/声明 tools 后点装）；②SKILL.md 文本默认可读加载；③references 内脚本默认不执行，AI 引用执行须二次确认；外部技能统一 external 信任级 |
| Q5 | 调用与发现 | **A：清单注入上下文**——已装技能（名称+一句话描述+来源）注入会话上下文（上限约 30 行），AI 可主动建议使用、用户确认后加载全文；/skill 路由不变，返回值标 source=builtin/local/external |
| Q6 | 架构落点 | **A：不另起炉灶**——SkillCatalog 双目录机制扩展第三类 external，落 `LocalAppData/XaiAssistant/skills`；「AI 技能包」卡升级为「技能管理」：已装列表三区（内置/导入/市场）+ 安装输入框 + 删除（内置不可删） |
| Q7 | 格式兼容 | **A：解析器宽容**——标准字段（name/description）必读；本项目自有字段（version/type/category/risk_policy/tools/triggers）缺失填默认值：type=external、risk_policy=强制 high、tools=[] 即纯知识技能；不要求外部技能适配本项目 |
| Q8 | 期次 | **A：P1 一批收口**（安装器+管理 UI+/skill 调用+清单注入）；P2（预览审计/更新检查/排行接入）另批 |

## 3. 功能需求

- **FR-S1 技能安装器（P1）**：三入口拉取 GitHub（repo zip）→ 解析 → zip-slip 校验 → front matter 解析（Q7 宽容口径）→ 兼容性三档标注（Q2）→ 确认门展示后落 `skills` 用户目录（external 类）。skills.sh 链接解析规则：页面 URL 提取 owner/repo。
- **FR-S2 技能管理 UI（P1）**：设置页「AI 技能包」卡升级「技能管理」——三区列表（内置 builtin / 导入 local / 市场 external，含名称/描述/来源/兼容档位）+ 安装输入框（owner/repo 或链接）+ 删除（仅非内置）。
- **FR-S3 聊天调用（P1）**：/skill 路由对 external 技能同口径生效；返回值与 AI 转述标注来源；加载全文即注入该技能 SKILL.md 指导内容。
- **FR-S4 技能清单注入（P1）**：会话上下文注入已装技能清单（名称+一句话描述+来源，上限约 30 行），AI 按语义匹配主动建议，用户确认后才加载全文（成本与安全双控）。
- **FR-S5 安全纪律（P1）**：external 信任级强制 risk_policy=high；references 脚本执行二次确认；**产品契约优先于技能内规范**（如外部技能写"字体用 SimHei"，产品中文字体探测链 msyh.ttc fail-closed 契约不因技能文本放松）；写操作门禁/回滚/auto-verify 全链不变。
- **FR-S6 预置推荐（P1）**：市场入口预推荐 4 个高价值技能（按 2026-09-15 检索实证）：`davila7/claude-code-templates` 的 **excel-analysis**（pandas 分析全流程/透视/清洗）、`langchain-ai/deepagents` 的 **data-visualization**（matplotlib/seaborn 出版级图表，含箱线/热力/多面板——补原生图表族缺口）、`beita6969/scienceclaw` 的 **matplotlib-viz**（出图美学规范）、`mattpocock/skills` 的 **grilling**（需求评审方法论）。预推荐仅是入口捷径，安装仍走确认门。

## 4. 兼容性三档口径（Q2 细化）

- **可用**：纯知识/方法论技能（front matter tools 为空或 ⊆ 本机注册表）——聊天直接生效；
- **部分可用**：声明工具部分缺失（如依赖 Python 通道的能力在本机已装则可用）——标注缺失项；
- **知识参考**：核心动作面完全不存在（web 开发/浏览器自动化/云 CLI 类）——只作文本参考，UI 明示"本机无法执行"。

## 5. 验收标准（AC）

1. AC-S1 真机安装闭环：从 skills.sh 实名技能（如 excel-analysis）以 owner/repo 入口安装成功，技能管理页三区正确显示，zip-slip 防护对恶意路径样本拦截；
2. AC-S2 聊天调用：安装后 /skill 可路由加载该技能，AI 回复体现技能指导内容且标注来源 external；
3. AC-S3 主动感知：新会话中 AI 能引用已装技能清单并主动建议（清单注入 ≤ 上限行数，超限截断）；
4. AC-S4 兼容标注：安装一个第三档技能（如 frontend-design）后标注为"知识参考"，不虚报可用；
5. AC-S5 契约优先：技能内与产品契约冲突的规范（字体链/统计口径/确认门）实测不生效；
6. AC-S6 回归：gate 全绿（单测/Excel 真机/WPS full/双泄漏），登记面（工具数、Schema、SKILL.md、文档）四件套同步，标准发版链走完整。

## 6. 里程碑

- **P1（本批）**：FR-S1~S6 全量，AC-S1~S6 验收，标准收口链（gate+bump+pack+安装+文档同步）。
- **P2（另批）**：安装前 SKILL.md 内容预览审计页、技能更新检查、skills.sh 排行/热度接入。

## 7. 已冻结的开放问题

无——八问全部拍板冻结；实施期新增暗礁按 grill-me 惯例补问。
