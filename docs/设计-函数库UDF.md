# 设计-函数库 UDF

> 版本：v0.4（需求草案，待评审收敛）
> 进度：**M1 已落地（2026-09-19，v3.8.1.97）**——JyyjtoolAlgorithms 上移 Core（namespace 不变零改动）、TryParseDate/IsDateLikeFormat 提公 Core TextDateParser（ExcelUtilityToolsService 转发单一来源）、TextDateParserTests 6 例回归锁定；M2~M6 待排期。
> 归属：A 类·新增功能（XLL / UDF 注册链路）
> 关联：侧边栏 `view-functions` 页签、`SidebarBridge.ListRegisteredFunctions()`、`docs/功能扩展调研报告.md`
> 状态：需求文档，**尚未进入实现与发版**。若随版本发布，须按红线同步 Addin.csproj / README / installer 三处版本号，并新增独立 UDF 项目的版本。

---

## 1. 背景与现状

- 侧边栏 `f(x)`·函数库页（`view-functions`）已搭好 UI 框架：列表 + 点击复制用法 + 空态文案。
- 但其数据源 `SidebarBridge.ListRegisteredFunctions()` 目前**硬编码返回空数组**，页面永远只显示「暂无已注册的自定义函数」。
- 页面文案称「UDF 注册属于独立机制（COM 函数暴露），不在本版本开启」——此处表述**有技术误导**：纯 COM Add-in 程序集默认**不能被工作表公式直接调用**，带点号前缀的 `JYYJ.*` 函数必须走 **XLL** 加载项（如 ExcelDNA）。

> 本设计以「带点 `JYYJ.*` + XLL(ExcelDNA)」为基线，同时澄清上述认知偏差（§6）。

另：竞品调研（方方格子 `FF.*` UDF 库）已把本项目定位为「形态缺失」，并建议「沉淀 8-10 个纯函数型工具为 `JYYJ.*`」，本设计与之对齐并落入可执行拆解。

## 2. 目标与非目标

### 目标
1. `f(x)`·函数库页展示 **真实可用的预置 `JYYJ.*` 函数目录**（分类 + 签名 + 中文说明 + 示例），点击复制用法，供用户在 Excel 单元格直接书写。
2. 建立一条 **XLL(ExcelDNA) 注册链路**，让 `=JYYJ.xxx(...)` 在单元格中真实可算。
3. **元数据单一来源**：函数目录与 XLL 内实现同源，杜绝「页面/文档/代码」三处漂移。

### 非目标（一期明确不做）
- 用户/开发者在面板手写自建函数、动态注册。
- 数组/区域（CSE）公式函数。
- 公式内调模型（如 `=JYYJ.AI(...)`）——涉及「计算时同步调模型」，独立论证，见调研报告 §UDF。
- 函数成为 AI 面板工具（双入口）。一期仅单元格公式用。

## 3. 核心决策（已与用户拷问收敛）

| # | 决策项 | 结论 |
|---|--------|------|
| D1 | 核心定位 | 普惠目录优先（先做真实可用的内置目录，管理体系后续再扩） |
| D2 | 命名/实现路线 | 带点 `JYYJ.*`，走 XLL(ExcelDNA) |
| D3 | 首版函数来源 | 预置一批内置函数 |
| D4 | 使用场景 | 仅 Excel 单元格公式 |
| D5 | 函数性质 | 纯计算（输入→输出，不访问 COM / 不读工作表状态，确定性、无副作用） |
| D6 | 实现归属 | 独立 UDF 项目产出 `.xll` |
| D7 | 安装分发 | 并入现有合并安装器，x86/x64 按 Excel PE 架构匹配加载 |
| D8 | 行为复用 | 复用工具层可无 COM 的纯逻辑，保证与面板行为一致 |
| D9 | 函数形态 | 单格标量优先；数组/区域后续 |
| D10 | 命名与分类 | 单点 `JYYJ.*` 平铺命名，前端按分类字段分组展示 |
| D11 | 加载状态 | 函数库页动态检测 xll 是否加载，未加载给可操作提示 |
| D12 | 验证门禁 | xUnit 单测 + 真机 Excel 单元格断言，均达门禁方可交付 |

## 4. 技术范式：XLL(ExcelDNA) 与 COM 隔离

### 4.1 为什么必须 XLL
- VBA 函数名不能含点，`JYYJ.xxx` 非法。
- 纯 COM Add-in 程序集不被工作表公式认可。
- 带点函数名的 UDF 由 **XLL 通过 C API 注册** 是正统路径；选 ExcelDNA（.NET Framework 4.8 兼容）将 `public static` 方法暴露为工作表 UDF。

### 4.2 UDF 运行红线（硬约束，写进代码评审清单）
- UDF 运行在**工作表重算线程**，禁止触碰 COM/Interop、禁止访问 Addin 状态、禁止 IO/网络/易失。
- 保持**确定性、无副作用**，与受控执行/安全分级机制共存（对齐调研报告约束）。
- 纯逻辑复用，**不得把 Interop 程序集引用进 UDF 项目**。

### 4.3 复用来源（已核实，指向真实代码）
- `ExcelAiAssistant.Core`、`ExcelAiAssistant.Excel` 两个 .csproj **均未引用 Microsoft.Office.Interop.Excel**（COM 集中在 Addin），为 UDF 项目引用纯逻辑提供了干净隔离点。
- **主复用类**：`ExcelAiAssistant.Excel.JyyjtoolAlgorithms`（`public static`，纯逻辑，注释明确「可脱离 Excel 单测」）：
  - `ExtractCharacters(string text, string kind)` → 提取数字/中文
  - `DeleteMatches(string text, string pattern)` → 正则删除（去非打印字符等）
  - `BankersRound(decimal value, int decimals)` → 银行家舍入
  - `NumberToUppercaseAmount(decimal value)` → 金额大写
  - `SplitSmart(...)` / `DetectDelimiter(...)` → 智能拆分
  - `CrossReferenceValues(...)` → 引用匹配
- **可抽离的纯内部方法**：`ExcelAiAssistant.Excel.ExcelUtilityToolsService` 内部静态 `TryParseDate` / `IsDateLikeFormat` / `ToDouble` / `CleanText` 等。
- **Core 纯逻辑**：`ExcelAiAssistant.Core.Tools.FormulaAnalyzer`（公式解析）、`Excel.RangeHelper`/`ExcelAddressParser`（地址解析）等。

**落地策略（防双份维护）**：UDF 项目——项目引用 `Core`；并把位于 `Excel` 程序集的纯静态逻辑（`JyyjtoolAlgorithms` 与 `TryParseDate` 等）**上移至 Core（或共享源文件）**，使「面板工具」与「UDF」共用同一份实现。若某纯逻辑当前为 `private/internal`，予以提公为 `public` 并补回归单测。

## 5. 架构总览

```
ExcelAiAssistant.Udf（新增独立项目，.NET Framework 4.8，ExcelDNA）
  ├─ 引用 ExcelAiAssistant.Core（纯逻辑，无 Interop）
  └─ 每函数以特性标记元数据：
       [JyyjFunction(Category="文本", Name="ExtractDigits", Signature=..., Example=..., Description=...)]
       public static string ExtractDigits(string text) => JyyjtoolAlgorithms.ExtractCharacters(text, "number");
       元数据采集 → 导出函数目录（JSON/程序集反射）

ExcelAiAssistant.Addin（改）
  ├─ ListRegisteredFunctions() → 从 UDF 元数据源读真实目录，不再硬编码空数组
  └─ 首次加载：检测 xll 是否已加载；未加载按 Excel PE 架构尝试 RegisterXLL，失败返回值给前端
Sidebar\index.html（改）
  └─ loadFunctions()：按分类分组渲染真目录；点击复制用法；动态加载状态 + 可操作提示
tools\installer\*.iss（改）
  └─ 并入 x86/x64 两枚 .xll，按架构匹配放置
```

## 6. 认知澄清（写入交接文档的必要性）

- 现有文档/页面普遍写「UDF 注册属于现成机制（COM 函数暴露）」——**不准确**。
- 带点 `JYYJ.*` 的单元格 UDF 无法由纯 COM Add-in 提供，必须 XLL。
- 本设计落地后，建议在 `docs/AI交接文档.md` 与函数库页文案中**更正这一表述**，改为「UDF 由独立 XLL 加载项提供」。

## 7. 第一批函数清单（初稿 12 个，待评审增减）

> 一期全部为单格标量。映射的"底层实现"均指向可无 COM 复用的纯逻辑；未列出现成的标 ★ 为新增纯静态方法（需补单测）。

| 分类 | 函数 | 语义 | 底层实现（复用/新增） |
|------|------|------|-----------------------|
| 文本 | `JYYJ.ExtractDigits` | 提取一串文本中的数字 | `JyyjtoolAlgorithms.ExtractCharacters(text,"number")` |
| 文本 | `JYYJ.ExtractChinese` | 只保留中文 | `ExtractCharacters(text,"chinese")` |
| 文本 | `JYYJ.JoinText` | 多参数文本按分隔符合并 | 新增 ★（变参标量） |
| 文本 | `JYYJ.Clean` | 去首尾空/非打印字符/统一内部空白 | `DeleteMatches` + trim（组合/新增） |
| 日期 | `JYYJ.ParseDate` | 容错解析手写日期 → 日期 | `ExcelUtilityToolsService.TryParseDate` 提公 |
| 日期 | `JYYJ.ToIsoDate` | 转 ISO 文本（配文本格式） | 同上组合 |
| 日期 | `JYYJ.CalcWorkAge` | 工龄/司龄（年.月） | 新增 ★ |
| 数值 | `JYYJ.ToNumber` | 文本/千分位/货币 → 数字 | `TryParse…/ToDouble` 提公 |
| 数值 | `JYYJ.Clamp` | 限制上下限 | 新增 ★ |
| 判定 | `JYYJ.IsEmail` | 正则校验邮箱 | 新增 ★ |
| 判定 | `JYYJ.IsPhone` | 正则校验手机号 | 新增 ★ |
| 判定 | `JYYJ.IsDateText` | 判断文本是否像日期 | `IsDateLikeFormat` 提公 |

## 8. 前端函数库页改造（`Sidebar/index.html` + `SidebarBridge.cs`）

1. `ListRegisteredFunctions()` → 从元数据源返回 `{ ok, loaded, functions:[{category,name,signature,example,description}], reason }`。
2. `loadFunctions()` 按 `category` 分组渲染；每项含名称 + 签名 + 中文说明 + 示例。
3. 点击复制：复制 `=JYYJ.xxx(示例参数)` 用法到剪贴板。
4. **动态加载状态**：新增桥接 `GetUdfStatus()`（或在返回值带 `loaded`）。未加载时：
   - 列表仍展示（标灰），顶部提示「未加载 UDF 加载项」；
   - 提供可操作按钮「重新加载」→ 触发 Addin 侧按架构 `RegisterXLL` 重试；
   - 加载失败返回 reason，指引到设置页/安装器。
5. 不再展示占位「暂无已注册的自定义函数」。

## 9. 安装分发

- `tools/installer/*.iss`（Inno Setup，合并 x64/x86）新增：把 `ExcelAiAssistant.Udf-x64.xll` / `-x86.xll` 复制到对应安装目录。
- Addin `OnConnection` 时检测 xll 是否已加载，未加载按**当前 Excel PE 架构（匹配位数）** `RegisterXLL`；失败不阻断主程序，仅返回 reason。
- 保留用户数据（红线：不删 `custom-tools` / `agent-settings.json`），xll 属应用目录可覆盖。

## 10. 验证与门禁（D12）

- **xUnit 单测**：UDF 项目的纯逻辑（含新增 ★）在现有 `Tests` 项目补齐，规则：每函数至少正向/反向/边界样例；无期望即 fail-closed。
- **真机 harness**：新增 `tools/VerifyUdf`，真实 Excel 单元格写入 `=JYYJ.ExtractDigits(A1)` 等并断言计算值；覆盖：12 个函数各至少一例、xll 未加载状态提示、RegisterXLL 失败路径。
- **覆盖审计**：UDF 属写能力吗？——纯计算**只读输入→输出**，不写工作表，按 `DocumentedWriteToolExemptions` 语境，不算写工具；但仍须进单元/真机验证，不放行裸壳。
- 交付门槛：单测全绿 + VerifyUdf 真机全过，才可打包。

## 11. 里程碑拆解

| 阶段 | 内容 | 出口 |
|------|------|------|
| M1 | 纯逻辑上移/提公：`JyyjtoolAlgorithms`、`TryParseDate` 等进 Core 并补回归单测 | 单测绿，无行为回归 |
| M2 | 建 UDF 独立项目（ExcelDNA），按 §7 实现 12 函数 + `[JyyjFunction]` 元数据 → 导出目录 | xll 可加载、函数可算 |
| M3 | 前端函数库页真实展示 + 分组 + 复制 + 动态加载状态 | 页面无空态、状态可操作 |
| M4 | 安装器并入 x86/x64 xll + OnConnection 加载检测/重试 | 装机后 `=JYYJ.*` 可用 |
| M5 | xUnit + `tools/VerifyUdf` 真机达门禁 → 评审 | 全绿 + 评审通过 |
| M6 | 交付文档：更正「COM 暴露」表述、同步 AI交接文档、若发版同步版本号 | 文档一致 |

## 12. 风险与边界

- **重算线程**：违反 D5 的 UDF 会拖慢重算甚至闪退，评审代码时强制检查「无 COM / 无 IO / 无 Addin 状态访问」。
- **架构位宽**：xll 位数必须匹配 Excel PE，否则 RegisterXLL 失败（走 reason 提示，不硬崩）。
- **ExcelDNA 依赖**：需评估离线 NuGet / 分发体积；若引入额外 exe/dll，走受控执行/技能包合规评估。
- **版本漂移红线**：若本期发版，`Addin.csproj` + `README.md` + `installer/*.iss` + 新 UDF 项目版本四处同步。

## 13. 开放问题（评审时确认）

1. §7 首批 12 个函数清单是否最终确认？需增/删/换。
2. 是否需要「函数可被 AI 面板当作工具双入口」？（D4 一期仅为单元格，D4 可后置）。
3. ExcelDNA 版本与离线打包策略。
4. 本期是否随版本号发布，还是先行功能内测。
5. 元数据存储形态：仅程序集反射导出 JSON（低依赖）vs 落 `custom-` 用户数据区（可改可扩）——倾向前者，防破坏用户数据。