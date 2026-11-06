# 设计文档：JYYJ 功能区一键命令「标准可配」与可见性提示

> 状态：已实现（增量 A/B 已落地并通过单测）；真机验收待补
> 日期：2026-09-06
> 涉及项目：`ExcelAiAssistant.Addin`（Ribbon 定义）、`ExcelAiAssistant.Excel`（执行体 + 规划器）、`ExcelAiAssistant.Tests`（单测）

---

## 1. 背景与目的

用户针对功能区截图质疑：**「清洗」等一键功能没有参数、没有清洗标准**。经 grill-me 三轮拷问收敛需求方向后，进入实现前做了代码现状核实，得到一个颠覆初始预期的结论——**现状并非"无参数"**。

本功能的目标仅是补齐两个真实的增量缺口，而不是从零造"参数框"（那会造成无用大工程与返工）。

---

## 2. 现状澄清（关键校正）

- **全部 24 个 Ribbon 命令按钮点开后均有参数框**，且标准可配：
  - 清洗：去空行 / 修日期（可选列）/ 填空白（可选填充值）
  - 去重：判重范围 / 判重列 / 保留首·末 / 含表头
  - 排序 / 筛选 / 转置 / 透视 / 拆分 / 合并 / 清理 / 填公式 / … 同样有完整字段
- 每个命令经 `RibbonCommandPlanner.BuildFields` 构建参数框字段、经 `BuildArgs` 把勾选值映射为真实工具参数，`RibbonCommandService.Execute` 统一走「参数框 → 采集 → BuildArgs → 工具执行 → 摘要」。
- **用户误判根因**：按钮为纯图标、零视觉提示，点开前用户不知道里面藏有可配置标准，也不清楚默认清洗规则是什么。

> 结论：缺少的不是"参数"，而是 **①按钮可见性提示** 与 **②清洗标准项不全（缺去空格/去控制字符）**。

---

## 3. 需求聚焦（grilling 收敛 + 现状协商）

经 grill-me 三轮 + 现状核实后与用户对齐，本次只做两个增量：

| 增量 | 说明 | 决策依据 |
|------|------|----------|
| **A. 可见性提示** | 给命令按钮加 hover 提示「点击配置…标准」，消除"看似无参"的误判 | 用户截图误判的根因，成本最低、收益最直接 |
| **B. 清洗标准补齐** | 清洗参数框增加「去首尾空格 / 去控制字符」勾选项，映射到真实执行 | 对齐清洗标准草案中现状缺失的两项 |

### 明确定界（不做）
- **记忆上次配置**：需引入持久化，本轮不纳入（用户选择暂缓）。
- **强制统一为"全部按钮弹确认框"**：视图/toggle/导航类（聚光灯/指针/打开面板）保持零摩擦，不做机械统一。
- **新增工具/改动既有工具执行体以外的范围**：不涉及。

---

## 4. 增量方案

### 4.1 可见性提示（Ribbon.xml）

给 `JYYJDataGroup / JYYJTextGroup / JYYJCommonGroup` 三个组的 18 个命令按钮加 `supertip` 静态属性（hover 显示）：

```
清洗：  "点击选择清洗标准：去空行 / 修日期 / 填空白 / 去空格 / 去控制字符"
去重：  "点击配置判重列与保留规则"
排序：  "点击配置排序列与升降序"
转置：  "点击配置写入位置"
筛选：  "点击配置条件列与运算符"
透视：  "点击配置行标签与数值列"
拆分：  "点击设置源列与分隔符"
合并：  "点击选择合并列与分隔符"
清理：  "点击选择清理项与大小写"
填公式： "点击填写公式模板"
插文本/删文本/随机数/序列/插空行/删空行/填下/合并相同：各配一句
```

「打开面板/聚光灯/指针」三个 toggle 已有动态 `getScreentip`，不改；「更多」组（宏收纳箱/函数库/SQL查询）为收纳导航类，不纳入本次。

### 4.2 清洗标准补齐（三段式落地）

| 层 | 文件 | 改动 |
|----|------|------|
| 算法层 | `DataTransformService` | 新增 `TrimWhitespace(object[,])`（去每格字符串首尾空白）、`RemoveNonPrintable(object[,])`（移除 `Char.IsControl` 控制字符）。均原地改写 + 返回 `TransformOutcome(values, changed)`，数字/日期/空单元格不触碰、不覆写为字符串 |
| 执行体 | `ExcelToolHost` clean 分发 | `loop over operations` 增加 `case "trim"`、`case "remove_non_printable"`，调用上两个方法，计 `changedCells` |
| 规划层 | `RibbonCommandPlanner` | `BuildFields("clean_excel_data")` 增加两个 `checkbox`（`trim`=去首尾空格、`remove_non_printable`=去控制字符，**默认 false 不改现状行为**）；`BuildArgs` 在勾选时向 `operations` 追加 `{tool:"trim"}` / `{tool:"remove_non_printable"}` |

> 注：`remove_blank_rows / fix_dates / fill_blanks` 三个既有子操作保持不变；新增两个操作挂进同一条 `operations` 链，`verify_data_clean` 对 `clean_excel_data` 的判定不变（该工具本就在覆盖审计内）。

---

## 5. 变更清单（文件级）

- `src/ExcelAiAssistant.Addin/Ribbon.xml` — 18 个按钮加 supertip
- `src/ExcelAiAssistant.Excel/DataTransformService.cs` — 新增 `TrimWhitespace` / `RemoveNonPrintable`
- `src/ExcelAiAssistant.Excel/ExcelToolHost.cs` — clean switch 增 `trim` / `remove_non_printable` case
- `src/ExcelAiAssistant.Excel/RibbonCommandPlanner.cs` — `BuildFields` + `BuildArgs` 增两项
- `src/ExcelAiAssistant.Tests/DataTransformServiceTests.cs` — 增 2 个测试
- `src/ExcelAiAssistant.Tests/RibbonCommandPlannerTests.cs` — BuildFields 计数 5→7、增 BuildArgs 映射测试

---

## 6. 验证状态

| 项目 | 结果 |
|------|------|
| 编译 `ExcelAiAssistant.sln`（Release，x64） | 通过，0 警告 0 错误 |
| 单测 | **829 全过**（806→829 累计新增 23） |
| 部署 | 新 `Addin.dll` / `Excel.dll` 已同步到 `%LOCALAPPDATA%\XaiAssistant\app\x64\` |
| **真机验收（悬停提示显示 + 清洗新勾选项执行去空格/去控制字符）** | **待补**（用户选择本轮跳过，留待后续） |

对抗审查（三角色）结论：无返工点。核心安全点——`TrimWhitespace`/`RemoveNonPrintable` 对非字符串单元格经 `s==trimmed` 判定不落写，数字/日期不会被转成字符串；新 case 在 switch 同一作用域独立变量名，编译通过；`trim` 默认 false，未勾选时 `operations` 与旧版完全一致。

---

## 7. 取舍与已知边界

- 「去首尾空格」作用于**整个选区所有列**（不提供单列限定）；「去控制字符」移除全部控制字符（含 `\t`）。用户可通过不勾选规避。
- 清洗默认仍只勾旧三项，新增标准项默认不启用——**不改变用户既有默认行为**，标准从"固定不可见"变为"可勾选透明"。
- 可见性提示为 hover 形态，不遮挡操作、不增加点击成本。
- 本次未做"记忆上次配置"（见 §3 定界），未来如需可基于 `BuildFields` 默认值持久化扩展。

---

## 8. 执行标准增强（v2.18.8，grill 拷问收敛）

| # | 标准 | 落点 |
|---|---|---|
| 1 | **执行期防重入**：busy 守卫忽略执行期新请求，杜绝连点导致同一工具串行重复执行 | `RibbonCommandService.Execute` |
| 2 | **覆盖预算告知**：Overwrite 类黄字显示「将覆盖原区域（约 N 个单元格）」（long 防 int 溢出） | `OverlayHint(safety, sel)` |
| 3 | **目标带数据守卫统一**：text_split 拆分带区 / text_merge 结果列 / transpose 指定位置，写前检查占用即拦截并提示 | `TryGetTargetRect` + `IsTargetOccupied` |
| 4 | **执行后自检 fail-closed**：transpose 维度+锚点格、删行数=执行前预估、随机数值域∈[min,max]（含舍入容差），不符弹红色告警 | `PostVerify` + `VerifyTransposeResult`/`VerifyRandomRange` |
| 5 | **参数级审计日志**：执行成功后选区地址 + 完整实参 JSON 落盘（兑现「运行日志全量留痕」，设置页可查） | `ExecuteTool` 双日志行 |
| 6 | **随机类验收锚点**：fill_random_numbers 自检通过后摘要附 1–3 个实际样本值，用户可核对本次生成值 | `VerifyRandom` out 样本 |
| 7 | **错误码文案映射补全**：INVALID_PATTERN/INVALID_CASE 带中文文案时展示文案而非裸码 | `BuildSummary` 错误码 case |
| 8 | **标准审计注册表**：`NeedsPostVerify` 单一来源声明自检覆盖清单，新增高风险工具必须表态（测试锁定） | `NeedsPostVerify` + 审计测试 |

- 边界：动态目标工具（pivot/filter 输出行数未知）与 transpose 原位写回不做目标守卫；守卫/自检均为保护性检查——读取异常不误拦、不阻断正常流程。
- 单测：TryParseAddress / TryGetTargetRect / VerifyTransposeResult / VerifyRandomRange / NeedsPostVerify / 错误码文案均有独立测试锁定；OverlayHint 文案带覆盖格数断言。累计 **855 用例**（v2.18.8）。
- 真机端到端（RealMachineVerify 用例 20）：守卫/自检/样本决策链路用真实 COM 数据验证——text_merge 目标列空区放行、有数据拦截；fill_random_numbers 回读值域自检通过、样本锚点提取 1–3 个、篡改越界值被自检抓到。弹窗展示本身需人工点击，故判定逻辑抽为纯函数 `MatrixHasData`（占用判定，0 基矩阵）+ `PickSampleValues`（样本提取），服务层 `BlockHasData`/`VerifyRandom` 委托之，单测锁定边界。