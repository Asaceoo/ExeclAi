# 设计-看板切片器探测harness（M-D0）

> 版本：v1.0（2026-09-10 探测完成并闭环）
> **探测结论：`slicer_native_ready = true`（本机 Excel 16.0 Build 20326，7/8 项 PASS，P8 由 P3 覆盖）——P3 可开发 native 切片器联动模式**
> 代码：`tools/SlicerProbe/`（独立探测项目）；报告：`tools/SlicerProbe/bin/Release/net48/reports/slicer_probe_16_0.json`
> 上游：`docs/设计-原生数据看板.md` §3③联动层、`docs/需求-原生数据看板.md` M-D0/NFR-1
> 性质：**只读探测 harness**（复用 tools/RealMachineVerify 模式），无写能力变更 → 不触发覆盖审计，但探测代码本身要过构建与现有单测基线

---

## 1. 探测目标与判定矩阵

回答一个核心问题：**本机 Excel 版本上，COM 晚绑定能否可靠创建并驱动原生切片器（Slicer）**。结果决定设计文档 §3③ 的 `linkage.mode` 默认值。

| # | 探测项 | 方法 | 判定 |
|---|---|---|---|
| P1 | API 存在性 | 反射/晚绑定调用 `Workbook.SlicerCaches` 属性 | 不存在 → native 直接出局，P3 走 cell |
| P2 | SlicerCaches.Add2 创建（基于透视表） | 建临时透视表 → `SlicerCaches.Add2(pivot, "字段")` | 抛错则记录 HRESULT 与异常类型 |
| P3 | Slicers.Add 放置切片器 | `cache.Slicers.Add`，设置 Top/Left/Width/Height | 几何属性可设性 |
| P4 | 切片器选择驱动联动 | 模拟选中单项 → 读透视表汇总值变化 | **联动闭环成立与否的关键项** |
| P5 | Report Connections 等效性 | 切片器缓存关联多透视表（`cache.PivotTables` 集合） | 多图联动的必要能力 |
| P6 | 检测现有切片器（只读回读） | 遍历现有 SlicerCaches/Slicers 属性 | 渲染引擎清理/幂等重生成所需 |
| P7 | 删除切片器 | `cache.Slicers` 逐个删除 + cache 删除 | 重新生成/删除看板的清理路径 |
| P8 | 显示位置独立 sheet | 切片器放在非透视表所在 sheet | 看板 sheet 与辅助 sheet 分离的前提 |

**判定输出**（写入探测报告 JSON，供 release/设计文档引用）：
- `slicer_native_ready: true|false`（P1-P5 全过才为 true）
- `partial_support`: 具体哪几项通过（如仅能创建不能驱动 = 无用）
- `excel_version`: `Application.Version` + Build
- `error_catalog`: 每项的 HRESULT/异常消息（沉淀踩坑手册）

## 2. 测试工作簿构造（探测代码内动态创建，用完即弃）

```
探测流程（全程内存中，不落盘用户目录）：
1. 新建临时 Workbook
2. 写入 20 行模拟销售明细（地区 3 值 × 品类 3 值 × 月份 2 值）
3. 建 PivotCache + 透视表（行=地区，值=销售额求和）
4. 依次执行 P2-P8
5. workbook.Close(SaveChanges:=false)（不保存，零残留）
```

防护要求：
- 全程 try/finally，异常路径也保证关闭临时工作簿（复用无状态 COM 纪律：`Release((object)x)`）；
- 探测期间捕获的 COM 异常**必须记录 HRESULT**（踩坑手册需要真实错误码，不是笼统"失败"）；
- 不触碰用户已打开的工作簿；探测入口断言 ActiveWorkbook 为临时簿。

## 3. 代码落位与形态

| 项 | 决定 |
|---|---|
| 位置 | `tools/RealMachineVerify/` 下新增独立探测用例组（M-D0 专用），不混入 22 用例组基线 |
| 形态 | 独立控制台入口（类似现有 harness），输出 JSON 探测报告到 `tools/RealMachineVerify/reports/slicer_probe_<version>.json` |
| 断言风格 | 探测失败不是测试失败——每项独立记录通过/失败+错误详情，报告以 `slicer_native_ready` 为结论 |
| 单测 | 探测代码的解析/报告生成部分（纯逻辑）可单测；COM 部分只能真机跑 |
| WPS | **不探测**（D11：看板仅 Excel） |

## 4. 判定后的行动映射

| 结果 | 行动 |
|---|---|
| `slicer_native_ready = true` | P3 开发 native 模式；设计文档 `linkage.mode` 默认改 `native`；FR-2 P3 验收启用 |
| `false`（API 不存在或创建失败） | P3 取消 native；`linkage.mode` 固定 `cell`；设计文档 §3③ 与 PRD FR-2 相应收敛；报告归档踩坑手册 |
| `partial`（能建不能驱动等） | 视具体项评估：不能驱动（P4 失败）等同 false；仅几何/清理缺陷可评估绕过 |

## 5. 验收结果（2026-09-10 实测）

- 探测代码构建 0 警告 0 错误 ✓；979 单测全绿无回退 ✓；报告 JSON 生成 ✓；零工作簿残留（临时 xlsx 已删）✓；复跑可复现 ✓。

## 6.x M-D1 真机验证结论（2026-09-10，SlicerProbe.exe md1；二轮审查后 10/10）

**10/10 全 PASS**（二轮对抗性审查补 T9 超界列拒绝无残留 / T10 文本型数字+合计行并行检出两个用例；修复渲染失败回滚缺失（AC-5）、扫描 break 不完整、越界静默钳制三处缺陷后回归通过）。原记录：（T1 数据门槛正向 / T2 合计行硬拦截 / T3 双轴组合图 / T4 verify 次轴断言 / T5 环形图 / T6 kpi_brief 渲染 / T7 守恒对账 / T8 幂等重生成），复跑两次稳定。主探针 M-D0 复跑 `slicer_native_ready=true` 不回退；单测 986 全绿；构建 0 警告 0 错误。

### 审查修复的真实缺陷（对 P2/P3 有实现指导意义）

1. **Dynamic 元组传染**：`CollectSeriesMeta(dynamic...)` 的返回 ValueTuple 在 dynamic 分派下 Item7/Item8 绑定失败（Item8 曾退化为嵌套 `ValueTuple<bool>`）——治本：签名改 object 入参 + 私有结果类 `SeriesMeta`，调用点字段访问全静态；
2. **RCW 断开陷阱（out-of-proc）**：`FinalReleaseComObject` 与 `ReleaseComObject` 减到 0 都会断开同对象**所有** RCW（按 IUnknown 缓存共享）。跨调用复用的 sheet 对象（看板/辅助表）在服务方法尾部**不得释放**，断开时机交给进程退出；
3. **公式双引号**：`EscapeSheet` 已含单引号，模板串再包一层会成 `''sheet''` → 0x800A03EC；
4. **NumberFormat 全角货币符**：`¥#,##0` 在部分 COM 环境抛 0x800A03EC → NumberFormat 赋值包 try/catch 降级 `#,##0`；
5. **幂等删除**：删除旧看板/辅助表包 try/catch 降级，`Worksheets.Add` 用无参形式（`Add(After: Item[Count])` 在释放链后抛 0x800A01A8）。

## 6. 探测实录与结论

| 项 | 结果 | 关键证据 |
|---|---|---|
| P1 API 存在性 | ✅ | `SlicerCaches` 可访问 |
| P2 创建 Cache | ✅ | `Add2(pivot, "字段")` 成功，name=切片器_地区 |
| P3 放置 | ✅ | 可放置于独立 sheet（P8 同步成立） |
| **P4 选择驱动联动** | ✅ | 选中"华北"后透视值 **19026 → 6342**（联动闭环成立） |
| P5 多透视表关联 | ✅ | **必须共享同一 PivotCache**（`pivotTable.PivotCache()` 复用），挂接后 Count=2 |
| P6 只读回读 | ✅ | caches/slicers 可枚举 |
| P7 删除清理 | ✅ | **`slicer.Delete()` 级联删除整个 cache**（删后 caches=0），无需单独删 cache |

**行动映射（设计文档 §4）**：`slicer_native_ready=true` → P3 开发 native 模式；`linkage.mode` 默认 `native`（cell 保留为兼容兜底）；FR-2 P3 验收启用。

### 6.1 探测过程踩出的坑（已具备实现指导意义）

1. **兼容模式陷阱**：`Workbooks.Add()` 未存盘工作簿=兼容模式，切片器被禁（0x800A03EC 明示）。**看板渲染引擎必须先把工作簿 SaveAs 为 xlsx（或探测 CompatibilityMode）再建切片器**；
2. **PivotCaches 必须用新式 `.Create()`**（旧 `.Add(SourceType:,SourceData:)` 命名参数在 dynamic 下绑定失败且触发兼容模式问题）；
3. **多透视表关联铁律**：切片器只能挂接**共享同一 PivotCache** 的透视表——渲染引擎建多图联动时必须复用同一 PivotCache（`pivot.PivotCache()` 获取）；
4. **删除语义**：`slicer.Delete()` 级联删除 cache，删除后旧 cache RCW 失效（0x800A01A8）——清理代码按"删 slicer → 读 caches.Count → 剩余才删"顺序编写；
5. **透视值读取**：联动验证需 `app.Calculate()` 后读 `GetPivotData`/单元格，COM variant 数值用 `is double` 模式解析（Convert.ToDouble 对部分 variant 抛 FormatException）。

> 以上 5 条实现约束应随 P3 native 开发任务书下发；踩坑 1/3/4 属通用 COM 坑，沉淀踩坑手册。
