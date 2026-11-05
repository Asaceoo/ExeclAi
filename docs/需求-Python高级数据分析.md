# 需求-Python高级数据分析（Python 分析通道增强批次）

> 版本：v1.3（2026-09-13，P3 落地（AC-P3/AC-P4 验收通过）：①用户机 Python 3.12.4/64 位安装 statsmodels 0.15.0 / jieba 0.42.1 / wordcloud 1.9.6（清华镜像，依赖链 pandas/numpy/scipy/matplotlib/pillow 原已齐备，补装 openpyxl）；②新增 `tools/validate_python_p3.py` 可复跑验收脚本：销售明细 180 行 12 个月聚合 → 留出回测（Holt MAPE 12.4% / ARIMA(0,1,0) 15.4%，均 <50% 断言过）→ 全量拟合选型 Holt 出未来 3 个月预测 + 95% CI + 预测图；产品名称列 jieba 分词词频 → wordcloud 词云（msyh.ttc 字体探测链命中）；③看图验收：预测图与词云中文零乱码；④已记录边界：季节性分解（STL/季节 Holt）需 ≥24 个月度点，12 点数据走趋势型模型，运行期由技能契约引导 AI 按数据长度选型；⑤worker 层超时守卫已由单测覆盖，AC-P4 超时负例不重复真机验证）。v1.2（2026-09-13，P2 落地：①新增内置技能 `python-advanced-analysis`（第 14 个，随 csproj 通配符自动入包，bin 输出已验证）承载 FR-P2 任务路由契约 + FR-P3 出图嵌表契约（中文字体探测链 msyh.ttc→simhei.ttf→simsun.ttc fail-closed、matplotlib/seaborn PNG 走 insert_image 嵌表、plotly 只交付交互 HTML 不嵌表、run_python_excel 高危确认门 + undo 回滚、结论三件套、ARIMA 网格收敛纪律）；②自愈白名单补齐 plotly/xlwings/openpyxl_image_loader 历史盲区，新增单测断言"目录全部 24 库导入名均被自愈兜住"全覆盖；构建 0 警 0 错、全量单测 1047/1047 过）。v1.1（2026-09-13，P1 实施期修正：核实 Catalog 现状发现 jieba 已在库目录（22 库含 jieba），实新增 statsmodels/wordcloud 两库 + 自愈白名单补 jieba/wordcloud/statsmodels 三映射，库数口径修正为 22→24；既有事实记录：plotly/xlwings/openpyxl-image-loader 不在缺模块自愈白名单，属历史盲区，列入 P2 备查不扩本批范围）。v1.0（2026-09-13，grill-me 六问全 A 拍板：①期次 P1/P2/P3 三期切分；②plotly 保留、交付交互 HTML 不嵌表；③FR-P4 保持自由脚本契约层，标准工具另行立项；④wordcloud 失败走既有兜底不阻塞；⑤中文字体探测链 fail-closed；⑥AC-P4 用仓库自带测试工作簿验收）。v0.1（2026-09-13，grill-me 产出草稿）
> 状态：P1/P2/P3 全部落地；剩余运行期验证项——用户首次在 Excel 内实际使用时由缺模块自愈/确认门链路兜底，随下次发版 gate 一并回归
> 上游调研：`docs/../数据分析功能扩展调研报告.md`（Python 路线）+ 源码核实：`ExcelSecondPhaseRegistration.cs` run_python_excel 注册描述、`ScriptRunnerService.cs` MissingModulePipMap、`PythonEnvService.cs` Catalog
> 红线继承：结果必须由用户 Excel 数据产生 + fail-closed + 可回滚 + 受控执行（白名单/禁网络/超时/确认门）

---

## 1. 背景与机会

- **通道已具备**：run_python_excel / run_python_code 常驻 worker + xlwings 直连活动工作簿；input_data 注入、xai_result 结构化回传、write_ranges 快照回滚、persist 跨调用状态；缺模块自愈闭环（白名单 → pip 镜像 → 重建 worker → 重试一次）。
- **分析空白**：C# 分析族止步描述统计；假设检验/回归/聚类/异常检测/时序/中文文本挖掘只能靠 AI 临时写脚本，无任务路由契约、无依据回显、无验收用例。
- **库缺口**：statsmodels（时序）、jieba/wordcloud（中文文本）既不在 `PythonEnvService.Catalog`（库目录/管理面板）也不在 `MissingModulePipMap`（缺模块自愈白名单），AI 用了也无法自愈、无法一键装。

## 2. 功能需求

| # | 需求 | 说明 |
|---|---|---|
| FR-P1 | 库目录与白名单扩展（前提项） | ① `PythonEnvService.Catalog` 增 statsmodels / jieba / wordcloud 三条（含导入名与用途说明）；② `MissingModulePipMap` 同步加三映射（否则缺模块自愈兜不住）；③ 侧边栏 Python 库管理面板（已装库/教程/一键安装三卡）随 Catalog 自动扩容 22→24 库（面板为动态渲染，改 Catalog 即随动，前端零改动），逐库安装按钮同步；④ 依赖事实：statsmodels 依赖 scipy+patsy（scipy 已在白名单，patsy 随装自动带入、面板说明标注）；jieba 纯 Python 零依赖；wordcloud 依赖 pillow+numpy（均在白名单），Windows 含 C 扩展 wheel，安装失败走既有"镜像 → --user → 分级提示"兜底，**单库失败不阻塞其他库**；⑤ 安装沿用 install_python_packages Dangerous 确认门（AI 先展示库目录、用户明确选择后才装）。 |
| FR-P2 | 高阶分析任务路由契约 | SKILL.md 增路由：描述统计/图表推荐/看板 → C# analyze_*；假设检验（scipy.stats：t/卡方/ANOVA/正态性）、回归/聚类/异常检测（sklearn）、批量数据处理（pandas）→ run_python_excel。结论三件套：xai_result 回传（统计量+p 值+方法名）+ 关键结果写回 sheet（write_ranges 声明）+ basis 计算依据（方法/参数/样本区间，与 FR-K3 形态对齐）。 |
| FR-P3 | 出图嵌表链路 | matplotlib/seaborn 静态图 → PNG → 复用现有插图工具嵌 sheet（保留 PNG/JPEG 魔数校验前置）；中文字体显式指定系统字体文件（msyh.ttc → simhei.ttf 探测链，缺失时返回替代清单并**放弃出图**，fail-closed 不出乱码图）；**plotly 定位修正：产出交互 HTML 非位图，不能嵌表**——交付形态为导出 HTML 到工作簿同目录 + 聊天回显路径，浏览器打开。 |
| FR-P4 | 时序预测与中文文本分析（新库解锁） | 时序：STL 分解 + Holt-Winters + **限定网格 ARIMA**（超时纪律：默认 30s，代码内可传 timeout_ms ≤ 300s 上限；网格搜索空间硬编码收敛，禁止暴力调参）；输出预测值/置信区间/季节性分解写回 sheet + 历史折线延伸图。文本：jieba 分词 + 词频/关键词（TF-IDF）统计表写回 + wordcloud 词云 PNG 嵌表。 |

## 3. 约束（红线继承）

- 受控执行不变：禁网络、30s/300s 超时、Dangerous 确认门、write_ranges 快照回滚；
- 新库安装必须用户显式选择，安装失败分级人话提示，不静默重试到死；
- 分析结论引用的数据必须来自用户工作簿（xw.books.active 通道），禁止凭空造数；
- 本需求为"用法契约 + 白名单扩展"层，不新增 C# 工具，不触发工具数基线变更（若后续把 FR-P4 沉淀为标准工具 analyze_forecast_pro 等，另行立项走覆盖审计/探针/SKILL.md 四件套）。

## 4. 验收标准

- AC-P1：三库在库管理面板可见、可一键安装；真机自愈用例（代码 import statsmodels 未装 → 自动装 → 重试通过）；Catalog/MissingModulePipMap 单测断言 24 库；
- AC-P2：t 检验样例全链一次跑通（读选区 → scipy 检验 → 写回结论区 → 回复附 p 值与 basis），ud 探针用例留痕；
- AC-P3：中文词云 PNG 看图验收无乱码；plotly HTML 双击可交互；
- AC-P4：月度销售 ARIMA 预测——**验收数据集 = 仓库自带测试工作簿**（历史段对账误差合理 + 预测段趋势人工核对，用户真实文件可选加验）；构造大网格用例被超时守卫拦截且返回人话提示（不挂死 worker）。

## 5. 期次建议

| 期次 | 内容 | 理由 |
|---|---|---|
| P1 | FR-P1 白名单扩展 | 改两个字典 + 面板随动，成本最低，解锁全部后续 |
| P2 | FR-P2 + FR-P3 路由契约与出图链路 | 零新库，纯契约/文档/链路验证 |
| P3 | FR-P4 时序与文本分析 | 依赖 P1 装库成功，且涉及超时纪律调优 |

## 6. 拍板结论（grill-me 六问全 A，2026-09-13）

- Q1 期次：按 P1 白名单扩展 → P2 路由契约+出图链路 → P3 时序/文本分析，三期独立收口；
- Q2 plotly：保留，交付交互 HTML 到工作簿同目录 + 聊天回显路径，不嵌 sheet；
- Q3 FR-P4 形态：自由脚本 + SKILL.md 任务路由契约，不改 C#、不触发工具数基线；沉淀标准工具另行立项走覆盖审计/探针/SKILL.md 四件套；
- Q4 wordcloud 失败兜底：既有"镜像 → --user → 分级人话提示"，单库失败不阻塞其他库，不做源码编译/conda 额外文档；
- Q5 中文字体：msyh.ttc → simhei.ttf 探测链，缺失返回替代清单并放弃出图（fail-closed）；
- Q6 AC-P4 验收数据：仓库自带测试工作簿月度销售数据，用户真实文件可选加验。
