---
name: python-advanced-analysis
description: Python 高阶数据分析路由助手。假设检验/回归聚类/异常检测/时序预测/中文文本挖掘与词云/统计可视化走 Python 通道，结论写回表格并附计算依据；描述统计与图表推荐仍走内置 C# 工具。
version: 2.4.0
type: orchestration
category: 分析
risk_policy: high
tools: [run_python_excel, run_python_code, insert_image, undo_excel_operation]
triggers: [假设检验, t检验, 卡方检验, 方差分析, 显著性, 回归分析, 聚类, 客户分群, 异常检测, 离群值, 时序预测, 销量预测, ARIMA, 趋势预测, 分词, 词频, 词云, 热力图, 相关性矩阵]
examples: [对比两组转化率差异是否显著, 对月度销售额做未来3个月预测, 给客户消费数据做RFM分群, 把评论列做词频统计并生成词云, 画各指标相关性热力图]
---

# Python 高阶数据分析路由技能

## 第零步：任务路由（先判断通道，不要用 Python 重做内置工具已覆盖的事）

| 任务 | 通道 |
|---|---|
| 描述统计（均值/中位数/占比/分布）、数值相关性（阈值可调 threshold）、轻量趋势预测（线性/移动平均）、图表推荐、看板 | 内置 C# 工具：analyze_range_statistics（含 IQR 离群点明细）/ analyze_column_correlation / analyze_value_distribution / analyze_trend_forecast / recommend_excel_chart / render_dashboard |
| 假设检验——常用四类（Welch 双样本 t/配对 t/单因素 ANOVA/卡方独立性） | 内置 C# 工具：analyze_hypothesis_test（精确 p 值+效应量+前提提醒） |
| 假设检验——非参检验（Mann-Whitney/Wilcoxon/Fisher 精确）/正态性/复杂设计 | run_python_excel + scipy.stats |
| 回归归因——线性/多元 OLS（R²+系数表+VIF/Cook's D 诊断） | 内置 C# 工具：analyze_regression（与 statsmodels 对齐，完全共线 fail-closed） |
| 回归归因——逻辑回归/正则化/非线性/稳健回归 | run_python_excel + statsmodels / scikit-learn |
| 分群画像——KMeans（数值特征，自动定 k） | 内置 C# 工具：analyze_cluster（标准化+轮廓系数+各簇画像） |
| 文本变体聚类——名称列错字/空白/全半角变体建议（探索式清洗） | 内置 C# 工具：analyze_text_variants（归一化+编辑距离≤阈值+首字符约束；只建议，合并走 write_excel_cells_batch 改值） |
| 文本变体进阶——拼音/语义相似、唯一值超 1500 的批量模糊匹配 | run_python_excel + rapidfuzz |
| 分群/分类——层次聚类/DBSCAN/监督分类 | run_python_excel + scikit-learn |
| 降维与结构（PCA/因子分析/相关矩阵热力图） | run_python_excel + scikit-learn / seaborn（PCA 前必须标准化） |
| 异常检测进阶（孤立森林/DBSCAN，配合内置 IQR 离群点） | run_python_excel + scikit-learn |
| 时序预测（STL 分解/Holt-Winters/ARIMA+预测区间） | run_python_excel + statsmodels |
| 实验与抽样（A/B 显著性/统计功效/样本量计算） | run_python_excel + statsmodels.stats |
| 优化与模拟（线性规划 scipy.optimize.linprog / 蒙特卡洛 numpy，须声明模拟次数与分布假设） | run_python_excel + scipy / numpy |
| 综合评价（加权评分/熵权法/TOPSIS+雷达图，先统一指标方向） | run_python_excel + numpy / matplotlib |
| 中文分词/词频/关键词/词云/情感分析（snownlp） | run_python_excel + jieba / wordcloud / snownlp |
| 批量数据处理（merge/去重/透视重塑） | run_python_excel + pandas |
| 纯计算不碰表格 | run_python_code（低风险） |

## 第一步：口径确认
1. 用只读工具读取数据结构与表头，确认分析列与样本区间。
2. 先向用户复述：用什么方法、检验/建模什么、结论写到哪里，确认后再执行。

## 第二步：Python 执行契约
- 数据进出：`xw.books.active` 读数 → pandas 分析 → `xai_result` 回传结构化结论（统计量/p 值/方法名/模型参数）。
- **通道铁律**：openpyxl 等磁盘引擎只处理未打开的文件；要写当前已打开的工作簿必须走 `xw.books.active` 直写——对已打开文件 `to_excel/save` 会报 PermissionError（Excel 进程持锁）。
- **边界铁律**：序列/字典下标访问前先做 `len()`/`in` 检查（IndexError/KeyError 高频源）；逐行处理前先确认行/列数一致。
- 写回：`write_ranges` 声明写入区域（自动快照、失败回滚）；结论写新区域或空白列，**不得覆盖用户原始数据**。
- 结论三件套缺一不可：①数值结论 ②计算依据（方法/参数/样本区间，如"双样本 t 检验，n1=86，n2=92，α=0.05"）③一句人话解读。
- 时序纪律：默认 30 秒超时，必要时传 timeout_ms（上限 300 秒）；ARIMA 网格搜索空间必须收敛（如 p,d,q 各 ≤2 的少量组合），禁止暴力调参。
- **统计口径零容忍纪律（v2.0）**：①检验前查前提——t 检验/ANOVA 前做正态性与方差齐性检查（不满足改非参检验并注明）；②分类变量先编码（LabelEncoder/OneHot）再入模；③聚类前必须标准化（StandardScaler），否则距离量纲失真；④回归须报告 R²/系数方向/残差诊断，不只给 p 值；⑤综合评价先统一指标方向（正向化/负向化）再算权。
- **文本情感分析契约（v2.3，snownlp）**：逐条 `SnowNLP(text).sentiments` 打情感分（0~1，越大越积极；建议阈值 ≥0.6 积极 / ≤0.4 消极 / 其余中性，报告须注明阈值为业务约定非统计检验）；批量分析结果经 write_ranges 写入右侧空白列（原文列旁：情感分+倾向标签两列，快照可回滚）；结论须给积极/消极占比与典型样本各 1~2 条；库缺失时经缺模块自愈兜底（白名单已含 snownlp）。情感结论可配内置工具 add_insight_card 生成洞察卡进看板（标题=一句结论，正文=占比+典型样本）。
- **大表抽样与超时分级契约（v2.4）**：
  - **抽样前置**：数据超过 5 万行时，建模类分析（聚类/回归/网格搜索/蒙特卡洛）不得直接全量拟合——先用 `df.sample(n=50000, random_state=42)` 随机抽样（有业务分层诉求时按关键维度 `groupby(...).apply(lambda g: g.sample(min(len(g), 每层上限), random_state=42))` 分层抽样），**抽样参数（抽样行数 n、random_state、分层列）必须写进结论三件套的「计算依据」**，明确告知用户"结论基于 X 万行中的 n 行样本"。全量可完成的统计（均值/占比/分布/情感打分/去重计数）不抽样。
  - **超时分级**：按场景显式传 `timeout_ms`（上限 300 秒）——统计检验/单次拟合类 60 秒；聚类/降维/时序建模类 120 秒；网格搜索/蒙特卡洛/大批量文本处理类顶格 300 秒。默认 30 秒只够轻量计算，大表建模不传 timeout_ms 必然超时并重置 worker（旧状态丢失）。调用前先估数据量（行数×列数×方法复杂度），宁可分级顶格也不反复重试。
- 缺库自愈：statsmodels/jieba/wordcloud/pandas/scipy/sklearn 等缺模块会自动 pip 安装并重试一次；wordcloud 安装失败不阻塞主分析（跳过词云并告知用户）。

## 第三步：出图嵌表契约
- **matplotlib/seaborn 静态图**：绘图前设置中文字体——按顺序探测 `C:\Windows\Fonts\msyh.ttc` → `simhei.ttf` → `simsun.ttc`，用 `matplotlib.font_manager.fontManager.addfont` + `rcParams['font.family']` 生效；全部缺失时**放弃出图并告知用户**，禁止输出乱码图。图存 PNG 到工作簿同目录，再用 `insert_image` 嵌到锚点单元格（不要用 Python 直接操作 Excel 内图片对象）。
- **plotly 交互图**：`fig.write_html` 存工作簿同目录，回复给出完整文件路径由用户浏览器打开；plotly 图是交互 HTML，**不能嵌入 sheet，不要尝试**。
- 生成后说明图片/HTML 的落盘路径与锚点位置。

## 第四步：汇报与回滚
- 汇报顺序：结论数值 → 计算依据 → 写回位置 → 图表位置/文件路径。
- run_python_excel 属高危工具（直写工作簿），执行前必须向用户说明要做什么、写哪里；完成后如需撤销用 `undo_excel_operation`。
