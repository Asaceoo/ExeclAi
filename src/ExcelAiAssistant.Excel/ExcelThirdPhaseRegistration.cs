using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第三期工具注册：导入外部文件/批量处理 与 数据分析/可视化。
    /// </summary>
    public static class ExcelThirdPhaseRegistration
    {
        private static readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> NoWriter =
            (_1, _2, _3, _4) => throw new InvalidOperationException("缺少写入回调。");

        public static void RegisterThirdPhaseTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer,
            OperationRecorder recorder)
        {
            var import = new ExcelImportService(application, writer ?? NoWriter, recorder);
            var analysis = new ExcelAnalysisService(application, recorder);
            var chart = new ExcelChartService(application, recorder);
            var report = new ExcelReportService(application);
            var formulaTools = new ExcelFormulaService(application);

            // ---- 导入外部文件与批量处理 ----
            // 高危：preview_external_file 读取任意外部文本路径并返回给模型，ReadOnly 会自动放行＝静默任意文件读取；须弹确认展示文件路径。
            Register(registry, "preview_external_file", "预览外部文本文件前若干行，自动识别分隔符与编码。", RiskLevel.Dangerous,
                () => import.PreviewExternalFile, ThirdPhaseSchemas.PreviewExternalFile());
            Register(registry, "import_delimited_file", "把分隔文本文件导入到工作表的指定区域。", RiskLevel.NormalWrite,
                () => import.ImportDelimitedFile, ThirdPhaseSchemas.ImportDelimitedFile());
            Register(registry, "import_excel_file_sheet", "导入另一个 Excel 文件的工作表（作为新表或取值复制）。", RiskLevel.StructuralChange,
                () => import.ImportExcelFileSheet, ThirdPhaseSchemas.ImportExcelFileSheet());
            Register(registry, "merge_delimited_files", "把多个分隔文本文件按表头合并输出到一个区域。", RiskLevel.NormalWrite,
                () => import.MergeDelimitedFiles, ThirdPhaseSchemas.MergeDelimitedFiles());

            // ---- 数据分析与可视化 ----
            Register(registry, "analyze_range_statistics", "对数据区域做数值/文本列的统计摘要（均值、中位数、标准差等）。返回附 basis 计算依据（来源区域/明细行区间/统计口径/公式示例）。", RiskLevel.ReadOnly,
                () => analysis.AnalyzeRangeStatistics, ThirdPhaseSchemas.RangeStatistics());
            Register(registry, "analyze_column_correlation", "计算数值列两两之间的相关系数（皮尔逊或秩相关）。成对完整观测（两列同一行均为数值），返回附 basis 计算依据（方法/口径/CORREL 公式示例/成对样本数 n）。", RiskLevel.ReadOnly,
                () => analysis.CorrelationAnalysis, ThirdPhaseSchemas.Correlation());
            Register(registry, "analyze_value_distribution", "统计某列取值分布与出现次数最多的前 N 个值。返回附 basis 计算依据（计数区间/口径/COUNTIF 公式示例）。", RiskLevel.ReadOnly,
                () => analysis.ColumnValueDistribution, ThirdPhaseSchemas.ValueDistribution());
            Register(registry, "analyze_column_profile", "列质量画像（Power Query 式单面板）：一次给出该列的质量（空值率/错误率/唯一值/类型占比）、分布（Top N 高频值）与剖析（数值列 min/max/均值/中位数/标准差；文本列长度与空白统计），并附质量提示（高空值/含错误/混合类型/带空格/恒定值）。数据清洗与选型前建议先对目标列跑一次。返回附 basis 计算依据（来源列区间/明细行区间/质量与统计口径）。", RiskLevel.ReadOnly,
                () => analysis.ColumnProfile, ThirdPhaseSchemas.ColumnProfile());
            Register(registry, "analyze_trend_forecast", "轻量趋势预测（只读）：对时间/序列标签列+数值列做线性回归或移动平均外推，返回未来 N 期预测值、趋势方向、斜率与拟合优度 r2，附 basis 计算依据（样本区间/SLOPE/INTERCEPT/FORECAST.LINEAR 公式示例）。有效数值点不足 6 个时拒绝预测（INSUFFICIENT_POINTS）；季节性 Holt-Winters/ARIMA 属高阶预测请走 python-advanced-analysis 技能（statsmodels）。", RiskLevel.ReadOnly,
                () => analysis.TrendForecast, ThirdPhaseSchemas.TrendForecast());
            Register(registry, "analyze_hypothesis_test", "假设检验标准工具（只读）：四类常用检验——two_sample_t（Welch 双样本 t）/paired_t（配对 t）/anova（单因素方差分析 F）/chi_square（卡方独立性）。返回统计量、自由度、精确 p 值、显著性判断（可调 α）、效应量（Cohen's d/η²/Cramér's V）、分组摘要与前提提醒（正态性/方差齐性/期望频数），附 basis（T.TEST/F.DIST.RT/CHISQ.TEST 公式示例）。统计口径纯 C# 可复算；非参检验与复杂设计走 python-advanced-analysis 技能。", RiskLevel.ReadOnly,
                () => analysis.HypothesisTest, ThirdPhaseSchemas.HypothesisTest());
            Register(registry, "analyze_cluster", "聚类分群标准工具（只读）：K-Means++ 初始化 + Lloyd 迭代 + Z-score 标准化，k 在 2~6 内按平均轮廓系数自动选定（可显式指定 k），返回各簇画像（样本数/特征均值）与逐行簇分派。完整观测行参与聚类（缺失行剔除并计数）；分派明细超 1000 行时省略并提示。DBSCAN/层次聚类等走 python-advanced-analysis 技能。", RiskLevel.ReadOnly,
                () => analysis.ClusterAnalysis, ThirdPhaseSchemas.ClusterAnalysis());
            Register(registry, "analyze_regression", "多元线性回归标准工具（只读）：普通最小二乘（OLS）拟合因变量与一个或多个特征列，返回系数表（估计/标准误/t/p）、R²/调整 R²/F 整体检验、残差标准误/Durbin-Watson/VIF/Cook's D 诊断与逐行拟合值/残差明细（超 1000 行省略）。完整观测行参与回归（缺失行剔除并计数）；完全共线或常数列 fail-closed 并提示。统计口径与 statsmodels 对齐、纯 C# 可复算；非线性/逻辑回归/正则化走 python-advanced-analysis 技能。", RiskLevel.ReadOnly,
                () => analysis.RegressionAnalysis, ThirdPhaseSchemas.RegressionAnalysis());
            Register(registry, "analyze_text_variants", "文本列变体聚类建议标准工具（只读，探索式清洗）：对文本列（如客户名/供应商名/品类）输出疑似变体组——空白/大小写/全半角归一化差异与编辑距离≤阈值的错字/后缀差异各成一组，每组给代表值、变体成员（含出现次数与匹配类型）与合并方向选项（取最长/取最短/取出现最多）。本工具只建议不合并：用户确认后合并动作须用 write_excel_cells_batch 将变体单元格改写为代表值（改值语义，走确认门可回滚；不得用 dedup_range——其语义为删除重复行）。唯一值超 1500 或需进阶模糊匹配走 python-advanced-analysis 技能。", RiskLevel.ReadOnly,
                () => analysis.TextVariants, ThirdPhaseSchemas.TextVariants());
            Register(registry, "add_insight_card", "看板洞察卡标准工具：把分析结论生成为圆角矩形洞察卡（标题加粗+正文白字，五色主题可选），落在指定工作表的锚点单元格处，卡高按文字量自动估算（正文超 8 行截断省略）。写入后自动读回文本校验（verdict=ok）。卡片文字必须来自工具真实分析结论（禁止编造数值）；replace_existing=true 可清空该表旧洞察卡重写（防反复调用堆积）；卡片是展示形状，随删表移除、不参与单元格快照回滚。适合配合情感分析/假设检验/聚类画像的结论做看板可视化。", RiskLevel.NormalWrite,
                () => analysis.AddInsightCard, ThirdPhaseSchemas.AddInsightCard());
            // ---- 公式构造与验证闭环（FormulaBuilder，2026-09-18 批次） ----
            Register(registry, "build_formula", "公式构造器标准工具（写前首选）：按模板结构化生成验证过的公式（18 模板：sum/count/max/min/average/sumifs/countifs/averageifs/weighted_avg/distinct_count/rank/cumulative_sum/percent_of_total/median/stdev/growth_rate/diff/ratio——rank/cumulative/percent/同环比为逐行混合引用形态，插入目标首格后向下填充联动）。内化表头行对齐（header_rows，默认 1，数据区自动从表头后开始）与引用卫生（同表裸地址/跨表规范限定）。写入前建议先用 eval_formula 试算。", RiskLevel.ReadOnly,
                () => formulaTools.BuildFormula, FormulaSchemas.BuildFormula());
            Register(registry, "eval_formula", "公式试算沙箱（写前验证）：用 Application.Evaluate 直接求值，不写单元格。返回真实结果或错误值映射（#NAME?/#REF!/#VALUE!…）。AI 生成公式后、写入前必须先试算——写错当场暴露。", RiskLevel.ReadOnly,
                () => formulaTools.EvalFormula, FormulaSchemas.EvalFormula());
            Register(registry, "diagnose_formula", "公式诊断标准工具（只读）：对区域（≤200 格）公式做静态体检五类根因——错误值(#NAME?/#REF!…)/悬空工作表引用/外部工作簿引用/自引用循环/公式内 #REF! 残码，逐格给人话修复建议。", RiskLevel.ReadOnly,
                () => formulaTools.DiagnoseFormula, FormulaSchemas.DiagnoseFormula());
            Register(registry, "rewrite_formulas", "公式批改器（NormalWrite，≤500 格护栏）：区域公式批量 to_absolute（引用转绝对）/ to_relative（转相对）/ wrap_iferror（整条包 IFERROR）。非公式格自动跳过。", RiskLevel.NormalWrite,
                () => formulaTools.RewriteFormulas, FormulaSchemas.RewriteFormulas());
            Register(registry, "read_business_glossary", "读取当前活动工作簿同目录的《_口径.md》业务口径规则文件（Markdown，每行一条「- 术语: 定义」，随业务文件走）。生成看板/分析结论/统计摘要前建议先调用本工具；回复中涉及口径的关键数字必须逐条引用条目原文；口径与用户当次指令冲突时用户指令优先，但必须向用户明示差异，不得静默择一。找不到文件时返回创建模板。隐私红线：文件仅本地读取，绝不外发。", RiskLevel.ReadOnly,
                () => analysis.ReadBusinessGlossary, ThirdPhaseSchemas.GlossaryRead());
            Register(registry, "recommend_excel_chart", "根据数据形状推荐合适的图表类型与坐标列。", RiskLevel.ReadOnly,
                () => analysis.RecommendChart, ThirdPhaseSchemas.RecommendChart());
            Register(registry, "create_excel_sparkline", "在目标单元格插入迷你图（折线/柱状/胜负）。", RiskLevel.NormalWrite,
                () => chart.CreateSparkline, ThirdPhaseSchemas.Sparkline());
            Register(registry, "export_excel_chart_image", "把图表导出为 PNG 图片文件。", RiskLevel.NormalWrite,
                () => chart.ExportChartAsImage, ThirdPhaseSchemas.ExportChartImage());
            Register(registry, "verify_chart", "图表自检：校验图表是否真正生成、数据系列/点数是否有效、数据源区域（来自 series 公式）与预期类型/标题/数据源/系列数是否相符，并对渲染图做像素级核验“数据是否真的显示出来”，返回渲染图片供核对。生成任何图表后都应调用一次。", RiskLevel.ReadOnly,
                () => chart.VerifyChart, ThirdPhaseSchemas.VerifyChart());
            var pivotChart = new ExcelPivotChartService(application, null, null);
            Register(registry, "diagnose_chart", "图表读回诊断（只读）：无期望全面体检——类型/系列/点数/空系列/数据源/标题/渲染像素核验/数据守恒对账，输出根因归类（root_cause）、诊断依据（findings）、修复建议（suggestions）与渲染截图（chart_image 供 AI 读图核验），诊断全程记录。生成图表后自检或排查图表问题时调用。", RiskLevel.ReadOnly,
                () => chart.DiagnoseChart, ThirdPhaseSchemas.ChartDiagnose());
            Register(registry, "repair_chart", "图表读回修复：按诊断结论修复——空系列/点数错乱重绑源数据、117~123 原生新图临时表重建、补标题；修复后自动复检并返回前后对比与截图（AI 读图）。与 diagnose_chart 配套使用。", RiskLevel.StructuralChange,
                () => (Func<Dictionary<string, object>, ToolResult>)(a => chart.RepairChart(a, pivotChart)), ThirdPhaseSchemas.ChartRepair());

            // ---- P4-C 数据分析与报告生成 ----
            Register(registry, "insert_recommended_chart", "按数据形状一键插入推荐图表（自动选型/坐标列/中文标题）。", RiskLevel.StructuralChange,
                () => analysis.InsertRecommendedChart, PhaseCSchemas.InsertRecommendedChart());
            Register(registry, "generate_analysis_report", "对数据区域生成模板化文字报告，可输出到工作表、Markdown 或 HTML。", RiskLevel.NormalWrite,
                () => report.GenerateReport, PhaseCSchemas.GenerateAnalysisReport());
        }

        private static void Register(
            ToolRegistry registry,
            string name,
            string description,
            RiskLevel risk,
            Func<Func<Dictionary<string, object>, ToolResult>> target,
            Dictionary<string, object> parameters)
        {
            var execute = target();
            registry.Register(new ToolDescriptor(name, description, risk, args => Execute(() => execute(args)), parameters));
        }

        private static ToolResult Execute(Func<ToolResult> action)
        {
            return ExcelToolHost.Execute(action);
        }
    }
}
