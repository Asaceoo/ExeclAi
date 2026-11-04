using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// M-D1：原生数据看板服务（P1 静态版）。
    /// kpi_brief 模板：隐藏 SUMIFS 辅助区 + KPI 卡片（三件套：数值 + 环比箭头 + IFERROR）。
    /// 设计约束（docs/设计-原生数据看板.md）：
    /// - 图表/卡片永不直连明细数据，全部引用辅助汇总区；
    /// - staging 公式限定实际数据区间（探测 max row），禁止整列引用（FR-3.4）；
    /// - 比率类 KPI 一律 IFERROR 包裹，由渲染引擎强制注入（§5.4）；
    /// - 渲染幂等：重生成 = 删旧 sheet 重建；失败 fail-closed 返回可读错误。
    /// </summary>
    public sealed partial class ExcelDashboardService
    {
        private readonly dynamic application;
        private readonly OperationRecorder recorder;

        public ExcelDashboardService(object application)
            : this(application, null)
        {
        }

        public ExcelDashboardService(object application, OperationRecorder recorder)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.recorder = recorder;
        }

        /// <summary>
        /// <summary>
        /// M-D2：render_dashboard 之 sales_overview 模板（视频同款）。
        /// 一次调用：控制单元格（地区下拉）+ KPI 卡 + 月度双轴趋势 + 品类环形。
        /// staging 全部公式引用控制单元格，切换自动重算（FR-2）；源数据只读；幂等；失败回滚。
        /// </summary>
        public ToolResult RenderSalesOverview(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 1);
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 2);
            var monthColumn = ExcelToolHostHelpers.GetIntSafe(args, "month_column", 3);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var quantityColumn = ExcelToolHostHelpers.GetIntSafe(args, "quantity_column", 5);
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "销售看板";
            var theme = (Optional(args, "theme") ?? "orange").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }

            var colParams = new[] { labelColumn, categoryColumn, monthColumn, valueColumn, quantityColumn };
            if (colParams.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic rangeSource = null;
            dynamic sheetOld = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeTmp = null;
            dynamic validation = null;
            object bulkGuard = null;
            ToolResult boardResult = null;
            string dataSheetName = null;
            int monthCount = 0, catCount = 0, labelCount = 0;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // FR-1.5 多看板共存：auto_rename=true 时重名自动「名 2/名 3」递增，实现共存；
                // 默认 false 保持幂等重生成语义（同名=覆盖旧看板）
                var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                // FR-6：支持显式 sheet 参数——渲染会激活看板表，重生成/连续生成时 ActiveSheet 不可靠
                // （真机 T11b 实证：不传 sheet 时把看板表当数据源，months=0/INVALID_FIELD）
                var sourceSheetName = Optional(args, "sheet");
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                rangeSource = sheetSource.Range[sourceAddress];

                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                // 块读取用绝对坐标——源区域锚不在 A 列时必须加首列偏移（测试2026.xlsx B 列锚实证）
                int srcFirstCol = Convert.ToInt32(rangeSource.Column);
                if (colParams.Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"列参数超界：区域列数 {srcCols}，最大传入 {colParams.Max()}。");
                }

                dataSheetName = Convert.ToString(sheetSource.Name);
                var addr = Convert.ToString(rangeSource.Address, CultureInfo.InvariantCulture);
                var bounds = ColumnLetters(addr, labelColumn, valueColumn, out int headerRow, out int lastRowAbs);
                string labelColL = bounds[0];
                string escSheet = EscapeSheet(dataSheetName);

                // 幂等删除旧看板/辅助表
                foreach (var name in new[] { dashboardSheet, "_" + dashboardSheet + "_calc" })
                {
                    if (SheetExists(workbook, name))
                    {
                        sheetOld = workbook.Worksheets.Item[name];
                        try { sheetOld.Delete(); } catch { }
                        sheetOld = null;
                    }
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add();
                sheetStaging.Name = "_" + dashboardSheet + "_calc";
                // M-D3 优化轮对照实验结论：守卫在 2000 行（公式 ~30 条）下因 COM 模式切换
                // 开销反而更慢（6.5s→10.4s），在 2 万行（公式 30+ 条 × 全簿重算）下快 37%。
                // 最终策略=行数阈值触发：≥5000 行启用批量守卫，小表走原生增量重算。
                bulkGuard = (lastRowAbs - headerRow) >= 5000 ? BeginBulkFormulaWrite() : null;

                string valColL = ColumnLetters(addr, labelColumn, valueColumn, out int hr1, out int lr1)[1];
                string qtyColL = ColumnLetters(addr, labelColumn, quantityColumn, out int hr2, out int lr2)[1];
                string catColL = ColumnLetters(addr, labelColumn, categoryColumn, out int hr3, out int lr3)[1];
                string monthColL = ColumnLetters(addr, labelColumn, monthColumn, out int hr4, out int lr4)[1];
                _ = hr1; _ = lr1; _ = hr2; _ = lr2; _ = hr3; _ = lr3; _ = hr4; _ = lr4;
                string dRef = $"{escSheet}!${valColL}${headerRow + 1}:${valColL}${lastRowAbs}";
                string qRef = $"{escSheet}!${qtyColL}${headerRow + 1}:${qtyColL}${lastRowAbs}";
                string lRef = $"{escSheet}!${labelColL}${headerRow + 1}:${labelColL}${lastRowAbs}";
                string cRef = $"{escSheet}!${catColL}${headerRow + 1}:${catColL}${lastRowAbs}";
                string mRef = $"{escSheet}!${monthColL}${headerRow + 1}:${monthColL}${lastRowAbs}";
                string ctrlRef = "'" + EscapeSheetInner(dashboardSheet) + "'!$B$2";
                string regionCond = $"IF({ctrlRef}=\"全部\",\"*\",{ctrlRef})";

                // ---- staging 维度去重集合：一次块读取（禁止逐格 COM 扫描，NFR-2 关键路径）----
                int minCol = Math.Min(Math.Min(labelColumn, categoryColumn), monthColumn);
                int maxCol = Math.Max(Math.Max(labelColumn, categoryColumn), monthColumn);
                int dimRowCount = lastRowAbs - headerRow;
                dynamic dimRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRow + 1, srcFirstCol + minCol - 1],
                    sheetSource.Cells.Item[lastRowAbs, srcFirstCol + maxCol - 1]];
                object[,] dimBlock = null;
                try
                {
                    dimBlock = dimRange.Value2 as object[,];
                }
                finally
                {
                    Release((object)dimRange);
                }
                // 工具分类审查批次防御：数据区仅 1 行 1 列（且三维度列同指一列）时 Value2 是标量非二维数组——
                // dimBlock[r, c] 直接索引会抛异常裸冒泡。包一层统一转 0 基二维。
                if (dimBlock == null)
                {
                    var scalar = sheetSource.Range[
                        sheetSource.Cells.Item[headerRow + 1, srcFirstCol + minCol - 1],
                        sheetSource.Cells.Item[lastRowAbs, srcFirstCol + maxCol - 1]].Value2;
                    dimBlock = new object[1, 1];
                    dimBlock[0, 0] = scalar is object[,] one ? one[1, 1] : scalar;
                }
                var monthSet = new SortedSet<string>();
                var catSet = new SortedSet<string>();
                var labelSet = new SortedSet<string>();
                // 0 基适配：下方 dimBlock[r, c] 原按 Value2 的 1 基下标读取，统一改为 0 基
                for (int r = 0; r < dimRowCount; r++)
                {
                    object lv = dimBlock[r, labelColumn - minCol];
                    if (lv != null) { labelSet.Add(Convert.ToString(lv)); }
                    object cv = dimBlock[r, categoryColumn - minCol];
                    if (cv != null) { catSet.Add(Convert.ToString(cv)); }
                    object mv = dimBlock[r, monthColumn - minCol];
                    if (mv != null) { monthSet.Add(Convert.ToString(mv)); }
                }

                // ---- staging：月度矩阵 C:E（SUMIFS 引用控制单元格，切换自动重算）----

                rangeTmp = sheetStaging.Range["C1"]; rangeTmp.Value2 = "月份"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["D1"]; rangeTmp.Value2 = "销售额"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["E1"]; rangeTmp.Value2 = "销量"; Release((object)rangeTmp);
                int mi = 2;
                foreach (var month in monthSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[mi, 3];
                    // 2026-09-12 用户文件实证：明细日期列为文本（"2025-11-24"）时，月份标签写入会被
                    // Excel 自动转成日期序列号 → 改写公式数值比较 ≠ 文本明细 → 月度矩阵全 0（趋势图空）。
                    // 文本标签先锁 NumberFormat=@ 保持与明细同型；数值标签（真日期）走数值比较不受影响。
                    if (month is string)
                    {
                        rangeTmp.NumberFormat = "@";
                    }
                    rangeTmp.Value2 = month;
                    Release((object)rangeTmp);
                    dynamic c1 = sheetStaging.Cells.Item[mi, 4];
                    c1.Formula = $"=SUMIFS({dRef},{mRef},C{mi},{lRef},{regionCond})";
                    Release((object)c1);
                    dynamic c2 = sheetStaging.Cells.Item[mi, 5];
                    c2.Formula = $"=SUMIFS({qRef},{mRef},C{mi},{lRef},{regionCond})";
                    Release((object)c2);
                    mi++;
                }
                int monthCount_ = monthSet.Count; monthCount = monthCount_;

                // ---- staging：品类构成 G:H（catSet 已由块读取收集）----

                rangeTmp = sheetStaging.Range["G1"]; rangeTmp.Value2 = "品类"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["H1"]; rangeTmp.Value2 = "销售额"; Release((object)rangeTmp);
                int ci = 2;
                foreach (var cat in catSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[ci, 7];
                    rangeTmp.Value2 = cat;
                    Release((object)rangeTmp);
                    dynamic c3 = sheetStaging.Cells.Item[ci, 8];
                    c3.Formula = $"=SUMIFS({dRef},{cRef},G{ci},{lRef},{regionCond})";
                    Release((object)c3);
                    ci++;
                }
                int catCount_ = catSet.Count; catCount = catCount_;

                // ---- staging：地区列表 J 列（控制单元格 B2 下拉数据源，首项「全部」；labelSet 已由块读取收集）----
                rangeTmp = sheetStaging.Range["J1"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                int li = 2;
                foreach (var label in labelSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[li, 10];
                    rangeTmp.Value2 = label;
                    Release((object)rangeTmp);
                    li++;
                }
                int labelCount_ = labelSet.Count; labelCount = labelCount_;
                // FR-6 防呆：维度为空说明布局与五列口径不符（如分块汇总表）——明确报错，绝不静默出错误看板
                if (monthCount == 0 || catCount == 0)
                {
                    return Invalid("NO_DIMENSION_DATA",
                        $"月份列有效值 {monthCount} 个、类别列有效值 {catCount} 个，与销售概览要求的「地区|类别|月份|数量」长表明细不符。" +
                        "请确认数据源是一行一条的明细记录（可先调 suggest_chart_mapping 分析布局，或让 AI 把分块汇总拼成长表）。");
                }

                // ---- staging：KPI 值 A 列（联动；比率类 IFERROR 强制）----
                var kpiFormulas = new Dictionary<string, string>
                {
                    ["销售总额"] = $"=SUMIFS({dRef},{lRef},{regionCond})",
                    ["成交笔数"] = $"=COUNTIFS({lRef},{regionCond})",
                    ["平均客单价"] = $"=IFERROR(AVERAGEIFS({dRef},{lRef},{regionCond}),0)",
                    // 最大单笔：MAX(IF()) 数组式在部分工作簿上下文重算为 #VALUE!（测试2026 实证）——
                    // 改 SUMPRODUCT(MAX(...)) 免 CSE 惯用法；「全部」分支用 1 兜底（"*"通配在数组等比较中不生效）
                    ["最大单笔"] = $"=SUMPRODUCT(MAX(IF({ctrlRef}=\"全部\",1,--({lRef}={ctrlRef}))*{dRef}))"
                };
                int ki = 1;
                foreach (var kv in kpiFormulas)
                {
                    // 2026-09-12 A1 坏缓存规避：KPI 公式写 B 列（A1 地址格在 Excel 求值缓存中存在
                    // "同名 sheet 旧 A1 坏缓存"实证，写新地址必然正确求值——md1 probe4 A5=2742 实证）
                    rangeTmp = sheetStaging.Cells.Item[ki, 2];
                    rangeTmp.Formula = kv.Value;
                    Release((object)rangeTmp);
                    ki++;
                }

                RenderSalesOverviewBoard(
                    dashboardSheet, sheetBoard, sheetStaging, themeColors,
                    monthCount, catCount, labelCount, args);
            }
            catch (InvalidOperationException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED", ex.Message);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                var hint = ex.HResult == unchecked((int)0x800A03EC)
                    ? "工作簿可能处于兼容模式，请另存为 .xlsx 后重试。"
                    : "已回滚，源数据未受影响。";
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "sales_overview 渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。" + hint);
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "sales_overview 渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                // M-D3 优化轮：恢复计算模式并只重算看板/辅助两表（省全簿扫描；
                // SUMIFS 依赖链指向明细表，Excel sheet 级 Calculate 会自动带上依赖）。
                // 异常路径也必须恢复，防遗留手动重算态。
                EndBulkFormulaWrite(bulkGuard, dashboardSheet, "_" + dashboardSheet + "_calc");
                Release((object)validation);
                Release((object)rangeTmp);
                Release((object)rangeSource);
                Release((object)sheetSource);
            }

            // FR-6 渲染增强：此刻批量重算已恢复，读取的是真实计算值。
            // a) auto_slicer（默认 true）：自动挂 native 切片器，失败回退渲染自带的 cell 下拉；
            // b) 渲染后即时自检，结论并入返回值回显聊天。
            if (boardResult == null)
            {
                bool autoSlicerOn = !string.Equals(Optional(args, "auto_slicer"), "false", StringComparison.OrdinalIgnoreCase);
                var selfcheck = RunPostRenderSelfCheck(dashboardSheet);
                (bool slicerOk, string slicerField, string slicerCache, int slicerRewritten, string slicerReason) slicer;
                if (autoSlicerOn)
                {
                    slicer = AttachAutoSlicerIfPossible(args, dashboardSheet, dataSheetName);
                }
                else
                {
                    slicer = (false, null, null, 0, "auto_slicer=false");
                }
                string selfcheckText = (selfcheck.Verdict == "ok" ? "自检通过 ✓ " : selfcheck.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheck.Detail;
                string linkageText = slicer.slicerOk
                    ? "native 切片器（" + slicer.slicerField + "）：点击选项联动 KPI 与图表"
                    : "看板!B2 下拉（(全部) 为全量视图），切换自动重算" + (autoSlicerOn ? "（切片器不可用已回退）" : "");
                recorder?.Invoke("render_dashboard", args, "dashboard-render",
                    new
                    {
                        template = "sales_overview",
                        sheet = dashboardSheet,
                        months = monthCount,
                        categories = catCount,
                        slicer = slicer.slicerOk ? slicer.slicerField : "cell",
                        selfcheck = selfcheck.Verdict
                    },
                    new { Sheet = dashboardSheet, Template = "sales_overview" });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "sales_overview",
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = "_" + dashboardSheet + "_calc",
                        kpi_count = 4,
                        months = monthCount,
                        categories = catCount,
                        linkage = linkageText,
                        slicer_mode = slicer.slicerOk ? "native" : (autoSlicerOn ? "cell_fallback" : "cell"),
                        slicer_field = slicer.slicerField,
                        slicer_cache = slicer.slicerCache,
                        slicer_reason = slicer.slicerOk ? null : slicer.slicerReason,
                        selfcheck = selfcheckText,
                        selfcheck_verdict = selfcheck.Verdict,
                        theme = theme,
                        created = true
                    }
                };
            }
            return boardResult;
        }

        /// <summary>sales_overview 后半段：控制单元格/数据验证 + KPI 卡 + 双轴/环形图表 + 布局（构建失败以异常冒泡，由外层统一回滚与报错）。</summary>
        private void RenderSalesOverviewBoard(
            string dashboardSheet, dynamic sheetBoard, dynamic sheetStaging,
            ThemePalette themeColors, int monthCount, int catCount, int labelCount,
            Dictionary<string, object> args)
        {
            dynamic rangeTmp = null;
            dynamic validation = null;
            try
            {
                // ---- 控制单元格 + 数据验证下拉（FR-2 联动源）----
                rangeTmp = sheetBoard.Range["A2"];
                rangeTmp.Value2 = "地区：";
                Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["B2"];
                rangeTmp.Value2 = "全部";
                Release((object)rangeTmp);
                validation = sheetBoard.Range["B2"].Validation;
                validation.Delete();
                validation.Add(Type: 3, Formula1: $"='_{EscapeSheetInner(dashboardSheet)}_calc'!$J$1:$J${labelCount + 1}");
                Release((object)validation);
                validation = null;

                // ---- KPI 卡（复用三件套布局，数值引用 staging A 列）----
                // 2026-09-12 截图视觉实证：22 号大字体的 11 位金额（¥352,949,424）在默认列宽下显示 ####
                sheetBoard.Range["B:D"].ColumnWidth = 12;
                var kpiDefs = DefaultKpis();
                string[] kpiTitles = { "销售总额", "成交笔数", "平均客单价", "最大单笔" };
                for (int i = 0; i < kpiDefs.Count; i++)
                {
                    int r1 = 4 + i * 4;
                    dynamic titleCell = sheetBoard.Range[sheetBoard.Cells.Item[r1, 2], sheetBoard.Cells.Item[r1, 4]];
                    titleCell.Merge();
                    titleCell.Value2 = kpiTitles[i];
                    titleCell.Font.Size = 11;
                    titleCell.Font.Bold = true;
                    titleCell.Font.Color = themeColors.TitleText;
                    Release((object)titleCell);

                    dynamic valueCell = sheetBoard.Range[sheetBoard.Cells.Item[r1 + 1, 2], sheetBoard.Cells.Item[r1 + 1, 4]];
                    valueCell.Merge();
                    // M-D3 审查⑨（自引用缺陷修复）：KPI 卡必须引用隐藏辅助表的 KPI 值
                    // （原实现误引看板 sheet 自身 A 列——看板 A 列无值，IFERROR 恒回退"-"，
                    //   KPI 卡实际显示占位符而非数值；断言此前只读 staging 未覆盖此缺陷）
                    valueCell.Formula = $"=IFERROR('{EscapeSheetInner("_" + dashboardSheet + "_calc")}'!$B${i + 1},\"-\")";  // 2026-09-12 KPI 值迁 B 列（A1 坏缓存规避）+ 2026-09-11 前导下划线修复
                    try { valueCell.NumberFormat = kpiDefs[i].Format; }
                    catch { valueCell.NumberFormat = "#,##0"; }
                    valueCell.Font.Size = 22;
                    valueCell.Font.Bold = true;
                    valueCell.Font.Color = themeColors.ValueText;
                    valueCell.Interior.Color = themeColors.CardBg;
                    valueCell.HorizontalAlignment = -4108;
                    Release((object)valueCell);
                }

                // ---- 图表：双轴趋势（staging C:E）+ 品类环形（G:H）----
                dynamic trendObj = null;
                dynamic trendChart = null;
                dynamic trendSeries = null;
                dynamic doughObj = null;
                dynamic doughChart = null;
                try
                {
                    trendObj = sheetBoard.ChartObjects().Add(320, 180, 480, 280);
                    trendChart = trendObj.Chart;
                    trendChart.ChartType = 51; // xlColumnClustered
                    trendChart.SetSourceData(sheetStaging.Range["C1:E" + (monthCount + 1)]);
                    trendSeries = trendChart.SeriesCollection();
                    if (Convert.ToInt32(trendSeries.Count) >= 2)
                    {
                        dynamic s2 = trendSeries.Item(2);
                        s2.AxisGroup = 2;  // xlSecondary：销量挂次轴
                        s2.ChartType = 4;  // xlLine
                        Release((object)s2);
                    }
                    trendChart.HasTitle = true;
                    trendChart.ChartTitle.Text = "月度销售趋势";
                }
                finally
                {
                    Release((object)trendSeries);
                    Release((object)trendChart);
                    Release((object)trendObj);
                }

                try
                {
                    doughObj = sheetBoard.ChartObjects().Add(840, 180, 360, 280);
                    doughChart = doughObj.Chart;
                    doughChart.ChartType = -4120; // xlDoughnut
                    doughChart.SetSourceData(sheetStaging.Range["G1:H" + (catCount + 1)]);
                    try { doughChart.ChartGroups().Item(1).DoughnutHoleSize = 45; } catch { }
                    doughChart.HasTitle = true;
                    doughChart.ChartTitle.Text = "品类销售构成";
                }
                finally
                {
                    Release((object)doughChart);
                    Release((object)doughObj);
                }

                // 布局收尾：关网格线 + 隐藏辅助表（recorder 与返回值由外层 RenderSalesOverview 统一处理）
                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }
                sheetStaging.Visible = 0;
            }
            finally
            {
                Release((object)validation);
                Release((object)rangeTmp);
            }
        }

        /// <summary>
        /// M-D3：set_linkage 看板联动设置。
        /// - mode=native（默认，仅 Excel 2013+）：基于明细数据在隐藏 staging sheet 上建 PivotTable，
        ///   SlicerCaches.Add2 创建切片器置于看板 sheet，并把既有透视表挂接同一 cache；
        ///   切片器选择变化 → staging SUMIFS 公式引用的控制单元格自动重算（FR-2.1/2.2）。
        /// - mode=cell：数据验证下拉写入控制单元格（Excel 2013 起全兼容，FR-2.3 兜底）。
        /// 幂等：已存在同名 SlicerCache/PivotTable 先清理再重建；失败回滚不残留半成品。
        /// </summary>
        public ToolResult SetLinkage(Dictionary<string, object> args)
        {
            var mode = (Optional(args, "mode") ?? "native").Trim().ToLowerInvariant();
            if (mode != "native" && mode != "cell")
            {
                return Invalid("INVALID_MODE", $"mode 仅支持 native/cell，当前为 {mode}。");
            }
            var sourceAddress = Required(args, "source_address");
            // 2026-09-12 多切片器联动：field_column 支持 "4,5"（逗号分隔多维度）
            // 工具分类审查批次修复：裸 int.Parse 遇 "a,b"/"4,,5"/空串抛 FormatException 裸异常冒泡
            // （报错人话性差）——改 TryParse 逐段解析，任何一段非法即回人话 INVALID_COLUMN。
            var fieldColumnRaw = Optional(args, "field_column")
                ?? Convert.ToString(ExcelToolHostHelpers.GetIntSafe(args, "field_column", 1));
            var fieldColumn = new List<int>();
            foreach (var seg in fieldColumnRaw.Split(','))
            {
                if (!int.TryParse(seg.Trim(), out var parsedSeg) || parsedSeg < 1)
                {
                    return Invalid("INVALID_COLUMN",
                        $"field_column 非法：「{fieldColumnRaw}」。应为单列号（如 4）或逗号分隔多列号（如 4,5），每段须为 ≥1 的整数。");
                }
                fieldColumn.Add(parsedSeg);
            }
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "销售看板";
            var sourceSheetName = Optional(args, "sheet");
            var monthColumn = ExcelToolHostHelpers.GetIntSafe(args, "month_column", 3);
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 2);
            var quantityColumn = ExcelToolHostHelpers.GetIntSafe(args, "quantity_column", 5);

            var cols = fieldColumn.Concat(new List<int> { valueColumn, monthColumn, categoryColumn, quantityColumn });
            if (cols.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic pivotCache = null;
            dynamic pivotTable = null;
            dynamic slicerCache = null;
            dynamic slicer = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // 源数据表：显式指定优先；缺省用活动表（注意 render 类工具会激活新表，调用方应显式传 sheet）
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                sheetBoard = ResolveSheetForLinkage(workbook, dashboardSheet);
                if (sheetBoard == null)
                {
                    return Invalid("DASHBOARD_NOT_FOUND",
                        $"找不到看板「{dashboardSheet}」。请先通过 render_dashboard 生成看板，再设置联动。");
                }
                sheetStaging = workbook.Worksheets.Item["_" + EscapeSheetInner(dashboardSheet) + "_calc"];

                // 源数据范围校验（列参数相对 source_address 首列，与 render_dashboard 口径一致）
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                int firstColAbs = Convert.ToInt32(rangeSource.Column);
                int headerRowAbs = Convert.ToInt32(rangeSource.Row);
                int lastRowAbs = headerRowAbs + Convert.ToInt32(rangeSource.Rows.Count) - 1;
                Release((object)rangeSource);
                var allCols = fieldColumn.Concat(new List<int> { valueColumn, monthColumn, categoryColumn, quantityColumn });
                if (allCols.Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"列参数超界：区域列数 {srcCols}，最大传入 {fieldColumn.Concat(new[] { valueColumn, monthColumn, categoryColumn, quantityColumn }).Max()}。");
                }
                var fieldColsAbs = fieldColumn.Select(fc => firstColAbs + fc - 1).ToList();
                int fieldColAbs = fieldColsAbs[0];
                int valueColAbs = firstColAbs + valueColumn - 1;
                string dataSheetName = Convert.ToString(sheetSource.Name);
                string dataRange = $"{EscapeSheet(dataSheetName)}!{sourceAddress}";

                if (mode == "native")
                {
                    // 幂等清理：同名 SlicerCache 与 staging 透视表先删后建
                    CleanupExistingLinkage(workbook, dashboardSheet);

                    // 探测结论（M-D0 P2）：新式 Create + Add2 链路；未存盘工作簿会兼容模式禁切片器
                    pivotCache = workbook.PivotCaches().Create(SourceType: 1, SourceData: dataRange);
                    string pivotName = "_linkage_pivot_" + EscapeSheetInner(dashboardSheet);
                    pivotTable = sheetStaging.PivotTables().Add(
                        PivotCache: pivotCache, TableDestination: sheetStaging.Range["Z1"]);
                    pivotTable.Name = pivotName;
                    // 2026-09-12 多切片器联动（方案 B）：每维度独立透视表+独立切片器。
                    // 放弃同透视多行字段+LayoutForm 表格布局——dynamic binder 下 LayoutForm
                    // 方法/属性均不可达（F1 真机两轮实证）；独立透视各自单字段布局稳定，
                    // 跨切片器筛选语义由多维掩码 AND（RewriteMulti 的 pTagRefs 每透视一个标签区）保证。
                    var fieldNames = new List<string>();
                    var linkagePivotNames = new List<string>();
                    foreach (var fc in fieldColumn)
                    {
                        int colAbs = firstColAbs + fc - 1;
                        string fName = Convert.ToString(sheetSource.Cells.Item[headerRowAbs, colAbs].Value2);
                        if (string.IsNullOrWhiteSpace(fName))
                        {
                            // M-D3 审查修正⑤：透视表已创建，提前返回必须回滚（AC-5 无半成品）
                            TryRollbackLinkage(workbook, dashboardSheet);
                            return Invalid("INVALID_FIELD", $"第 {fc} 列表头为空，无法作为筛选维度。已回滚。");
                        }
                        fieldNames.Add(fName);
                    }

                    double slicerLeft0 = Convert.ToDouble(sheetBoard.Range["F1"].Left);
                    var slicerInfos = new List<string>();
                    for (int si = 0; si < fieldColumn.Count; si++)
                    {
                        int colAbs = firstColAbs + fieldColumn[si] - 1;
                        string fName = fieldNames[si];
                        string pivotNameSi = "_linkage_pivot_" + EscapeSheetInner(dashboardSheet) + "_" + si;
                        // 幂等：先清扫 staging 上本看板的旧联动透视（含旧版无 _si 后缀的 Z1 遗留，
                        // md1 实证 _linkage_pivot_销售看板B 残留 $Z$1:$AB$18），防位置占用与重复缓存
                        try
                        {
                            var legacyNames = new List<string>();
                            dynamic ptsAll = sheetStaging.PivotTables();
                            for (int pi2 = 1; pi2 <= Convert.ToInt32(ptsAll.Count); pi2++)
                            {
                                string pn2 = Convert.ToString(ptsAll.Item(pi2).Name);
                                if (pn2.StartsWith("_linkage_pivot_", StringComparison.Ordinal)) { legacyNames.Add(pn2); }
                            }
                            Release((object)ptsAll);
                            foreach (var pn2 in legacyNames)
                            {
                                try { sheetStaging.PivotTables(pn2).Delete(); } catch { }
                            }
                        }
                        catch { }
                        // 位置放第 200 列起（远离 Z 区历史残留；透视表在隐藏 staging 深处，不影响任何视图）
                        dynamic pivotSi = sheetStaging.PivotTables().Add(
                            PivotCache: pivotCache, TableDestination: sheetStaging.Cells.Item[1, 200 + si * 25]);
                        pivotSi.Name = pivotNameSi;
                        pivotSi.PivotFields(fName).Orientation = 1; // xlRowField（单字段透视，压缩式=表格式无差异）
                        linkagePivotNames.Add(pivotNameSi);

                        dynamic sc = workbook.SlicerCaches.Add2(pivotSi, fName);
                        dynamic sl = sc.Slicers().Add(
                            SlicerDestination: sheetBoard, Top: 6,
                            Left: slicerLeft0 + si * 154, Width: 144, Height: 192);
                        sl.Caption = fName;
                        slicerInfos.Add(fName + ":" + Convert.ToString(sc.Name));
                        if (si == 0) { slicerCache = sc; slicer = sl; pivotTable = pivotSi; }
                        else { Release((object)sl); Release((object)sc); Release((object)pivotSi); }
                    }

                    recorder?.Invoke("set_linkage", args, "linkage-set",
                        new { mode = "native", field = string.Join(",", fieldNames), sheet = dashboardSheet, slicers = slicerInfos.Count },
                        new { Sheet = dashboardSheet, Mode = "native", Fields = string.Join(",", fieldNames) });

                    // native 联动闭环：staging 公式从「引用 B2 控制单元格」切换为
                    // 「SUMPRODUCT 引用透视表可见标签区」——切片器选择变化 → 透视行区变化 → 公式重算。
                    // 透视行区含「总计」行（标签无明细匹配，SUMIF 贡献 0，语义自洽），无需排除。
                    int kpiRewritten = RewriteStagingFormulasForNativeMulti(
                        workbook, dashboardSheet, linkagePivotNames,
                        fieldColsAbs, valueColAbs,
                        firstColAbs + monthColumn - 1, firstColAbs + categoryColumn - 1,
                        firstColAbs + quantityColumn - 1,
                        dataSheetName, lastRowAbs - headerRowAbs, headerRowAbs + 1);

                    // 2026-09-12 真机实证（md1 T11c 判定实验）：改写公式写入时透视尚未完成首次求值，
                    // Excel 将 KPI 缓存为 0 且 CalculateFull 不清该缓存（重写公式才触发重算，probe3 实证）。
                    // CalculateFullRebuild 重建依赖树并全算——根治"切片器挂载后 KPI 恒 0"。
                    // CalculateFullRebuild 二分排查：暂时禁用（T11c 二分）
                    // try { application.CalculateFullRebuild(); } catch { try { application.CalculateFull(); } catch { } }

                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            mode = "native",
                            dashboard_sheet = dashboardSheet,
                            field = string.Join(",", fieldNames),
                            slicer_cache = Convert.ToString(slicerCache.Name),
                            staging_formulas_rewritten = kpiRewritten,
                            placement = "看板 sheet 左上（Top=24,Left=24）",
                            created = true
                        }
                    };
                }

                // mode=cell：数据验证下拉（与 render_dashboard B2 同构；列表源=源数据该列去重值 + 首项「全部」）
                string ctrlCell = Optional(args, "control_cell") ?? "B2";
                var values = new SortedSet<string>();
                object[,] block = sheetSource.Range[
                    sheetSource.Cells.Item[headerRowAbs + 1, fieldColAbs],
                    sheetSource.Cells.Item[lastRowAbs, fieldColAbs]].Value2;
                for (int r = 1; r <= block.GetLength(0); r++)
                {
                    if (block[r, 1] != null) { values.Add(Convert.ToString(block[r, 1])); }
                }
                if (values.Count == 0)
                {
                    return Invalid("NO_DIMENSION_VALUES", $"第 {fieldColumn} 列无有效维度值。");
                }
                // M-D3 审查修正②：下拉首项「全部」（与 staging B2 缺省值一致，用户可切回全量视图）；
                // 不覆盖当前值（若当前值不在列表中则复位为「全部」）
                string listFormula = "\"(全部)," + string.Join(",", values) + "\"";
                dynamic validation = sheetBoard.Range[ctrlCell].Validation;
                validation.Delete();
                validation.Add(Type: 3, Formula1: listFormula);
                object cur = sheetBoard.Range[ctrlCell].Value2;
                string curText = cur == null ? string.Empty : Convert.ToString(cur);
                if (!curText.Equals("(全部)", StringComparison.Ordinal) && !values.Contains(curText))
                {
                    sheetBoard.Range[ctrlCell].Value2 = "(全部)";
                }
                Release((object)validation);

                recorder?.Invoke("set_linkage", args, "linkage-set",
                    new { mode = "cell", field_column = fieldColumn, sheet = dashboardSheet, control_cell = ctrlCell, options = values.Count },
                    new { Sheet = dashboardSheet, Mode = "cell", FieldColumn = fieldColumn });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        mode = "cell",
                        dashboard_sheet = dashboardSheet,
                        control_cell = ctrlCell,
                        options = values.Count,
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryRollbackLinkage(workbook, dashboardSheet);
                string hint = ex.HResult == unchecked((int)0x800A03EC)
                    ? "工作簿可能处于兼容模式（未存盘），切片器被禁用；请先另存为 .xlsx 再试，或改用 mode=cell。"
                    : "已回滚，看板与源数据未受影响。";
                return Invalid("LINKAGE_SET_FAILED",
                    "联动设置 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。" + hint);
            }
            catch (InvalidOperationException ex)
            {
                TryRollbackLinkage(workbook, dashboardSheet);
                return Invalid("LINKAGE_SET_FAILED", ex.Message + "。已回滚，看板与源数据未受影响。");
            }
            catch (Exception ex)
            {
                TryRollbackLinkage(workbook, dashboardSheet);
                return Invalid("LINKAGE_SET_FAILED", "联动设置失败：" + ex.Message + "。已回滚，看板与源数据未受影响。");
            }
            finally
            {
                Release((object)slicer);
                Release((object)slicerCache);
                Release((object)pivotTable);
                Release((object)pivotCache);
                Release((object)sheetBoard);
                Release((object)sheetStaging);
                Release((object)sheetSource);
            }
        }

        /// <summary>
        /// native 联动闭环（M-D3）：把 staging 的筛选条件从「B2 控制单元格」切换为
        /// 「SUMPRODUCT(SUMIF/COUNTIF 引用透视表可见标签区)」——切片器选择变化 →
        /// 透视行区变化 → 公式自动重算（FR-2.1）；「总计」行标签无明细匹配，SUMIF 贡献 0。
        /// 覆盖范围：A 列 KPI×4、C:E 月度矩阵（D/E 列）、G:H 品类矩阵（H 列）。
        /// 返回改写的公式数（诊断用）。
        /// </summary>
        private int RewriteStagingFormulasForNative(
            dynamic workbook, string dashboardSheet, dynamic pivotTable,
            int fieldColAbs, int valueColAbs, int monthColAbs, int catColAbs,
            int qtyColAbs, string dataSheetName, int dataRowCount, int firstDataRowAbs)
        {
            return RewriteStagingFormulasForNativeMulti(workbook, dashboardSheet,
                new List<string> { Convert.ToString(pivotTable.Name) },
                new List<int> { fieldColAbs }, valueColAbs, monthColAbs, catColAbs,
                qtyColAbs, dataSheetName, dataRowCount, firstDataRowAbs);
        }

        /// <summary>多维版本：fieldColsAbs 每列一个切片器维度，M 掩码=各维 COUNTIF 乘积（AND 语义）。</summary>
        private int RewriteStagingFormulasForNativeMulti(
            dynamic workbook, string dashboardSheet, List<string> pivotNames,
            List<int> fieldColsAbs, int valueColAbs, int monthColAbs, int catColAbs,
            int qtyColAbs, string dataSheetName, int dataRowCount, int firstDataRowAbs)
        {
            string esc = EscapeSheet(dataSheetName);
            string valL = ColumnLetter(valueColAbs);
            string monL = ColumnLetter(monthColAbs);
            string catL = ColumnLetter(catColAbs);
            // 行区间精确到实际首末行（firstDataRowAbs：明细首行绝对行号——测试2026 实证明细从第 10 行起，
            // 旧实现硬编码第 2 行起，锚不在 A1/表头不在第 1 行时改写公式全错位）
            var dimLs = new List<string>();
            var dimRefs = new List<string>();
            foreach (var fc in fieldColsAbs)
            {
                string dl = ColumnLetter(fc);
                dimLs.Add(dl);
                dimRefs.Add($"{esc}!${dl}${firstDataRowAbs}:${dl}${firstDataRowAbs + dataRowCount - 1}");
            }
            string dimRef = dimRefs[0];  // 兼容既有引用（A4/单维路径）
            string valRef = $"{esc}!${valL}${firstDataRowAbs}:${valL}${firstDataRowAbs + dataRowCount - 1}";
            string monRef = $"{esc}!${monL}${firstDataRowAbs}:${monL}${firstDataRowAbs + dataRowCount - 1}";
            string catRef = $"{esc}!${catL}${firstDataRowAbs}:${catL}${firstDataRowAbs + dataRowCount - 1}";

            // 透视表行区标签引用（每透视行区首列；地址由下方 pTagRefs 循环逐透视生成）
            dynamic stagingForPivot = workbook.Worksheets.Item["_" + EscapeSheetInner(dashboardSheet) + "_calc"];
            dynamic ptFirst = stagingForPivot.PivotTables().Item(pivotNames[0]);
            dynamic rowRange = ptFirst.TableRange1;
            int pRow = Convert.ToInt32(rowRange.Row);
            int pRows = Convert.ToInt32(rowRange.Rows.Count);
            int pCol = Convert.ToInt32(rowRange.Column);
            Release((object)rowRange);
            // 多维（独立透视方案）：每维度透视行区地址不同（diag3 实证 GR/HQ）——
            // pTagRefs 必须逐透视取实际行区首列标签区，统一地址会让非首维 COUNTIF 恒 0
            var pTagRefs = new List<string>();
            foreach (var pnDim in pivotNames)
            {
                dynamic ptDim = stagingForPivot.PivotTables().Item(pnDim);
                dynamic trD = ptDim.TableRange1;
                int rD = Convert.ToInt32(trD.Row);
                int rowsD = Convert.ToInt32(trD.Rows.Count);
                int cD = Convert.ToInt32(trD.Column);
                // 2026-09-12 根治（md1 T11c 铁证）：标签区引用必须【裸地址】——公式与透视同写在本
                // staging sheet；此前写 '工作簿名'!$GR$1:$GR$5（Parent RCW 分离后拿不到 sheet 名的
                // 过渡方案），Excel 会把自工作簿单限定引用规范化成【第一张 sheet】
                //（'Md2_xxx.xlsx'!$GR$1:$GR$5 → _销售看板_calc!$GR$1:$GR$5 实测），旧表 GR 区为空
                // → SUMIF/COUNTIF 恒 0、切片器联动 KPI 全 0。裸地址同表自引用无此问题。
                pTagRefs.Add($"${ColumnLetter(cD)}${rD}:${ColumnLetter(cD)}${rD + rowsD - 1}");
                Release((object)trD);
                Release((object)ptDim);
            }
            string pTagRef = pTagRefs[0];  // 兼容单维引用
            // 选中掩码：逐行标量 COUNTIF 写入辅助列（L 列）——SUMPRODUCT(MAX(数组)) 在部分工作簿
            // 重算 #VALUE!（测试2026 实证，U5 值级验证首次覆盖改写后 KPI）；A4 改为纯 MAX(辅助列区间)。
            // 行内 COUNTIF(标签区, 单个维度值) 是标量语义，全版本安全。

            dynamic staging = workbook.Worksheets.Item["_" + EscapeSheetInner(dashboardSheet) + "_calc"];
            int rewritten = 0;
            var writtenAddrs = new List<string>();
            void SetFormula(string addr, string formula)
            {
                dynamic cell = staging.Range[addr];
                cell.Formula = formula;
                Release((object)cell);
                writtenAddrs.Add(addr);
                rewritten++;
            }
            // 2026-09-12 真机实证（md1 T11c 判定实验 probe3）：写入时透视未就绪 → Excel 把公式
            // 求值缓存为 0，CalculateFull/Rebuild 均不清；对公式格"读回再写回"必然触发正确重算。
            void ForceRecalcWrittenCells()
            {
                // 2026-09-12 升级：同名 sheet 重建后同地址格有求值坏缓存（md1 T12/T11c/T14b 实证，
                // CalculateFull/Rebuild/写回 Formula 均无效）——改用 Range.Calculate() 单格显式重算 API
                foreach (var addr in writtenAddrs)
                {
                    try
                    {
                        dynamic c = staging.Range[addr];
                        c.Calculate();
                        Release((object)c);
                    }
                    catch { }
                }
            }

            // A 列 KPI×4（与 DefaultKpis 缺省四卡对应）。
            // 2026-09-12 分叉（T11c 深层求值异常止损）：单字段维持 v3.7.1.30 已实证 58/58 的 SUMIF 形态
            //（多维 COUNTIF 泛化版在单字段下触发 Excel 求值缓存异常：同公式探针 2742 而 A1=0 且
            //  CalculateFull/Rebuild/公式重写均无法清除——已记录为已知研究项）；多维路径保留泛化版。
            // 2026-09-12 A1 坏缓存规避：联动 KPI 公式写 B1-B3（A1 地址格存在 Excel 求值坏缓存实证）
            if (fieldColsAbs.Count == 1)
            {
                string sumSel = $"SUMPRODUCT(SUMIF({dimRef},{pTagRef},{valRef}))";
                string cntSel = $"SUMPRODUCT(COUNTIF({dimRef},{pTagRef}))";
                SetFormula("B1", $"={sumSel}");
                SetFormula("B2", $"={cntSel}");
                SetFormula("B3", $"=IFERROR({sumSel}/{cntSel},0)");
            }
            else
            {
                var dimMatchParts = new List<string>();
                for (int fi = 0; fi < fieldColsAbs.Count; fi++)
                {
                    dimMatchParts.Add($"(COUNTIF({pTagRefs[fi]},{dimRefs[fi]})>0)");
                }
                string dimMatch = string.Join("*", dimMatchParts);
                string sumSel = $"SUMPRODUCT({dimMatch}*{valRef})";
                string cntSel = $"SUMPRODUCT({dimMatch}*({valRef}>0)*1)";
                SetFormula("B1", $"={sumSel}");
                SetFormula("B2", $"={cntSel}");
                SetFormula("B3", $"=IFERROR({sumSel}/{cntSel},0)");
            }

            // 逐行辅助列（全标量公式，全版本安全）：M=选中掩码 0/1、L=选中值（掩码×金额）
            // A4=MAX(L 区间)。禁用数组/MAX(IF) 族——测试2026 实证：SUMPRODUCT(MAX(数组)) 与
            // MAX(IF(数组)) 在用户工作簿重算 #VALUE!（U5 值级验证首次覆盖改写后 KPI）
            // 同上短名原则：M/L/A4 互相引用本 staging sheet，用裸地址
            for (int i = 0; i < dataRowCount; i++)
            {
                int srcRow = firstDataRowAbs + i;
                if (fieldColsAbs.Count == 1)
                {
                    SetFormula($"M{i + 2}", $"=IF(COUNTIF({pTagRef},{esc}!${dimLs[0]}${srcRow})>0,1,0)");
                }
                else
                {
                    var cntParts = new List<string>();
                    for (int fi = 0; fi < fieldColsAbs.Count; fi++)
                    {
                        cntParts.Add($"COUNTIF({pTagRefs[fi]},{esc}!${dimLs[fi]}${srcRow})");
                    }
                    SetFormula($"M{i + 2}", $"=IF({string.Join("*", cntParts)}>0,1,0)");
                }
                SetFormula($"L{i + 2}", $"=$M${i + 2}*{esc}!${valL}${srcRow}");
            }
            SetFormula("A4", $"=MAX($L$2:$L${dataRowCount + 1})");

            // D/E 列月度矩阵（标签在 C 列，即快照第 1 列）与 H 列品类矩阵（标签在 G 列，即快照第 5 列）：
            // SUMPRODUCT((月=标签)*(inSel)*值)；品类同理
            dynamic used = staging.UsedRange;
            int uRows = Convert.ToInt32(used.Rows.Count);
            int uCols = Convert.ToInt32(used.Columns.Count);
            Release((object)used);
            object[,] snap = staging.Range[
                staging.Cells.Item[1, 3], staging.Cells.Item[Math.Max(uRows, 2), Math.Max(uCols, 8)]].Value2;
            // 2026-09-12 用户文件实证修复：qtyRef 行区间原硬编码 $2 起，表头不在第 1 行时（测试2026 第 10 行）
            // 与 valRef($11:$27) 错位——区间里混入表头文本 → SUMPRODUCT #VALUE! → 销量折线全空。
            // 对齐 monRef/valRef 口径：firstDataRowAbs 起。
            string qtyRef = $"{esc}!${ColumnLetter(qtyColAbs)}${firstDataRowAbs}:${ColumnLetter(qtyColAbs)}${firstDataRowAbs + dataRowCount - 1}";
            string maskRef = "$M$2:$M$" + (dataRowCount + 1);

            for (int r = 2; r <= Math.Min(uRows, snap.GetLength(0)); r++)
            {
                // 月度矩阵：C 列标签 + D/E 公式列；品类矩阵：G 列标签 + H 公式列
                object monthTag = snap[r, 1];
                if (monthTag != null)
                {
                    // 2026-09-12 用户文件实证修复：日期型月份列的标签是序列号 double，
                    // 旧实现统一拼成 "45674" 文本比较 → 明细里是日期值永不相等 → 月度矩阵全 0 → 趋势图空。
                    // 数值型标签不加引号（数值比较），文本标签维持引号。
                    string t = Convert.ToString(monthTag).Replace("\"", "\"\"");
                    string mCmp = monthTag is double ? t : "\"" + t + "\"";
                    SetFormula($"D{r}", $"=SUMPRODUCT(({monRef}={mCmp})*{maskRef}*{valRef})");
                    SetFormula($"E{r}", $"=SUMPRODUCT(({monRef}={mCmp})*{maskRef}*{qtyRef})");
                }
                object catTag = snap[r, 5];
                if (catTag != null && Convert.ToString(catTag) != "品类")
                {
                    string t = Convert.ToString(catTag).Replace("\"", "\"\"");
                    string cCmp = catTag is double ? t : "\"" + t + "\"";
                    SetFormula($"H{r}", $"=SUMPRODUCT(({catRef}={cCmp})*{maskRef}*{valRef})");
                }
            }

            ForceRecalcWrittenCells();
            return rewritten;
        }

        /// <summary>
        /// 批量写公式性能守卫（M-D3 优化轮）：渲染前关屏幕刷新 + 切手动重算，
        /// EndBulkFormulaWrite 恢复原状态并统一重算一次。避免 staging 每写一个公式
        /// 触发全簿自动重算（2 万行实测曾因此 14.7s 超标）。
        /// </summary>
        private object BeginBulkFormulaWrite()
        {
            var state = new object[2];
            try { state[0] = application.ScreenUpdating; } catch { state[0] = true; }
            try { state[1] = application.Calculation; } catch { state[1] = -4105; }
            try { application.ScreenUpdating = false; } catch { }
            try { application.Calculation = -4135; } catch { } // xlCalculationManual
            return state;
        }

        private void EndBulkFormulaWrite(object guard, params string[] sheetNames)
        {
            if (!(guard is object[] state))
            {
                return;
            }
            try { application.Calculation = state[1]; } catch { }
            // M-D3 优化：指定 sheet 时用 sheet 级 Calculate（省全簿扫描）；
            // 但依赖链跨表时 Excel 会自动扩展计算范围，语义安全。
            try
            {
                if (sheetNames != null && sheetNames.Length > 0)
                {
                    foreach (var name in sheetNames)
                    {
                        try { application.ActiveWorkbook.Worksheets.Item[name].Calculate(); } catch { }
                    }
                }
                else
                {
                    application.Calculate();
                }
            }
            catch { }
            try { application.ScreenUpdating = state[0]; } catch { }
        }

        /// <summary>
        /// FR-6.3 自检回显：看板渲染成功后立即跑一次 verify_dashboard，把结论并入渲染返回值——
        /// 用户在聊天里能直接看到「自检通过 ✓」，而不是埋在 auto-verify 观察消息里。
        /// 返回（verdict, 可读明细）；任何内部异常都不抛出（自检失败不否定渲染本身，auto-verify 仍会兜底）。
        /// </summary>
        private (string Verdict, string Detail) RunPostRenderSelfCheck(string dashboardSheet)
        {
            try
            {
                var v = VerifyDashboard(new Dictionary<string, object> { ["dashboard_sheet"] = dashboardSheet });
                if (v != null && v.Ok && v.Data != null)
                {
                    dynamic d = v.Data;
                    return (Convert.ToString(d.verdict), Convert.ToString(d.selfcheck));
                }
            }
            catch { }
            return ("not_run", "自检未执行（内部异常，auto-verify 仍会兜底）");
        }

        /// <summary>
        /// FR-6.1 默认切片器：渲染成功后自动挂 native 切片器（复用 set_linkage native 全路径：
        /// 透视表 + SlicerCache + staging 公式改写）。任何失败（兼容模式禁切片器/低版本等）都
        /// 回退到渲染已建好的 cell 下拉，绝不因切片器失败否定整次渲染。
        /// 返回（ok, 字段名, SlicerCache 名, 改写公式数, 失败原因）。
        /// </summary>
        private (bool Ok, string Field, string Cache, int Rewritten, string Reason) AttachAutoSlicerIfPossible(
            Dictionary<string, object> renderArgs, string dashboardSheet, string sourceSheetName)
        {
            try
            {
                var linkArgs = new Dictionary<string, object>(renderArgs);
                linkArgs["mode"] = "native";
                linkArgs["dashboard_sheet"] = dashboardSheet;
                // render 的 label_column 即联动筛选维度列——映射为 set_linkage 的 field_column
                // （测试2026 实证：不映射时切片器错误建在缺省第 1 列）
                if (renderArgs.TryGetValue("label_column", out var labelColObj))
                {
                    linkArgs["field_column"] = labelColObj;
                }
                // 渲染会激活新表，ActiveSheet 不可靠——源数据表名必须显式传（M-D1 教训）
                linkArgs["sheet"] = sourceSheetName;
                var link = SetLinkage(linkArgs);
                if (link != null && link.Ok && link.Data != null)
                {
                    dynamic ld = link.Data;
                    return (true, Convert.ToString(ld.field), Convert.ToString(ld.slicer_cache),
                            Convert.ToInt32(ld.staging_formulas_rewritten), null);
                }
                return (false, null, null, 0,
                    (link == null ? "UNKNOWN" : Convert.ToString(link.ErrorCode)) + " " +
                    Convert.ToString(link == null ? null : link.Data));
            }
            catch (Exception ex)
            {
                return (false, null, null, 0, ex.Message);
            }
        }

        // ---- FR-6.2 图表选型建议（suggest_chart_mapping，只读启发式）----

        /// <summary>列画像：选型启发式与单测共用的纯数据载体。</summary>
        public sealed class ChartColumnProfile
        {
            public string Name;
            public int Rows;
            public int NonEmpty;
            public int Numeric;
            public int DateLike;
            public int DistinctText;
            public bool HasNegative;
            public double NumericRatio;
            public bool IsMostlyNumeric;
            public bool IsMostlyDateLike;
            public List<double> NumericValues = new List<double>();
        }

        /// <summary>单条选型建议：工具 + 图表类型 + 理由 + 数据要求（AI 拿去向用户复述并确认映射）。</summary>
        public sealed class ChartSuggestion
        {
            public string Tool;
            public string ChartType;
            public string Reason;
            public string Requires;
        }

        /// <summary>
        /// 选型启发式纯函数（可单测）：block 第 0 行为表头，其后为数据。
        /// 规则与 dashboard-chart-advisor 技能同源：数值含负→瀑布；逐级递减→漏斗；双数值列→组合图；
        /// 两文本维度→交叉/层级；类别+时间+数值→排名动态；单文本单数值→环形/柱形；纯数值→直方图。
        /// 每条建议必须带理由（拍板：建议和理由缺一不可）。
        /// </summary>
        public static List<ChartSuggestion> AnalyzeChartMapping(object[,] block, out List<ChartColumnProfile> columns)
        {
            columns = new List<ChartColumnProfile>();
            int colCount = block.GetLength(1);
            int rowCount = block.GetLength(0) - 1;
            if (colCount < 1) { return new List<ChartSuggestion>(); }
            if (rowCount < 1)
            {
                // 表头-only：列画像照常返回（名称/零行），建议为空（T 用例边界）
                for (int c = 0; c < colCount; c++)
                {
                    columns.Add(new ChartColumnProfile { Name = Convert.ToString(block[0, c]), Rows = 0 });
                }
                return new List<ChartSuggestion>();
            }
            for (int c = 0; c < colCount; c++)
            {
                var p = new ChartColumnProfile { Name = Convert.ToString(block[0, c]), Rows = rowCount };
                var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int nonEmpty = 0, numeric = 0, dateLike = 0;
                for (int r = 1; r <= rowCount; r++)
                {
                    object v = block[r, c];
                    if (v == null) { continue; }
                    nonEmpty++;
                    if (v is double d)
                    {
                        numeric++;
                        if (d < 0) { p.HasNegative = true; }
                        // 日期序列号粗判（2000~2119 年）；纯业务数值落在该区间会误判，仅作辅助特征
                        if (d >= 36526 && d <= 80000) { dateLike++; }
                        if (p.NumericValues.Count < 5000) { p.NumericValues.Add(d); }
                    }
                    string t = Convert.ToString(v);
                    if (!string.IsNullOrEmpty(t) && distinct.Count < 500) { distinct.Add(t); }
                }
                p.NonEmpty = nonEmpty; p.Numeric = numeric; p.DateLike = dateLike;
                p.DistinctText = distinct.Count;
                p.NumericRatio = nonEmpty == 0 ? 0 : (double)numeric / nonEmpty;
                p.IsMostlyNumeric = p.NumericRatio >= 0.6;
                p.IsMostlyDateLike = nonEmpty > 0 && (double)dateLike / nonEmpty >= 0.6;
                columns.Add(p);
            }

            var textCols = columns.Where(c => !c.IsMostlyNumeric).ToList();
            var numCols = columns.Where(c => c.IsMostlyNumeric).ToList();
            var suggestions = new List<ChartSuggestion>();
            void Add(string tool, string chartType, string reason, string requires)
            {
                if (!suggestions.Any(s => s.Tool == tool && s.ChartType == chartType))
                {
                    suggestions.Add(new ChartSuggestion { Tool = tool, ChartType = chartType, Reason = reason, Requires = requires });
                }
            }

            // 1) 财务增减：数值列含负数 → 瀑布图
            if (numCols.Any(c => c.HasNegative))
            {
                Add("create_chart_advanced", "waterfall",
                    "数值列含负数（正负冲减），瀑布图最能表达「收入→成本→利润」式累计增减",
                    "类别列 + 含正负的数值列；Excel 2016+（低版本可改垫底柱堆积模拟）");
            }

            // 2) 转化漏斗：第一数值列沿行序逐级递减（≥3 行且至少一次下降）
            if (numCols.Count >= 1 && rowCount >= 3 && numCols[0].NumericValues.Count >= 3)
            {
                var vals = numCols[0].NumericValues;
                bool anyDrop = false, monotonic = true;
                for (int i = 1; i < vals.Count; i++)
                {
                    if (vals[i] > vals[i - 1]) { monotonic = false; break; }
                    if (vals[i] < vals[i - 1]) { anyDrop = true; }
                }
                if (monotonic && anyDrop)
                {
                    Add("create_chart_advanced", "funnel",
                        "数值沿行序逐级递减，符合转化流程形态（如 访问→下单→支付）",
                        "阶段列 + 单数值列；Excel 2013+ 技巧型漏斗全版本兼容");
                }
            }

            // 3) 双数值列 + 类别 → 组合图（次轴）
            if (numCols.Count >= 2 && textCols.Count >= 1)
            {
                Add("create_combo_chart", "column+line",
                    "检测到两列数值——量纲不同时组合图（柱形 + 次轴折线）比两张图并排更易读",
                    "共享类别列 + 两个数值列；次轴挂数量/比率类第二指标");
            }

            // 4) 两个文本维度 + 数值 → 交叉看板 / 层级图
            if (textCols.Count >= 2 && numCols.Count >= 1)
            {
                if (textCols.Count(c => c.DistinctText >= 2 && c.DistinctText <= 15) >= 2)
                {
                    Add("render_dim_cross", "matrix",
                        "两个低基数文本维度 + 数值——多维交叉矩阵 + 双下拉联动最适合『A×B 看指标』式分析",
                        "行维度列 + 列维度列 + 数值列；交叉组合 ≤400 单元");
                }
                if (textCols.Select(c => c.DistinctText).Distinct().Count() >= 2)
                {
                    Add("create_chart_advanced", "treemap",
                        "文本列基数不同（可能存在层级嵌套），树状图/旭日图表达层级占比最直观",
                        "层级类别列 + 数值列；Excel 2016+（树状/旭日建议更高版本，先跑 check_chart_version）");
                }
            }

            // 5) 类别 + 时间 + 数值 → 排名动态看板
            var timeCol = columns.FirstOrDefault(c => c.IsMostlyDateLike) ??
                          textCols.FirstOrDefault(c => c.DistinctText >= 2 && c.DistinctText <= 31 &&
                              c.Name != null && (c.Name.Contains("月") || c.Name.Contains("时间") || c.Name.Contains("日期") || c.Name.Contains("季度")));
            if (textCols.Count >= 1 && numCols.Count >= 1 && timeCol != null && !ReferenceEquals(timeCol, textCols.FirstOrDefault()))
            {
                Add("render_rank_dynamic", "topN+timeslicer",
                    "检测到类别 + 时间 + 数值三要素——排名动态看板支持按时间下拉切换 Top N 排行",
                    "类别列 + 时间列 + 数值列（列位可传参重指）");
            }

            // 6) 四列以上多文本多数值 → 销售总览看板
            if (textCols.Count >= 2 && numCols.Count >= 2 && columns.Count >= 4)
            {
                Add("render_dashboard", "sales_overview",
                    "四列以上且多文本多数值——销售总览看板（KPI 卡 + 趋势 + 品类环形 + 联动筛选）一次成型",
                    "地区/类别/月份/销售额/销量 五列口径（缺列会报数据门槛，建议先补齐）");
            }

            // 7) 单文本 + 单数值 → 环形（≤5 类）或柱形
            if (textCols.Count == 1 && numCols.Count == 1)
            {
                if (textCols[0].DistinctText <= 5)
                {
                    Add("create_doughnut", "doughnut",
                        "类别数 ≤5，环形图表达构成占比最简洁（官方建议饼图 ≤7 类）",
                        "类别列 + 单数值列；多系列时每环一列");
                }
                else
                {
                    Add("create_excel_chart", "column",
                        "常规类别对比，簇状柱形是最稳的默认选择；类别多/标签长时改条形图",
                        "类别列 + 数值列");
                }
            }

            // 8) 纯数值（无有效文本列）→ 直方图
            if (textCols.Count == 0 && numCols.Count >= 1)
            {
                Add("create_chart_advanced", "histogram",
                    "没有类别列、只有数值——直方图分箱看分布形态（成绩/时长类数据）",
                    "单数值列；Excel 2016+；可调箱宽/箱数");
            }

            // 9) 兜底：任何数据都能画的默认项
            Add("create_excel_chart", "column",
                "未命中更强特征时的稳妥默认项：簇状柱形先看整体，再按需换型",
                "类别列 + 数值列（首列默认作类别）");
            return suggestions;
        }

        // ---- FR-ED 深度核验（verify_dashboard_deep）：检查是否“生成对”，依据全量回显 ----

        /// <summary>CopyPicture 带一次重试：剪贴板偶发占用/渲染竞态导致 0x800A03EC，延迟后重试可解。</summary>
        private static void CopyPictureWithRetry(dynamic range, int appearance, int format)
        {
            try
            {
                range.CopyPicture(Appearance: appearance, Format: format);
                return;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                System.Threading.Thread.Sleep(300);
            }
            range.CopyPicture(Appearance: appearance, Format: format);  // 二次失败原样抛出
        }

        /// <summary>
        /// 剪贴板直取：从剪贴板读取 Bitmap 并保存为 PNG（独立 STA 线程；CopyPicture 在调用线程完成）。
        /// 返回 true=已保存；失败不抛出（由调用方降级到 Activate+Paste 路线）。
        /// </summary>
        private static bool TrySaveClipboardBitmapToPng(string pngPath)
        {
            bool ok = false;
            Exception threadError = null;
            var t = new System.Threading.Thread(() =>
            {
                try
                {
                    System.Threading.Thread.Sleep(150);
                    var data = System.Windows.Forms.Clipboard.GetDataObject();
                    if (data != null && data.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap))
                    {
                        using (var bmp = (System.Drawing.Bitmap)data.GetData(System.Windows.Forms.DataFormats.Bitmap))
                        {
                            bmp.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
                            ok = true;
                        }
                    }
                }
                catch (Exception ex) { threadError = ex; }
            });
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
            t.Join(10000);
            return ok;
        }

        /// <summary>
        /// 路线二（降级，已实证可用）：EMF 复制 + 临时 ChartObject 粘贴 + Export PNG。
        /// 2026-09-13 从看板路线抽出共用：地址截图路线在剪贴板直取失败时同样降级到这里
        ///（真机 udchart 套件实证：剪贴板瞬态竞争下地址路线此前直接外抛、无任何降级；
        /// 看板路线首拷异常也会绕过本路线直接外抛）。失败时清理临时图后原样抛出。
        /// </summary>
        private void ExportRangeViaTempChart(dynamic board, dynamic shotRange, string pngPath, double w, double h)
        {
            // 首拷异常（剪贴板/COM 瞬态）直接外抛：调用方决定降级路线，本方法不吞。
            shotRange.CopyPicture(Appearance: 1, Format: -4147);
            dynamic co = board.ChartObjects().Add(0, 0, w, h);
            dynamic chartD = null;
            try
            {
                chartD = co.Chart;
                try { co.Activate(); } catch { }   // 隐藏实例 Activate 可能失败（SB3 教训），Paste 前尽力激活
                System.Threading.Thread.Sleep(250);
                chartD.Paste();
                chartD.Export(pngPath, "PNG");
            }
            catch (Exception ex)
            {
                AppLog.Warn("dashboard", "ExcelDashboardService.cs:ExportRangeViaTempChart 静默兜底触发: " + ex.Message);
                try { co.Delete(); } catch { }
                throw;
            }
            finally
            {
                if (chartD != null) { Release((object)chartD); }
                Release((object)co);
            }
        }

        /// <summary>
        /// 截图缩略图 DataUri（2026-09-13）：降采样到宽 ≤720px 的 PNG，
        /// 让加载项 AI 能像 auto_check.chart_image 一样直接读图做视觉自审（用户核心诉求：AI 读截图→自审→自调）。
        /// 失败返回 null（不影响截图主产物）。
        /// </summary>
        private static string TryBuildImageThumbDataUri(string pngPath)
        {
            try
            {
                using (var src = new System.Drawing.Bitmap(pngPath))
                {
                    int maxW = 720;
                    int tw = src.Width, th = src.Height;
                    if (tw > maxW)
                    {
                        tw = maxW;
                        th = Math.Max(1, (int)Math.Round(src.Height * (maxW / (double)src.Width)));
                    }
                    using (var dst = new System.Drawing.Bitmap(tw, th))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(dst))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.DrawImage(src, 0, 0, tw, th);
                        }
                        using (var ms = new System.IO.MemoryStream())
                        {
                            dst.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            return "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
                        }
                    }
                }
            }
            catch
            {
                AppLog.Warn("dashboard", "ExcelDashboardService.cs:TryBuildImageThumbDataUri 静默兜底触发（读图缩略失败不影响截图主产物）");
                return null;
            }
        }

        /// <summary>
        /// FR-6.5：screenshot_dashboard 截图工具（只读）。
        /// 看板级：区域 CopyPicture → 临时 ChartObject 粘贴 → Export PNG；图表级：按 chart_name 直接 Export。
        /// 返回截图文件路径 + 像素统计，供用户直接查看整体视觉效果
        ///（grilling 决策 2026-09-11：AI 无法直接看图，视觉判断以像素统计 + 用户反馈为准；宿主图片消息支持单独立项）。
        /// </summary>
        public ToolResult ScreenshotDashboard(Dictionary<string, object> args)
        {
            // 2026-09-12 扩展（用户实证反馈：侧边栏 AI 不知道能截图/只能截看板）：
            // 优先级 address(任意选区截图) > chart_name(单图表) > dashboard_sheet(看板整体)；
            // dashboard_sheet 改可选（缺省活动表），描述强化"用户要求截图/预览时调用"。
            var dashboardSheet = Optional(args, "dashboard_sheet");
            var chartName = Optional(args, "chart_name");
            var addrArg = Optional(args, "address");
            var sheetArg = Optional(args, "sheet");
            var outDirArg = Optional(args, "out_dir");
            dynamic workbook = null;
            dynamic board = null;
            string screenshotRoute = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // "随时随地截图"：什么都不传 = 截当前活动选区（Selection）——用户选中哪儿截哪儿
                if (string.IsNullOrEmpty(dashboardSheet) && string.IsNullOrEmpty(sheetArg) &&
                    string.IsNullOrEmpty(addrArg) && string.IsNullOrEmpty(chartName))
                {
                    dynamic sel = application.ActiveWindow.RangeSelection;
                    if (sel != null)
                    {
                        addrArg = Convert.ToString(sel.Address);
                        sheetArg = Convert.ToString(sel.Worksheet.Name);
                        ExcelTrace.Info("screenshot_dashboard", "未传参：缺省截活动选区 " + addrArg + " @ " + sheetArg);
                    }
                    Release((object)sel);
                }
                board = workbook.Worksheets.Item[dashboardSheet ?? sheetArg ?? Convert.ToString(workbook.ActiveSheet.Name)];

                string outDir = outDirArg;
                if (string.IsNullOrWhiteSpace(outDir))
                {
                    string wbPath = Convert.ToString(workbook.Path);
                    outDir = string.IsNullOrEmpty(wbPath)
                        ? System.IO.Path.Combine(System.IO.Path.GetTempPath(), "看板截图")
                        : System.IO.Path.Combine(wbPath, "_看板截图");
                }
                System.IO.Directory.CreateDirectory(outDir);
                string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                // 2026-09-13 文件名碰撞修复：同 sheet 多图（如高级图A 的 bullet/waffle/funnel/butterfly）
                // 旧命名 sheet_chart_时间戳 同秒互相覆盖丢图——图表级把 chart_name 编入文件名。
                string namePart = string.IsNullOrEmpty(chartName)
                    ? "board"
                    : System.Text.RegularExpressions.Regex.Replace(chartName, "[\\\\/:*?\"<>|]", "_");
                string pngPath = System.IO.Path.Combine(outDir,
                    dashboardSheet + "_" + (string.IsNullOrEmpty(chartName) ? "board" : "chart") + "_" + stamp + ".png");

                if (!string.IsNullOrEmpty(chartName))
                {
                    // 2026-09-13 存在性预检：未找到时给人话报错，不再回裸 COM HRESULT（AI 使用方视角审查项）
                    dynamic cosP = board.ChartObjects();
                    bool found = false;
                    try
                    {
                        int cntP = Convert.ToInt32(cosP.Count);
                        for (int ip = 1; ip <= cntP; ip++)
                        {
                            dynamic itP = cosP.Item(ip);
                            try { if (Convert.ToString(itP.Name) == chartName) { found = true; } }
                            finally { Release((object)itP); }
                            if (found) { break; }
                        }
                    }
                    finally { Release((object)cosP); }
                    if (!found)
                    {
                        return Invalid("CHART_NOT_FOUND",
                            "未找到图表「" + chartName + "」。请先用 list_dashboard_charts 或 verify_chart 获取真实图表名（名称区分大小写与全半角）。");
                    }
                    dynamic coC = board.ChartObjects().Item(chartName);
                    dynamic chartC = coC.Chart;
                    chartC.Export(pngPath, "PNG");
                    Release((object)chartC);
                    Release((object)coC);
                    screenshotRoute = "chart_export";
                }
                else if (!string.IsNullOrEmpty(addrArg))
                {
                    // 任意选区截图（用户"截当前选区数据"场景）：address 区域直接导出
                    dynamic addrRange = board.Range[addrArg];
                    CopyPictureWithRetry(addrRange, 1, 2);  // xlBitmap 供剪贴板直取（含重试）
                    Release((object)addrRange);
                    if (!TrySaveClipboardBitmapToPng(pngPath))
                    {
                        return Invalid("SCREENSHOT_FAILED", "选区截图失败（剪贴板直取不可用）。请让 Excel 窗口可见后重试。");
                    }
                    screenshotRoute = "address_clipboard";
                }
                else
                {
                    // 截图区域 = UsedRange ∪ 全部浮动对象（图表/切片器）边界——
                    // 第五轮视觉实证：UsedRange 只覆盖 KPI 单元格区（A~D 列），图表在区域外截不到
                    dynamic used = board.UsedRange;
                    double w = Convert.ToDouble(used.Width);
                    double h = Convert.ToDouble(used.Height);
                    Release((object)used);
                    dynamic cosB = board.ChartObjects();
                    int cnB = Convert.ToInt32(cosB.Count);
                    for (int i = 1; i <= cnB; i++)
                    {
                        dynamic cB = cosB.Item(i);
                        double rr = Convert.ToDouble(cB.Left) + Convert.ToDouble(cB.Width);
                        double bb = Convert.ToDouble(cB.Top) + Convert.ToDouble(cB.Height);
                        if (rr > w) { w = rr; }
                        if (bb > h) { h = bb; }
                        Release((object)cB);
                    }
                    Release((object)cosB);
                    if (w < 10 || h < 10)
                    {
                        return Invalid("EMPTY_BOARD", "看板区域为空，无法截图。");
                    }
                    w = Math.Min(w, 12000);
                    h = Math.Min(h, 12000);
                    // 像素→单元格：CopyPicture 区域必须覆盖浮动对象所在行列，否则图表被裁掉
                    int endCol = 1, endRow = 1;
                    double accW = 0, accH = 0;
                    while (accW < w && endCol < 120)
                    {
                        dynamic colO = board.Columns.Item[endCol];
                        accW += Convert.ToDouble(colO.Width);
                        Release((object)colO);
                        endCol++;
                    }
                    while (accH < h && endRow < 600)
                    {
                        dynamic rowO = board.Rows.Item[endRow];
                        accH += Convert.ToDouble(rowO.Height);
                        Release((object)rowO);
                        endRow++;
                    }
                    dynamic shotRange = board.Range[board.Cells.Item[1, 1], board.Cells.Item[endRow, endCol]];
                    // 路线一（2026-09-12 迭代，借鉴开源 Clipboard 直取方案）：CopyPicture(xlBitmap) 后
                    // 从剪贴板直取 Bitmap 存 PNG——绕过 chart.Paste 的 Activate 依赖（隐藏实例最脆弱一环）；
                    // 剪贴板读取放独立 STA 线程（COM 对象留主线程保线程亲和）。
                    bool gridWasOn = false;
                    try
                    {
                        dynamic winG = application.ActiveWindow;
                        gridWasOn = Convert.ToBoolean(winG.DisplayGridlines);
                        if (gridWasOn) { winG.DisplayGridlines = false; }
                        Release((object)winG);
                    }
                    catch { }
                    CopyPictureWithRetry(shotRange, 1, 2);  // xlScreen + xlBitmap（剪贴板直取格式，含重试）
                    bool viaClipboard = TrySaveClipboardBitmapToPng(pngPath);
                    screenshotRoute = viaClipboard ? "clipboard" : "chart_paste";
                    if (!viaClipboard)
                    {
                        // 路线二（降级，已实证可用）：xlPicture(EMF) 重复制 + Activate + Paste + Export
                        shotRange.CopyPicture(Appearance: 1, Format: -4147);
                        dynamic co = board.ChartObjects().Add(0, 0, w, h);
                        try { co.Border.LineStyle = 0; } catch { }
                        dynamic chartD = co.Chart;
                        try
                        {
                            try { co.Activate(); } catch { }
                            System.Threading.Thread.Sleep(250);
                            chartD.Paste();
                            chartD.Export(pngPath, "PNG");
                        }
                        catch
                        {
                            try { co.Delete(); } catch { }
                            throw;
                        }
                        co.Delete();
                        Release((object)chartD);
                        Release((object)co);
                    }
                    // 真机实证（2026-09-12 grilling 批次）：隐藏实例上 xlBitmap + 直接 Paste 会产出空白 PNG
                    //（假成功）。修复：xlPicture(EMF) + ChartObject.Activate + 粘贴延迟；并在导出后做
                    // 空白 fail-closed 防御（非白占比过低即报错，绝不返回假成功）。
                    if (gridWasOn)
                    {
                        try
                        {
                            dynamic winR = application.ActiveWindow;
                            winR.DisplayGridlines = true;
                            Release((object)winR);
                        }
                        catch { }
                    }
                    var statsGuard = ExcelChartService.ComputeRenderStats(System.IO.File.ReadAllBytes(pngPath));
                    if (!statsGuard.Item1 || statsGuard.Item2 < 0.005)
                    {
                        try { if (System.IO.File.Exists(pngPath)) { System.IO.File.Delete(pngPath); } } catch { }
                        return Invalid("SCREENSHOT_BLANK",
                            "截图导出为空白（非白占比 " + Math.Round(statsGuard.Item2 * 100, 2) + "% < 0.5%），已删除产物。" +
                            "常见原因：宿主实例隐藏导致 CopyPicture/Paste 渲染失败——请让 Excel 窗口可见后重试。");
                    }
                }

                var stats = ExcelChartService.ComputeRenderStats(System.IO.File.ReadAllBytes(pngPath));
                var fi = new System.IO.FileInfo(pngPath);
                recorder?.Invoke("screenshot_dashboard", args, "dashboard-screenshot",
                    new { sheet = dashboardSheet, chart = chartName, non_blank = Math.Round(stats.Item2, 4) }, null);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        sheet = sheetArg,
                        address = addrArg,
                        chart_name = chartName,
                        screenshot_path = pngPath,
                        route = screenshotRoute ?? "chart_export",
                        file_size_kb = Math.Round(fi.Length / 1024.0, 1),
                        data_rendered = stats.Item1,
                        non_blank_ratio = Math.Round(stats.Item2, 4),
                        sampled_pixels = stats.Item3,
                        hint = "截图已保存，请打开 PNG 查看看板整体效果；AI 无法直接看图，数据正确性依据见 verify_dashboard_deep 的 checks。"
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return Invalid("SCREENSHOT_FAILED", "截图 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message);
            }
            catch (Exception ex)
            {
                return Invalid("SCREENSHOT_FAILED", "截图失败：" + ex.Message);
            }
            finally
            {
                // board/workbook 为应用级共享对象：不释放（M-D1 RCW 纪律，同 VerifyDashboard）
            }
        }


        /// <summary>单项核验结果：状态（pass/fail/warn/skip）+ 期望 + 实际 + 依据（判定规则与对比值）。</summary>
        public sealed class DeepCheckItem
        {
            public string Item;
            public string Status;
            public string Expected;
            public string Actual;
            public string Basis;
        }

        private static readonly System.Text.RegularExpressions.Regex StagingRefRegex =
            new System.Text.RegularExpressions.Regex(@"'?([^'!=\r\n()]+)'?!(\$?[A-Z]{1,3}\$?[0-9]+)");

        private static readonly System.Text.RegularExpressions.Regex StagingRangeRegex =
            new System.Text.RegularExpressions.Regex(@"'?([^'!=\r\n()]+)'?!(\$?[A-Z]{1,3}\$?[0-9]+):(\$?[A-Z]{1,3}\$?[0-9]+)");

        /// <summary>公式引用的表名是否指向辅助表（兼容规范化四形态：带引号/裸名/剥首下划线）。</summary>
        /// <summary>松解析格式化显示值："93.7%"→0.937、"1,405,797"→1405797、"¥1,234.5"→1234.5。</summary>
        private static bool TryLooseDisplayNumeric(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) { return false; }
            string t = text.Trim().Replace(",", "").Replace("¥", "").Replace(" ", "");
            bool pct = t.EndsWith("%");
            if (pct) { t = t.TrimEnd('%'); }
            if (!double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) { return false; }
            if (pct) { value /= 100.0; }
            return true;
        }

        private static bool RefTargetsStaging(string sheetPart, string stagingName)
        {
            if (string.IsNullOrEmpty(sheetPart) || string.IsNullOrEmpty(stagingName)) { return false; }
            string t = sheetPart.Trim('\'', ' ');
            if (string.Equals(t, stagingName, StringComparison.OrdinalIgnoreCase)) { return true; }
            if (t.Length > 1 && t[0] == '_' && string.Equals(t.Substring(1), stagingName, StringComparison.OrdinalIgnoreCase)) { return true; }
            if (stagingName[0] == '_' && string.Equals(t, stagingName.Substring(1), StringComparison.OrdinalIgnoreCase)) { return true; }
            return false;
        }

        private static string NormCellText(object v)
        {
            if (v == null) { return string.Empty; }
            if (v is double d) { return d.ToString("0.######", CultureInfo.InvariantCulture); }
            return Convert.ToString(v).Trim();
        }

        /// <summary>按 checks 汇总生成深度核验返回值（verdict：任一 fail=mismatch，否则 ok）。</summary>
        private ToolResult DeepResult(string dashboardSheet, List<DeepCheckItem> checks, string themeName)
        {
            string verdict = checks.Any(c => c.Status == "fail") ? "mismatch" : "ok";
            var summaryLines = checks.Select(c =>
                $"[{c.Status.ToUpperInvariant()}] {c.Item}｜期望：{c.Expected}｜实际：{c.Actual}｜依据：{c.Basis}");
            recorder?.Invoke("verify_dashboard_deep", new Dictionary<string, object> { ["dashboard_sheet"] = dashboardSheet },
                "dashboard-deep-verify", new { sheet = dashboardSheet, verdict, checks = checks.Count }, null);
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    dashboard_sheet = dashboardSheet,
                    theme = themeName,
                    verdict,
                    checks,
                    selfcheck = verdict == "ok"
                        ? "深度核验通过 ✓ " + string.Join("；", checks.Select(c => c.Item + " " + c.Status))
                        : "深度核验发现异常 ✗ " + string.Join("；", checks.Where(c => c.Status == "fail").Select(c => c.Item + "：" + c.Actual)),
                    usage = "AI 回复必须逐项转述 checks 的期望/实际/依据，不得只说「正确」"
                }
            };
        }
        /// <summary>
        /// FR-ED：verify_dashboard_deep 看板深度核验（只读，六项检查各带期望/实际/依据）：
        /// ①看板与辅助表存在；②辅助表 KPI 重算全数值；③看板卡片显示值守恒（显示值==辅助表被引用值，
        /// 历史缺陷审查⑨同形态必抓）；④联动控件（下拉/切片器，信息项）；⑤数值守恒对账（全量视图下
        /// KPI 重算值 vs 源数据列直和）；⑥图表渲染像素核验（空白图如实暴露，层级图已知引擎限制标 warn）。
        /// </summary>
        public ToolResult VerifyDashboardDeep(Dictionary<string, object> args)
        {
            var dashboardSheet = Required(args, "dashboard_sheet");
            var checkPixels = !string.Equals(Optional(args, "check_pixels"), "false", StringComparison.OrdinalIgnoreCase);
            var checks = new List<DeepCheckItem>();
            string stagingName = "_" + EscapeSheetInner(dashboardSheet) + "_calc";
            string themeName = "unknown";
            dynamic workbook = null;
            dynamic board = null;
            dynamic staging = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");

                // ① 看板与辅助表存在
                bool boardOk = SheetExists(workbook, dashboardSheet);
                bool stgOk = boardOk && SheetExists(workbook, stagingName);
                checks.Add(new DeepCheckItem
                {
                    Item = "看板与辅助表存在",
                    Status = boardOk && stgOk ? "pass" : "fail",
                    Expected = "看板 sheet 与 _看板名_calc 隐藏辅助表均存在",
                    Actual = $"看板={boardOk}，辅助表={stgOk}",
                    Basis = "看板由可见看板 sheet + 隐藏 _calc 公式辅助表构成，缺一即不完整"
                });
                if (!boardOk || !stgOk)
                {
                    return DeepResult(dashboardSheet, checks, themeName);
                }
                board = workbook.Worksheets.Item[dashboardSheet];
                staging = workbook.Worksheets.Item[stagingName];

                var theme = DetectBoardTheme(board);
                themeName = ReferenceEquals(theme, ThemeColors["orange"]) ? "orange"
                    : ReferenceEquals(theme, ThemeColors["blue"]) ? "blue"
                    : ReferenceEquals(theme, ThemeColors["graygreen"]) ? "graygreen" : "unknown";

                // ② 辅助表 KPI 重算（A 列聚合公式全数值）
                int kpiChecked = 0, kpiBad = 0;
                string firstKpiFormula = null;
                double firstKpiValue = 0;
                for (int r = 1; r <= 16; r++)
                {
                    // KPI 公式列已迁 B（A1 坏缓存规避）；先 B 后 A 兼容旧看板
                    dynamic cell = staging.Cells.Item[r, 2];
                    object v = cell.Value2;
                    string f = null;
                    try { f = Convert.ToString(cell.Formula); } catch { }
                    Release((object)cell);
                    if (string.IsNullOrEmpty(f))
                    {
                        cell = staging.Cells.Item[r, 1];
                        v = cell.Value2;
                        try { f = Convert.ToString(cell.Formula); } catch { }
                        Release((object)cell);
                    }
                    if (v == null && string.IsNullOrEmpty(f)) { break; }
                    bool hasFormula = !string.IsNullOrEmpty(f) && f.StartsWith("=", StringComparison.Ordinal);
                    bool pureRef = hasFormula && f.IndexOf("(") < 0;
                    if (hasFormula && !pureRef)
                    {
                        kpiChecked++;
                        if (firstKpiFormula == null) { firstKpiFormula = f; firstKpiValue = v is double dv ? dv : double.NaN; }
                        if (!(v is double)) { kpiBad++; }
                    }
                }
                checks.Add(new DeepCheckItem
                {
                    Item = "辅助表 KPI 重算",
                    Status = kpiBad == 0 ? "pass" : "fail",
                    Expected = "全部聚合公式重算为数值",
                    Actual = $"检查 {kpiChecked} 个 KPI 公式，异常 {kpiBad} 个",
                    Basis = "辅助表 SUMIFS/SUMPRODUCT 聚合公式重算值必须为数值（文本/错误值=联动公式失效）"
                });

                // ③ 看板卡片显示值守恒（凡引用辅助表的看板单元格，显示值==辅助表值）
                var used = board.UsedRange;
                int rows = Math.Min(Convert.ToInt32(used.Rows.Count), 60);
                int cols = Math.Min(Convert.ToInt32(used.Columns.Count), 30);
                Release((object)used);
                int cardRefs = 0, cardMismatch = 0;
                var mismatchSamples = new List<string>();
                int tamperSuspects = 0;
                var tamperSamples = new List<string>();
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        dynamic cell = board.Cells.Item[r, c];
                        string f = null;
                        object disp = null;
                        string numFmt = null;
                        try { f = Convert.ToString(cell.Formula); disp = cell.Value2; numFmt = Convert.ToString(cell.NumberFormat); } catch { }
                        Release((object)cell);
                        if (string.IsNullOrEmpty(f) || !f.StartsWith("=", StringComparison.Ordinal))
                        {
                            // 常量覆盖扫描：数字格式卡片被文本常量覆盖 = 公式被毁的篡改痕迹（③ 引用扫描对无公式格失明）
                            if (disp is string txt && txt.Trim().Length > 0
                                && !string.IsNullOrEmpty(numFmt) && numFmt != "@"
                                && (numFmt.IndexOf('0') >= 0 || numFmt.IndexOf('#') >= 0))
                            {
                                tamperSuspects++;
                                if (tamperSamples.Count < 3)
                                {
                                    string colL2 = c <= 26 ? ((char)(64 + c)).ToString() : "A" + (char)(64 + c - 26);
                                    tamperSamples.Add($"{colL2}{r} 常量「{txt}」（数字格式 {numFmt}）");
                                }
                            }
                            continue;
                        }
                        // 先剥离区间引用（带冒号），单格正则不会误配区间前半段造成假失配
                        string fScan = System.Text.RegularExpressions.Regex.Replace(f, @"'?[^'!\r\n:]+'?!(\$?[A-Z]{1,3}\$?[0-9]+):(\$?[A-Z]{1,3}\$?[0-9]+)", "");
                        foreach (System.Text.RegularExpressions.Match m in StagingRefRegex.Matches(fScan))
                        {
                            string refSheet = m.Groups[1].Value.Trim(new char[] { (char)39, ' ' });
                            bool targetsStaging = RefTargetsStaging(refSheet, stagingName);
                            bool calcStyle = refSheet.EndsWith("_calc", StringComparison.OrdinalIgnoreCase);
                            if (!targetsStaging && !calcStyle) { continue; }  // 非 calc 引用（标题/普通数据）不参与判定
                            cardRefs++;
                            if (!targetsStaging)
                            {
                                // 悬空 calc 引用：卡片引用了 *_calc 形态的名字但与实际辅助表不符 → fail
                                //（2026-09-11 实证：KPI 公式漏前导下划线→保存为外部簿引用[1]→显示恒"-"，
                                //  旧逻辑把不匹配归为"未发现引用"放行 WARN，缺陷穿透三层防线）
                                cardMismatch++;
                                if (mismatchSamples.Count < 3)
                                {
                                    string colL3 = c <= 26 ? ((char)(64 + c)).ToString() : "A" + (char)(64 + c - 26);
                                    mismatchSamples.Add($"{colL3}{r} 引用「{refSheet}」与辅助表「{stagingName}」不符（悬空引用）");
                                }
                                break;
                            }
                            {
                                string refAddr = m.Groups[2].Value.Replace("$", "");
                                dynamic tgt = staging.Range[refAddr];
                                object expectedV = tgt.Value2;
                                Release((object)tgt);
                                // 空位镜像特判：交叉表左上角等结构性空格被卡片引用时，Excel 显示 0/空属正常渲染
                                // （U-full B6 实证：辅助表 B2 空、镜像卡显示 0，非失真）
                                if (expectedV == null && (disp == null || (disp is double dz && dz == 0)))
                                {
                                    cardRefs++;
                                    break;
                                }
                                string e = NormCellText(expectedV);
                                string a = NormCellText(disp);
                                bool same = string.Equals(e, a, StringComparison.Ordinal)
                                    || (double.TryParse(e, NumberStyles.Any, CultureInfo.InvariantCulture, out var de)
                                        && double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out var da)
                                        && Math.Abs(de - da) < 0.0001);
                                if (!same && TryLooseDisplayNumeric(a, out var da2) && double.TryParse(e, NumberStyles.Any, CultureInfo.InvariantCulture, out var de2))
                                {
                                    // 第三通道：格式化显示值（TEXT 卡/百分比/千分位，U31 实证 TEXT(...,"0.0%") 卡
                                    // 显示 93.7% vs 辅助表 0.9372——四舍五入损失属正常渲染，非失真）
                                    double tol = a.Contains("%") ? Math.Max(0.0011, Math.Abs(de2) * 0.0011) : 0.011;
                                    same = Math.Abs(de2 - da2) <= tol;
                                }
                                if (!same)
                                {
                                    cardMismatch++;
                                    if (mismatchSamples.Count < 3)
                                    {
                                        string colL = c <= 26 ? ((char)(64 + c)).ToString() : "A" + (char)(64 + c - 26);
                                        mismatchSamples.Add($"{colL}{r} 显示「{a}」≠ 辅助表 {refAddr}「{e}」");
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
                checks.Add(new DeepCheckItem
                {
                    Item = "看板卡片显示值守恒",
                    Status = cardRefs == 0 ? "warn" : (cardMismatch == 0 ? "pass" : "fail"),
                    Expected = "看板卡片显示值 == 辅助表被引用值",
                    Actual = cardRefs == 0 ? "未发现引用辅助表的卡片" : $"引用 {cardRefs} 处，不一致 {cardMismatch} 处" + (mismatchSamples.Count > 0 ? "；例：" + string.Join("；", mismatchSamples) : ""),
                    Basis = "卡片公式引用辅助表单元格，显示值必须等于辅助表值；IFERROR 回退「-」也算失真（历史缺陷审查⑨同形态）"
                });
                checks.Add(new DeepCheckItem
                {
                    Item = "数字卡片篡改扫描",
                    Status = tamperSuspects == 0 ? "pass" : "fail",
                    Expected = "数字格式的看板单元格必须是公式或数值，不得为文本常量",
                    Actual = tamperSuspects == 0 ? "未发现常量覆盖" : $"疑似篡改 {tamperSuspects} 处；例：" + string.Join("；", tamperSamples),
                    Basis = "卡片公式被手工常量覆盖后公式消失，③ 的引用扫描对其失明；数字格式+文本常量是公式被毁的直接痕迹"
                });
                // ④ 联动控件（信息项：下拉/切片器存在性）
                bool hasDropdown = false;
                try
                {
                    dynamic ctrl = board.Range["B2"].Validation;
                    hasDropdown = ctrl != null;
                    Release((object)ctrl);
                }
                catch { }
                int slicerCount = 0;
                try
                {
                    dynamic slicers = board.Slicers();
                    slicerCount = Convert.ToInt32(slicers.Count);
                    Release((object)slicers);
                }
                catch { }
                checks.Add(new DeepCheckItem
                {
                    Item = "联动控件",
                    Status = (hasDropdown || slicerCount > 0) ? "pass" : "warn",
                    Expected = "下拉验证或切片器至少其一（部分模板无联动）",
                    Actual = $"下拉={hasDropdown}，切片器={slicerCount} 个",
                    Basis = "联动控件让 KPI/图表随筛选重算；无联动仅提示，不判失败"
                });

                // ⑤ 数值守恒对账：控制视图为全量时，首个 KPI 重算值 == 源数据引用区域直和
                try
                {
                    if (!string.IsNullOrEmpty(firstKpiFormula) && !double.IsNaN(firstKpiValue))
                    {
                        // 从 KPI 公式提取源数据区域引用（非辅助表的跨表区间）
                        var srcMatches = StagingRangeRegex.Matches(firstKpiFormula)
                            .Cast<System.Text.RegularExpressions.Match>()
                            .Where(m => !RefTargetsStaging(m.Groups[1].Value, stagingName))
                            .Select(m => new { Sheet = m.Groups[1].Value.Trim('\'', ' '), Addr = m.Groups[2].Value + ":" + m.Groups[3].Value })
                            .Distinct()
                            .ToList();
                        // 控制视图判断：找公式中的控制单元格引用（看板!$B$2 类）读其值
                        string controlValue = "全部";
                        var ctrlMatch = System.Text.RegularExpressions.Regex.Match(firstKpiFormula, @"'([^']+)'!(\$B\$2|\$B\$3)");
                        if (ctrlMatch.Success)
                        {
                            dynamic ctrlCell = workbook.Worksheets.Item[ctrlMatch.Groups[1].Value].Range[ctrlMatch.Groups[2].Value.Replace("$", "")];
                            controlValue = Convert.ToString(ctrlCell.Value2);
                            Release((object)ctrlCell);
                        }
                        if (controlValue == "全部" || string.IsNullOrEmpty(controlValue) || controlValue == "(全部)")
                        {
                            bool conserved = false;
                            string usedRangeBasis = "无候选区域";
                            foreach (var cand in srcMatches.Take(3))
                            {
                                dynamic wsSrc = workbook.Worksheets.Item[cand.Sheet];
                                dynamic rng = wsSrc.Range[cand.Addr];
                                object raw = rng.Value2;
                                Release((object)rng);
                                Release((object)wsSrc);
                                if (!(raw is object[,] arr2))
                                {
                                    if (raw is double scalar) { arr2 = new object[1, 1] { { scalar } }; }
                                    else { continue; }
                                }
                                double sum = 0;
                                // COM Value2 数组 1 基（标量包装为 0 基 [1,1]）——GetLowerBound 自适应；
                                // 原实现 arr2[r2-1, c2-1] 只累加了首元素（区间直和恒错）
                                for (int r2 = arr2.GetLowerBound(0); r2 <= arr2.GetUpperBound(0); r2++)
                                {
                                    for (int c2 = arr2.GetLowerBound(1); c2 <= arr2.GetUpperBound(1); c2++)
                                    {
                                        if (arr2[r2, c2] is double d2) { sum += d2; }
                                    }
                                }
                                usedRangeBasis = cand.Sheet + "!" + cand.Addr + " 直和=" + Math.Round(sum, 4);
                                if (Math.Abs(sum - firstKpiValue) < 0.01) { conserved = true; break; }
                            }
                            checks.Add(new DeepCheckItem
                            {
                                Item = "数值守恒对账（全量视图）",
                                Status = conserved ? "pass" : "fail",
                                Expected = "KPI 重算值 " + Math.Round(firstKpiValue, 4) + " == 源数据区域直和",
                                Actual = usedRangeBasis + (conserved ? "（一致）" : "（不一致）"),
                                Basis = "全量视图下，辅助表聚合值必须等于源数据数值列直和（数据门槛级对账，防『看板数字与源数据不符』）"
                            });
                        }
                        else
                        {
                            checks.Add(new DeepCheckItem
                            {
                                Item = "数值守恒对账",
                                Status = "skip",
                                Expected = "控制视图=全量时对账",
                                Actual = "当前为筛选视图（" + controlValue + "），跳过直和比对",
                                Basis = "筛选语义下的聚合值≠全量直和属正常，不判失败"
                            });
                        }
                    }
                }
                catch (Exception exCons)
                {
                    checks.Add(new DeepCheckItem { Item = "数值守恒对账", Status = "skip", Expected = "-", Actual = "对账内部异常：" + exCons.Message, Basis = "异常不阻断其余检查" });
                }

                // ⑥ 图表渲染像素核验（空白图如实暴露）
                if (checkPixels)
                {
                    dynamic chartObjects = board.ChartObjects();
                    int n = Convert.ToInt32(chartObjects.Count);
                    Release((object)chartObjects);
                    for (int i = 1; i <= Math.Min(n, 6); i++)
                    {
                        dynamic co = null;
                        dynamic chart = null;
                        string pngPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "deep_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
                        try
                        {
                            co = board.ChartObjects().Item(i);
                            chart = co.Chart;
                            chart.Export(pngPath, "PNG");
                            var stats = ExcelChartService.ComputeRenderStats(System.IO.File.ReadAllBytes(pngPath));
                            bool rendered = stats.Item1;
                            double ratio = stats.Item2;
                            checks.Add(new DeepCheckItem
                            {
                                Item = $"图表渲染核验（{Convert.ToString(co.Name)}）",
                                Status = rendered ? "pass" : "warn",
                                Expected = "非白内容像素占比 ≥1%",
                                Actual = $"非白占比 {ratio:0.0%}（采样 {stats.Item3} 点）",
                                Basis = rendered
                                    ? "像素级渲染核验：图表真实画出内容"
                                    : "像素级渲染核验：图表内容几乎空白——数据可能未真实显示（层级图树状/旭日为已知引擎限制，建议改条形 Top10/环形替代）"
                            });
                        }
                        catch (Exception exPix)
                        {
                            checks.Add(new DeepCheckItem { Item = $"图表渲染核验（第 {i} 图）", Status = "skip", Expected = "-", Actual = "导出失败：" + exPix.Message, Basis = "导出失败不影响其余检查" });
                        }
                        finally
                        {
                            Release((object)chart);
                            Release((object)co);
                            try { if (System.IO.File.Exists(pngPath)) { System.IO.File.Delete(pngPath); } } catch { }
                        }
                    }
                }

                ExcelTrace.Info("verify_dashboard_deep", "核验完成 verdict=" + (checks.Any(c => c.Status == "fail") ? "mismatch" : "ok") + " checks=" + checks.Count);
                return DeepResult(dashboardSheet, checks, themeName);
            }
            finally
            {
                // board/staging/workbook 为应用级共享对象：不释放（M-D1 RCW 纪律，同 VerifyDashboard）
            }
        }
        /// <summary>
        /// FR-6.2：suggest_chart_mapping 图表选型建议（只读）。
        /// 分析选区列画像（数值比/去重基数/日期特征/负值），输出建议清单，每条附理由与数据要求；
        /// AI 应在生成图表/看板前调用，把建议与理由复述给用户并确认「哪个做数据源、哪个做行、哪个做列」。
        /// 采样上限 2000 行（性能护栏，NFR-2 同源纪律）。
        /// </summary>
        public ToolResult SuggestChartMapping(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var sheetName = Optional(args, "sheet");
            var maxSuggestions = ExcelToolHostHelpers.GetIntSafe(args, "max_suggestions", 5);
            dynamic workbook = null;
            dynamic sheetSource = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheetSource = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int headerRowAbs = Convert.ToInt32(rangeSource.Row);
                int firstColAbs = Convert.ToInt32(rangeSource.Column);
                int colCount = Convert.ToInt32(rangeSource.Columns.Count);
                int dataRows = Convert.ToInt32(rangeSource.Rows.Count) - 1;
                Release((object)rangeSource);
                if (dataRows < 1)
                {
                    return Invalid("NO_DATA_ROWS", "区域只有表头行（或为空），无数据可分析。");
                }
                int sampledRows = Math.Min(dataRows, 2000);
                dynamic blkRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRowAbs + 1, firstColAbs],
                    sheetSource.Cells.Item[headerRowAbs + sampledRows, firstColAbs + colCount - 1]];
                object raw = blkRange.Value2;
                Release((object)blkRange);
                // 单格区域 Value2 返回标量——防御性包装（UsedRange 块读同款教训）
                object[,] blk = raw is object[,] arr ? arr : new object[1, 1] { { raw } };
                var full = new object[sampledRows + 1, colCount];
                for (int c = 1; c <= colCount; c++)
                {
                    dynamic hc = sheetSource.Cells.Item[headerRowAbs, firstColAbs + c - 1];
                    full[0, c - 1] = hc.Value2;
                    Release((object)hc);
                }
                for (int r = 1; r <= sampledRows; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        full[r, c - 1] = blk[r, c];
                    }
                }
                var suggestions = AnalyzeChartMapping(full, out var cols);
                string note = sampledRows < dataRows
                    ? $"数据共 {dataRows} 行，建议基于前 {sampledRows} 行采样（统计口径，不影响后续渲染全量读取）"
                    : $"数据共 {dataRows} 行";
                recorder?.Invoke("suggest_chart_mapping", args, "mapping-suggest",
                    new { sheet = Convert.ToString(sheetSource.Name), columns = colCount, rows = sampledRows, suggestions = suggestions.Count },
                    null);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        source_address = sourceAddress,
                        sheet = Convert.ToString(sheetSource.Name),
                        sampled_rows = sampledRows,
                        columns = cols.Select(p => new
                        {
                            name = p.Name,
                            rows = p.Rows,
                            non_empty = p.NonEmpty,
                            numeric = p.Numeric,
                            numeric_ratio = Math.Round(p.NumericRatio, 2),
                            distinct_text = p.DistinctText,
                            date_like = p.DateLike,
                            has_negative = p.HasNegative
                        }),
                        suggestions = suggestions.Take(Math.Max(1, maxSuggestions)).Select(s => new
                        {
                            tool = s.Tool,
                            chart_type = s.ChartType,
                            reason = s.Reason,
                            requires = s.Requires
                        }),
                        note,
                        usage = "把建议与理由复述给用户，确认数据源/行/列映射后再调用渲染工具"
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return Invalid("MAPPING_FAILED",
                    "选型分析 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message);
            }
            catch (Exception ex)
            {
                return Invalid("MAPPING_FAILED", "选型分析失败：" + ex.Message);
            }
            finally
            {
                Release((object)sheetSource);
                Release((object)workbook);
            }
        }

        // ---- FR-ED 看板编辑器 P1（docs/需求-看板编辑器.md v0.2）----

        /// <summary>看板身份校验：看板 sheet 存在且配套隐藏辅助表存在（防误改普通工作表）。</summary>
        private ToolResult CheckEditableBoard(dynamic workbook, string dashboardSheet)
        {
            if (!SheetExists(workbook, dashboardSheet))
            {
                return Invalid("DASHBOARD_NOT_FOUND", $"找不到看板「{dashboardSheet}」。");
            }
            if (!SheetExists(workbook, "_" + EscapeSheetInner(dashboardSheet) + "_calc"))
            {
                return Invalid("NOT_A_DASHBOARD",
                    $"「{dashboardSheet}」没有配套隐藏辅助表（_看板名_calc），不是本助手生成的看板，编辑器拒绝操作以防误改普通工作表。");
            }
            return null;
        }

        /// <summary>识别看板当前主题：扫描 UsedRange 特征色（CardBg/TitleText/ValueText），命中最多且 ≥3 的主题；无法识别返回 null。</summary>
        private ThemePalette DetectBoardTheme(dynamic sheetBoard)
        {
            var used = sheetBoard.UsedRange;
            int rows = Math.Min(Convert.ToInt32(used.Rows.Count), 40);
            int cols = Math.Min(Convert.ToInt32(used.Columns.Count), 20);
            Release((object)used);
            var scores = new Dictionary<ThemePalette, int>();
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    dynamic cell = sheetBoard.Cells.Item[r, c];
                    int interior = 0, font = 0;
                    try { interior = Convert.ToInt32(cell.Interior.Color); } catch { }
                    try { font = Convert.ToInt32(cell.Font.Color); } catch { }
                    Release((object)cell);
                    foreach (var kv in ThemeColors)
                    {
                        int hit = 0;
                        if (interior == kv.Value.CardBg) { hit++; }
                        if (font == kv.Value.TitleText || font == kv.Value.ValueText) { hit++; }
                        if (hit > 0)
                        {
                            scores.TryGetValue(kv.Value, out int prev);
                            scores[kv.Value] = prev + hit;
                        }
                    }
                }
            }
            ThemePalette best = null;
            int bestScore = 2; // 至少 3 个特征命中才认定（防手工改色的偶然碰撞）
            foreach (var kv in scores)
            {
                if (kv.Value > bestScore) { bestScore = kv.Value; best = kv.Key; }
            }
            return best;
        }

        /// <summary>按调色板映射重刷看板特征色（CardBg/TitleText/ValueText/NoteText），返回改写属性数。图表系列配色不在范围（A2 拍板）。</summary>
        private int RecolorBoard(dynamic sheetBoard, ThemePalette from, ThemePalette to)
        {
            var map = new Dictionary<int, int>
            {
                [from.CardBg] = to.CardBg,
                [from.TitleText] = to.TitleText,
                [from.ValueText] = to.ValueText,
                [from.NoteText] = to.NoteText
            };
            var used = sheetBoard.UsedRange;
            int rows = Math.Min(Convert.ToInt32(used.Rows.Count), 40);
            int cols = Math.Min(Convert.ToInt32(used.Columns.Count), 20);
            Release((object)used);
            int changed = 0;
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    dynamic cell = sheetBoard.Cells.Item[r, c];
                    try
                    {
                        int interior = Convert.ToInt32(cell.Interior.Color);
                        if (map.TryGetValue(interior, out int newInterior))
                        {
                            cell.Interior.Color = newInterior;
                            changed++;
                        }
                        int font = Convert.ToInt32(cell.Font.Color);
                        if (map.TryGetValue(font, out int newFont))
                        {
                            cell.Font.Color = newFont;
                            changed++;
                        }
                    }
                    catch { }
                    Release((object)cell);
                }
            }
            return changed;
        }
        /// <summary>FR-E0：list_dashboard_charts 看板图表清单（只读）。编辑前必调——AI 据此向用户复述现状。</summary>
        public ToolResult ListDashboardCharts(Dictionary<string, object> args)
        {
            var dashboardSheet = Required(args, "dashboard_sheet");
            dynamic workbook = null;
            dynamic sheetBoard = null;
            dynamic chartObjects = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var gate = CheckEditableBoard(workbook, dashboardSheet);
                if (gate != null) { return gate; }
                sheetBoard = workbook.Worksheets.Item[dashboardSheet];
                chartObjects = sheetBoard.ChartObjects();
                int n = Convert.ToInt32(chartObjects.Count);
                var charts = new List<object>();
                for (int i = 1; i <= n; i++)
                {
                    dynamic co = chartObjects.Item(i);
                    int chartType = 0, seriesCount = 0;
                    string anchor = "";
                    try
                    {
                        dynamic chart = co.Chart;
                        chartType = Convert.ToInt32(chart.ChartType);
                        dynamic sc = chart.SeriesCollection();
                        seriesCount = Convert.ToInt32(sc.Count);
                        Release((object)sc);
                        dynamic tl = co.TopLeftCell;
                        anchor = Convert.ToString(tl.Address);
                        Release((object)tl);
                    }
                    catch { }
                    charts.Add(new { name = Convert.ToString(co.Name), chart_type = chartType, series_count = seriesCount, anchor });
                    Release((object)co);
                }
                var theme = DetectBoardTheme(sheetBoard);
                string themeName = ReferenceEquals(theme, ThemeColors["orange"]) ? "orange"
                    : ReferenceEquals(theme, ThemeColors["blue"]) ? "blue"
                    : ReferenceEquals(theme, ThemeColors["graygreen"]) ? "graygreen" : "unknown";
                recorder?.Invoke("list_dashboard_charts", args, "dashboard-inspect",
                    new { sheet = dashboardSheet, charts = n, theme = themeName }, null);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        theme = themeName,
                        chart_count = n,
                        charts,
                        usage = "把图表清单复述给用户，删除时 chart_name 必须逐字取自本清单"
                    }
                };
            }
            finally
            {
                Release((object)chartObjects);
                Release((object)sheetBoard);
                Release((object)workbook);
            }
        }

        /// <summary>FR-E1：edit_dashboard_theme 换主题（A2 拍板：只刷卡片底色/字色，图表系列配色不在范围）。特征色映射重刷，不动任何公式与数值。</summary>
        public ToolResult EditDashboardTheme(Dictionary<string, object> args)
        {
            var dashboardSheet = Required(args, "dashboard_sheet");
            var theme = (Required(args, "theme") ?? string.Empty).Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var newPalette))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            dynamic workbook = null;
            dynamic sheetBoard = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var gate = CheckEditableBoard(workbook, dashboardSheet);
                if (gate != null) { return gate; }
                sheetBoard = workbook.Worksheets.Item[dashboardSheet];
                var oldPalette = DetectBoardTheme(sheetBoard);
                if (oldPalette == null)
                {
                    return Invalid("THEME_NOT_RECOGNIZED",
                        "看板当前配色不属于内置主题（可能已被手工修改），换主题工具拒绝执行以防误刷。可整板重生成恢复标准主题。");
                }
                string oldName = ReferenceEquals(oldPalette, ThemeColors["orange"]) ? "orange"
                    : ReferenceEquals(oldPalette, ThemeColors["blue"]) ? "blue"
                    : ReferenceEquals(oldPalette, ThemeColors["graygreen"]) ? "graygreen" : "unknown";
                if (ReferenceEquals(oldPalette, newPalette))
                {
                    return Invalid("THEME_SAME", $"看板已是 {theme} 主题，无需变更。");
                }
                int changed = RecolorBoard(sheetBoard, oldPalette, newPalette);
                var selfcheck = RunPostRenderSelfCheck(dashboardSheet);
                recorder?.Invoke("edit_dashboard_theme", args, "dashboard-edit",
                    new { sheet = dashboardSheet, from = oldName, to = theme, cells_changed = changed },
                    new { Sheet = dashboardSheet, Theme = theme });
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        old_theme = oldName,
                        new_theme = theme,
                        cells_changed = changed,
                        note = "仅重刷卡片底色/字色；图表系列配色与全部数值不受影响（AC-E7 数值守恒）",
                        selfcheck = (selfcheck.Verdict == "ok" ? "自检通过 ✓ " : "自检异常 ✗ ") + selfcheck.Detail,
                        selfcheck_verdict = selfcheck.Verdict,
                        changed = true
                    }
                };
            }
            finally
            {
                Release((object)sheetBoard);
                Release((object)workbook);
            }
        }
        /// <summary>FR-E3：delete_dashboard_chart 删单图（A4 拍板：只删 ChartObject，切片器/下拉不在此范围）。删后强制 AC-E7 守恒检查（恰好 -1）+ 自检回显。</summary>
        public ToolResult DeleteDashboardChart(Dictionary<string, object> args)
        {
            var dashboardSheet = Required(args, "dashboard_sheet");
            var chartName = (Required(args, "chart_name") ?? string.Empty).Trim();
            dynamic workbook = null;
            dynamic sheetBoard = null;
            dynamic chartObjects = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var gate = CheckEditableBoard(workbook, dashboardSheet);
                if (gate != null) { return gate; }
                sheetBoard = workbook.Worksheets.Item[dashboardSheet];
                chartObjects = sheetBoard.ChartObjects();
                int before = Convert.ToInt32(chartObjects.Count);
                if (before < 1)
                {
                    return Invalid("CHART_NOT_FOUND", $"看板「{dashboardSheet}」上没有图表。");
                }
                dynamic target = null;
                try { target = sheetBoard.ChartObjects(chartName); } catch { }
                if (target == null)
                {
                    return Invalid("CHART_NOT_FOUND",
                        $"看板上没有名为「{chartName}」的图表。请先调 list_dashboard_charts 枚举，chart_name 必须逐字取自清单。");
                }
                target.Delete();
                Release((object)target);
                dynamic afterColl = sheetBoard.ChartObjects();
                int after = Convert.ToInt32(afterColl.Count);
                Release((object)afterColl);
                // AC-E7 守恒：删除必须恰好 -1，否则报错（不留半成品状态）
                if (after != before - 1)
                {
                    return Invalid("DELETE_COUNT_MISMATCH",
                        $"删除后图表数 {after} ≠ 删除前 {before} - 1，违反守恒断言，请检查看板状态。");
                }
                var selfcheck = RunPostRenderSelfCheck(dashboardSheet);
                recorder?.Invoke("delete_dashboard_chart", args, "dashboard-edit",
                    new { sheet = dashboardSheet, chart = chartName, before, after },
                    new { Sheet = dashboardSheet, ChartName = chartName, Deleted = true });
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        chart_name = chartName,
                        charts_before = before,
                        charts_after = after,
                        note = "仅删除该图表对象；切片器/下拉与辅助表不在范围（整板清理走 delete_dashboard）",
                        selfcheck = (selfcheck.Verdict == "ok" ? "自检通过 ✓ " : "自检异常 ✗ ") + selfcheck.Detail,
                        selfcheck_verdict = selfcheck.Verdict,
                        deleted = true
                    }
                };
            }
            finally
            {
                Release((object)chartObjects);
                Release((object)sheetBoard);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// FR-5.1：pnl_brief 财务损益看板。
        /// 数据约定：source_address 三列——科目 | 金额 | 类型（type：收入 revenue / 成本 cost，首行表头）。
        /// 输出：KPI 卡 4 张（总收入/总成本/费用合计/净利润）+ 科目明细表（收入绿色/成本红色数据条）
        /// + 净利润瀑布式堆叠条（收入正段+成本负段同轴对比）。源只读；幂等；失败回滚。
        /// </summary>
        /// <summary>
        /// 超级数据表（用户拍板方向，2026-09-12）：看板数据底座——把明细源规范化为一张可见 sheet：
        /// A:E 明细长表（类型规整：月份列锁文本防类型漂移、文本型数字数值化）+ 右侧派生汇总区
        /// （G:I 月度、K:M 品类、O:P KPI、R:S 地区，全部同表裸地址 SUMIF/SUM/COUNT 公式，
        /// 引用一律裸地址——工作簿名/跨表限定自引用会被 Excel 规范化指错 sheet，见踩坑 9.x.7）。
        /// 后续 render_dashboard 的 source_address 指向本表明细区，多看板共享同一数据源；
        /// set_linkage 切片器语义不变（透视源=本表明细区）。源数据只读。
        /// </summary>
        public ToolResult CreateSuperTable(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 1);
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 2);
            var monthColumn = ExcelToolHostHelpers.GetIntSafe(args, "month_column", 3);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var quantityColumn = ExcelToolHostHelpers.GetIntSafe(args, "quantity_column", 5);
            var superSheet = (Optional(args, "super_sheet") ?? "超级数据表").Trim();
            if (string.IsNullOrWhiteSpace(superSheet))
            {
                return Invalid("INVALID_NAME", "super_sheet 不能为空。");
            }
            var colParams = new[] { labelColumn, categoryColumn, monthColumn, valueColumn, quantityColumn };
            if (colParams.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic rangeSource = null;
            dynamic sheetOld = null;
            dynamic sheetSuper = null;
            dynamic rangeTmp = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sourceSheetName = Optional(args, "sheet");
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                rangeSource = sheetSource.Range[sourceAddress];

                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                if (colParams.Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"列参数超界：区域列数 {srcCols}，最大传入 {colParams.Max()}。");
                }
                int rows = Convert.ToInt32(rangeSource.Rows.Count) - 1;  // 去表头行
                if (rows < 1)
                {
                    return Invalid("NO_DATA", "数据区域只有表头行，没有明细数据。");
                }

                // 幂等重建：旧超级数据表（标记 A1=地区）删除覆盖；名字被其他表占用则拒绝（防误删用户表）
                if (SheetExists(workbook, superSheet))
                {
                    sheetOld = workbook.Worksheets.Item[superSheet];
                    string marker = Convert.ToString(sheetOld.Range["A1"].Value2);
                    if (marker != "地区")
                    {
                        return Invalid("SUPER_TABLE_NAME_TAKEN",
                            $"工作簿已存在同名工作表「{superSheet}」且不是超级数据表。请换 super_sheet 名称，或确认后手动删除该表再重试。");
                    }
                    sheetOld.Delete();
                    sheetOld = null;
                }

                // 源块读（NFR-2 纪律：禁逐格 COM 扫描）
                object[,] src = rangeSource.Value2 as object[,];
                if (src == null)
                {
                    return Invalid("NO_DATA", "数据区域为空或只有单个单元格。");
                }

                // 规范化明细（内存）：null→空、文本型数字数值化、维度值保持原类型
                var detail = new object[rows, 5];
                var monthSet = new SortedSet<string>(StringComparer.Ordinal);
                var catSet = new SortedSet<string>(StringComparer.Ordinal);
                var labelSet = new SortedSet<string>(StringComparer.Ordinal);
                bool monthsAllText = true;
                double sumValue = 0, sumQty = 0, maxValue = 0;
                long cntValue = 0;
                for (int r = 2; r <= rows + 1; r++)  // src 1 基：src[1]=表头行，明细从 src[2] 起（1 基错位曾致表头入明细+末行丢失）
                {
                    detail[r - 2, 0] = src[r, labelColumn] ?? "";
                    detail[r - 2, 1] = src[r, categoryColumn] ?? "";
                    detail[r - 2, 2] = src[r, monthColumn] ?? "";
                    if (detail[r - 2, 2] is string) { monthSet.Add((string)detail[r - 2, 2]); }
                    else if (detail[r - 2, 2] != null && !"".Equals(detail[r - 2, 2])) { monthSet.Add(Convert.ToString(detail[r - 2, 2], CultureInfo.InvariantCulture)); monthsAllText = false; }
                    else { monthSet.Add(""); }
                    var cv = detail[r - 2, 1];
                    catSet.Add(cv == null ? "" : Convert.ToString(cv, CultureInfo.InvariantCulture));
                    var lv = detail[r - 2, 0];
                    labelSet.Add(lv == null ? "" : Convert.ToString(lv, CultureInfo.InvariantCulture));
                    detail[r - 2, 3] = NormalizeNumber(src[r, valueColumn], ref sumValue, ref cntValue, ref maxValue);
                    detail[r - 2, 4] = NormalizeNumber(src[r, quantityColumn], ref sumQty, ref cntValue, ref maxValue);
                }

                sheetSuper = workbook.Worksheets.Add();
                sheetSuper.Name = superSheet;

                // 表头（标记格 A1=地区，幂等重建识别用）
                sheetSuper.Range["A1:E1"].Value2 = new object[,] { { "地区", "品类", "月份", "销售额", "销量" } };
                // 月份列锁文本（类型漂移教训：写入前锁 NumberFormat=@，防 "2025-11" 被自动转日期序列号）
                if (monthsAllText)
                {
                    rangeTmp = sheetSuper.Range["C2:C" + (rows + 1)];
                    rangeTmp.NumberFormat = "@";
                    Release((object)rangeTmp); rangeTmp = null;
                }
                // 明细块写（rows==1 时 1×N 区块写 object[,] 兼容性差，逐格写）
                if (rows == 1)
                {
                    for (int c = 0; c < 5; c++) { sheetSuper.Cells.Item[2, c + 1].Value2 = detail[0, c]; }
                }
                else
                {
                    sheetSuper.Range["A2:E" + (rows + 1)].Value2 = detail;
                }

                // ---- 派生汇总区（同表裸地址公式；SUMIF criteria 用单元格引用保持同型匹配）----
                string dRef = $"$D$2:$D${rows + 1}";
                string qRef = $"$E$2:$E${rows + 1}";
                string lRef = $"$A$2:$A${rows + 1}";
                string cRef = $"$B$2:$B${rows + 1}";
                string mRef = $"$C$2:$C${rows + 1}";

                // 月度汇总 G:I
                sheetSuper.Range["G1:I1"].Value2 = new object[,] { { "月份", "销售额", "销量" } };
                int mi2 = 2;
                foreach (var m in monthSet)
                {
                    if (m == "") { continue; }
                    sheetSuper.Cells.Item[mi2, 7].Value2 = monthsAllText ? (object)m : (object)ParseDoubleSafe(m);
                    sheetSuper.Cells.Item[mi2, 8].Formula = $"=SUMIF({mRef},$G${mi2},{dRef})";
                    sheetSuper.Cells.Item[mi2, 9].Formula = $"=SUMIF({mRef},$G${mi2},{qRef})";
                    mi2++;
                }
                // 品类汇总 K:M
                sheetSuper.Range["K1:M1"].Value2 = new object[,] { { "品类", "销售额", "占比" } };
                int ci2 = 2;
                foreach (var c in catSet)
                {
                    if (c == "") { continue; }
                    sheetSuper.Cells.Item[ci2, 11].Value2 = c;
                    sheetSuper.Cells.Item[ci2, 12].Formula = $"=SUMIF({cRef},$K${ci2},{dRef})";
                    sheetSuper.Cells.Item[ci2, 13].Formula = $"=IFERROR(L{ci2}/SUM({dRef}),0)";
                    ci2++;
                }
                if (ci2 > 2)
                {
                    rangeTmp = sheetSuper.Range["M2:M" + (ci2 - 1)];
                    rangeTmp.NumberFormat = "0.0%";
                    Release((object)rangeTmp); rangeTmp = null;
                }
                // KPI 汇总 O:P（全量口径；最大单笔=明细最大值，零数组写法）
                sheetSuper.Range["O1:P1"].Value2 = new object[,] { { "指标", "值" } };
                string[] kpiNames = { "销售总额", "成交笔数", "平均客单价", "最大单笔" };
                string[] kpiFmls =
                {
                    $"=SUM({dRef})",
                    $"=COUNT({dRef})",
                    $"=IFERROR(SUM({dRef})/COUNT({dRef}),0)",
                    $"=MAX({dRef})"
                };
                for (int k = 0; k < 4; k++)
                {
                    sheetSuper.Cells.Item[k + 2, 15].Value2 = kpiNames[k];
                    sheetSuper.Cells.Item[k + 2, 16].Formula = kpiFmls[k];
                }
                // 地区清单 R:S
                sheetSuper.Range["R1:S1"].Value2 = new object[,] { { "地区", "销售额" } };
                int ri2 = 2;
                foreach (var l in labelSet)
                {
                    if (l == "") { continue; }
                    sheetSuper.Cells.Item[ri2, 18].Value2 = l;
                    sheetSuper.Cells.Item[ri2, 19].Formula = $"=SUMIF({lRef},$R${ri2},{dRef})";
                    ri2++;
                }

                // 表头样式 + 列宽
                foreach (var addr in new[] { "A1:E1", "G1:I1", "K1:M1", "O1:P1", "R1:S1" })
                {
                    rangeTmp = sheetSuper.Range[addr];
                    rangeTmp.Font.Bold = true;
                    rangeTmp.Interior.Color = 0xF0DCB8;  // 浅色主题底（BGR）
                    Release((object)rangeTmp); rangeTmp = null;
                }
                rangeTmp = sheetSuper.UsedRange;
                rangeTmp.Columns.AutoFit();
                Release((object)rangeTmp); rangeTmp = null;

                double avg = cntValue > 0 ? sumValue / cntValue : 0;
                ExcelTrace.Info("create_super_table",
                    $"rows={rows} months={monthSet.Count} categories={catSet.Count} regions={labelSet.Count} total={sumValue}");

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "super_table",
                        super_sheet = superSheet,
                        detail_range = $"A1:E{rows + 1}",
                        rows = rows,
                        months = monthSet.Count,
                        categories = catSet.Count,
                        regions = labelSet.Count,
                        summary_ranges = new
                        {
                            monthly = $"G1:I{Math.Max(2, mi2 - 1)}",
                            category = $"K1:M{Math.Max(2, ci2 - 1)}",
                            kpi = "O1:P5",
                            region = $"R1:S{Math.Max(2, ri2 - 1)}"
                        },
                        kpis = new { total = sumValue, count = cntValue, avg = avg, max = maxValue },
                        usage = $"render_dashboard 传 sheet=\"{superSheet}\"、source_address=\"A1:E{rows + 1}\" 即可让看板共享本表数据源",
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetSuper);
                return Invalid("SUPER_TABLE_FAILED",
                    "超级数据表生成 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message);
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetSuper);
                ExcelTrace.Error("create_super_table", "CreateSuperTable", ex);
                return Invalid("SUPER_TABLE_FAILED", "超级数据表生成失败：" + ex.Message + "。源数据未受影响。");
            }
            finally
            {
                Release((object)rangeTmp);
                Release((object)rangeSource);
                Release((object)sheetSource);
                Release((object)sheetOld);
                Release((object)sheetSuper);
            }
        }

        /// <summary>超级数据表：单值规范化——文本型数字转 double（失败保留原文），并累计 KPI 统计。</summary>
        private static object NormalizeNumber(object v, ref double sum, ref long cnt, ref double max)
        {
            if (v == null) { return ""; }
            if (v is double d)
            {
                sum += d; cnt++; if (d > max) { max = d; }
                return d;
            }
            if (double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture),
                System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
            {
                sum += parsed; cnt++; if (parsed > max) { max = parsed; }
                return parsed;
            }
            return v;  // 非数值文本原样保留（可见可查，SUM/COUNT 自动忽略）
        }

        /// <summary>超级数据表：月份标签文本还原 double（数值型月份写回用）。</summary>
        private static object ParseDoubleSafe(string s)
        {
            return double.TryParse(s, System.Globalization.NumberStyles.Any,
                CultureInfo.InvariantCulture, out double d) ? (object)d : s;
        }

        public ToolResult RenderPnlBrief(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var sheetName = Optional(args, "sheet");
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "财务损益看板";
            var theme = (Optional(args, "theme") ?? "blue").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeCell = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                sheetSource = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int headerRow = Convert.ToInt32(rangeSource.Row);
                int lastRow = headerRow + Convert.ToInt32(rangeSource.Rows.Count) - 1;
                int colA = Convert.ToInt32(rangeSource.Column);
                Release((object)rangeSource);

                // 块读取：科目/金额/类型
                dynamic blkRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRow + 1, colA],
                    sheetSource.Cells.Item[lastRow, colA + 2]];
                object[,] blk = blkRange.Value2;
                Release((object)blkRange);
                var rows = new List<Tuple<string, double, string>>();
                for (int r = 1; r <= blk.GetLength(0); r++)
                {
                    object name = blk[r, 1];
                    object amt = blk[r, 2];
                    object typ = blk[r, 3];
                    if (name == null || amt == null || typ == null) { continue; }
                    if (!(amt is double d)) { continue; }
                    string ts = Convert.ToString(typ).Trim().ToLowerInvariant();
                    if (ts != "revenue" && ts != "cost")
                    {
                        return Invalid("INVALID_PNL_TYPE",
                            $"第 {r} 行类型「{ts}」无效：类型列仅支持 revenue（收入）/cost（成本）。");
                    }
                    rows.Add(Tuple.Create(Convert.ToString(name), d, ts));
                }
                if (rows.Count == 0)
                {
                    return Invalid("NO_NUMERIC_DATA",
                        "财务损益看板需三列（科目/金额/类型），类型列取值 revenue 或 cost，且至少一行有效。");
                }

                string stagingName = "_" + EscapeSheetInner(dashboardSheet) + "_calc";
                foreach (var nm in new[] { dashboardSheet, stagingName })
                {
                    if (SheetExists(workbook, nm))
                    {
                        dynamic old = workbook.Worksheets.Item[nm];
                        try { old.Delete(); } catch { }
                        Release((object)old);
                    }
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add();
                sheetStaging.Name = stagingName;

                double totalRevenue = rows.Where(x => x.Item3 == "revenue").Sum(x => x.Item2);
                double totalCost = rows.Where(x => x.Item3 == "cost").Sum(x => x.Item2);
                double profit = totalRevenue - totalCost;

                // staging：A 科目/B 金额/C 类型/D 占比（占各自类型合计）；G1~G4 KPI 值
                rangeCell = sheetStaging.Range["A1"]; rangeCell.Value2 = "科目"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["B1"]; rangeCell.Value2 = "金额"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["C1"]; rangeCell.Value2 = "类型"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["D1"]; rangeCell.Value2 = "占比"; Release((object)rangeCell);
                for (int i = 0; i < rows.Count; i++)
                {
                    int rr = i + 2;
                    double denom = rows[i].Item3 == "revenue" ? totalRevenue : totalCost;
                    // 2026-09-13 实证修正（AI 读图自审发现）：兜底条件误用 denom>0——负记账成本合计(-4350)
                    // 被当成除零写死分母 1，占比列退化为金额原值。带符号相除天然出正占比，仅需防真除零。
                    rangeCell = sheetStaging.Cells.Item[rr, 1]; rangeCell.Value2 = rows[i].Item1; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 2]; rangeCell.Value2 = rows[i].Item2; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 3]; rangeCell.Value2 = rows[i].Item3; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 4];
                    rangeCell.Formula = $"=IFERROR(B{rr}/{(denom > 0 ? denom.ToString(System.Globalization.CultureInfo.InvariantCulture) : "1")},0)";
                    Release((object)rangeCell);
                }
                rangeCell = sheetStaging.Range["G1"]; rangeCell.Value2 = totalRevenue; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["G2"]; rangeCell.Value2 = totalCost; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["G3"]; rangeCell.Value2 = totalCost; // 费用合计=成本合计（本模板口径）
                Release((object)rangeCell);
                rangeCell = sheetStaging.Range["G4"]; rangeCell.Value2 = profit; Release((object)rangeCell);

                // 看板：标题 + KPI 卡 4 张
                rangeCell = sheetBoard.Range["A1"]; rangeCell.Value2 = "财务损益看板"; Release((object)rangeCell);
                string[] kpiTitles = { "总收入", "总成本", "费用合计", "净利润" };
                string[] kpiFormulas = { "G1", "G2", "G3", "G4" };
                for (int i = 0; i < 4; i++)
                {
                    int r1 = 3 + i * 4;
                    dynamic titleCell = sheetBoard.Range[sheetBoard.Cells.Item[r1, 2], sheetBoard.Cells.Item[r1, 4]];
                    titleCell.Merge();
                    titleCell.Value2 = kpiTitles[i];
                    titleCell.Font.Bold = true;
                    titleCell.Font.Color = themeColors.TitleText;
                    Release((object)titleCell);
                    dynamic valueCell = sheetBoard.Range[sheetBoard.Cells.Item[r1 + 1, 2], sheetBoard.Cells.Item[r1 + 1, 4]];
                    valueCell.Merge();
                    valueCell.Formula = $"=IFERROR('{EscapeSheetInner(stagingName)}'!${kpiFormulas[i]},\"-\")";
                    valueCell.NumberFormat = "¥#,##0";
                    valueCell.Font.Size = 20;
                    valueCell.Font.Bold = true;
                    valueCell.Interior.Color = themeColors.CardBg;
                    Release((object)valueCell);
                }

                // 科目明细表（16 行起）：科目/金额/类型/占比，金额列数据条
                int headRow = 22;
                string[] headTexts = { "科目", "金额", "类型", "占比" };
                for (int c = 0; c < 4; c++)
                {
                    rangeCell = sheetBoard.Cells.Item[headRow, 2 + c];
                    rangeCell.Value2 = headTexts[c];
                    rangeCell.Font.Bold = true;
                    Release((object)rangeCell);
                }
                for (int i = 0; i < rows.Count; i++)
                {
                    int rr = headRow + 1 + i;
                    for (int c = 0; c < 4; c++)
                    {
                        string colL = new[] { "A", "B", "C", "D" }[c];
                        rangeCell = sheetBoard.Cells.Item[rr, 2 + c];
                        rangeCell.Formula = $"=IFERROR('{EscapeSheetInner(stagingName)}'!${colL}${i + 2},\"-\")";
                        Release((object)rangeCell);
                    }
                }
                dynamic amtColumn = sheetBoard.Range[sheetBoard.Cells.Item[headRow + 1, 3], sheetBoard.Cells.Item[headRow + rows.Count, 3]];
                dynamic fcAmt = amtColumn.FormatConditions.AddDatabar();
                try { fcAmt.BarColor.Color = 0x7B636B; } finally { Release((object)fcAmt); }
                Release((object)amtColumn);
                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }

                recorder?.Invoke("render_pnl_brief", args, "dashboard-render",
                    new { template = "pnl_brief", sheet = dashboardSheet, items = rows.Count, revenue = totalRevenue, cost = totalCost, profit },
                    new { Sheet = dashboardSheet, Template = "pnl_brief" });

                var selfcheckInfo = RunPostRenderSelfCheck(dashboardSheet);
                string selfcheckText = (selfcheckInfo.Verdict == "ok" ? "自检通过 ✓ " : selfcheckInfo.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheckInfo.Detail;
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "pnl_brief",
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = stagingName,
                        items = rows.Count,
                        total_revenue = totalRevenue,
                        total_cost = totalCost,
                        profit = profit,
                        theme = theme,
                        selfcheck = selfcheckText,
                        selfcheck_verdict = selfcheckInfo.Verdict,
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "pnl_brief 渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。已回滚。");
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED", "pnl_brief 渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                Release((object)rangeCell);
                Release((object)sheetBoard);
                Release((object)sheetStaging);
                Release((object)sheetSource);
            }
        }

        /// <summary>
        /// M-D3 补齐：delete_dashboard 一键删除看板（FR-1.4 删除侧 / AC-10 防误删）。
        /// 删除前回读断言目标 sheet 具备看板结构特征（存在配对隐藏辅助表 + 看板含引用该辅助表的公式
        /// 或控制单元格标记），不匹配则拒绝删除——防止把普通工作表当看板误删。
        /// 同时清理挂接在该看板的联动资产（SlicerCache，避免悬空切片器引用失效 sheet）。
        /// Dangerous 级：过确认门（NFR-3）。
        /// </summary>
        public ToolResult DeleteDashboard(Dictionary<string, object> args)
        {
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "销售看板";

            dynamic workbook = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");

                string stagingName = "_" + EscapeSheetInner(dashboardSheet) + "_calc";
                if (!SheetExists(workbook, dashboardSheet) || !SheetExists(workbook, stagingName))
                {
                    return Invalid("DASHBOARD_NOT_FOUND",
                        $"找不到看板「{dashboardSheet}」（需要看板 sheet 与隐藏辅助表 {stagingName} 同时存在）。");
                }

                // AC-10 回读断言：看板结构特征 = 看板 sheet 至少 1 个公式引用辅助表，且辅助表存在 KPI/矩阵内容。
                // 遍历用防御式枚举（M-D3 审查⑩：UsedRange.Formula 返回 1 基 COM 数组且可能非 object[,]，
                // 直接 cast 曾走 catch 单格回退而 A1 为空 → 断言恒 false 误拒真看板）。
                sheetBoard = workbook.Worksheets.Item[dashboardSheet];
                sheetStaging = workbook.Worksheets.Item[stagingName];
                bool boardLooksLikeDashboard = false;
                {
                    dynamic used = null;
                    try
                    {
                        used = sheetBoard.UsedRange;
                        // Excel 会把公式中带引号的 '_销售看板_calc' 规范化为 销售看板_calc（剥引号），
                        // 匹配须同时容忍两种形态：按「表名去首下划线」与「原名」各构造一个 needle
                        string needleQuoted = "'" + EscapeSheetInner(stagingName) + "'!";
                        string needleBare = EscapeSheetInner(stagingName).TrimStart('_') + "!";
                        string needleFull = EscapeSheetInner(stagingName) + "!";
                        // 带空格/特殊字符的 sheet 名：Excel 保留引号但仍剥首下划线 → '名 2_calc'!
                        string needleQuotedBare = "'" + EscapeSheetInner(stagingName).TrimStart('_') + "'!";
                        dynamic cells = used.Cells;
                        int uRows = Convert.ToInt32(used.Rows.Count);
                        int uCols = Convert.ToInt32(used.Columns.Count);
                        int scan = Math.Min(uRows * uCols, 5000); // 防御上限，看板元素有限
                        for (int i = 1; i <= scan && !boardLooksLikeDashboard; i++)
                        {
                            int r = (i - 1) / uCols + 1;
                            int c = (i - 1) % uCols + 1;
                            dynamic cell = null;
                            try
                            {
                                cell = cells.Item[r, c];
                                object f = cell.Formula;
                                if (f is string s && s.Length > 0 && s[0] == '='
                                    && (s.IndexOf(needleQuoted, StringComparison.Ordinal) >= 0
                                        || s.IndexOf(needleBare, StringComparison.Ordinal) >= 0
                                        || s.IndexOf(needleFull, StringComparison.Ordinal) >= 0
                                        || s.IndexOf(needleQuotedBare, StringComparison.Ordinal) >= 0))
                                {
                                    boardLooksLikeDashboard = true;
                                }
                            }
                            finally { Release((object)cell); }
                        }
                        Release((object)cells);
                    }
                    finally { Release((object)used); }
                }

                // 2026-09-12 KPI 公式迁 B 列后，A1 可能为空——A1 或 B1 有内容即视为看板
                dynamic stgA1 = sheetStaging.Cells.Item[1, 1].Value2;
                dynamic stgB1 = sheetStaging.Cells.Item[1, 2].Value2;
                Release((object)stgA1); Release((object)stgB1);
                bool stagingHasContent = stgA1 != null || stgB1 != null;
                if (!boardLooksLikeDashboard || !stagingHasContent)
                {
                    return Invalid("NOT_A_DASHBOARD",
                        $"「{dashboardSheet}」不具备看板结构特征（看板内无引用 {stagingName} 的公式，或辅助表无内容），已拒绝删除以保护普通工作表。"
                        + "如确需删除，请用 delete_excel_worksheet 并自行确认。");
                }

                // 清理挂接在该看板的切片器资产（native 联动创建的 SlicerCache）
                CleanupExistingLinkage(workbook, dashboardSheet);

                // 删除顺序：先看板后辅助表（DisplayAlerts 已由宿主关，Delete 不弹窗）
                sheetBoard.Delete();
                Release((object)sheetBoard);
                sheetBoard = null;
                sheetStaging.Delete();
                Release((object)sheetStaging);
                sheetStaging = null;

                recorder?.Invoke("delete_dashboard", args, "dashboard-delete",
                    new { sheet = dashboardSheet, staging = stagingName },
                    new { Sheet = dashboardSheet, Deleted = true });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = stagingName,
                        linkage_assets_cleaned = true,
                        deleted = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return Invalid("DASHBOARD_DELETE_FAILED",
                    "删除看板 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message
                    + "。看板可能处于保护态或被其他进程占用。");
            }
            catch (Exception ex)
            {
                return Invalid("DASHBOARD_DELETE_FAILED", "删除看板失败：" + ex.Message);
            }
            finally
            {
                Release((object)sheetBoard);
                Release((object)sheetStaging);
            }
        }

        /// <summary>
        /// FR-5.1：target_track 目标达成看板。
        /// 数据约定：source_address 三列——指标名 | 实际值 | 目标值（首行表头）。
        /// 输出：KPI 卡 3 张（总体达成率/达标项数/最佳指标）+ 逐指标完成率表（数据条进度）。
        /// 源只读；幂等；失败回滚。
        /// </summary>
        public ToolResult RenderTargetTrack(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var sheetName = Optional(args, "sheet");
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "目标达成看板";
            var theme = (Optional(args, "theme") ?? "graygreen").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeCell = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                sheetSource = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int headerRow = Convert.ToInt32(rangeSource.Row);
                int lastRow = headerRow + Convert.ToInt32(rangeSource.Rows.Count) - 1;
                int colA = Convert.ToInt32(rangeSource.Column);
                Release((object)rangeSource);

                // 块读取：指标名/实际/目标（块读防逐格 COM，NFR-2 口径）
                dynamic blkRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRow + 1, colA],
                    sheetSource.Cells.Item[lastRow, colA + 2]];
                object[,] blk = blkRange.Value2;
                Release((object)blkRange);
                var items = new List<Tuple<string, double, double>>();
                for (int r = 1; r <= blk.GetLength(0); r++)
                {
                    object name = blk[r, 1];
                    object act = blk[r, 2];
                    object tgt = blk[r, 3];
                    if (name == null) { continue; }
                    double a = act is double d1 ? d1 : double.NaN;
                    double t = tgt is double d2 ? d2 : double.NaN;
                    if (double.IsNaN(a) || double.IsNaN(t)) { continue; }
                    items.Add(Tuple.Create(Convert.ToString(name), a, t));
                }
                if (items.Count == 0)
                {
                    return Invalid("NO_NUMERIC_DATA",
                        "目标达成看板需三列（指标名/实际值/目标值）且至少一行数值有效。");
                }

                // 幂等删除旧看板/辅助表
                string stagingName = "_" + EscapeSheetInner(dashboardSheet) + "_calc";
                foreach (var nm in new[] { dashboardSheet, stagingName })
                {
                    if (SheetExists(workbook, nm))
                    {
                        dynamic old = workbook.Worksheets.Item[nm];
                        try { old.Delete(); } catch { }
                        Release((object)old);
                    }
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add();
                sheetStaging.Name = stagingName;

                // staging：A 指标/B 实际/C 目标/D 完成率/E 差额 + G1~G3 KPI 值
                rangeCell = sheetStaging.Range["A1"]; rangeCell.Value2 = "指标"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["B1"]; rangeCell.Value2 = "实际"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["C1"]; rangeCell.Value2 = "目标"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["D1"]; rangeCell.Value2 = "完成率"; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["E1"]; rangeCell.Value2 = "差额"; Release((object)rangeCell);
                for (int i = 0; i < items.Count; i++)
                {
                    int rr = i + 2;
                    rangeCell = sheetStaging.Cells.Item[rr, 1]; rangeCell.Value2 = items[i].Item1; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 2]; rangeCell.Value2 = items[i].Item2; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 3]; rangeCell.Value2 = items[i].Item3; Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 4];
                    // 完成率：目标>0 用 实际/目标；目标<0（成本类）用 目标/实际；零目标防除零
                    rangeCell.Formula = items[i].Item3 > 0
                        ? $"=IFERROR(B{rr}/C{rr},0)"
                        : $"=IFERROR(IF(C{rr}=0,0,C{rr}/B{rr}),0)";
                    Release((object)rangeCell);
                    rangeCell = sheetStaging.Cells.Item[rr, 5];
                    rangeCell.Formula = $"=B{rr}-C{rr}";
                    Release((object)rangeCell);
                }

                double sumA = items.Sum(x => x.Item2), sumT = items.Sum(x => x.Item3);
                int metCount = items.Count(x => (x.Item3 > 0 ? x.Item2 >= x.Item3 : x.Item2 <= x.Item3));
                string bestName = items
                    .OrderByDescending(x => x.Item3 > 0 ? (double.IsNaN(x.Item2 / x.Item3) ? 0 : x.Item2 / x.Item3) : (x.Item2 == 0 ? 0 : (double.IsNaN(x.Item3 / x.Item2) ? 0 : x.Item3 / x.Item2)))
                    .First().Item1;
                rangeCell = sheetStaging.Range["G1"]; rangeCell.Value2 = sumT > 0 ? Math.Round(sumA / sumT, 4) : 0; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["G2"]; rangeCell.Value2 = metCount; Release((object)rangeCell);
                rangeCell = sheetStaging.Range["G3"]; rangeCell.Value2 = bestName; Release((object)rangeCell);

                // 看板：标题 + KPI 卡 3 张 + 明细进度表（完成率列数据条）
                rangeCell = sheetBoard.Range["A1"]; rangeCell.Value2 = "目标达成看板"; Release((object)rangeCell);
                string[] kpiTitles = { "总体达成率", "达标项数", "最佳指标" };
                for (int i = 0; i < 3; i++)
                {
                    int r1 = 3 + i * 4;
                    dynamic titleCell = sheetBoard.Range[sheetBoard.Cells.Item[r1, 2], sheetBoard.Cells.Item[r1, 4]];
                    titleCell.Merge();
                    titleCell.Value2 = kpiTitles[i];
                    titleCell.Font.Bold = true;
                    titleCell.Font.Color = themeColors.TitleText;
                    Release((object)titleCell);
                    dynamic valueCell = sheetBoard.Range[sheetBoard.Cells.Item[r1 + 1, 2], sheetBoard.Cells.Item[r1 + 1, 4]];
                    valueCell.Merge();
                    // 引用 staging（禁自引用，审查⑨同款纪律）
                    valueCell.Formula = i == 2
                        ? $"=IFERROR('{EscapeSheetInner(stagingName)}'!$G${i + 1},\"-\")"
                        : $"=IFERROR(TEXT('{EscapeSheetInner(stagingName)}'!$G${i + 1},{(i == 0 ? "\"0.0%\"" : "\"0\"")}),\"-\")";
                    valueCell.Font.Size = 20;
                    valueCell.Font.Bold = true;
                    valueCell.Interior.Color = themeColors.CardBg;
                    Release((object)valueCell);
                }

                int headRow = 16;
                string[] headTexts = { "指标", "实际", "目标", "完成率", "差额" };
                for (int c = 0; c < 5; c++)
                {
                    rangeCell = sheetBoard.Cells.Item[headRow, 2 + c];
                    rangeCell.Value2 = headTexts[c];
                    rangeCell.Font.Bold = true;
                    Release((object)rangeCell);
                }
                for (int i = 0; i < items.Count; i++)
                {
                    int rr = headRow + 1 + i;
                    for (int c = 0; c < 5; c++)
                    {
                        string colL = new[] { "A", "B", "C", "D", "E" }[c];
                        rangeCell = sheetBoard.Cells.Item[rr, 2 + c];
                        rangeCell.Formula = $"=IFERROR('{EscapeSheetInner(stagingName)}'!${colL}${i + 2},\"-\")";
                        Release((object)rangeCell);
                    }
                }
                // 数据条加在完成率列（看板侧）
                dynamic rateColumn = sheetBoard.Range[sheetBoard.Cells.Item[headRow + 1, 5], sheetBoard.Cells.Item[headRow + items.Count, 5]];
                dynamic fc = rateColumn.FormatConditions.AddDatabar();
                try { fc.BarColor.Color = 0x7B636B; } finally { Release((object)fc); }
                Release((object)rateColumn);
                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }

                recorder?.Invoke("render_target_track", args, "dashboard-render",
                    new { template = "target_track", sheet = dashboardSheet, metrics = items.Count, met = metCount },
                    new { Sheet = dashboardSheet, Template = "target_track" });

                var selfcheckInfo = RunPostRenderSelfCheck(dashboardSheet);
                string selfcheckText = (selfcheckInfo.Verdict == "ok" ? "自检通过 ✓ " : selfcheckInfo.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheckInfo.Detail;
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "target_track",
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = stagingName,
                        metrics = items.Count,
                        met_count = metCount,
                        overall_rate = sumT > 0 ? Math.Round(sumA / sumT, 4) : 0,
                        theme = theme,
                        selfcheck = selfcheckText,
                        selfcheck_verdict = selfcheckInfo.Verdict,
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "target_track 渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。已回滚。");
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED", "target_track 渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                Release((object)rangeCell);
                Release((object)sheetBoard);
                Release((object)sheetStaging);
                Release((object)sheetSource);
            }
        }

        /// <summary>
        /// FR-5.1 收口：rank_dynamic 排名动态看板（设计 §5.1 首批六类最后一类）。
        /// 数据约定：source_address 三列——类别 | 时间 | 数值（首行表头；列位可传参重指）。
        /// 输出：控制单元格（时间下拉，「全部」为全量）+ KPI 卡 3 张（榜首类别/Top N 合计/全量合计）
        /// + Top N 排行条形图。staging 公式引用控制单元格，切换自动重算（FR-2 同源）；
        /// 源只读；幂等；失败回滚。排名用 RANK+COUNTIF 去重（2013 兼容，无动态数组）。
        /// </summary>
        public ToolResult RenderRankDynamic(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var sheetName = Optional(args, "sheet");
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "排名动态看板";
            var theme = (Optional(args, "theme") ?? "blue").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            var topN = ExcelToolHostHelpers.GetIntSafe(args, "top_n", 10);
            if (topN < 1 || topN > 100)
            {
                return Invalid("INVALID_TOP_N", $"top_n 需在 1~100，当前为 {topN}。");
            }
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 1);
            var timeColumn = ExcelToolHostHelpers.GetIntSafe(args, "time_column", 2);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 3);
            var colParams = new[] { categoryColumn, timeColumn, valueColumn };
            if (colParams.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }
            if (categoryColumn == timeColumn || categoryColumn == valueColumn || timeColumn == valueColumn)
            {
                return Invalid("INVALID_COLUMN", "类别/时间/数值三列参数需互不相同。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeTmp = null;
            dynamic chartObj = null;
            dynamic chart = null;
            dynamic chartAxis = null;
            object bulkGuard = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                sheetSource = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                // Cells.Item 是绝对坐标——区域相对列号必须加源区域首列偏移（T28a 真机教训：
                // 源锚在 AA 列时 min/max 相对列直接索引会读到 A:C，数值判定全空 → NO_NUMERIC_DATA 误报）
                int srcFirstCol = Convert.ToInt32(rangeSource.Column);
                string addr = Convert.ToString(rangeSource.Address, CultureInfo.InvariantCulture);
                Release((object)rangeSource);
                if (colParams.Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"列参数超界：区域列数 {srcCols}，最大传入 {colParams.Max()}。");
                }

                string catColL = ColumnLetters(addr, categoryColumn, categoryColumn, out int headerRow, out int lastRowAbs)[0];
                string timeColL = ColumnLetters(addr, categoryColumn, timeColumn, out _, out _)[1];
                string valColL = ColumnLetters(addr, categoryColumn, valueColumn, out _, out _)[1];
                string escSheet = EscapeSheet(Convert.ToString(sheetSource.Name));
                string catRef = $"{escSheet}!${catColL}${headerRow + 1}:${catColL}${lastRowAbs}";
                string timeRef = $"{escSheet}!${timeColL}${headerRow + 1}:${timeColL}${lastRowAbs}";
                string valRef = $"{escSheet}!${valColL}${headerRow + 1}:${valColL}${lastRowAbs}";
                string ctrlRef = "'" + EscapeSheetInner(dashboardSheet) + "'!$B$2";

                // 块读取：类别/时间去重 + 数值有效性（禁止逐格 COM 扫描，NFR-2 口径）
                int minCol = Math.Min(Math.Min(categoryColumn, timeColumn), valueColumn);
                int maxCol = Math.Max(Math.Max(categoryColumn, timeColumn), valueColumn);
                dynamic dimRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRow + 1, srcFirstCol + minCol - 1],
                    sheetSource.Cells.Item[lastRowAbs, srcFirstCol + maxCol - 1]];
                object[,] dimBlock = dimRange.Value2;
                Release((object)dimRange);
                var catSet = new SortedSet<string>();
                var timeSet = new SortedSet<string>();
                bool hasNumeric = false;
                for (int r = 1; r <= dimBlock.GetLength(0); r++)
                {
                    object cv = dimBlock[r, categoryColumn - minCol + 1];
                    if (cv != null) { catSet.Add(Convert.ToString(cv)); }
                    object tv = dimBlock[r, timeColumn - minCol + 1];
                    if (tv != null) { timeSet.Add(Convert.ToString(tv)); }
                    if (dimBlock[r, valueColumn - minCol + 1] is double) { hasNumeric = true; }
                }
                int catCount = catSet.Count;
                if (catCount == 0 || !hasNumeric || timeSet.Count == 0)
                {
                    return Invalid("NO_NUMERIC_DATA",
                        "排名动态看板需三列（类别/时间/数值）且类别、时间列有值、数值列至少一行数字。");
                }
                int topNEff = Math.Min(topN, catCount);

                // 幂等删除旧看板/辅助表
                string stagingName = "_" + EscapeSheetInner(dashboardSheet) + "_calc";
                foreach (var nm in new[] { dashboardSheet, stagingName })
                {
                    if (SheetExists(workbook, nm))
                    {
                        dynamic old = workbook.Worksheets.Item[nm];
                        try { old.Delete(); } catch { }
                        Release((object)old);
                    }
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add();
                sheetStaging.Name = stagingName;
                bulkGuard = (lastRowAbs - headerRow) >= 5000 ? BeginBulkFormulaWrite() : null;
                // staging：A 类别 / B 数值（SUMIFS 引用控制单元格）/ C 唯一排名 / J 时间下拉源 / E KPI 值
                rangeTmp = sheetStaging.Range["A1"]; rangeTmp.Value2 = "类别"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["B1"]; rangeTmp.Value2 = "数值"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["C1"]; rangeTmp.Value2 = "排名"; Release((object)rangeTmp);
                int ri = 2;
                foreach (var cat in catSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[ri, 1]; rangeTmp.Value2 = cat; Release((object)rangeTmp);
                    dynamic cVal = sheetStaging.Cells.Item[ri, 2];
                    // 「全部」= 全量求和；选定时间 = 加时间条件（分支写法兼容数值型时间，"*" 通配只匹配文本）
                    cVal.Formula = $"=IF({ctrlRef}=\"全部\",SUMIFS({valRef},{catRef},$A{ri}),SUMIFS({valRef},{catRef},$A{ri},{timeRef},{ctrlRef}))";
                    Release((object)cVal);
                    dynamic cRank = sheetStaging.Cells.Item[ri, 3];
                    // RANK + COUNTIF 同值去重 → 唯一排名（INDEX/MATCH 可精确命中；2013 兼容）
                    cRank.Formula = $"=RANK(B{ri},B$2:B${catCount + 1})+COUNTIF(B$2:B{ri},B{ri})-1";
                    Release((object)cRank);
                    ri++;
                }
                rangeTmp = sheetStaging.Range["J1"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                int ti = 2;
                foreach (var t in timeSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[ti, 10]; rangeTmp.Value2 = t; Release((object)rangeTmp);
                    ti++;
                }
                int timeCount = timeSet.Count;
                // KPI 值（E 列；verify_dashboard 只扫 A 列公式，E 列聚合值不参与对账——与 sales_overview G/H 同理）
                string stgInner = EscapeSheetInner(stagingName);
                dynamic k1 = sheetStaging.Range["E1"];
                k1.Formula = $"=INDEX($A$2:$A${catCount + 1},MATCH(MAX($B$2:$B${catCount + 1}),$B$2:$B${catCount + 1},0))";
                Release((object)k1);
                dynamic k2 = sheetStaging.Range["E2"];
                k2.Formula = $"=SUMIF($C$2:$C${catCount + 1},\"<=\"&{topNEff},$B$2:$B${catCount + 1})";
                Release((object)k2);
                dynamic k3 = sheetStaging.Range["E3"];
                k3.Formula = $"=SUM($B$2:$B${catCount + 1})";
                Release((object)k3);

                // 看板：标题 + 控制单元格 + KPI 卡 3 张 + Top N 明细表（降序）+ 排行条形图
                rangeTmp = sheetBoard.Range["A1"]; rangeTmp.Value2 = "排名动态看板"; Release((object)rangeTmp);
                dynamic titleCell0 = sheetBoard.Range["A1"];
                titleCell0.Font.Bold = true;
                titleCell0.Font.Size = 14;
                titleCell0.Font.Color = themeColors.TitleText;
                Release((object)titleCell0);
                rangeTmp = sheetBoard.Range["A2"]; rangeTmp.Value2 = "时间："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["B2"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                dynamic validation = sheetBoard.Range["B2"].Validation;
                try
                {
                    validation.Delete();
                    validation.Add(Type: 3, Formula1: $"='{stgInner}'!$J$1:$J${timeCount + 1}");
                }
                finally { Release((object)validation); }

                string[] kpiTitles = { "榜首类别", $"Top {topNEff} 合计", "全量合计" };
                for (int i = 0; i < 3; i++)
                {
                    int r1 = 4 + i * 4;
                    dynamic kt = sheetBoard.Range[sheetBoard.Cells.Item[r1, 2], sheetBoard.Cells.Item[r1, 4]];
                    kt.Merge();
                    kt.Value2 = kpiTitles[i];
                    kt.Font.Bold = true;
                    kt.Font.Color = themeColors.TitleText;
                    Release((object)kt);
                    dynamic kv = sheetBoard.Range[sheetBoard.Cells.Item[r1 + 1, 2], sheetBoard.Cells.Item[r1 + 1, 4]];
                    kv.Merge();
                    // 引用 staging E 列 KPI 值（禁自引用，审查⑨同款纪律）
                    kv.Formula = i == 0
                        ? $"=IFERROR('{stgInner}'!$E$1,\"-\")"
                        : $"=IFERROR(TEXT('{stgInner}'!$E${i + 1},\"#,##0\"),\"-\")";
                    kv.Font.Size = 20;
                    kv.Font.Bold = true;
                    kv.Font.Color = themeColors.ValueText;
                    kv.Interior.Color = themeColors.CardBg;
                    Release((object)kv);
                }

                int headRow = 16;
                string[] headTexts = { "排名", "类别", "数值" };
                for (int c = 0; c < 3; c++)
                {
                    rangeTmp = sheetBoard.Cells.Item[headRow, 2 + c];
                    rangeTmp.Value2 = headTexts[c];
                    rangeTmp.Font.Bold = true;
                    Release((object)rangeTmp);
                }
                for (int i = 1; i <= topNEff; i++)
                {
                    int rr = headRow + i;
                    rangeTmp = sheetBoard.Cells.Item[rr, 2]; rangeTmp.Value2 = i; Release((object)rangeTmp);
                    dynamic cc = sheetBoard.Cells.Item[rr, 3];
                    cc.Formula = $"=IFERROR(INDEX('{stgInner}'!$A$2:$A${catCount + 1},MATCH({i},'{stgInner}'!$C$2:$C${catCount + 1},0)),\"-\")";
                    Release((object)cc);
                    dynamic cd = sheetBoard.Cells.Item[rr, 4];
                    cd.Formula = $"=IFERROR(INDEX('{stgInner}'!$B$2:$B${catCount + 1},MATCH({i},'{stgInner}'!$C$2:$C${catCount + 1},0)),\"-\")";
                    cd.NumberFormat = "#,##0";
                    Release((object)cd);
                }
                // 数据条加在数值列（看板侧，target_track 同款）
                dynamic valColumn = sheetBoard.Range[sheetBoard.Cells.Item[headRow + 1, 4], sheetBoard.Cells.Item[headRow + topNEff, 4]];
                dynamic fc = valColumn.FormatConditions.AddDatabar();
                try { fc.BarColor.Color = 0x7B636B; } finally { Release((object)fc); }
                Release((object)valColumn);

                // Top N 排行条形图（锚定 F2 右侧空白区；明细表 B:D 不会被覆盖）
                try
                {
                    dynamic anchor = sheetBoard.Range["F2"];
                    double left = Convert.ToDouble(anchor.Left);
                    double top = Convert.ToDouble(anchor.Top);
                    Release((object)anchor);
                    chartObj = sheetBoard.ChartObjects().Add(left, top, 480, 380);
                    chart = chartObj.Chart;
                    chart.ChartType = 57; // xlBarClustered
                    chart.SetSourceData(sheetBoard.Range[sheetBoard.Cells.Item[headRow, 3], sheetBoard.Cells.Item[headRow + topNEff, 4]]);
                    try
                    {
                        chartAxis = chart.Axes(1); // xlCategory
                        chartAxis.ReversePlotOrder = true; // 第一名显示在最上
                    }
                    catch { }
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = $"Top {topNEff} 排行";
                }
                finally
                {
                    Release((object)chartAxis);
                    Release((object)chart);
                    Release((object)chartObj);
                }

                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }
                sheetStaging.Visible = 0;

                recorder?.Invoke("render_rank_dynamic", args, "dashboard-render",
                    new { template = "rank_dynamic", sheet = dashboardSheet, categories = catCount, top_n = topNEff },
                    new { Sheet = dashboardSheet, Template = "rank_dynamic" });

                // 批量守卫激活（≥5000 行）时手动重算未恢复，读值会误报——交由渲染后 auto-verify 兜底
                var selfcheckInfo = bulkGuard == null
                    ? RunPostRenderSelfCheck(dashboardSheet)
                    : (Verdict: "deferred", Detail: "大批量重算中，自检由 auto-verify 在重算完成后执行");
                string selfcheckText = (selfcheckInfo.Verdict == "ok" ? "自检通过 ✓ " : selfcheckInfo.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheckInfo.Detail;
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "rank_dynamic",
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = stagingName,
                        categories = catCount,
                        times = timeCount,
                        top_n = topNEff,
                        linkage = dashboardSheet + "!B2 下拉（全部/各期），切换自动重算",
                        theme = theme,
                        selfcheck = selfcheckText,
                        selfcheck_verdict = selfcheckInfo.Verdict,
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "rank_dynamic 渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED", "rank_dynamic 渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                EndBulkFormulaWrite(bulkGuard, dashboardSheet, "_" + EscapeSheetInner(dashboardSheet) + "_calc");
                Release((object)rangeTmp);
                Release((object)sheetBoard);
                Release((object)sheetStaging);
                Release((object)sheetSource);
            }
        }

        /// <summary>
        /// M-D3 FR-5.1：dim_cross 多维交叉看板。
        /// 行维度 × 列维度交叉矩阵（SUMIFS），双 cell 下拉联动（行/列各一个控制单元格），
        /// 交叉矩阵条形图。源数据只读；幂等；失败回滚；staging 公式限定实际区间（FR-3.4）。
        /// </summary>
        public ToolResult RenderDimCross(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var rowDimColumn = ExcelToolHostHelpers.GetIntSafe(args, "row_dim_column", 1);
            var colDimColumn = ExcelToolHostHelpers.GetIntSafe(args, "col_dim_column", 2);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "多维交叉看板";
            var sourceSheetName = Optional(args, "sheet");
            var theme = (Optional(args, "theme") ?? "blue").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            if (rowDimColumn == colDimColumn)
            {
                return Invalid("INVALID_DIMENSION", "行维度列与列维度列不能相同。");
            }
            var cols = new[] { rowDimColumn, colDimColumn, valueColumn };
            if (cols.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeTmp = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // FR-1.5 多看板共存：auto_rename=true 时重名自动「名 2/名 3」递增
                var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                // 源数据表：显式指定优先（set_linkage 等前置工具会激活看板表，ActiveSheet 不可靠）
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                dynamic rangeSource = sheetSource.Range[sourceAddress];
                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                int srcFirstCol = Convert.ToInt32(rangeSource.Column);
                int headerRowAbs = Convert.ToInt32(rangeSource.Row);
                int lastRowAbs = headerRowAbs + Convert.ToInt32(rangeSource.Rows.Count) - 1;
                Release((object)rangeSource);
                if (cols.Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"列参数超界：区域列数 {srcCols}，最大传入 {cols.Max()}。");
                }

                // 块读取维度去重（禁止逐格 COM 扫描，NFR-2 口径）
                int minCol = Math.Min(Math.Min(rowDimColumn, colDimColumn), valueColumn);
                int maxCol = Math.Max(Math.Max(rowDimColumn, colDimColumn), valueColumn);
                dynamic dimRange = sheetSource.Range[
                    sheetSource.Cells.Item[headerRowAbs + 1, srcFirstCol + minCol - 1],
                    sheetSource.Cells.Item[lastRowAbs, srcFirstCol + maxCol - 1]];
                object[,] dimBlock = dimRange.Value2;
                Release((object)dimRange);
                var rowSet = new SortedSet<string>();
                var colSet = new SortedSet<string>();
                for (int r = 1; r <= dimBlock.GetLength(0); r++)
                {
                    object rv = dimBlock[r, rowDimColumn - minCol + 1];
                    if (rv != null) { rowSet.Add(Convert.ToString(rv)); }
                    object cv = dimBlock[r, colDimColumn - minCol + 1];
                    if (cv != null) { colSet.Add(Convert.ToString(cv)); }
                }
                if (rowSet.Count == 0 || colSet.Count == 0)
                {
                    return Invalid("NO_DIMENSION_VALUES", "行/列维度列无有效值，无法构建交叉矩阵。");
                }

                // 矩阵布局防呆（JYYJ图表生成测试数据 看板03 真机实证）：真实长表的维度列是文本标签；
                // 列维度列的值大量为数值 = 用户把「第一列=行维度、右侧各列=各维度值的数值」交叉矩阵当长表传——
                // 此时列分类把数值当维度值，产出语义错误的看板且守恒对账在同一错误口径内自洽无法发现，必须拦截。
                if (dimBlock.GetLength(0) >= 3)
                {
                    int numericCount = 0, nonEmpty = 0;
                    for (int r = 1; r <= dimBlock.GetLength(0); r++)
                    {
                        object cv = dimBlock[r, colDimColumn - minCol + 1];
                        if (cv == null) { continue; }
                        nonEmpty++;
                        if (cv is double) { numericCount++; }
                    }
                    if (nonEmpty > 0 && numericCount / (double)nonEmpty > 0.6)
                    {
                        return Invalid("MATRIX_LAYOUT",
                            "检测到交叉矩阵布局：列维度列的值绝大多数是数值，说明数据是「第一列=行维度、右侧各列=各维度值对应的数值」的矩阵，" +
                            "而不是「行维度|列维度|数值」三列长表。请先把矩阵转成长表（每行：行维度值|列维度值|数值），再渲染多维交叉看板。");
                    }
                }
                if (rowSet.Count * colSet.Count > 400)
                {
                    return Invalid("MATRIX_TOO_LARGE",
                        $"交叉矩阵 {rowSet.Count}×{colSet.Count} 超过 400 单元上限；请先聚合维度或分组 Top N。");
                }

                string dataSheetName = Convert.ToString(sheetSource.Name);
                string esc = EscapeSheet(dataSheetName);
                string rowRef = $"{esc}!${ColumnLetter(srcFirstCol + rowDimColumn - 1)}$2:${ColumnLetter(srcFirstCol + rowDimColumn - 1)}${lastRowAbs}";
                string colRef = $"{esc}!${ColumnLetter(srcFirstCol + colDimColumn - 1)}$2:${ColumnLetter(srcFirstCol + colDimColumn - 1)}${lastRowAbs}";
                string valRef = $"{esc}!${ColumnLetter(srcFirstCol + valueColumn - 1)}$2:${ColumnLetter(srcFirstCol + valueColumn - 1)}${lastRowAbs}";

                // 幂等删除旧看板/辅助表
                foreach (var name in new[] { dashboardSheet, "_" + dashboardSheet + "_calc" })
                {
                    if (SheetExists(workbook, name))
                    {
                        dynamic old = workbook.Worksheets.Item[name];
                        try { old.Delete(); } catch { }
                        Release((object)old);
                    }
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add();
                sheetStaging.Name = "_" + dashboardSheet + "_calc";

                // staging 布局：A1=行控制 C1=列控制（公式引看板 B3/D3）；B2 起交叉矩阵（B 列行标签，第 2 行列标签）
                rangeTmp = sheetStaging.Range["A1"]; rangeTmp.Formula = null; rangeTmp.Value2 = "(全部)"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["C1"]; rangeTmp.Formula = null; rangeTmp.Value2 = "(全部)"; Release((object)rangeTmp);
                int ci0 = 0;
                foreach (var cTag in colSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[2, 3 + ci0];
                    rangeTmp.Value2 = cTag;
                    Release((object)rangeTmp);
                    ci0++;
                }
                int ri0 = 3;
                foreach (var rTag in rowSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[ri0, 2];
                    rangeTmp.Value2 = rTag;
                    Release((object)rangeTmp);
                    int cj = 0;
                    foreach (var cTag in colSet)
                    {
                        string colLetter = ColumnLetter(3 + cj);
                        dynamic cell = sheetStaging.Cells.Item[ri0, 3 + cj];
                        cell.Formula =
                            $"=SUMIFS({valRef},{rowRef},$B{ri0},{colRef},{colLetter}$2)" +
                            $"*IF($A$1=\"(全部)\",1,--($B{ri0}=$A$1))*IF($C$1=\"(全部)\",1,--({colLetter}$2=$C$1))";
                        Release((object)cell);
                        cj++;
                    }
                    ri0++;
                }

                // 看板 sheet：标题 + 双下拉控制单元格 + 矩阵引用视图 + 条形图
                rangeTmp = sheetBoard.Range["A1"]; rangeTmp.Value2 = "多维交叉看板"; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["A3"]; rangeTmp.Value2 = "行筛选："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["B3"]; rangeTmp.Value2 = "(全部)"; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["C3"]; rangeTmp.Value2 = "列筛选："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["D3"]; rangeTmp.Value2 = "(全部)"; Release((object)rangeTmp);
                dynamic vRow = sheetBoard.Range["B3"].Validation;
                vRow.Add(Type: 3, Formula1: "\"(全部)," + string.Join(",", rowSet) + "\"");
                Release((object)vRow);
                dynamic vCol = sheetBoard.Range["D3"].Validation;
                vCol.Add(Type: 3, Formula1: "\"(全部)," + string.Join(",", colSet) + "\"");
                Release((object)vCol);
                rangeTmp = sheetStaging.Range["A1"]; rangeTmp.Formula = $"='{EscapeSheetInner(dashboardSheet)}'!$B$3"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["C1"]; rangeTmp.Formula = $"='{EscapeSheetInner(dashboardSheet)}'!$D$3"; Release((object)rangeTmp);
                // 矩阵引用视图（看板 ← staging）
                for (int r = 0; r <= rowSet.Count; r++)
                {
                    for (int c = 0; c <= colSet.Count; c++)
                    {
                        dynamic srcCell = sheetStaging.Cells.Item[2 + r, 2 + c];
                        string addr = Convert.ToString(srcCell.Address(false, false));
                        Release((object)srcCell);
                        dynamic dst = sheetBoard.Cells.Item[6 + r, 2 + c];
                        dst.Formula = $"='{EscapeSheetInner("_" + dashboardSheet + "_calc")}'!{addr}";  // 2026-09-11 修复：此前漏前导 _ → 引用悬空 → 卡片全显 #NAME?
                        Release((object)dst);
                    }
                }
                // 条形图（引用 staging 矩阵，含行列标签）
                dynamic chartObj = null;
                dynamic chart = null;
                try
                {
                    chartObj = sheetBoard.ChartObjects().Add(420, 120, 560, 340);
                    chart = chartObj.Chart;
                    chart.ChartType = 57; // xlBarClustered
                    chart.SetSourceData(sheetStaging.Range[
                        sheetStaging.Cells.Item[2, 2], sheetStaging.Cells.Item[2 + rowSet.Count, 2 + colSet.Count]],
                        1); // PlotBy=xlRows：行=分类
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "维度交叉矩阵";
                }
                finally
                {
                    Release((object)chart);
                    Release((object)chartObj);
                }
                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }

                recorder?.Invoke("render_dim_cross", args, "dashboard-render",
                    new { template = "dim_cross", sheet = dashboardSheet, rows = rowSet.Count, cols = colSet.Count },
                    new { Sheet = dashboardSheet, Template = "dim_cross" });

                var selfcheckInfo = RunPostRenderSelfCheck(dashboardSheet);
                string selfcheckText = (selfcheckInfo.Verdict == "ok" ? "自检通过 ✓ " : selfcheckInfo.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheckInfo.Detail;
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        template = "dim_cross",
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = "_" + dashboardSheet + "_calc",
                        row_dim_values = rowSet.Count,
                        col_dim_values = colSet.Count,
                        linkage = "看板!B3(行)/D3(列) 双下拉，(全部) 为全量视图",
                        theme = theme,
                        selfcheck = selfcheckText,
                        selfcheck_verdict = selfcheckInfo.Verdict,
                        created = true
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                string hint = ex.HResult == unchecked((int)0x800A03EC)
                    ? "工作簿可能处于兼容模式，请另存为 .xlsx 后重试。"
                    : "已回滚，源数据未受影响。";
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "dim_cross 渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。" + hint);
            }
            catch (Exception ex)
            {
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED", "dim_cross 渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                Release((object)rangeTmp);
                Release((object)sheetBoard);
                Release((object)sheetStaging);
                Release((object)sheetSource);
            }
        }

        /// <summary>1 基列号 → 列字母。</summary>
        private static string ColumnLetter(int column)
        {
            string s = string.Empty;
            while (column > 0)
            {
                int m = (column - 1) % 26;
                s = (char)('A' + m) + s;
                column = (column - 1) / 26;
            }
            return s;
        }

        /// <summary>
        /// M-D3 FR-4.3/4.4：图表版本徽章探测（只读）。
        /// 检测当前 Excel 版本是否支持指定版本敏感图表类型；不可用时给替代建议（拒绝而非静默降级原则）。
        /// 图表类型白名单：treemap/sunburst/funnel（2021+）、histogram/pareto/waterfall（2016+）、boxwhisker（2016+）。
        /// </summary>
        public ToolResult CheckChartVersion(Dictionary<string, object> args)
        {
            var chartType = (Required(args, "chart_type") ?? string.Empty).Trim().ToLowerInvariant();
            var required = new Dictionary<string, double>
            {
                ["treemap"] = 21.0, ["sunburst"] = 21.0, ["funnel"] = 21.0,
                ["histogram"] = 16.0, ["pareto"] = 16.0, ["waterfall"] = 16.0, ["boxwhisker"] = 16.0
            };
            var substitutes = new Dictionary<string, string>
            {
                ["treemap"] = "可改用「环形图（≤5 类）」或「条形 Top10 排名」表达层级构成。",
                ["sunburst"] = "可改用「环形图 + 明细分层条形」组合表达层级。",
                ["funnel"] = "可改用「堆积条形（各级转化）」模拟漏斗。",
                ["histogram"] = "可改用「柱形图 + FREQUENCY 辅助列」模拟分布。",
                ["pareto"] = "可改用「排序条形 + 累计占比折线（双轴）」模拟排列图。",
                ["waterfall"] = "可改用「垫底柱模拟瀑布（首段透明堆积柱）」，全版本兼容。",
                ["boxwhisker"] = "可改用「散点 + 误差线」模拟箱形，或改用均值/分位数卡片。"
            };
            var labels = new Dictionary<string, string>
            {
                ["treemap"] = "树状图", ["sunburst"] = "旭日图", ["funnel"] = "漏斗图",
                ["histogram"] = "直方图", ["pareto"] = "排列图", ["waterfall"] = "瀑布图", ["boxwhisker"] = "箱形图"
            };

            if (!required.TryGetValue(chartType, out var minVersion))
            {
                return Invalid("UNKNOWN_CHART_TYPE",
                    "不支持的图表类型查询：" + chartType + "。可选：treemap/sunburst/funnel/histogram/pareto/waterfall/boxwhisker。");
            }

            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                double ver = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
                bool available = ver >= minVersion;
                string verLabel = minVersion >= 21.0 ? "2021/365" : "2016+";
                string badge = $"[{verLabel} 徽章]";

                recorder?.Invoke("check_chart_version", args, "version-probe",
                    new { chart_type = chartType, excel_version = ver, available },
                    new { ChartType = chartType, Available = available });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_type = chartType,
                        chart_label = labels[chartType],
                        excel_version = ver,
                        min_version = minVersion,
                        version_label = verLabel,
                        badge = badge,
                        available = available,
                        suggestion = available
                            ? $"当前 Excel {ver} 支持{labels[chartType]}，可直接生成。"
                            : $"当前 Excel {ver} 不支持{labels[chartType]}（需 {verLabel}）。{substitutes[chartType]}"
                    }
                };
            }
            catch (InvalidOperationException ex)
            {
                return Invalid("VERSION_PROBE_FAILED", ex.Message);
            }
            catch (Exception ex)
            {
                return Invalid("VERSION_PROBE_FAILED", "版本探测失败：" + ex.Message);
            }
            finally
            {
                Release((object)workbook);
            }
        }


        /// <summary>按名解析看板 sheet；不存在返回 null（不抛异常，交由调用方给可读错误）。</summary>
        private dynamic ResolveSheetForLinkage(dynamic workbook, string dashboardSheet)
        {
            return SheetExists(workbook, dashboardSheet) ? workbook.Worksheets.Item[dashboardSheet] : null;
        }

        /// <summary>幂等清理既有联动资产（同名 SlicerCache + staging 透视表），供重建使用。</summary>
        private void CleanupExistingLinkage(dynamic workbook, string dashboardSheet)
        {
            // SlicerCache 名称含维度字段不可预知，按 Caption/关联透视表名清理：遍历 caches 删除挂在 _linkage_pivot_* 的 cache
            dynamic caches = workbook.SlicerCaches;
            for (int i = Convert.ToInt32(caches.Count); i >= 1; i--)
            {
                dynamic cache = null;
                try
                {
                    cache = caches.Item(i);
                    dynamic cachePivots = cache.PivotTables;
                    for (int p = Convert.ToInt32(cachePivots.Count); p >= 1; p--)
                    {
                        string pname = Convert.ToString(cachePivots.Item(p).Name);
                        if (pname.StartsWith("_linkage_pivot_", StringComparison.Ordinal))
                        {
                            try { cache.Delete(); } catch { }
                            break;
                        }
                    }
                    Release((object)cachePivots);
                }
                catch { }
                finally { Release((object)cache); }
            }

            // staging 上的遗留透视表
            try
            {
                dynamic staging = workbook.Worksheets.Item["_" + EscapeSheetInner(dashboardSheet) + "_calc"];
                dynamic pivots = staging.PivotTables();
                for (int p = Convert.ToInt32(pivots.Count); p >= 1; p--)
                {
                    string pname = Convert.ToString(pivots.Item(p).Name);
                    if (pname.StartsWith("_linkage_pivot_", StringComparison.Ordinal))
                    {
                        try { pivots.Item(p).TableRange2.Clear(); } catch { }
                    }
                }
                Release((object)pivots);
                Release((object)staging);
            }
            catch { }
        }

        /// <summary>联动设置失败回滚：清理本次可能新建的切片器/透视表资产（AC-5 语义）。</summary>
        private void TryRollbackLinkage(dynamic workbook, string dashboardSheet)
        {
            try
            {
                if (workbook != null) { CleanupExistingLinkage(workbook, dashboardSheet); }
            }
            catch { }
        }

        /// <summary>
        /// 参数：
        /// - source_address（必填）：明细数据区域（首行表头；须先通过数据门槛探测）；
        /// - value_column：数值列（1 基列号，相对 source_address 首列），默认 4（D 列相对）；
        /// - label_column：聚合维度列（1 基列号），默认 2；用于辅助区的分类小计/最大分类指标；
        /// - kpis：KPI 定义数组（title/aggregation/format），缺省 = 总额/计数/均值/最大单笔 四卡；
        ///   aggregation ∈ sum|count|average|max|min|ratio_max_min；
        /// - dashboard_sheet：看板 sheet 名，默认「KPI看板」；已存在则幂等重建（先删后建）；
        /// - theme：主题色 orange|blue|graygreen，默认 orange。
        /// </summary>
        public ToolResult CreateKpiBrief(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 2);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "KPI看板";
            var theme = (Optional(args, "theme") ?? "orange").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }

            if (labelColumn < 1 || valueColumn < 1)
            {
                return Invalid("INVALID_COLUMN", $"label_column/value_column 需 ≥1（1 基），当前 {labelColumn}/{valueColumn}。");
            }

            if (!args.TryGetValue("kpis", out var kpisRaw) || !(kpisRaw is IEnumerable<object> kpiList) || IsEmpty(kpiList))
            {
                kpiList = DefaultKpis();
            }

            // 解析 KPI 定义（支持 JArray 或 object[]）
            var kpis = new List<KpiDef>();
            foreach (var item in kpiList)
            {
                var def = ParseKpi(item);
                if (def == null)
                {
                    return Invalid("INVALID_KPI", "kpis 数组元素需含 title 与 aggregation 字段（aggregation ∈ sum/count/average/max/min）。");
                }

                if (kpis.Count >= 8)
                {
                    return Invalid("TOO_MANY_KPIS", "KPI 卡最多 8 张（PRD FR-3.1）。");
                }

                kpis.Add(def);
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic rangeSource = null;
            dynamic sheetOld = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeCell = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // FR-1.5 多看板共存：auto_rename=true 时重名自动「名 2/名 3」递增
                var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                // FR-6：支持显式 sheet 参数——渲染会激活看板表，重生成/连续生成时 ActiveSheet 不可靠
                // （真机 T11b 实证：不传 sheet 时把看板表当数据源，months=0/INVALID_FIELD）
                var sourceSheetName = Optional(args, "sheet");
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                rangeSource = sheetSource.Range[sourceAddress];

                // 实际数据边界（FR-3.4：staging 公式限定实际区间，禁整列引用）
                int lastRow = Convert.ToInt32(rangeSource.Rows.Count);
                var sourceCols = Convert.ToInt32(rangeSource.Columns.Count);
                // 2026-09-13 单列明细支持（图表测试集实证）：单列数值源（"单列明细 → 自动探测数值"场景）
                // 强制 label/value 同列——聚合公式只用数值列（BuildKpiFormula 未使用 labelRef），语义为纯数值汇总。
                if (sourceCols == 1) { labelColumn = 1; valueColumn = 1; }
                if (valueColumn > sourceCols || labelColumn > sourceCols)
                {
                    return Invalid("COLUMN_OUT_OF_RANGE",
                        $"label_column={labelColumn}/value_column={valueColumn} 超出数据区域列数 {sourceCols}。");
                }

                // 幂等：删旧看板 sheet（含其隐藏辅助表）
                if (SheetExists(workbook, dashboardSheet))
                {
                    sheetOld = workbook.Worksheets.Item[dashboardSheet];
                    try { sheetOld.Delete(); } catch { /* 删除降级：存在即删，失败不阻断渲染 */ }
                    sheetOld = null;
                }

                var stagingSheet = "_" + dashboardSheet + "_calc";
                if (SheetExists(workbook, stagingSheet))
                {
                    sheetOld = workbook.Worksheets.Item[stagingSheet];
                    try { sheetOld.Delete(); } catch { }
                    sheetOld = null;
                }

                sheetBoard = workbook.Worksheets.Add(); // 无参 Add（避开采集合 Item[Count] 的 0x800A01A8）
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add(After: sheetBoard);
                sheetStaging.Name = stagingSheet;

                // ---- 辅助区（隐藏 sheet）：KPI 数值全部由公式计算 ----
                // 地址化：明细表引用（数据实际区间，禁整列）
                string dataSheetName = Convert.ToString(sheetSource.Name);
                var addr = Convert.ToString(rangeSource.Address, CultureInfo.InvariantCulture);
                // addr 形如 $A$1:$F$21；解析出首列字母与末行号
                var colLetters = ColumnLetters(addr, labelColumn, valueColumn, out int headerRow, out int lastRowAbs);
                string labelCol = colLetters[0];
                string valueCol = colLetters[1];
                string dataRef = $"{EscapeSheet(dataSheetName)}!${valueCol}${headerRow + 1}:${valueCol}${lastRowAbs}";  // EscapeSheet 已含引号
                string labelRef = $"{EscapeSheet(dataSheetName)}!${labelCol}${headerRow + 1}:${labelCol}${lastRowAbs}";

                int writeRow = 1;
                foreach (var kpi in kpis)
                {
                    string formula = BuildKpiFormula(kpi.Aggregation, dataRef, labelRef);
                    rangeCell = sheetStaging.Cells.Item[writeRow, 1];
                    rangeCell.Formula = formula;
                    Release((object)rangeCell);
                    rangeCell = null;
                    writeRow++;
                }

                // ---- 看板 sheet：KPI 卡片（合并单元格 + 大字号 + 主题色 + 三件套）----
                // 布局网格：每卡占 3 列 × 3 行，卡间 1 列留白；每行 3 张卡
                int cardsPerRow = 3;
                int cardCols = 3;
                int cardRows = 3;
                for (int i = 0; i < kpis.Count; i++)
                {
                    int rowBlock = i / cardsPerRow;
                    int colBlock = i % cardsPerRow;
                    int r1 = 2 + (rowBlock * (cardRows + 1));
                    int c1 = 2 + (colBlock * (cardCols + 1));
                    int r2 = r1 + cardRows - 1;
                    int c2 = c1 + cardCols - 1;

                    dynamic titleCell = sheetBoard.Range[sheetBoard.Cells.Item[r1, c1], sheetBoard.Cells.Item[r1, c2]];
                    titleCell.Merge();
                    titleCell.Value2 = kpis[i].Title;
                    titleCell.Font.Size = 11;
                    titleCell.Font.Bold = true;
                    titleCell.Font.Color = themeColors.TitleText;
                    titleCell.HorizontalAlignment = -4108; // xlCenter
                    Release((object)titleCell);

                    dynamic valueCell = sheetBoard.Range[sheetBoard.Cells.Item[r1 + 1, c1], sheetBoard.Cells.Item[r1 + 1, c2]];
                    try
                    {
                        string f = $"=IFERROR({EscapeSheet(stagingSheet)}!$B${i + 1},\"-\")";  // EscapeSheet 已含引号；2026-09-12 KPI 值迁 B 列（A1 坏缓存规避）
                        valueCell.Formula = f;
                    }
                    catch (Exception fEx)
                    {
                        throw new Exception("Formula 写入失败: " + fEx.Message);
                    }
                    try
                    {
                        valueCell.NumberFormat = kpis[i].Format;
                    }
                    catch
                    {
                        valueCell.NumberFormat = "#,##0";
                    }
                    valueCell.Font.Size = 22;
                    valueCell.Font.Bold = true;
                    valueCell.Font.Color = themeColors.ValueText;
                    valueCell.Interior.Color = themeColors.CardBg;
                    valueCell.HorizontalAlignment = -4108;
                    Release((object)valueCell);

                    dynamic unitCell = sheetBoard.Range[sheetBoard.Cells.Item[r2, c1], sheetBoard.Cells.Item[r2, c2]];
                    unitCell.Merge();
                    // 三件套②：环比箭头位（P1 静态版为占位说明，环比需时序聚合，P2 扩展）
                    unitCell.Value2 = kpis[i].Note ?? "—";
                    unitCell.Font.Size = 9;
                    unitCell.Font.Color = themeColors.NoteText;
                    unitCell.HorizontalAlignment = -4108; // xlCenter
                    Release((object)unitCell);
                }

                // 列宽微调 + 关闭网格线（布局层）
                sheetBoard.Columns.AutoFit();
                try
                {
                    workbook.Windows.Item(1).DisplayGridlines = false;
                }
                catch
                {
                    // 无窗口句柄场景（无头）忽略
                }

                sheetStaging.Visible = 0; // xlSheetHidden

                recorder?.Invoke(
                    "create_dashboard_kpi_brief",
                    args,
                    "dashboard-render",
                    new { sheet = dashboardSheet, staging = stagingSheet, kpi_count = kpis.Count, theme },
                    new { Sheet = dashboardSheet, KpiCount = kpis.Count });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        staging_sheet = stagingSheet,
                        kpi_count = kpis.Count,
                        kpis = kpis.ConvertAll(k => new { k.Title, k.Aggregation, k.Format }),
                        theme,
                        data_range = addr,
                        created = true
                    }
                };
            }
            catch (InvalidOperationException ex)
            {
                return Invalid("DASHBOARD_RENDER_FAILED", ex.Message);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // AC-5：渲染失败即清理本次已创建的半成品（新增 sheet），源数据不动
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                var hint = ex.HResult == unchecked((int)0x800A03EC)
                    ? "工作簿可能处于兼容模式（未保存为 .xlsx），原生看板元素被禁用——请先将工作簿另存为 .xlsx 后重试。"
                    : "已回滚，源数据未受影响。";
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "看板渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。" + hint);
            }
            catch (Exception ex)
            {
                // M-D1 二轮审查补漏：非 COM 异常（如公式构造失败）同样必须回滚，
                // 否则半成品残留且裸异常逃逸到宿主层（原实现违反 AC-5）
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("DASHBOARD_RENDER_FAILED",
                    "看板渲染失败：" + ex.Message + "。已回滚（清理本次创建的看板/辅助 sheet），源数据未受影响。");
            }
            finally
            {
                Release((object)rangeCell);
                Release((object)rangeSource);
                // sheetBoard/sheetStaging/sheetOld 不在此释放：幂等重生成跨调用复用同一底层 sheet，
                // ReleaseComObject/FinalReleaseComObject 都会断开 RCW 导致后续调用 0x800A01A8（M-D1 审查沉淀）。
                // 断开时机交给进程退出，无泄漏窗口（out-of-proc 独立实例）。
            }
        }

        /// <summary>
        /// 看板选区（需求①）：识别当前工作簿内全部看板及其明细数据源。
        /// 识别口径与 delete_dashboard 一致——sheet 名以「_」开头且以「_calc」结尾的隐藏辅助表；
        /// 数据源从 staging 首行 KPI 公式中解析明细引用（SUMIFS/SUMPRODUCT 第一参数），
        /// 兼容 Excel 公式规范化形态（引号剥离/首下划线剥离，踩坑手册 §9 公式规范化族）。
        /// 只读：不写任何单元格、不创建对象。
        /// </summary>
        public ToolResult GetDashboardSources(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");

                // 内部名集合（staging + 看板自身）：解析数据源时排除，避免把控制单元格镜像当数据源
                var internalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic wsScan in workbook.Worksheets)
                {
                    try
                    {
                        string n = Convert.ToString(wsScan.Name);
                        if (n.StartsWith("_", StringComparison.Ordinal) && n.EndsWith("_calc", StringComparison.Ordinal))
                        {
                            internalNames.Add(n);
                            internalNames.Add(ExtractDashboardName(n));
                        }
                    }
                    finally { Release((object)wsScan); }
                }

                var dashboards = new List<object>();
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    string sheetName = null;
                    dynamic used = null;
                    dynamic firstCell = null;
                    try
                    {
                        sheetName = Convert.ToString(worksheet.Name);
                        // 看板辅助表命名约定：_X_calc（与渲染/删除工具同源口径）
                        if (!sheetName.StartsWith("_", StringComparison.Ordinal) ||
                            !sheetName.EndsWith("_calc", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        used = worksheet.UsedRange;
                        firstCell = used.Cells.Item[1, 1];
                        string f = null;
                        try { f = Convert.ToString(firstCell.Formula); } catch { }

                        // A1 无 KPI 聚合公式 → 扫描前 16 行找首个含明细引用的公式。
                        // 各模板 staging 布局差异（真机核实）：sales_overview/kpi_brief KPI 公式在 A 列；
                        // dim_cross A1 是控制单元格引用（无"("，需跳过）；pnl/target 的 D 列完成率公式
                        // 只引用自身 B/C 列（KPI 总额是 C# 预聚合值）——这类模板从 staging 解析不出明细源，
                        // 走看板 sheet 公式反查（看板明细行引用 staging，其中带源数据 sheet 的公式可兜底）。
                        if (string.IsNullOrEmpty(f) || !f.StartsWith("=", StringComparison.Ordinal) || !f.Contains("("))
                        {
                            f = FindFirstKpiFormula(worksheet, internalNames);
                        }

                        string sourceRef = ParseSourceFromFormula(f, internalNames);
                        if (sourceRef == null)
                        {
                            // 兜底：从可见看板表（去 _calc 前后缀）的公式反查——看板明细行会引用
                            // 「源数据表」（值直拷的模板如 pnl/target 检索不到，如实跳过不误报）
                            sourceRef = FindSourceFromBoardSheet(workbook, ExtractDashboardName(sheetName), internalNames);
                        }
                        if (sourceRef == null)
                        {
                            continue; // 两个途径都解析不出 → 不当作可定位数据源的看板（宁漏报不误报）
                        }

                        dashboards.Add(new
                        {
                            dashboard_sheet = ExtractDashboardName(sheetName),
                            staging_sheet = sheetName,
                            source = sourceRef
                        });
                    }
                    finally
                    {
                        Release((object)firstCell);
                        Release((object)used);
                        Release((object)worksheet);
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { dashboard_count = dashboards.Count, dashboards }
                };
            }
            finally
            {
                Release((object)workbook);
            }
        }

        /// <summary>
        /// 在 staging 前 16 行 × 6 列内找首个「含明细数据源引用」的公式。
        /// 各模板布局差异（真机核实）：sales_overview/kpi_brief KPI 公式在 A 列；
        /// dim_cross 的 SUMIFS 矩阵公式在 B2 起的 B 列——只扫 A 列会漏（真机 T27b 教训）。
        /// 判据委托 ParseSourceFromFormula：能解析出非 _calc 族的明细引用才算命中。
        /// </summary>
        private string FindFirstKpiFormula(dynamic staging, HashSet<string> excludeSheets)
        {
            for (int row = 1; row <= 16; row++)
            {
                for (int col = 1; col <= 6; col++)
                {
                    dynamic cell = null;
                    try
                    {
                        cell = staging.Cells.Item[row, col];
                        string f = null;
                        try { f = Convert.ToString(cell.Formula); } catch { }
                        if (!string.IsNullOrEmpty(f) && f.StartsWith("=", StringComparison.Ordinal) &&
                            ParseSourceFromFormula(f, excludeSheets) != null)
                        {
                            return f;
                        }
                    }
                    finally
                    {
                        Release((object)cell);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// 从可见看板表（staging 名去 _calc 前后缀）的公式反查明细数据源。
        /// 看板表上的 KPI 卡/明细行公式引用 staging；若看板表存在直接引用源数据表的公式
        /// （如 kpi_brief 卡片引源表），从中提取。找不到返回 null（调用方如实跳过）。
        /// </summary>
        private string FindSourceFromBoardSheet(dynamic workbook, string boardName, HashSet<string> excludeSheets)
        {
            dynamic board = null;
            dynamic used = null;
            try
            {
                bool exists = false;
                try { exists = workbook.Worksheets.Item[boardName] != null; } catch { }
                if (!exists) { return null; }
                board = workbook.Worksheets.Item[boardName];
                used = board.UsedRange;
                int rows = Convert.ToInt32(used.Rows.Count);
                int cols = Convert.ToInt32(used.Columns.Count);
                if (rows < 1 || cols < 1) { return null; }
                if (rows > 40) { rows = 40; }   // KPI 卡/标题集中在顶部，防大表全扫
                if (cols > 12) { cols = 12; }

                // 块读公式矩阵（Formula 二维属性；单格退化由防御分支兜底）
                object matrix = null;
                try { matrix = used.Formula; } catch { }
                if (matrix is object[,] arr)
                {
                    for (int r = arr.GetLowerBound(0); r <= Math.Min(arr.GetUpperBound(0), rows); r++)
                    {
                        for (int c = arr.GetLowerBound(1); c <= Math.Min(arr.GetUpperBound(1), cols); c++)
                        {
                            if (arr.GetValue(r, c) is string s && s.StartsWith("=", StringComparison.Ordinal))
                            {
                                var parsed = ParseSourceFromFormula(s, excludeSheets);
                                if (parsed != null) { return parsed; }
                            }
                        }
                    }
                }
                else if (matrix is string single && single.StartsWith("=", StringComparison.Ordinal))
                {
                    return ParseSourceFromFormula(single);
                }
                return null;
            }
            finally
            {
                Release((object)used);
                Release((object)board);
            }
        }

        /// <summary>
        /// 从 staging KPI 公式解析明细数据源地址。
        /// SUMIFS 取「汇总范围」（第一参数），SUMPRODUCT/COUNTIFS 等取第一个 Range 引用；
        /// 兼容公式规范化四种形态：'源'!A1:B9（带引号原名）/ 源!A1:B9（裸名）/
        /// '_xx_calc'! 与 xx_calc!（下划线开头 sheet 引号与首下划线被剥离）。
        /// 返回「表名!地址」完全限定引用；解析不出返回 null。
        /// </summary>
        internal static string ParseSourceFromFormula(string formula)
        {
            return ParseSourceFromFormula(formula, null);
        }

        /// <param name="excludeSheets">需排除的 sheet 名集合（staging/看板自身等内部名；null 时仅按 _calc 族判定）。</param>
        internal static string ParseSourceFromFormula(string formula, HashSet<string> excludeSheets)
        {
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=", StringComparison.Ordinal))
            {
                return null;
            }

            // 排除自身族引用：公式只引用 staging 自身（_X_calc）说明明细引用在别处，按未解析处理
            var matches = Regex.Matches(formula, @"(?:'([^']+)'|([A-Za-z0-9_\u4e00-\u9fa5]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)");
            foreach (Match m in matches)
            {
                string sheet = m.Groups[1].Success ? m.Groups[1].Value.Replace("''", "'") : m.Groups[2].Value;
                string addr = m.Groups[3].Value;
                // 引用指向看板辅助表族 → 跳过，继续找真正的明细引用。
                // 族形态四变体（踩坑手册公式规范化族）：'_X_calc'（引号原名）/ X_calc（裸名规范化剥引号剥下划线）/
                // 'X_calc'（带空格保留引号）/ _X_calc（裸名带下划线）。统一口径：引用名（去引号）以 _calc 结尾即辅助表族。
                if (sheet.Replace("'", "").EndsWith("_calc", StringComparison.Ordinal))
                {
                    continue;
                }
                if (excludeSheets != null && excludeSheets.Contains(sheet))
                {
                    // 看板自身/其他内部 sheet 引用（如 dim_cross staging A1 镜像看板 B3 控制单元格）：跳过
                    continue;
                }

                return sheet + "!" + addr;
            }
            return null;
        }

        /// <summary>从辅助表名 _X_calc 还原看板名 X（auto_rename 带空格名同样适配：'_名 2_calc'→「名 2」）。</summary>
        private static string ExtractDashboardName(string stagingName)
        {
            string core = stagingName.Substring(1, stagingName.Length - 1 - 5); // 去 1 前导下划线与 5 尾 _calc
            return core;
        }

        /// <summary>verify_dashboard 最小断言（P1）：看板 sheet 存在、辅助区 KPI 值与明细重算一致（守恒对账）。</summary>
        public ToolResult VerifyDashboard(Dictionary<string, object> args)
        {
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "KPI看板";
            var stagingSheet = Optional(args, "staging_sheet");
            var expectedMin = ExcelToolHostHelpers.GetIntSafe(args, "expected_kpi_count", 0);

            dynamic workbook = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");

                bool boardExists = SheetExists(workbook, dashboardSheet);
                var issues = new List<string>();
                if (!boardExists)
                {
                    issues.Add($"看板工作表「{dashboardSheet}」不存在。");
                }
                else
                {
                    sheetBoard = workbook.Worksheets.Item[dashboardSheet];
                }

                string resolvedStaging = stagingSheet ?? "_" + dashboardSheet + "_calc";
                if (!SheetExists(workbook, resolvedStaging))
                {
                    issues.Add($"辅助汇总表「{resolvedStaging}」不存在（KPI 数值无来源）。");
                }
                else
                {
                    sheetStaging = workbook.Worksheets.Item[resolvedStaging];
                }

                if (sheetStaging != null)
                {
                    // 守恒对账：辅助区每个 KPI 公式重算值必须为有限数值（SUMIFS 语义成立）。
                    // M-D3 审查修正①：按「公式存在性」判定 KPI 单元而非盲扫 A 列——
                    // dim_cross 布局 A1 是控制单元格文本（无公式），不应误报 mismatch。
                    int row = 1;
                    int checkedCount = 0;
                    while (true)
                    {
                        // KPI 公式列已迁 B（A1 坏缓存规避）；dim_cross 等无 B 列 KPI 的模板回退 A 列
                        dynamic cell = sheetStaging.Cells.Item[row, 2];
                        if (Convert.ToString(cell.Formula) == string.Empty)
                        {
                            Release((object)cell);
                            cell = sheetStaging.Cells.Item[row, 1];
                        }
                        object v = cell.Value2;
                        string f = null;
                        try { f = Convert.ToString(cell.Formula); } catch { }
                        Release((object)cell);
                        bool hasFormula = !string.IsNullOrEmpty(f) && f.StartsWith("=", StringComparison.Ordinal);
                        // M-D3 审查修正①（二轮收敛）：dim_cross 的 A1 是「纯引用公式」（=看板!$B$3，
                        // 镜像控制单元格），其值恒为文本「(全部)」等——纯引用（无函数括号）不是 KPI，跳过；
                        // KPI 聚合公式必含函数调用 "("，非数值重算值才报 mismatch。
                        bool isPureReference = hasFormula && !f.Contains("(");
                        if (v == null && !hasFormula)
                        {
                            break;
                        }
                        if (v == null && hasFormula)
                        {
                            // 公式存在但结果为空（如空矩阵）：不计入 KPI，继续扫描防御上限
                            row++;
                            if (row > 16) { break; }
                            continue;
                        }
                        if (hasFormula && !isPureReference && !(v is double))
                        {
                            issues.Add($"辅助区 A{row} 重算值非数值（{v}），SUMIFS/聚合公式可能失效。");
                        }
                        if (!isPureReference)
                        {
                            checkedCount++;
                        }
                        row++;
                        if (row > 16)
                        {
                            break; // KPI 上限 8，防御性上限 16
                        }
                    }

                    if (expectedMin > 0 && checkedCount < expectedMin)
                    {
                        issues.Add($"辅助区 KPI 数（{checkedCount}）少于预期（{expectedMin}）。");
                    }
                }

                bool ok = issues.Count == 0;
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        dashboard_sheet = dashboardSheet,
                        board_exists = boardExists,
                        verdict = ok ? "ok" : "mismatch",
                        issues,
                        selfcheck = ok ? "看板结构存在、辅助区 KPI 重算全部为数值（对账通过）" : string.Join("；", issues)
                    }
                };
            }
            finally
            {
                // sheetStaging/sheetBoard 跨调用复用，不释放（M-D1 RCW 沉淀：FinalRelease/ReleaseComObject
                // 都会断开同对象所有 RCW，幂等重生成与后续 verify 会 0x800A01A8）
            }
        }

        // ---- 内部类型与辅助 ----

        private sealed class KpiDef
        {
            public string Title;
            public string Aggregation;
            public string Format;
            public string Note;
        }

        private sealed class ThemePalette
        {
            public int CardBg;
            public int TitleText;
            public int ValueText;
            public int NoteText;
        }

        private static readonly Dictionary<string, ThemePalette> ThemeColors = new Dictionary<string, ThemePalette>
        {
            // COM Color = 0x00BBGGRR
            ["orange"] = new ThemePalette { CardBg = 0x00ADD7ED, TitleText = 0x2F1E0D, ValueText = 0x0F59A0, NoteText = 0x757575 },
            ["blue"] = new ThemePalette { CardBg = 0xF5DCD1, TitleText = 0x302010, ValueText = 0x8B572A, NoteText = 0x757575 },
            ["graygreen"] = new ThemePalette { CardBg = 0xE3EFE5, TitleText = 0x203020, ValueText = 0x58804E, NoteText = 0x757575 }
        };

        private static List<KpiDef> DefaultKpis()
        {
            return new List<KpiDef>
            {
                new KpiDef { Title = "销售总额", Aggregation = "sum", Format = "¥#,##0" },
                new KpiDef { Title = "成交笔数", Aggregation = "count", Format = "#,##0" },
                new KpiDef { Title = "平均客单价", Aggregation = "average", Format = "¥#,##0" },
                new KpiDef { Title = "最大单笔", Aggregation = "max", Format = "¥#,##0" }
            };
        }

        private static KpiDef ParseKpi(object item)
        {
            string title = null;
            string agg = null;
            string format = null;
            string note = null;

            if (item is Newtonsoft.Json.Linq.JObject jo)
            {
                title = jo.Value<string>("title");
                agg = jo.Value<string>("aggregation");
                format = jo.Value<string>("format");
                note = jo.Value<string>("note");
            }
            else if (item is Dictionary<string, object> dict)
            {
                title = dict.TryGetValue("title", out var t1) ? Convert.ToString(t1) : null;
                agg = dict.TryGetValue("aggregation", out var a1) ? Convert.ToString(a1) : null;
                format = dict.TryGetValue("format", out var f1) ? Convert.ToString(f1) : null;
                note = dict.TryGetValue("note", out var n1) ? Convert.ToString(n1) : null;
            }
            else if (item is KpiDef predefined)
            {
                // DefaultKpis() 的强类型直通（M-D1 真机审查缺陷修复：原 ParseKpi 不识别 KpiDef，
                // 导致缺省 KPI 路径反而报 INVALID_KPI）
                title = predefined.Title;
                agg = predefined.Aggregation;
                format = predefined.Format;
                note = predefined.Note;
            }

            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(agg))
            {
                return null;
            }

            agg = agg.Trim().ToLowerInvariant();
            switch (agg)
            {
                case "sum":
                case "count":
                case "average":
                case "max":
                case "min":
                    break;
                default:
                    return null;
            }

            return new KpiDef
            {
                Title = title,
                Aggregation = agg,
                Format = string.IsNullOrWhiteSpace(format) ? "#,##0" : format,
                Note = note
            };
        }

        /// <summary>按聚合类型生成辅助区公式（数值/维度引用均为实际区间，禁整列）。</summary>
        private static string BuildKpiFormula(string aggregation, string dataRef, string labelRef)
        {
            switch (aggregation)
            {
                case "sum": return $"=SUM({dataRef})";
                case "count": return $"=COUNT({dataRef})";
                case "average": return $"=IFERROR(AVERAGE({dataRef}),0)";   // 三件套：比率类强制 IFERROR
                case "max": return $"=MAX({dataRef})";
                case "min": return $"=MIN({dataRef})";
                default: return $"=SUM({dataRef})";
            }
        }

        /// <summary>从源区域地址解析出 label/value 列字母与首末行号。</summary>
        private static string[] ColumnLetters(string absAddress, int labelColumn, int valueColumn, out int headerRow, out int lastRow)
        {
            // absAddress 形如 $A$1:$F$21
            var parts = absAddress.Split(':');
            string first = parts[0].Trim('$');
            string last = parts.Length > 1 ? parts[1].Trim('$') : parts[0].Trim('$');
            headerRow = int.Parse(RegexDigits(first), CultureInfo.InvariantCulture);
            lastRow = int.Parse(RegexDigits(last), CultureInfo.InvariantCulture);
            string firstColLetters = RegexLetters(first);
            return new[]
            {
                ShiftColumn(firstColLetters, labelColumn - 1),
                ShiftColumn(firstColLetters, valueColumn - 1)
            };
        }

        private static string RegexDigits(string cell)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var c in cell)
            {
                if (char.IsDigit(c))
                {
                    sb.Append(c);
                }
            }

            return sb.Length > 0 ? sb.ToString() : "1";
        }

        private static string RegexLetters(string cell)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var c in cell)
            {
                if (char.IsLetter(c))
                {
                    sb.Append(c);
                }
            }

            return sb.Length > 0 ? sb.ToString() : "A";
        }

        /// <summary>列字母右移 n 列（支持多字母）。</summary>
        private static string ShiftColumn(string letters, int shift)
        {
            int value = 0;
            foreach (var c in letters.ToUpperInvariant())
            {
                value = (value * 26) + (c - 'A' + 1);
            }

            value += shift;
            var sb = new System.Text.StringBuilder();
            while (value > 0)
            {
                int rem = (value - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                value = (value - 1) / 26;
            }

            return sb.ToString();
        }

        private static string EscapeSheet(string sheetName)
        {
            return "'" + sheetName.Replace("'", "''") + "'";
        }

        /// <summary>sheet 名规范化（不加引号包裹，供拼接 '_名_calc' 等复合引用）。</summary>
        private static string EscapeSheetInner(string sheetName)
        {
            return sheetName.Replace("'", "''");
        }

        /// <summary>
        /// FR-1.5 多看板共存（AC-11）：为看板 sheet 解析可用名称。
        /// auto_rename=false（默认）：沿用请求名（同名的旧看板走幂等重建，语义=重生成覆盖）；
        /// auto_rename=true：名称被占用时按「名 2」「名 3」…递增序号，实现多看板共存互不影响。
        /// </summary>
        private static string ResolveDashboardName(dynamic workbook, string requested, bool autoRename)
        {
            // 工具分类审查总复盘修复（命名预校验家族单点收口）：render 系全部看板工具与
            // CreateSuperTable 共用本方法——用户可控 dashboard_sheet/super_sheet 此前直达
            // sheetBoard.Name 无合法性校验（带斜杠/超 31 字符抛裸 COM 异常）。在此单点接入
            // ValidateSheetName（人话 ArgumentException，宿主统一归类 INVALID_ARGS），
            // 全部调用方一次性收口，不再逐入口补丁。
            var nameError = ExcelStructureService.ValidateSheetName(requested);
            if (nameError != null)
            {
                throw new ArgumentException("dashboard_sheet/super_sheet " + nameError);
            }

            if (!autoRename || !SheetExists(workbook, requested))
            {
                return requested;
            }
            for (int seq = 2; seq < 1000; seq++)
            {
                string candidate = requested + " " + seq;
                if (!SheetExists(workbook, candidate))
                {
                    return candidate;
                }
            }
            return requested; // 防御上限内未找到（几乎不可能），回退幂等语义
        }

        private static bool SheetExists(dynamic workbook, string name)
        {
            try
            {
                dynamic s = workbook.Worksheets.Item[name];
                bool exists = s != null;
                Release((object)s);
                return exists;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsEmpty(IEnumerable<object> source)
        {
            foreach (var _ in source)
            {
                return false;
            }

            return true;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        /// <summary>
        /// 数据门槛探测（PRD §3.1，D-① 决策：拒绝生成 + 一键转清洗）。
        /// 硬拦截项：疑似合计行/小计行（文本标签 + 数值为同维度部分和）、合并单元格/多级表头、
        /// 文本型数字、空列。任一硬拦截 → ready=false，给出可读理由与清洗建议。
        /// ReadOnly 工具，不改任何数据。
        /// </summary>
        public ToolResult DetectDataReadiness(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var sheetName = Optional(args, "sheet");
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var sampleRows = ExcelToolHostHelpers.GetIntSafe(args, "sample_rows", 500);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic rangeSource = null;
            dynamic cellProbe = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                // FR-6：支持显式 sheet——其他工作表激活后 ActiveSheet 不可靠（测试2026 实证）
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                rangeSource = sheet.Range[sourceAddress];

                var issues = new List<string>();
                var blockers = new List<string>();

                int rowCount = Convert.ToInt32(rangeSource.Rows.Count);
                int colCount = Convert.ToInt32(rangeSource.Columns.Count);
                if (rowCount > sampleRows)
                {
                    rowCount = sampleRows; // 大表抽样探测前 N 行
                }

                if (valueColumn > colCount)
                {
                    // 二轮审查修复：原实现对超界列静默钳制到最后一列，会在错误列上做数值检查
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            source_address = sourceAddress,
                            ready = false,
                            verdict = "blocked",
                            blockers = new List<string> { $"value_column={valueColumn} 超出数据区域列数 {colCount}。" },
                            suggestion = "请修正 value_column 后重测。"
                        }
                    };
                }

                int absRowStart = Convert.ToInt32(rangeSource.Row); // 绝对行号基准（用户按 Excel 行号定位）
                int firstColAbs = Convert.ToInt32(rangeSource.Column); // 绝对列号基准（FR-6 布局提示用）

                // ① 合并单元格（探测区域内任一合并 → 拦截）
                try
                {
                    if (Convert.ToInt32(rangeSource.MergeCells) != 0)
                    {
                        blockers.Add("区域含合并单元格（多级表头/合并区会破坏 SUMIFS 聚合）。");
                    }
                }
                catch
                {
                    // MergeCells 返回 null 表示混合状态，同样拦截
                    blockers.Add("区域合并状态混合（部分单元格被合并），需先取消合并。");
                }

                // ② 表头行有效性：首行应有 ≥2 个非空文本
                int nonEmptyHeader = 0;
                for (int c = 1; c <= colCount; c++)
                {
                    cellProbe = rangeSource.Cells.Item[1, c];
                    object hv = cellProbe.Value2;
                    Release((object)cellProbe);
                    cellProbe = null;
                    if (hv != null && !string.IsNullOrWhiteSpace(Convert.ToString(hv)))
                    {
                        nonEmptyHeader++;
                    }
                }

                if (nonEmptyHeader < 2)
                {
                    blockers.Add($"首行仅 {nonEmptyHeader} 个非空表头（需 ≥2），疑似缺表头或多级表头。");
                }

                // ③ 明细行扫描：合计行 / 文本型数字 / 空列
                var emptyColCount = new int[colCount + 1];
                var suspectTotalRows = new List<int>();
                var textNumRows = new List<int>();
                var unparseableRows = new List<int>();
                int dataRows = 0;
                double colSum = 0;
                int colNumericCount = 0;
                for (int r = 2; r <= rowCount; r++)
                {
                    dynamic labelCell = rangeSource.Cells.Item[r, 1];
                    object label = labelCell.Value2;
                    Release((object)labelCell);
                    labelCell = null;

                    dynamic valueCell = rangeSource.Cells.Item[r, Math.Min(valueColumn, colCount)];
                    object v = valueCell.Value2;
                    Release((object)valueCell);
                    valueCell = null;

                    bool rowEmpty = label == null && v == null;
                    if (rowEmpty)
                    {
                        continue;
                    }

                    dataRows++;
                    string labelText = label == null ? string.Empty : Convert.ToString(label).Trim();

                    // 空列统计
                    for (int c = 1; c <= colCount; c++)
                    {
                        dynamic cc = rangeSource.Cells.Item[r, c];
                        object cv = cc.Value2;
                        Release((object)cc);
                        if (cv != null && !string.IsNullOrWhiteSpace(Convert.ToString(cv)))
                        {
                            emptyColCount[c]++;
                        }
                    }

                    // 数值列：数值性 & 文本型数字判别
                    if (v != null)
                    {
                        if (v is double dv)
                        {
                            colSum += dv;
                            colNumericCount++;
                        }
                        else if (double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                        {
                            // 文本型数字：Value2 读出字符串但可解析为数字（二轮审查修复：不再 break，
                            // 否则后续行的合计行/空列检测被跳过，错误报告不完整）
                            colSum += parsed;
                            colNumericCount++;
                            textNumRows.Add(absRowStart + r - 1);
                        }
                        else
                        {
                            unparseableRows.Add(absRowStart + r - 1);
                        }
                    }

                    // 疑似合计/小计行：标签含关键字
                    var lower = labelText.ToLowerInvariant();
                    if (lower.Contains("合计") || lower.Contains("总计") || lower.Contains("小计")
                        || lower.Contains("total") || lower.Contains("subtotal"))
                    {
                        suspectTotalRows.Add(r);
                    }
                }

                if (suspectTotalRows.Count > 0)
                {
                    blockers.Add(
                        "明细混入合计/小计行（行 " + string.Join("、", suspectTotalRows) +
                        "）——SUMIFS 会把汇总行重复计入，导致看板数值虚高（行业实证可放大 +161%）。请先删除汇总行或用数据清洗处理。");
                }

                if (textNumRows.Count > 0)
                {
                    blockers.Add(
                        "明细存在文本型数字（行 " + string.Join("、", textNumRows) +
                        "）——数值被存成文本会让聚合结果失真。请用数据清洗修正为数值。");
                }

                if (unparseableRows.Count > 0)
                {
                    blockers.Add(
                        "数值列存在不可解析值（行 " + string.Join("、", unparseableRows) + "）。请清洗或修正后重测。");
                }

                if (dataRows == 0)
                {
                    blockers.Add("区域内没有数据行（仅表头或全空）。");
                }

                // 空列
                for (int c = 1; c <= colCount; c++)
                {
                    if (emptyColCount[c] == 0)
                    {
                        blockers.Add($"第 {c} 列全空，请确认数据区域范围。");
                        break;
                    }
                }

                bool ready = blockers.Count == 0;
                var all = new List<string>();
                all.AddRange(blockers);
                all.AddRange(issues);

                // FR-6.2 扩展：映射建议并入门槛结果——AI 跑门槛必得建议（修「没有弹出选型与映射建议」）
                var mappingSuggestions = new List<object>();
                var columnProfiles = new List<object>();
                string layoutHint = null;
                try
                {
                    if (absRowStart != 1 || firstColAbs != 1)
                    {
                        layoutHint = "数据区未从 A1 开始（实际起点第 " + absRowStart + " 行第 " + firstColAbs +
                            " 列）——按 Excel 实际行列核对选区，避免把空行/空列计为数据。";
                    }
                    // 块读采样区（含表头行）复用选型启发式；失败不阻断门槛结论
                    dynamic mapRange = rangeSource.Range[
                        rangeSource.Cells.Item[1, 1],
                        rangeSource.Cells.Item[Math.Min(rowCount, sampleRows), colCount]];
                    object rawMap = mapRange.Value2;
                    Release((object)mapRange);
                    object[,] mapBlk = rawMap is object[,] mArr ? mArr : new object[1, 1] { { rawMap } };
                    var sugg = AnalyzeChartMapping(mapBlk, out var colProfiles);
                    foreach (var pr in colProfiles)
                    {
                        columnProfiles.Add(new
                        {
                            name = pr.Name,
                            distinct_text = pr.DistinctText,
                            numeric_ratio = Math.Round(pr.NumericRatio, 2),
                            date_like = pr.DateLike,
                            has_negative = pr.HasNegative
                        });
                    }
                    foreach (var s in sugg.Take(5))
                    {
                        mappingSuggestions.Add(new { tool = s.Tool, chart_type = s.ChartType, reason = s.Reason, requires = s.Requires });
                    }
                }
                catch { /* 建议失败不影响门槛判定 */ }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        source_address = Convert.ToString(rangeSource.Address),
                        data_rows = dataRows,
                        value_col_numeric = colNumericCount,
                        value_col_sum = colSum,
                        ready,
                        verdict = ready ? "ok" : "blocked",
                        blockers,
                        layout_hint = layoutHint,
                        mapping_suggestions = mappingSuggestions,
                        column_profiles = columnProfiles,
                        suggestion = ready
                            ? "数据门槛通过，可调用 create_dashboard_kpi_brief 生成看板。"
                            : "请先用 clean_excel_data / 数据清洗流程处理后重测；处理完成后再生成看板。"
                    }
                };
            }
            finally
            {
                Release((object)cellProbe);
                Release((object)rangeSource);
                // sheet 不释放（跨调用复用，同上）
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        /// <summary>渲染失败回滚：删除本次已创建的看板/辅助 sheet（AC-5：失败无半成品残留）。</summary>
        private static void TryDeleteSheet(dynamic sheet)
        {
            if (sheet == null)
            {
                return;
            }

            try { sheet.Delete(); } catch { }
        }

        private static void Release(object value)
        {
            // M-D1 审查修正：DashboardService 的对象可能被宿主/调用方继续引用（幂等重生成跨调用），
            // FinalReleaseComObject 会断开同对象全部 RCW（out-of-proc RCW 按 IUnknown 缓存），
            // 改用减计数 ReleaseComObject，且容错（计数归零后再次调用会抛，吞掉即可）。
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
            }
        }
    }
}
