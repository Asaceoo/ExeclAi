using Newtonsoft.Json.Linq;
﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    public sealed class ExcelPivotChartService
    {
        private readonly dynamic application;
        private readonly OperationRecorder recorder;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;

        public ExcelPivotChartService(object application)
            : this(application, null, null)
        {
        }

        public ExcelPivotChartService(object application, OperationRecorder recorder)
            : this(application, recorder, null)
        {
        }

        public ExcelPivotChartService(
            object application,
            OperationRecorder recorder,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.recorder = recorder;
            this.writer = writer;
        }

        public ToolResult CreatePivot(Dictionary<string, object> args)
        {
            var sourceAddress = Optional(args, "source_address");
            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                // 未显式指定源区域时（如内置工作流步骤），回退到当前多格选区，避免
                // run_workflow 缺参直接报「缺少必填参数」。
                sourceAddress = CurrentSelectionAddress();
            }

            if (string.IsNullOrWhiteSpace(sourceAddress))
            {
                throw new ArgumentException("缺少必填参数：source_address（且当前无多格选区可作为数据源）");
            }

            var targetAddress = Optional(args, "target_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 0);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 0);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];
                var rowsCount = Convert.ToInt32(source.Rows.Count);
                var columnsCount = Convert.ToInt32(source.Columns.Count);
                var values = ExcelToolHostHelpers.FlattenValues(
                    source.Value2,
                    rowsCount,
                    columnsCount);

                // 默认列推断：标签列=第 1 列；数值列=数据区最后一个含数值的列（兜底最后一列）。
                if (labelColumn < 1) labelColumn = 1;
                if (valueColumn < 1) valueColumn = DetectNumericColumn(values, columnsCount);
                if (valueColumn < 1) valueColumn = columnsCount;
                if (labelColumn > columnsCount) labelColumn = columnsCount;
                if (valueColumn > columnsCount) valueColumn = columnsCount;

                var summary = PivotTableService.SumByRow(values, labelColumn, valueColumn);

                // 写入统一走 writer（WriteGenericRange：按结果尺寸写回 + 记录撤销快照），而不是直写 Value2——
                // 直写会绕过撤销系统，使透视表对目标区的一次性覆盖不可回滚，与"全程可回滚"契约冲突。
                if (writer == null)
                {
                    return new ToolResult { Ok = false, ErrorCode = "WRITER_UNAVAILABLE", Data = "当前上下文未挂接写入器，无法落盘透视表。" };
                }

                // 未指定输出位置时放在源区域下方 2 行、与源同列起始，避免覆盖源数据。
                if (string.IsNullOrWhiteSpace(targetAddress))
                {
                    var top = Convert.ToInt32(source.Row);
                    var left = Convert.ToInt32(source.Column);
                    targetAddress = Convert.ToString(sheet.Cells[top + rowsCount + 1, left].Address);
                }

                // 已从 summary 取完全部需要的数据：先释放本地共享 COM 句柄，避免 writer 的 FinalRelease
                // 与本方法 finally 的 Release 竞态（同一 Worksheet RCW）。
                var usedSheetName = Convert.ToString(sheet.Name);
                Release((object)source); source = null;
                Release((object)sheet); sheet = null;
                Release((object)workbook); workbook = null;

                var writeArgs = new Dictionary<string, object>
                {
                    ["sheet"] = Optional(args, "sheet"),
                    ["address"] = targetAddress
                };
                var res = writer(writeArgs, summary.Values, "value", "create_excel_pivot_summary");
                if (!res.Ok)
                {
                    return res;
                }

                var targetAddressOut = targetAddress;
                if (res.AffectedRanges != null && res.AffectedRanges.Count > 0)
                {
                    var af = res.AffectedRanges[0];
                    var bang = af.LastIndexOf('!');
                    targetAddressOut = bang >= 0 ? af.Substring(bang + 1) : af;
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{usedSheetName}!{targetAddressOut}" },
                    Data = new
                    {
                        source_rows = summary.SourceRows,
                        row_count = summary.RowLabels.Length,
                        target_address = targetAddressOut
                    }
                };
            }
            finally
            {
                Release((object)source);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult CreateChart(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var chartType = Optional(args, "chart_type") ?? "ColumnClustered";
            if (!ChartTypeMapper.IsKnown(chartType))
            {
                return Invalid("INVALID_CHART_TYPE", $"chart_type 仅支持 ColumnClustered/BarClustered/Line/Pie/Area/Scatter/Histogram，当前为 {chartType}。");
            }

            var title = Optional(args, "title");
            var anchor = Optional(args, "anchor");
            var width = ExcelToolHostHelpers.GetIntSafe(args, "width", 420);
            var height = ExcelToolHostHelpers.GetIntSafe(args, "height", 260);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Convert.ToString(sheet.Name) + " 分析";
                }

                double left;
                double top;
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    left = Convert.ToDouble(anchorRange.Left);
                    top = Convert.ToDouble(anchorRange.Top);
                }
                else
                {
                    left = Convert.ToDouble(source.Left) + Convert.ToDouble(source.Width) + 12;
                    top = Convert.ToDouble(source.Top);
                }

                var resolvedType = ChartTypeMapper.Resolve(chartType);
                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;

                if (ChartTypeMapper.IsHistogram(chartType))
                {
                    if (!TryNativeHistogram(chart, source))
                    {
                        // 旧版 Excel 无原生直方图时回退到分箱柱状图。
                        dynamic histRange = sheet.Range[sourceAddress];
                        bool binned;
                        try
                        {
                            // M-D3 二轮审查⑧：无数值行时回退路径残留空图表，改为显式拒绝并清理
                            binned = CreateBinnedHistogram(histRange, chart);
                        }
                        finally
                        {
                            // M16 全量审查：直方图路径的 Range RCW 释放（原实现从不释放）。
                            Release((object)histRange);
                        }

                        if (!binned)
                        {
                            TryDeleteChartObject(chartObject);
                            chartObject = null;
                            chart = null;
                            return Invalid("NO_NUMERIC_DATA",
                                "数据源无数值行，无法生成直方图。请检查 source_address 是否包含数值列。");
                        }
                    }
                }
                else
                {
                    chart.ChartType = ChartTypeMapper.ResolveXlChartType(chartType);
                    chart.SetSourceData(source);
                    ApplyChartVariant(chart, chartType);
                }

                chart.HasTitle = true;
                chart.ChartTitle.Text = title;

                // v3.7.1.41 可读性五件套②：柱/条/折线/雷达/饼环自动数据标签（点数少才开，避免拥挤），
                // 用户不必对照表格估读数值；data_labels=on/off 可强制覆盖，缺省 auto。
                var dataLabelsMode = (Optional(args, "data_labels") ?? "auto").Trim().ToLowerInvariant();
                bool dataLabelsOn;
                if (dataLabelsMode == "on")
                {
                    dataLabelsOn = true;
                }
                else if (dataLabelsMode == "off")
                {
                    dataLabelsOn = false;
                }
                else
                {
                    dataLabelsOn = ShouldAutoDataLabels(chart, chartType);
                }

                if (dataLabelsOn)
                {
                    TryApplyDataLabels(chart, chartType);
                }

                var sheetName = Convert.ToString(sheet.Name);
                var chartName = Convert.ToString(chartObject.Name);
                var sourceAddressForRecord = Convert.ToString(source.Address);
                recorder?.Invoke(
                    "create_excel_chart",
                    args,
                    "chart-insert",
                    new { sheet = sheetName, chart_name = chartName, source_address = sourceAddressForRecord, chart_type = resolvedType },
                    new ChartInsertPayload { Sheet = sheetName, ChartName = chartName });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartName,
                        chart_type = resolvedType,
                        source_address = Convert.ToString(source.Address),
                        anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                        title,
                        data_labels = dataLabelsOn ? "on" : "off"
                    }
                };
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                // 2026-09-12：不再释放 source/sheet/workbook——调用方（repair_chart 等）共享同身份 RCW，
                // 减计数断开会让调用方后续操作炸 InvalidComObjectException（M-D1 纪律），GC 兜底即可。
            }
        }

        /// <summary>数据标签 auto 规则：饼/环恒开（百分比）；柱/条/折线/雷达在「系列数≤3 且每系列点数≤12」时开
        /// （标签少才有可读性）；面积/散点/直方及超限图不开。任何 COM 读取失败按保守不开处理。</summary>
        private bool ShouldAutoDataLabels(dynamic chart, string chartType)
        {
            var isPie = chartType.IndexOf("Pie", StringComparison.OrdinalIgnoreCase) >= 0
                        || chartType.IndexOf("Doughnut", StringComparison.OrdinalIgnoreCase) >= 0;
            var isBarLike = chartType.IndexOf("Column", StringComparison.OrdinalIgnoreCase) >= 0
                            || chartType.IndexOf("Bar", StringComparison.OrdinalIgnoreCase) >= 0
                            || chartType.IndexOf("Line", StringComparison.OrdinalIgnoreCase) >= 0
                            || chartType.IndexOf("Radar", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!isPie && !isBarLike)
            {
                return false;
            }

            dynamic seriesCollection = null;
            try
            {
                seriesCollection = chart.SeriesCollection();
                var seriesCount = Convert.ToInt32(seriesCollection.Count);
                if (isPie)
                {
                    return seriesCount >= 1;
                }

                if (seriesCount < 1 || seriesCount > 3)
                {
                    return false;
                }

                for (var i = 1; i <= seriesCount; i++)
                {
                    dynamic s = null;
                    try
                    {
                        s = seriesCollection.Item(i);
                        if (Convert.ToInt32(s.Points.Count) > 12)
                        {
                            return false;
                        }
                    }
                    finally
                    {
                        Release((object)s);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                Release((object)seriesCollection);
            }
        }

        /// <summary>实际打开数据标签：饼/环显示百分比，其余显示值。增强性功能，失败静默降级不影响出图。</summary>
        private void TryApplyDataLabels(dynamic chart, string chartType)
        {
            var isPie = chartType.IndexOf("Pie", StringComparison.OrdinalIgnoreCase) >= 0
                        || chartType.IndexOf("Doughnut", StringComparison.OrdinalIgnoreCase) >= 0;
            dynamic seriesCollection = null;
            try
            {
                seriesCollection = chart.SeriesCollection();
                for (var i = 1; i <= Convert.ToInt32(seriesCollection.Count); i++)
                {
                    dynamic s = null;
                    dynamic labels = null;
                    try
                    {
                        s = seriesCollection.Item(i);
                        s.HasDataLabels = true;
                        if (isPie)
                        {
                            labels = s.DataLabels;
                            labels.ShowPercentage = true;
                            labels.ShowValue = false;
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Release((object)labels);
                        Release((object)s);
                    }
                }
            }
            catch (Exception dlEx)
            {
                // 数据标签是增强性功能：失败降级为无标签出图，但降级必须留痕（可读性五件套回归排查依据）。
                ExcelTrace.Warn("chart_tool", $"数据标签应用失败已降级（不影响出图）：" + dlEx.Message);
            }
            finally
            {
                Release((object)seriesCollection);
            }
        }

        /// <summary>
        /// 图表变体后处理（2026-09-11）：XlChartType 无分离环形/填充雷达/标记雷达枚举，
        /// 以基础类型 + 系列级格式实现同 UI 效果；失败不阻断（保底为可用的基础图）。
        /// </summary>
        private static void ApplyChartVariant(dynamic chart, string chartType)
        {
            try
            {
                if (chartType == "DoughnutExploded")
                {
                    dynamic series = chart.SeriesCollection().Item(1);
                    dynamic pts = series.Points();
                    int n = Convert.ToInt32(pts.Count);
                    for (int i = 1; i <= n; i++)
                    {
                        dynamic pt = pts.Item(i);
                        try { pt.Explosion = 18; } catch { }
                        Release((object)pt);
                    }
                    Release((object)pts);
                    Release((object)series);
                }
                else if (chartType == "RadarFilled")
                {
                    dynamic sc = chart.SeriesCollection();
                    for (int i = 1; i <= Convert.ToInt32(sc.Count); i++)
                    {
                        dynamic sr = sc.Item(i);
                        try
                        {
                            sr.Format.Fill.Visible = -1;          // msoTrue
                            sr.Format.Fill.Transparency = 0.35;   // 半透明防系列互遮
                            sr.MarkerStyle = -4142;               // xlMarkerStyleNone
                        }
                        catch { }
                        Release((object)sr);
                    }
                    Release((object)sc);
                }
                else if (chartType == "RadarMarkers")
                {
                    dynamic sc2 = chart.SeriesCollection();
                    for (int i = 1; i <= Convert.ToInt32(sc2.Count); i++)
                    {
                        dynamic sr2 = sc2.Item(i);
                        try { sr2.MarkerStyle = 8; sr2.MarkerSize = 7; } catch { }  // xlMarkerStyleCircle
                        Release((object)sr2);
                    }
                    Release((object)sc2);
                }
            }
            catch { }
        }

        /// <summary>尝试设置原生直方图类型并绑定数据源；返回 false 表示当前 Excel 不支持（回退分箱）。
        /// 工具分类审查批次双重真 bug 修复：①xlHistogram 权威值 = 118（2016+ 新图表族 117~123 实证域，
        /// 与 create_chart_advanced nativeTypes 同源）——原 58 是 xlBarStacked（横向堆积条形图），合法值
        /// 赋值不抛错致「恒成功」假象，直方图从未真正渲染；②原实现只赋 ChartType 未绑数据源（空图）。</summary>
        private static bool TryNativeHistogram(dynamic chart, dynamic source)
        {
            try
            {
                chart.ChartType = 118; // xlHistogram（与 create_chart_advanced nativeTypes["histogram"] 对齐）
                chart.SetSourceData(source);
                return true;
            }
            catch
            {
                try
                {
                    chart.ChartType = ChartTypeMapper.ResolveXlChartType("Column");
                }
                catch
                {
                    // 兜底：仍尝试柱状图。
                }
                return false;
            }
        }

        /// <summary>
        /// 直方图回退：抽取区域内数值→分箱→写辅助网格→柱状图。辅助数据写入图表所在表的空白区域。
        /// M-D3 二轮审查⑧：返回 false 表示区域内无数值行（调用方须清理空图表并显式拒绝，不静默残留）。
        /// </summary>
        private static bool CreateBinnedHistogram(dynamic range, dynamic chart)
        {
            var values = ExcelAiAssistant.Excel.ExcelToolHostHelpers.FlattenValues(
                range.Value2,
                Convert.ToInt32(range.Rows.Count),
                Convert.ToInt32(range.Columns.Count));
            var numbers = ExcelAiAssistant.Excel.HistogramBinner.BuildBins(
                ExtractNumbers(values),
                8);

            // 注意：range 被重新赋值为其所属 Worksheet 后，Rows.Count 返回整表最大行
            // （Excel 2016+ 为 1,048,576），绝不能直接 +N 作为锚点（必然越界）。
            // 改用 UsedRange 的下侧空白定位辅助数据起始行，并夹紧到工作表行上限内。
            dynamic sheet = range.Worksheet;
            dynamic used = null;
            dynamic target = null;
            try
            {
                used = sheet.UsedRange;
                var lastUsedRow = Convert.ToInt32(used.Row) + Convert.ToInt32(used.Rows.Count) - 1;
                var maxRows = Convert.ToInt32(sheet.Rows.Count);
                var anchorRow = Math.Min(Math.Max(lastUsedRow + 4, 1), maxRows);
                if (numbers.Length == 0)
                {
                    // M-D3 二轮审查⑧：无数值行显式告知调用方，不再静默 return
                    return false;
                }

                var endRow = Math.Min(anchorRow + numbers.GetLength(0) - 1, maxRows);
                dynamic cellA = sheet.Cells[anchorRow, 1];
                dynamic cellB = sheet.Cells[endRow, 2];
                dynamic dataRange = null;
                try
                {
                    dataRange = sheet.Range[cellA, cellB];
                    dataRange.Value2 = numbers;
                    chart.SetSourceData(dataRange);
                    chart.ChartType = ChartTypeMapper.ResolveXlChartType("Column");
                }
                finally
                {
                    // M16 全量审查：内联 Cells/Range 临时 RCW 释放（原实现每次泄漏 3 个）。
                    Release((object)dataRange);
                    Release((object)cellA);
                    Release((object)cellB);
                }
                return true;
            }
            finally
            {
                Release((object)target);
                Release((object)used);
                Release((object)sheet);
            }
        }

        private static double[] ExtractNumbers(object[,] values)
        {
            var list = new List<double>();
            // 跳过首行表头，遍历所有单元格收集可解析数值，避免仅 1 列时 values[row,1] 越界。
            for (var row = 1; row < values.GetLength(0); row++)
            {
                for (var col = 0; col < values.GetLength(1); col++)
                {
                    var cell = values[row, col];
                    // CVErr 错误格归一为错误文本，防止 "-2146826246" 被解析成数字进直方图（M16）。
                    var text = cell == null ? null : Convert.ToString(ExcelErrorValues.Normalize(cell));
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (double.TryParse(text.Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var number))
                    {
                        list.Add(number);
                    }
                }
            }

            return list.ToArray();
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        /// <summary>取当前多格选区地址；无选区或选区为单格时返回 null（单格透视无意义）。</summary>
        private string CurrentSelectionAddress()
        {
            try
            {
                dynamic selection = application.Selection;
                if (selection == null)
                {
                    return null;
                }

                var rows = Convert.ToInt32(selection.Rows.Count);
                var cols = Convert.ToInt32(selection.Columns.Count);
                return rows > 1 || cols > 1 ? Convert.ToString(selection.Address) : null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>从数据区（跳过标题行）自右向左探测第一个含数值的列号（1 基），无则返回 0。</summary>
        private static int DetectNumericColumn(object[,] values, int columnCount)
        {
            if (values == null || columnCount < 1)
            {
                return 0;
            }

            var rows = values.GetLength(0);
            for (var column = columnCount; column >= 1; column--)
            {
                for (var row = 1; row < rows; row++)
                {
                    var value = values[row, column - 1];
                    if (value == null)
                    {
                        continue;
                    }

                    if (double.TryParse(
                        Convert.ToString(value, CultureInfo.InvariantCulture),
                        NumberStyles.Float | NumberStyles.AllowThousands,
                        CultureInfo.InvariantCulture,
                        out _))
                    {
                        return column;
                    }
                }
            }

            return 0;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        /// <summary>
        /// M-D1 P1：双轴组合图（柱形主轴 + 折线次轴）。
        /// source_address 需含 ≥2 个数值系列（第 1 列为类别，其余列逐一成系列）；
        /// secondary_series 指定哪些系列（1 基系列序号，可多个）挂次轴并改为折线，缺省最后一个系列。
        /// 复用 M-D0 探测结论：逐系列 SeriesCollection 编辑 + AxisGroup=xlSecondary(2)。
        /// </summary>
        public ToolResult CreateComboChart(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var title = Optional(args, "title");
            var anchor = Optional(args, "anchor");
            var width = ExcelToolHostHelpers.GetIntSafe(args, "width", 480);
            var height = ExcelToolHostHelpers.GetIntSafe(args, "height", 280);
            var secondary = ParseSecondarySeries(Optional(args, "secondary_series"));

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            dynamic seriesColl = null;
            var releasedSeries = new List<object>();
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Convert.ToString(sheet.Name) + " 组合分析";
                }

                double left;
                double top;
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    left = Convert.ToDouble(anchorRange.Left);
                    top = Convert.ToDouble(anchorRange.Top);
                }
                else
                {
                    left = Convert.ToDouble(source.Left) + Convert.ToDouble(source.Width) + 12;
                    top = Convert.ToDouble(source.Top);
                }

                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;
                chart.ChartType = 51; // xlColumnClustered 主轴柱形
                chart.SetSourceData(source);

                seriesColl = chart.SeriesCollection();
                int seriesCount = Convert.ToInt32(seriesColl.Count);
                if (seriesCount < 2)
                {
                    // M-D3 二轮审查⑥：错误路径清理已创建的空图表（AC-5 无半成品）
                    TryDeleteChartObject(chartObject);
                    chartObject = null;
                    chart = null;
                    return Invalid("COMBO_NEEDS_TWO_SERIES",
                        $"双轴组合图需要 ≥2 个数值系列（首列为类别），当前仅解析出 {seriesCount} 个。请检查 source_address 是否包含多列数值。");
                }

                // secondary_series 缺省=最后一个系列挂次轴折线
                var secondarySet = secondary.Count > 0
                    ? secondary
                    : new HashSet<int> { seriesCount };
                foreach (var idx in secondarySet)
                {
                    if (idx < 1 || idx > seriesCount)
                    {
                        // M-D3 二轮审查⑥：同上，先清理再返回
                        TryDeleteChartObject(chartObject);
                        chartObject = null;
                        chart = null;
                        return Invalid("INVALID_SECONDARY_SERIES",
                            $"secondary_series 序号 {idx} 超出范围（1..{seriesCount}）。");
                    }
                }

                foreach (var idx in secondarySet)
                {
                    dynamic series = seriesColl.Item(idx);
                    releasedSeries.Add(series);
                    series.AxisGroup = 2;   // xlSecondary
                    series.ChartType = 4;   // xlLine
                }

                chart.HasTitle = true;
                chart.ChartTitle.Text = title;


                var sheetName = Convert.ToString(sheet.Name);
                var chartName = Convert.ToString(chartObject.Name);
                recorder?.Invoke(
                    "create_combo_chart",
                    args,
                    "chart-insert",
                    new { sheet = sheetName, chart_name = chartName, source_address = Convert.ToString(source.Address), chart_type = "ComboDualAxis", secondary_series = secondarySet },
                    new ChartInsertPayload { Sheet = sheetName, ChartName = chartName });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartName,
                        chart_type = "ComboDualAxis",
                        series_count = seriesCount,
                        secondary_series = secondarySet,
                        source_address = Convert.ToString(source.Address),
                        anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                        title
                    }
                };
            }
            finally
            {
                foreach (var s in releasedSeries)
                {
                    Release(s);
                }
                Release((object)seriesColl);
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                Release((object)source);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// M-D3 二轮审查⑦：series.Values（COM variant 数组）→ double[]；非数值/缺失点位记为 NaN，
        /// 截断比较 NaN>cap 恒 false，天然跳过。COM 数组是 1 基下界（1..N），必须按 LowerBound 遍历
        /// （0 基 GetValue 曾致 T13/T20 全挂「索引超出了数组界限」）。
        /// </summary>
        private static double[] ToDoubleArray(object rawValues)
        {
            if (!(rawValues is Array arr))
            {
                return Array.Empty<double>();
            }
            int lb = arr.GetLowerBound(0);
            int ub = arr.GetUpperBound(0);
            var result = new double[ub - lb + 1];
            for (int i = lb; i <= ub; i++)
            {
                object v = arr.GetValue(i);
                result[i - lb] = v is double d ? d : double.NaN;
            }
            return result;
        }

        /// <summary>解析 secondary_series 参数："2" / "2,3" / [2,3] 均可；空返回空集合。</summary>
        private static HashSet<int> ParseSecondarySeries(string raw)
        {
            var result = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return result;
            }

            foreach (var part in raw.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (int.TryParse(part.Trim(), out var n))
                {
                    result.Add(n);
                }
            }

            return result;
        }

        /// <summary>
        /// M-D3 二轮审查⑥：错误路径清理已创建的 ChartObject（AC-5 无半成品语义）。
        /// 调用后调用方应将自身引用置 null，避免 finally 重复 Release 已失效 RCW。
        /// </summary>
        private static void TryDeleteChartObject(dynamic chartObject)
        {
            try { chartObject?.Delete(); } catch { }
        }

        /// <summary>
        /// M-D1 P1：环形图（xlDoughnut=-4120 / 分离型 xlDoughnutExploded=-4121），支持 hole_size 孔径（10-90，默认 50）。
        /// </summary>
        public ToolResult CreateDoughnut(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var chartType = (Optional(args, "chart_type") ?? "doughnut").Trim().ToLowerInvariant();
            if (chartType != "doughnut" && chartType != "doughnutexploded")
            {
                return Invalid("INVALID_CHART_TYPE", $"环形图 chart_type 仅支持 doughnut/doughnutexploded，当前为 {chartType}。");
            }

            var holeSize = ExcelToolHostHelpers.GetIntSafe(args, "hole_size", 50);
            if (holeSize < 10 || holeSize > 90)
            {
                return Invalid("INVALID_HOLE_SIZE", $"hole_size 需在 10-90 之间，当前为 {holeSize}。");
            }

            var title = Optional(args, "title");
            var anchor = Optional(args, "anchor");
            var width = ExcelToolHostHelpers.GetIntSafe(args, "width", 360);
            var height = ExcelToolHostHelpers.GetIntSafe(args, "height", 300);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Convert.ToString(sheet.Name) + " 构成";
                }

                double left;
                double top;
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    left = Convert.ToDouble(anchorRange.Left);
                    top = Convert.ToDouble(anchorRange.Top);
                }
                else
                {
                    left = Convert.ToDouble(source.Left) + Convert.ToDouble(source.Width) + 12;
                    top = Convert.ToDouble(source.Top);
                }

                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;
                // 2026-09-12 enumprobe 实证：XlChartType 无 -4121（xlDoughnutExploded 手抄臆造值，
                // 赋值即 E_FAIL——AI 按 Schema 传 doughnutexploded 必踩）。改 xlDoughnut + Explosion 变体后处理，
                // 与 create_excel_chart 的 ApplyChartVariant 同口径。
                chart.ChartType = -4120;
                chart.SetSourceData(source);
                if (chartType == "doughnutexploded") { ApplyChartVariant(chart, "DoughnutExploded"); }
                try
                {
                    // PlotArea 内 DoughnutHoleSize 仅对 xlDoughnut 有效（分离型由 ChartType 自带形态）
                    if (chartType == "doughnut")
                    {
                        chart.ChartGroups().Item(1).DoughnutHoleSize = holeSize;
                    }
                }
                catch (Exception holeEx)
                {
                    // 旧版本 Excel 对 DoughnutHoleSize 支持差异：不致命，记录后继续
                    Console.WriteLine("[create_doughnut] hole_size 设置降级：" + holeEx.Message);
                }

                chart.HasTitle = true;
                chart.ChartTitle.Text = title;


                var sheetName = Convert.ToString(sheet.Name);
                var chartName = Convert.ToString(chartObject.Name);
                recorder?.Invoke(
                    "create_doughnut",
                    args,
                    "chart-insert",
                    new { sheet = sheetName, chart_name = chartName, source_address = Convert.ToString(source.Address), chart_type = chartType, hole_size = holeSize },
                    new ChartInsertPayload { Sheet = sheetName, ChartName = chartName });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartName,
                        chart_type = chartType,
                        hole_size = holeSize,
                        source_address = Convert.ToString(source.Address),
                        anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                        title
                    }
                };
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                // 2026-09-12：不再释放 source/sheet/workbook——调用方（repair_chart 等）共享同身份 RCW，
                // 减计数断开会让调用方后续操作炸 InvalidComObjectException（M-D1 纪律），GC 兜底即可。
            }
        }

        /// <summary>
        /// M-D2 技巧型图表：特大值截断柱（辅助序列 + 双系列同色叠加 + 柱顶自定义标签标真实值）。
        /// 数据源首列类别 + 1 数值列；特大值超出均值 × truncate_factor 时截断显示，
        /// 柱顶数据标签显示真实值（“值(截断)”样式）。
        /// </summary>
        public ToolResult CreateTruncatedColumn(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var factor = 3.0;
            if (args.TryGetValue("truncate_factor", out var tfRaw) && tfRaw != null)
            {
                double.TryParse(Convert.ToString(tfRaw), NumberStyles.Any, CultureInfo.InvariantCulture, out factor);
            }
            var title = Optional(args, "title") ?? "特大值截断柱";
            var anchor = Optional(args, "anchor");
            var width = ExcelToolHostHelpers.GetIntSafe(args, "width", 460);
            var height = ExcelToolHostHelpers.GetIntSafe(args, "height", 280);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            dynamic seriesColl = null;
            dynamic series = null;
            dynamic points = null;
            dynamic point = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];

                // 读取数值列，判断截断阈值
                int rowCount = Convert.ToInt32(source.Rows.Count);
                double max = double.MinValue;
                double sum = 0;
                int n = 0;
                // M-D3 二轮审查⑦：记录「实际行号→值」，Points 索引 = 行号-1（跳过表头）。
                // 原实现用 values 列表序号对齐 Points，源数据混入非数值行时标签错位。
                var numericRows = new List<KeyValuePair<int, double>>();
                for (int r = 2; r <= rowCount; r++)
                {
                    dynamic c = source.Cells.Item[r, 2];
                    object v = c.Value2;
                    Release((object)c);
                    if (v is double d)
                    {
                        numericRows.Add(new KeyValuePair<int, double>(r, d));
                        max = Math.Max(max, d);
                        sum += d;
                        n++;
                    }
                }
                if (n == 0)
                {
                    return Invalid("NO_NUMERIC_DATA", "数据源无数值行，无法生成截断柱。");
                }

                double mean = sum / n;
                bool needsTruncate = max > mean * factor;
                double cap = needsTruncate ? mean * factor * 0.95 : max;

                double left;
                double top;
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    left = Convert.ToDouble(anchorRange.Left);
                    top = Convert.ToDouble(anchorRange.Top);
                }
                else
                {
                    left = Convert.ToDouble(source.Left) + Convert.ToDouble(source.Width) + 12;
                    top = Convert.ToDouble(source.Top);
                }

                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;
                chart.ChartType = 51; // xlColumnClustered
                chart.SetSourceData(source);

                seriesColl = chart.SeriesCollection();
                series = seriesColl.Item(1);
                if (needsTruncate)
                {
                    // 截断：把超出 cap 的点值改写为 cap（只改图内显示；辅助标签标注真实值）。
                    // M-D3 二轮审查⑦（二轮收敛）：按「图内实际点值」对齐标签——Excel 对混入文本的
                    // 源区域，其 Points 点序与源行号的映射不可靠（行号-1 对齐实测 0x80070057 参数无效）。
                    // series.Values 是图内实际绘制值数组，与 Points 一一对应，语义最稳。
                    points = series.Points();
                    object rawValues = series.Values;
                    double[] plotValues = ToDoubleArray(rawValues);
                    Release((object)rawValues);
                    int pointCount = Math.Min(plotValues.Length, Convert.ToInt32(points.Count));
                    for (int pi = 1; pi <= pointCount; pi++)
                    {
                        double pv = plotValues[pi - 1];
                        if (pv > cap)
                        {
                            point = points.Item(pi);
                            point.HasDataLabel = true;
                            point.DataLabel.Text = pv.ToString("#,##0") + "(截断)";
                            Release((object)point);
                            point = null;
                        }
                    }
                }
                else
                {
                    // 无特大值：不截断，仅常规柱形 + 数值标签
                    series.HasDataLabels = true;
                }

                chart.HasTitle = true;
                chart.ChartTitle.Text = needsTruncate ? title + "（已截断）" : title;

                var sheetName = Convert.ToString(sheet.Name);
                var chartName = Convert.ToString(chartObject.Name);
                recorder?.Invoke(
                    "create_truncated_column",
                    args,
                    "chart-insert",
                    new { sheet = sheetName, chart_name = chartName, truncated = needsTruncate, cap, max },
                    new { Sheet = sheetName, ChartName = chartName });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartName,
                        truncated = needsTruncate,
                        cap,
                        max,
                        series_count = 1,
                        source_address = Convert.ToString(source.Address),
                        title
                    }
                };
            }
            finally
            {
                Release((object)point);
                Release((object)points);
                Release((object)series);
                Release((object)seriesColl);
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                Release((object)source);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>
        /// M-D3 Q2 决策交付：create_chart_advanced 技巧型图表族单入口。
        /// type 枚举：bullet 子弹图 / waffle 华夫图 / slider 滑珠图 / funnel 漏斗图 /
        /// butterfly 蝴蝶图 / gantt 甘特图。全部「堆积柱/条 + 辅助序列 + 纯 COM」拼装，
        /// Excel 2013 兼容；辅助数据写源区域右侧空白列（只读源数据，NFR-5）；
        /// 失败清理半成品（AC-5 同源）；错误口径统一（NO_NUMERIC_DATA/INVALID_ADV_TYPE）。
        /// 参数约定：source_address 首列=类别标签；bullet/slider 需 target 列（目标值），
        /// gantt 需 start_days + duration_days 两数值列，其余单数值列。
        /// </summary>
        /// <summary>列号 → 列字母（1 基；原生图表临时表块写地址用）。</summary>
        private static string ColumnLetterSafe(int col)
        {
            string s = "";
            while (col > 0) { int m = (col - 1) % 26; s = (char)('A' + m) + s; col = (col - 1) / 26; }
            return s;
        }

        /// <summary>
        /// 创建图表自动体检闭环（v3.7.1.37）：生成成功后自动 verify_chart（结构+渲染像素核验），
        /// 发现空系列/空白渲染且提供了源区域时自动重绑/重建一次，复检结果与截图（chart_image）
        /// 以 auto_check 并入返回值——AI 与用户无需再次调用即可获得"图对不对"的依据。
        /// 自检失败不影响主结果（旁路），但全程记录 ExcelTrace。
        /// </summary>
        private ToolResult AttachAutoCheck(ToolResult result, dynamic sheet, string chartName, string sourceAddress, string advType)
        {
            if (result == null || !result.Ok) { return result; }
            try
            {
                string sheetName = Convert.ToString(sheet.Name);
                var diag = new ExcelChartService(application, null);
                Dictionary<string, object> VArgs() => new Dictionary<string, object>
                {
                    ["sheet"] = sheetName,
                    ["chart_name"] = chartName
                };
                var v1 = diag.VerifyChart(VArgs());
                JObject jv = v1.Data == null ? new JObject() : JObject.FromObject(v1.Data);
                string verdict = jv.Value<string>("verdict") ?? "unknown";
                bool needRepair = verdict.Contains("empty") || verdict.Contains("blank")
                    || (jv.Value<int?>("point_total") ?? 0) == 0
                    || jv.Value<bool?>("data_rendered") == false;
                bool repaired = false;
                string repairNote = null;
                if (needRepair && !string.IsNullOrWhiteSpace(sourceAddress))
                {
                    var rArgs = new Dictionary<string, object>
                    {
                        ["sheet"] = sheetName,
                        ["chart_name"] = chartName,
                        ["source_address"] = sourceAddress
                    };
                    if (!string.IsNullOrWhiteSpace(advType)) { rArgs["type"] = advType; }
                    var rr = diag.RepairChart(rArgs, this);
                    repaired = rr.Ok;
                    repairNote = rr.Ok ? "自动重绑/重建完成" : ("自动修复失败：" + Convert.ToString(rr.Data));
                    if (rr.Ok)
                    {
                        v1 = diag.VerifyChart(VArgs());
                        jv = v1.Data == null ? new JObject() : JObject.FromObject(v1.Data);
                        verdict = jv.Value<string>("verdict") ?? verdict;
                    }
                }
                var autoCheck = new JObject
                {
                    ["verdict"] = verdict,
                    ["point_total"] = jv.Value<int?>("point_total"),
                    ["nonblank_ratio"] = jv.Value<double?>("nonblank_ratio"),
                    ["chart_image"] = jv.Value<string>("chart_image"),
                    ["repaired"] = repaired,
                    ["repair_note"] = repairNote
                };
                JObject data = result.Data == null ? new JObject() : JObject.FromObject(result.Data);
                data["auto_check"] = autoCheck;
                result.Data = data;
                ExcelTrace.Info("create_chart_advanced", "auto_check verdict=" + verdict + " repaired=" + repaired);
            }
            catch (Exception ex)
            {
                ExcelTrace.Error("create_chart_advanced", "AttachAutoCheck", ex);
            }
            return result;
        }

        public ToolResult CreateChartAdvanced(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var type = (Optional(args, "type") ?? "bullet").Trim().ToLowerInvariant();
            // 2016+ 原生新图表四类（FR-4.3）：waterfall/treemap/sunburst/pareto，
            // 与旧六类共用 source_address/title/anchor/width/height 参数。
            var validTypes = new[] { "bullet", "waffle", "slider", "funnel", "butterfly", "gantt",
                                     "waterfall", "treemap", "sunburst", "pareto", "histogram" };
            if (!validTypes.Contains(type))
            {
                return Invalid("INVALID_ADV_TYPE",
                    "type 仅支持 " + string.Join("/", validTypes) + "，当前为 " + type + "。");
            }
            // 原生新图表类型 → XlChartType 枚举值（类型库权威值，见踩坑手册真机探测记录）
            var nativeTypes = new Dictionary<string, int>
            {
                ["waterfall"] = 119, ["treemap"] = 117, ["sunburst"] = 120, ["pareto"] = 122,
                ["histogram"] = 118
            };
            var title = Optional(args, "title");
            var anchor = Optional(args, "anchor");
            var width = ExcelToolHostHelpers.GetIntSafe(args, "width", 520);
            var height = ExcelToolHostHelpers.GetIntSafe(args, "height", 320);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            var cleanupRanges = new List<dynamic>();
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                source = sheet.Range[sourceAddress];
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = Convert.ToString(sheet.Name) + " 高级图表";
                }
                int rowCount = Convert.ToInt32(source.Rows.Count);
                int colCount = Convert.ToInt32(source.Columns.Count);
                int firstCol = Convert.ToInt32(source.Column);
                int firstRow = Convert.ToInt32(source.Row);

                // 读明细：labels + 全部数值列（1 基相对列号→值），失败口径统一 NO_NUMERIC_DATA
                var labels = new List<string>();
                var numericCols = new Dictionary<int, List<double>>(); // 相对列号 → 按行数值
                // histogram 单列观测值支持（Schema 宣称"仅需 1 列原始观测值"）：值挂伪第 2 列位，
                // native 分支 valueColIdx=1 取真值（2026-09-11 修复：原读取循环 c 从 2 起，
                // colCount=1 时收集不到数值而误报 NO_NUMERIC_DATA，单列路径从未真正可达）
                bool histSingleCol = type == "histogram" && colCount == 1;
                if (histSingleCol) { numericCols[2] = new List<double>(); }
                for (int c = 2; c <= colCount; c++) { numericCols[c] = new List<double>(); }
                for (int r = 2; r <= rowCount; r++)
                {
                    dynamic lc = source.Cells.Item[r, 1];
                    object lv = lc.Value2;
                    Release((object)lc);
                    if (lv == null) { continue; }
                    labels.Add(Convert.ToString(lv));
                    if (histSingleCol)
                    {
                        numericCols[2].Add(lv is double d1v ? d1v : double.NaN);
                        continue;
                    }
                    foreach (var key in numericCols.Keys.ToList())
                    {
                        dynamic vc = source.Cells.Item[r, key];
                        object vv = vc.Value2;
                        Release((object)vc);
                        numericCols[key].Add(vv is double d ? d : double.NaN);
                    }
                }
                int nRows = labels.Count;
                bool hasAnyNumber = numericCols.Values.Any(v => v.Any(x => !double.IsNaN(x)));
                if (nRows == 0 || !hasAnyNumber)
                {
                    return Invalid("NO_NUMERIC_DATA",
                        "数据源无类别行或无数值，无法生成 " + type + " 图表。请检查 source_address。");
                }

                // 辅助数据锚点：UsedRange 右侧 +4 列（夹紧行上限；只读源数据，辅助列可随后清理）
                dynamic usedRange = sheet.UsedRange;
                int usedLastCol = Convert.ToInt32(usedRange.Column) + Convert.ToInt32(usedRange.Columns.Count) - 1;
                int usedLastRow = Convert.ToInt32(usedRange.Row) + Convert.ToInt32(usedRange.Rows.Count) - 1;
                Release((object)usedRange);
                int auxCol = usedLastCol + 4;
                int maxRows = Convert.ToInt32(sheet.Rows.Count);
                if (auxCol > Convert.ToInt32(sheet.Columns.Count) - 8)
                {
                    return Invalid("NO_AUX_SPACE", "源区域右侧无足够空白列放置辅助数据，请清理右侧内容后重试。");
                }

                double left, top;
                if (!string.IsNullOrWhiteSpace(anchor))
                {
                    anchorRange = sheet.Range[anchor];
                    left = Convert.ToDouble(anchorRange.Left);
                    top = Convert.ToDouble(anchorRange.Top);
                }
                else
                {
                    left = Convert.ToDouble(source.Left) + Convert.ToDouble(source.Width) + 12;
                    top = Convert.ToDouble(source.Top);
                }

                // ---- 2016+ 原生新图表（waterfall/treemap/sunburst/pareto）----
                // 真机结论（踩坑手册）：① Chart.ChartType 赋值对新类型枚举（117~123）报 0x80070057，
                // 必须走 Shapes.AddChart2(-1, type, ...) 创建；② AddChart2 默认 Series 存在但源数据为空，
                // 需手动给默认 Series 赋 Values/XValues（SetSourceData 对新图表报 E_FAIL/NotImplemented）；
                // ③ 前置版本守卫（Excel 16+），低版本此处显式报错而非 COM 深处报错。
                if (nativeTypes.TryGetValue(type, out int nativeChartType))
                {
                    double ver = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
                    if (ver < 16.0)
                    {
                        TryDeleteChartObject(chartObject);
                        chartObject = null;
                        CleanupAuxRanges(sheet, cleanupRanges);
                        return Invalid("CHART_VERSION_UNSUPPORTED",
                            type + " 图表需要 Excel 2016+（当前 " + ver.ToString("0.0", CultureInfo.InvariantCulture) +
                            "）。可先调用 check_chart_version 查看替代方案（如瀑布图可用垫底柱堆积模拟）。");
                    }

                    dynamic shape = null;
                    try
                    {
                        // AddChart2 返回 Shape；C# dynamic 下 shape.DrawingObject 解析不出 .Chart
                        // （真机 T26 教训：PS 可行但 C# IDispatch 报「未包含 Chart 的定义」）。
                        // 2026-09-12：用户 sheet（外来引擎生成）上【禁止直接建原生图】——绑定恒真失败，
                        // 改为全部在会话内新建的临时 sheet 上建图+绑定（见下方 W1~W8 判别实验结论）。
                        shape = null;

                        // 2026-09-12 根治（用户实证四类全坏 + 判别实验 PNG 视觉核验）：
                        // 旧 Series.Values/XValues 逐点赋值对新图表类型（117~123）不成立——
                        // Series.Formula 读不出（空）、Points 错乱（瀑布 6 点只读回 1）、渲染空白/异常
                        //（treemap/sunburst"引擎限制"系错误归因）。正确路径=SetSourceData(完整源区域
                        // 含表头, xlColumns)：COM 返回 E_FAIL 但【实际生效】（假失败——判别实验四类
                        // PNG 全部正确渲染：瀑布升降柱/排列双轴折线/树状面积块/旭日环），吞掉 E_FAIL
                        // 后以系列数结构兜底；多级层级（旭日/树状）整块交给 Excel 解析（末列=值、前列=层级）。
                        // 2026-09-12 最终修复（用户四类全坏实证 + W1~W8 判别实验）：决定因素不是绑定 API，
                        // 而是【图表所在 sheet 的出身】——非 Excel 引擎（openpyxl 等 AI 工具）生成的 xlsx 里，
                        // 其原有 sheet 上直接创建的原生新图表 SetSourceData 恒真失败（E_FAIL、pts 0~3、渲染
                        // 空白，激活/重命名/换行位均无效）；会话内新建 sheet 上同绑定 100% 生效。
                        // 方案：临时 sheet 建图+绑定（生效）→ Cut/Paste 回用户 sheet → 源固定指向隐藏临时表
                        //（数据快照，语义同看板 _calc 隐藏表）。粘贴后源不可再改绑（W8 实证），故临时表保留。
                        string hashTag = Math.Abs(StringComparer.Ordinal.GetHashCode(Convert.ToString(sheet.Name)))
                            .ToString("X", CultureInfo.InvariantCulture);
                        hashTag = hashTag.Length > 5 ? hashTag.Substring(0, 5) : hashTag.PadLeft(5, '0');
                        string tmpSheetName = ("_advsrc_" + type + "_" + hashTag);
                        if (tmpSheetName.Length > 31) { tmpSheetName = tmpSheetName.Substring(0, 31); }
                        // 幂等：清理同名旧临时表（重建场景）
                        ExcelTrace.Info("create_chart_advanced", "native W: enter type=" + type);
                        var staleNames = new List<string>();
                        // 注意：枚举元素与外部持有引用（sheet 等）共享同身份 RCW，绝不可 Release(sIter)
                        //（曾把 sheet 的 RCW 减断，后续 sheet.Activate() 即抛 InvalidComObjectException）
                        foreach (dynamic sIter in workbook.Worksheets)
                        {
                            string sn = Convert.ToString(sIter.Name);
                            if (sn == tmpSheetName) { staleNames.Add(sn); }
                        }
                        foreach (var sn in staleNames)
                        {
                            dynamic stale = workbook.Worksheets.Item[sn];
                            try { stale.Visible = -1; stale.Delete(); } catch { }
                            Release((object)stale);
                        }
                        ExcelTrace.Info("create_chart_advanced", "native W: stale cleared, adding tmp " + tmpSheetName);
                        dynamic tmpSheet = workbook.Worksheets.Add();
                        tmpSheet.Name = tmpSheetName;
                        ExcelTrace.Info("create_chart_advanced", "native W: tmp added+named");
                        try
                        {
                            // 源区域整块复制到临时表 A1 起（1 基 object[,] 直接块写；单行/单列退化逐格）
                            object[,] block = source.Value2;
                            string lastColL = ColumnLetterSafe(colCount);
                            if (rowCount == 1 || colCount == 1)
                            {
                                for (int r = 1; r <= rowCount; r++)
                                    for (int c = 1; c <= colCount; c++)
                                        tmpSheet.Cells.Item[r, c].Value2 = block[r, c];
                            }
                            else
                            {
                                tmpSheet.Range["A1:" + lastColL + rowCount].Value2 = block;
                            }
                            Release((object)block);
                            ExcelTrace.Info("create_chart_advanced", "native W: data copied");
                            // Range 用字符串地址形式（Cells.Item 两参形式实测触发 0x800A03EC）
                            dynamic rngTmp = tmpSheet.Range["A1:" + lastColL + rowCount];
                            ExcelTrace.Info("create_chart_advanced", "native W: rng ready " + "A1:" + lastColL + rowCount + ", addchart2 begin");
                            dynamic shapeT = tmpSheet.Shapes.AddChart2(-1, nativeChartType, 100, 100, width, height);
                            ExcelTrace.Info("create_chart_advanced", "native W: addchart2 done, setsrc begin");
                            ExcelTrace.Info("create_chart_advanced", "native W: shape name=" + Convert.ToString(shapeT.Name));
                            string shapeTName = Convert.ToString(shapeT.Name);
                            dynamic coTmp = tmpSheet.ChartObjects().Item(shapeTName);
                            dynamic chartT = coTmp.Chart;
                            ExcelTrace.Info("create_chart_advanced", "native W: coTmp+chart resolved");
                            try
                            {
                                try { chartT.SetSourceData(rngTmp, 2);
                                ExcelTrace.Info("create_chart_advanced", "native W: setsrc NO-EXC"); }
                                catch (Exception exSS) { ExcelTrace.Info("create_chart_advanced", "native W: setsrc EXC " + exSS.GetType().Name + " " + exSS.Message); }
                                try
                                {
                                    dynamic scCheck = chartT.SeriesCollection();
                                    int scCount = Convert.ToInt32(scCheck.Count);
                                    Release((object)scCheck);
                                    if (scCount < 1)
                                    {
                                        throw new InvalidOperationException("临时表上 SetSourceData 后无系列，原生图表绑定失败。");
                                    }
                                }
                                catch (System.Runtime.InteropServices.COMException) { /* 结构读回失败不在此判死，粘贴后 pts 校验兜底 */ }
                                ExcelTrace.Info("create_chart_advanced", "native W: sccheck finished");
                                // 移回用户 sheet：Cut → 用户表激活 → Paste → 重设位置
                                var beforeNames = new HashSet<string>();
                                ExcelTrace.Info("create_chart_advanced", "W-move: beforenames begin");
                                // 隐藏实例上对非活动 sheet 调 ChartObjects() 会抛 0x800A03EC（实测）——先激活
                                sheet.Activate();
                                dynamic coBefore = sheet.ChartObjects();
                                ExcelTrace.Info("create_chart_advanced", "W-move: coBefore got count=" + Convert.ToInt32(coBefore.Count));
                                for (int cb = 1; cb <= Convert.ToInt32(coBefore.Count); cb++)
                                { beforeNames.Add(Convert.ToString(coBefore.Item(cb).Name)); }
                                Release((object)coBefore);
                                ExcelTrace.Info("create_chart_advanced", "W-move: beforenames done");
                                ExcelTrace.Info("create_chart_advanced", "W-move: cut begin");
                                ExcelTrace.Info("create_chart_advanced", "native W: cut begin");
                                coTmp.Cut();
                                ExcelTrace.Info("create_chart_advanced", "native W: cut done");
                                ExcelTrace.Info("create_chart_advanced", "W-move: cut done, activate");
                                sheet.Activate();
                                ExcelTrace.Info("create_chart_advanced", "W-move: activated, paste begin");
                                // Paste 偶发 0x800A03EC（剪贴板竞态，同 CopyPicture 先例）：延迟重试 ×3
                                for (int pr = 1; ; pr++)
                                {
                                    try { sheet.Paste(); ExcelTrace.Info("create_chart_advanced", "W-move: paste ok try=" + pr); break; }
                                    catch (System.Runtime.InteropServices.COMException exP) when (pr < 3 && exP.HResult == unchecked((int)0x800A03EC))
                                    {
                                        ExcelTrace.Info("create_chart_advanced", "W-move: paste retry " + pr + " " + exP.Message);
                                        System.Threading.Thread.Sleep(400);
                                        sheet.Activate();
                                    }
                                }
                                dynamic coAfter = sheet.ChartObjects();
                                string pastedName = null;
                                for (int ca = 1; ca <= Convert.ToInt32(coAfter.Count); ca++)
                                {
                                    string nm = Convert.ToString(coAfter.Item(ca).Name);
                                    if (!beforeNames.Contains(nm)) { pastedName = nm; }
                                    Release((object)coAfter.Item(ca));
                                }
                                Release((object)coAfter);
                                if (pastedName == null)
                                {
                                    throw new InvalidOperationException("图表移回用户工作表失败（粘贴后未找到新图表对象）。");
                                }
                                chartObject = sheet.ChartObjects().Item(pastedName);
                                chart = chartObject.Chart;
                                chartObject.Left = left;
                                chartObject.Top = top;
                                ExcelTrace.Info("create_chart_advanced", "native W: pasted+positioned");
                                chartObject.Width = width;
                                chartObject.Height = height;
                            }
                            finally
                            {
                                Release((object)chartT);
                                Release((object)coTmp);
                                Release((object)shapeT);
                                Release((object)rngTmp);
                            }
                            tmpSheet.Visible = 0;  // xlSheetHidden：图表源指向本表，保留（语义同 _calc 隐藏表）
                        }
                        catch
                        {
                            // 半成品清理：临时表删除（含未隐藏失败的场景）
                            try { tmpSheet.Visible = -1; tmpSheet.Delete(); } catch { }
                            throw;
                        }

                        // pareto 排列图默认降序——若源数据未排好序，提示信息中说明（原生排序由 Excel 自动处理）
                        string advInfoNative = type + " 原生图表（Excel 2016+，SetSourceData 绑定），行数=" + nRows;
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = title;

                        var sheetNameN = Convert.ToString(sheet.Name);
                        var chartNameN = Convert.ToString(chartObject.Name);
                        recorder?.Invoke("create_chart_advanced", args, "chart-insert",
                            new { sheet = sheetNameN, chart_name = chartNameN, adv_type = type, rows = nRows, info = advInfoNative },
                            new ChartInsertPayload { Sheet = sheetNameN, ChartName = chartNameN });

                        var nativeResult = new ToolResult
                        {
                            Ok = true,
                            Data = new
                            {
                                chart_name = chartNameN,
                                adv_type = type,
                                rows = nRows,
                                info = advInfoNative,
                                source_address = Convert.ToString(source.Address),
                                anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                                title
                            }
                        };
                        return AttachAutoCheck(nativeResult, sheet, chartNameN, sourceAddress, type);
                    }
                    catch
                    {
                        // 半成品清理（Shape.Delete 与 ChartObject.Delete 等价，TryDeleteChartObject 兼容两种）
                        TryDeleteChartObject(chartObject);
                        chartObject = null;
                        chart = null;
                        throw;
                    }
                    finally
                    {
                        Release((object)shape);
                    }
                }

                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;
                string advInfo;
                ToolResult renderResult;
                switch (type)
                {
                    case "bullet":
                        renderResult = RenderBullet(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                    case "waffle":
                        renderResult = RenderWaffle(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                    case "slider":
                        renderResult = RenderSlider(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                    case "funnel":
                        renderResult = RenderFunnel(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                    case "butterfly":
                        renderResult = RenderButterfly(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                    default: // gantt
                        renderResult = RenderGantt(chart, sheet, labels, numericCols, auxCol, maxRows, cleanupRanges, out advInfo);
                        break;
                }

                if (!renderResult.Ok)
                {
                    // AC-5：渲染分支失败 → 清理已创建图表与已写辅助区
                    TryDeleteChartObject(chartObject);
                    chartObject = null;
                    chart = null;
                    CleanupAuxRanges(sheet, cleanupRanges);
                    return AttachAutoCheck(renderResult, sheet, null, sourceAddress, type);
                }
                advInfo = renderResult.Data?.ToString() ?? advInfo;

                chart.HasTitle = true;
                chart.ChartTitle.Text = title;

                try { chart.HasLegend = type == "bullet" || type == "slider" || type == "gantt"; } catch { }

                var sheetName = Convert.ToString(sheet.Name);
                var chartName = Convert.ToString(chartObject.Name);
                recorder?.Invoke("create_chart_advanced", args, "chart-insert",
                    new { sheet = sheetName, chart_name = chartName, adv_type = type, rows = nRows, info = advInfo },
                    new ChartInsertPayload { Sheet = sheetName, ChartName = chartName });

                // v3.7.1.44：技巧图分支同样并入 auto_check（此前仅原生图分支有自检回显，
                // 漏斗/子弹/华夫/滑珠/蝴蝶/甘特六类 AI 拿不到自检结果——口径统一为全部高级图必附自检）
                var trickResult = new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartName,
                        adv_type = type,
                        rows = nRows,
                        aux_col = auxCol,
                        info = advInfo,
                        source_address = Convert.ToString(source.Address),
                        anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                        title
                    }
                };
                return AttachAutoCheck(trickResult, sheet, chartName, sourceAddress, type);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                TryDeleteChartObject(chartObject);
                chartObject = null;
                chart = null;
                CleanupAuxRanges(sheet, cleanupRanges);
                return Invalid("ADV_CHART_FAILED",
                    "技巧型图表生成 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。已清理半成品。");
            }
            catch (Exception ex)
            {
                ExcelTrace.Error("chart_tool", "图表创建失败（已清理半成品）", ex);
                TryDeleteChartObject(chartObject);
                chartObject = null;
                chart = null;
                CleanupAuxRanges(sheet, cleanupRanges);
                return Invalid("ADV_CHART_FAILED", "技巧型图表生成失败：" + ex.Message + "。已清理半成品。");
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                // 2026-09-12：不再释放 source/sheet/workbook——调用方（repair_chart 等）共享同身份 RCW，
                // 减计数断开会让调用方后续操作炸 InvalidComObjectException（M-D1 纪律），GC 兜底即可。
            }
        }

        /// <summary>写辅助区：值块 + 公式块均可；写入范围登记到 cleanupRanges 供失败清理。</summary>
        private static void WriteAuxBlock(dynamic sheet, List<dynamic> cleanupRanges, int row, int col, object[,] values)
        {
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            dynamic a1 = sheet.Cells.Item[row, col];
            dynamic a2 = sheet.Cells.Item[row + rows - 1, col + cols - 1];
            dynamic rng = null;
            try
            {
                rng = sheet.Range[a1, a2];
                rng.Value2 = values;
                cleanupRanges.Add(rng);
                rng = null; // 所有权移交 cleanupRanges
            }
            finally
            {
                Release((object)a1);
                Release((object)a2);
                if (rng != null) { Release((object)rng); }
            }
        }

        /// <summary>失败路径清理辅助区（成功路径保留——图表引用辅助数据，删除会断链）。</summary>
        private static void CleanupAuxRanges(dynamic sheet, List<dynamic> cleanupRanges)
        {
            foreach (var rng in cleanupRanges)
            {
                try { rng.Clear(); } catch { }
                Release((object)rng);
            }
            cleanupRanges.Clear();
        }

        /// <summary>
        /// 子弹图：三档背景条（60%/80%/100% 目标分段）+ 实际值柱。
        /// 数据约定：B 列=实际值，C 列=目标值。
        /// </summary>
        private ToolResult RenderBullet(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var actual = cols.ContainsKey(2) ? cols[2] : new List<double>();
            var target = cols.ContainsKey(3) ? cols[3] : new List<double>();
            if (actual.Count == 0 || target.Count == 0)
            {
                return Invalid("MISSING_TARGET_COLUMN", "子弹图需两列数值：实际值列 + 目标值列（source_address 共 3 列：类别/实际/目标）。");
            }

            int rows = labels.Count;
            var block = new object[rows, 4];
            for (int i = 0; i < rows; i++)
            {
                double a = double.IsNaN(actual[i]) ? 0 : actual[i];
                double t = double.IsNaN(target[i]) ? 0 : target[i];
                if (t <= 0) { t = a > 0 ? a : 1; }
                double poor = Math.Round(t * 0.6, 4);
                double fair = Math.Round(t * 0.8, 4);
                block[i, 0] = a;
                block[i, 1] = Math.Min(a, poor);
                block[i, 2] = Math.Max(Math.Min(a, fair) - Math.Min(a, poor), 0);
                block[i, 3] = Math.Max(a - Math.Min(a, fair), 0);
            }
            var header = new object[1, 4] { { "实际", "差", "良", "优" } };
            dynamic h1 = sheet.Cells.Item[1, auxCol];
            dynamic h4 = sheet.Cells.Item[1, auxCol + 3];
            dynamic hRng = null;
            try { hRng = sheet.Range[h1, h4]; hRng.Value2 = header; } finally { Release((object)h1); Release((object)h4); if (hRng != null) Release((object)hRng); }
            WriteAuxBlock(sheet, cleanupRanges, 2, auxCol, block);

            chart.ChartType = 58; // xlBarStacked
            dynamic srcRange = null;
            dynamic srcA = sheet.Cells.Item[1, auxCol];
            dynamic srcB = sheet.Cells.Item[rows + 1, auxCol + 3];
            try { srcRange = sheet.Range[srcA, srcB]; chart.SetSourceData(srcRange, 2); }
            finally { Release((object)srcA); Release((object)srcB); Release((object)srcRange); }

            info = "三档背景带(60%/80%/100% 目标) + 实际值柱";
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>
        /// 华夫图：10×10 格子，完成率格子数 = round(rate*100)。散点两系列（已填充/未填充）。
        /// 数据约定：B 列=完成率（0-1 或 0-100 自动识别），取首个非空值。
        /// </summary>
        private ToolResult RenderWaffle(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var vals = cols.ContainsKey(2) ? cols[2] : new List<double>();
            double rate = vals.FirstOrDefault(v => !double.IsNaN(v));
            if (double.IsNaN(rate))
            {
                return Invalid("NO_NUMERIC_DATA", "华夫图需一个完成率数值（B 列首行，0-1 或 0-100）。");
            }
            if (rate > 1) { rate = rate / 100.0; }
            rate = Math.Max(0, Math.Min(1, rate));
            int filled = (int)Math.Round(rate * 100);

            var doneList = new List<object[]>();
            var todoList = new List<object[]>();
            for (int cell = 1; cell <= 100; cell++)
            {
                int row = (cell - 1) / 10 + 1;
                int col = (cell - 1) % 10 + 1;
                if (cell <= filled) { doneList.Add(new object[] { col, 11 - row }); }
                else { todoList.Add(new object[] { col, 11 - row }); }
            }
            WriteAuxBlock(sheet, cleanupRanges, 2, auxCol, To2D(doneList));
            if (todoList.Count > 0) { WriteAuxBlock(sheet, cleanupRanges, 2, auxCol + 3, To2D(todoList)); }

            chart.ChartType = -4169; // xlXYScatter
            dynamic seriesColl = chart.SeriesCollection();
            while (Convert.ToInt32(seriesColl.Count) > 0)
            {
                dynamic last = seriesColl.Item(Convert.ToInt32(seriesColl.Count));
                last.Delete();
                Release((object)last);
            }
            dynamic sDone = seriesColl.NewSeries();
            sDone.Name = "已完成 " + Math.Round(rate * 100) + "%";
            sDone.ChartType = -4169;
            sDone.MarkerStyle = 8; // 方点
            sDone.MarkerSize = 12;
            BindXY(sDone, sheet, 2, auxCol, doneList.Count);
            dynamic sTodo = seriesColl.NewSeries();
            sTodo.Name = "未完成";
            sTodo.ChartType = -4169;
            sTodo.MarkerStyle = 8;
            sTodo.MarkerSize = 12;
            if (todoList.Count > 0) { BindXY(sTodo, sheet, 2, auxCol + 3, todoList.Count); }
            Release((object)sDone);
            Release((object)sTodo);
            Release((object)seriesColl);

            info = "10×10 华夫格，完成率 " + Math.Round(rate * 100) + "%（填充 " + filled + " 格）";
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>
        /// 滑珠图：水平轨道 + 实际/目标两系列圆点。
        /// 数据约定：B 列=实际，C 列=目标。
        /// </summary>
        private ToolResult RenderSlider(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var actual = cols.ContainsKey(2) ? cols[2] : new List<double>();
            var target = cols.ContainsKey(3) ? cols[3] : new List<double>();
            if (actual.Count == 0 || target.Count == 0)
            {
                return Invalid("MISSING_TARGET_COLUMN", "滑珠图需两列数值：实际值列 + 目标值列。");
            }

            int rows = labels.Count;
            var track = new object[rows, 1];
            for (int i = 0; i < rows; i++)
            {
                double a = double.IsNaN(actual[i]) ? 0 : actual[i];
                double t = double.IsNaN(target[i]) ? 0 : target[i];
                track[i, 0] = Math.Round(Math.Max(a, t) * 1.1, 4);
            }
            WriteAuxBlock(sheet, cleanupRanges, 2, auxCol, track);

            chart.ChartType = 57; // xlBarClustered
            dynamic srcRange = null;
            dynamic srcA = sheet.Cells.Item[1, auxCol];
            dynamic srcB = sheet.Cells.Item[rows + 1, auxCol];
            try
            {
                srcRange = sheet.Range[srcA, srcB];
                chart.SetSourceData(srcRange, 2);
                dynamic seriesColl = chart.SeriesCollection();
                dynamic sTrack = seriesColl.Item(1);
                sTrack.Name = "轨道";
                Release((object)sTrack);
                dynamic sActual = seriesColl.NewSeries();
                sActual.Name = "实际";
                sActual.ChartType = 57;
                sActual.MarkerStyle = 8;
                sActual.MarkerSize = 10;
                dynamic vA = sheet.Range[sheet.Cells.Item[2, 2], sheet.Cells.Item[rows + 1, 2]];
                sActual.Values = vA;
                Release((object)vA);
                dynamic sTarget = seriesColl.NewSeries();
                sTarget.Name = "目标";
                sTarget.ChartType = 57;
                sTarget.MarkerStyle = 8;
                sTarget.MarkerSize = 10;
                dynamic vT = sheet.Range[sheet.Cells.Item[2, 3], sheet.Cells.Item[rows + 1, 3]];
                sTarget.Values = vT;
                Release((object)vT);
                Release((object)sActual);
                Release((object)sTarget);
                Release((object)seriesColl);
            }
            finally { Release((object)srcA); Release((object)srcB); Release((object)srcRange); }

            info = "水平轨道 + 实际/目标滑珠";
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>
        /// 漏斗图：条形按数值降序 + 各级占比数据标签（ xlBarClustered 倒序）。
        /// 数据约定：B 列=各级数值（如各阶段人数）。
        /// </summary>
        private ToolResult RenderFunnel(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var vals = cols.ContainsKey(2) ? cols[2] : new List<double>();
            double maxV = vals.Where(v => !double.IsNaN(v)).DefaultIfEmpty(0).Max();
            if (maxV <= 0)
            {
                return Invalid("NO_NUMERIC_DATA", "漏斗图需各级数值（B 列），且最大值须大于 0。");
            }

            int rows = labels.Count;
            // 辅助区：占位差（max - 值）写左侧，实现条形「右对齐漏斗」；占比写标签列
            var pad = new object[rows, 1];
            for (int i = 0; i < rows; i++)
            {
                double v = double.IsNaN(vals[i]) ? 0 : vals[i];
                pad[i, 0] = Math.Max(maxV - v, 0);
            }
            WriteAuxBlock(sheet, cleanupRanges, 2, auxCol, pad);

            chart.ChartType = 57; // xlBarClustered
            dynamic srcRange = null;
            dynamic srcA = sheet.Cells.Item[1, auxCol];
            dynamic srcB = sheet.Cells.Item[rows + 1, auxCol];
            try
            {
                srcRange = sheet.Range[srcA, srcB];
                chart.SetSourceData(srcRange, 2);
                dynamic seriesColl = chart.SeriesCollection();
                dynamic sPad = seriesColl.Item(1);
                sPad.Name = "占位";
                Release((object)sPad);
                dynamic sVal = seriesColl.NewSeries();
                sVal.Name = "数值";
                dynamic vB = sheet.Range[sheet.Cells.Item[2, 2], sheet.Cells.Item[rows + 1, 2]];
                sVal.Values = vB;
                Release((object)vB);
                Release((object)sVal);
                Release((object)seriesColl);
            }
            finally { Release((object)srcA); Release((object)srcB); Release((object)srcRange); }

            info = "倒序条形漏斗（占位列对齐），最大值=" + maxV;
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>
        /// 蝴蝶图：左右背靠背条形（B 列取负镜像=左侧，C 列=右侧）。
        /// 数据约定：B 列=左侧数值，C 列=右侧数值。
        /// </summary>
        private ToolResult RenderButterfly(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var left = cols.ContainsKey(2) ? cols[2] : new List<double>();
            var right = cols.ContainsKey(3) ? cols[3] : new List<double>();
            if (left.Count == 0 || right.Count == 0)
            {
                return Invalid("MISSING_TARGET_COLUMN", "蝴蝶图需两列数值：左翼列 + 右翼列。");
            }

            int rows = labels.Count;
            var mirror = new object[rows, 1];
            for (int i = 0; i < rows; i++)
            {
                double v = double.IsNaN(left[i]) ? 0 : left[i];
                mirror[i, 0] = -Math.Abs(v); // 左翼取负镜像
            }
            WriteAuxBlock(sheet, cleanupRanges, 2, auxCol, mirror);

            chart.ChartType = 57;
            dynamic srcRange = null;
            dynamic srcA = sheet.Cells.Item[1, auxCol];
            dynamic srcB = sheet.Cells.Item[rows + 1, auxCol];
            try
            {
                srcRange = sheet.Range[srcA, srcB];
                chart.SetSourceData(srcRange, 2);
                dynamic seriesColl = chart.SeriesCollection();
                dynamic sLeft = seriesColl.Item(1);
                sLeft.Name = "左翼";
                Release((object)sLeft);
                dynamic sRight = seriesColl.NewSeries();
                sRight.Name = "右翼";
                dynamic vR = sheet.Range[sheet.Cells.Item[2, 3], sheet.Cells.Item[rows + 1, 3]];
                sRight.Values = vR;
                Release((object)vR);
                Release((object)sRight);
                Release((object)seriesColl);
                // 数值轴刻度改绝对值显示（负轴标签取正），失败不致命
                try { chart.Axes(1).TickLabels.NumberFormat = "0;0"; } catch { }
            }
            finally { Release((object)srcA); Release((object)srcB); Release((object)srcRange); }

            info = "背靠背条形（左翼负值镜像 + 右翼正向）";
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>
        /// 甘特图：堆积条形 + 首段（开始天数）透明垫底 + 持续天数段。
        /// 数据约定：B 列=开始（天，相对数），C 列=持续（天）。
        /// </summary>
        private ToolResult RenderGantt(dynamic chart, dynamic sheet, List<string> labels,
            Dictionary<int, List<double>> cols, int auxCol, int maxRows,
            List<dynamic> cleanupRanges, out string info)
        {
            info = string.Empty;
            var start = cols.ContainsKey(2) ? cols[2] : new List<double>();
            var duration = cols.ContainsKey(3) ? cols[3] : new List<double>();
            if (start.Count == 0 || duration.Count == 0)
            {
                return Invalid("MISSING_TARGET_COLUMN", "甘特图需两列数值：开始天数 + 持续天数。");
            }

            int rows = labels.Count;
            chart.ChartType = 57; // xlBarClustered
            // 直接以源 B/C 两列做堆积条形（首段垫底后透明由系列 1 无填充实现）
            dynamic srcRange = null;
            dynamic srcA = sheet.Cells.Item[2, 2];
            dynamic srcB = sheet.Cells.Item[rows + 1, 3];
            try
            {
                srcRange = sheet.Range[srcA, srcB];
                chart.SetSourceData(srcRange, 1);
                dynamic seriesColl = chart.SeriesCollection();
                if (Convert.ToInt32(seriesColl.Count) >= 2)
                {
                    dynamic sStart = seriesColl.Item(1);
                    sStart.Name = "开始(垫底)";
                    try { sStart.Format.Fill.Visible = 0; } catch { } // xlNone 透明垫底
                    Release((object)sStart);
                    dynamic sDur = seriesColl.Item(2);
                    sDur.Name = "持续";
                    Release((object)sDur);
                }
                Release((object)seriesColl);
                try { chart.Axes(1).ReverseOrder = 1; } catch { } // 任务自上而下
            }
            finally { Release((object)srcA); Release((object)srcB); Release((object)srcRange); }

            info = "堆积条形甘特（垫底透明 + 持续段），任务数=" + rows;
            return new ToolResult { Ok = true, Data = info };
        }

        /// <summary>List&lt;object[]&gt; → object[,]（辅助区块写用）。</summary>
        private static object[,] To2D(List<object[]> rows)
        {
            var result = new object[rows.Count, rows.Count > 0 ? rows[0].Length : 1];
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < rows[r].Length; c++) { result[r, c] = rows[r][c]; }
            }
            return result;
        }

        /// <summary>为散点系列绑定 X/Y（辅助区列地址，1 基行起）。</summary>
        private static void BindXY(dynamic series, dynamic sheet, int startRow, int col, int count)
        {
            dynamic x = sheet.Range[sheet.Cells.Item[startRow, col], sheet.Cells.Item[startRow + count - 1, col]];
            dynamic y = sheet.Range[sheet.Cells.Item[startRow, col + 1], sheet.Cells.Item[startRow + count - 1, col + 1]];
            try
            {
                series.XValues = x;
                series.Values = y;
            }
            finally
            {
                Release((object)x);
                Release((object)y);
            }
        }

        /// <summary>需求②b：verify_chart 已由 ExcelChartService.VerifyChart 提供成熟实现
        /// （含像素级渲染核验），此处不重复实现；对新图表类型（waterfall/treemap/sunburst/pareto）
        /// 的支持通过扩展 ExcelChartService.ChartTypeMap 与 NormalizeChartFamily 完成。</summary>
        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
        }
    }
}
