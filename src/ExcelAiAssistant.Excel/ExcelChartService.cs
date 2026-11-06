using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第三期-可视化：迷你图插入、图表导出为图片。
    /// </summary>
    public sealed partial class ExcelChartService
    {
        private readonly dynamic application;
        private readonly OperationRecorder recorder;

        public ExcelChartService(object application)
            : this(application, null)
        {
        }

        public ExcelChartService(object application, OperationRecorder recorder)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.recorder = recorder;
        }

        public ToolResult CreateSparkline(Dictionary<string, object> args)
        {
            var dataRange = Required(args, "data_range");
            var destRange = Required(args, "dest_range");
            var typeName = GetString(args, "type", "line").ToLowerInvariant();
            if (typeName != "line" && typeName != "column" && typeName != "winloss")
            {
                return new ToolResult { Ok = false, ErrorCode = "INVALID_TYPE", Data = $"type 仅支持 line/column/winloss，当前为 {typeName}。" };
            }

            var color = GetString(args, "color", null);
            long? colorValue = null;
            if (!string.IsNullOrWhiteSpace(color))
            {
                try
                {
                    colorValue = ParseColor(color);
                }
                catch (Exception)
                {
                    return Invalid("INVALID_COLOR", $"迷你图颜色格式不合法：{color}，需为 #RRGGBB 或 #AARRGGBB。");
                }
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic data = null;
            dynamic dest = null;
            dynamic groups = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                data = sheet.Range[dataRange];
                dest = sheet.Range[destRange];
                var sparkType = typeName switch
                {
                    "column" => 2,
                    "winloss" => 3,
                    _ => 1
                };
                groups = dest.SparklineGroups.Add(sparkType, Convert.ToString(data.Address));
                if (colorValue.HasValue)
                {
                    try
                    {
                        groups.SeriesColor.Color = colorValue.Value;
                    }
                    catch (Exception xaiEx) {
                        AppLog.Warn("chart", "ExcelChartService.cs:CreateSparkline 静默兜底触发: " + xaiEx.Message);
                        // 颜色设置失败时忽略，保留默认色（不阻断迷你图插入）。
                    }
                }

                var sheetName = Convert.ToString(sheet.Name);
                var destAddress = Convert.ToString(dest.Address);
                recorder?.Invoke(
                    "create_excel_sparkline",
                    args,
                    "sparkline-insert",
                    new { sheet = sheetName, dest_range = destAddress },
                    new SparklineInsertPayload { Sheet = sheetName, DestRange = destAddress });

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!{destAddress}" },
                    Data = new
                    {
                        type = typeName,
                        data_range = Convert.ToString(data.Address),
                        dest_range = destAddress,
                        created = true
                    }
                };
            }
            finally
            {
                Release((object)groups);
                Release((object)dest);
                Release((object)data);
                // sheet/workbook 是 application 的共享 RCW（ActiveSheet/ActiveWorkbook），
                // FinalRelease 会连带销毁调用方持有的同身份 RCW——只置空交 GC。
                sheet = null;
                workbook = null;
            }
        }

        /// <summary>M28 审查（第三轮）：布尔参数严格解析——只接受 true/false/1/0。</summary>
        private static bool TryReadBoolArg(Dictionary<string, object> args, string key, out bool value, out string error)
        {
            value = false;
            error = null;
            if (!args.TryGetValue(key, out var raw) || raw == null)
            {
                return true;
            }

            if (raw is Newtonsoft.Json.Linq.JValue jv) { raw = jv.Value; }
            if (raw is bool boolean) { value = boolean; return true; }
            var text = Convert.ToString(raw)?.Trim();
            if (bool.TryParse(text, out value)) { return true; }
            if (text == "1") { value = true; return true; }
            if (text == "0") { value = false; return true; }
            error = "参数 " + key + " 需布尔值（true/false，或 1/0），当前：" + Convert.ToString(raw);
            return false;
        }

        public ToolResult ExportChartAsImage(Dictionary<string, object> args)
        {
            var filePath = ExcelAiAssistant.Core.Security.PathGuard.GuardAndComplete(
                Required(args, "file_path"), ".png");
            // M28 审查：导出不得静默覆盖既有图片文件（布尔参数严格解析，非法值报参数错误）
            if (!TryReadBoolArg(args, "overwrite", out var overwriteExport, out var overwriteError))
            {
                return Invalid("INVALID_ARGS", overwriteError);
            }

            if (System.IO.File.Exists(filePath) && !overwriteExport)
            {
                return Invalid("OVERWRITE_DENIED",
                    "目标文件已存在：" + filePath + "。为防止覆盖既有文件，请显式传 overwrite=true（或换一个 file_path）。");
            }

            // M28（第三轮续）：覆盖既有图片前先备份（可恢复），备份失败则中止
            string backupPath = null;
            if (System.IO.File.Exists(filePath))
            {
                try { backupPath = ExcelAiAssistant.Core.Security.PathGuard.CreateBackup(filePath); }
                catch (Exception backupException)
                {
                    return Invalid("BACKUP_FAILED", "覆盖前备份既有图片失败，已中止导出：" + backupException.Message);
                }
            }

            var width = GetInt(args, "width", 0);
            var height = GetInt(args, "height", 0);
            var chartName = Optional(args, "chart_name");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic chartObject = null;
            dynamic chart = null;
            var origWidth = -1.0;
            var origHeight = -1.0;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                chartObject = FindChartObject(sheet, chartName);
                if (chartObject == null)
                {
                    return Invalid("CHART_NOT_FOUND", chartName == null ? "当前工作表没有可导出的图表。" : $"未找到图表：{chartName}。");
                }

                // width/height 仅用于“导出时的像素尺寸”，导出后还原，避免持久改动用户图表 Object 的真实尺寸。
                if (width > 0)
                {
                    origWidth = Convert.ToDouble(chartObject.Width);
                    chartObject.Width = width;
                }

                if (height > 0)
                {
                    origHeight = Convert.ToDouble(chartObject.Height);
                    chartObject.Height = height;
                }

                chart = chartObject.Chart;
                chart.Export(filePath, "PNG");
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        file_path = filePath,
                        chart_name = Convert.ToString(chartObject.Name),
                        width = Convert.ToInt32(chartObject.Width),
                        height = Convert.ToInt32(chartObject.Height),
                        exported = true,
                        backup_path = backupPath
                    }
                };
            }
            finally
            {
                TryRestoreSize(chartObject, "Width", origWidth);
                TryRestoreSize(chartObject, "Height", origHeight);
                Release((object)chart);
                Release((object)chartObject);
                // sheet/workbook 是 application 的共享 RCW（ActiveSheet/ActiveWorkbook），
                // FinalRelease 会连带销毁调用方持有的同身份 RCW——只置空交 GC。
                sheet = null;
                workbook = null;
            }
        }

        private static void TryRestoreSize(dynamic chartObject, string property, double original)
        {
            if (original < 0 || chartObject == null)
            {
                return;
            }

            try
            {
                if (property == "Width")
                {
                    chartObject.Width = original;
                }
                else
                {
                    chartObject.Height = original;
                }
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:TryRestoreSize 静默兜底触发: " + xaiEx.Message);
            }
        }

        /// <summary>把「当前活动工作表最近一张图表」导出为 PNG 的 data URI（供 AI 回复里直接显示图片）。
        /// 无图表时返回 ok=true 且 chart_image=null；导出失败返回 ok=false。</summary>
        public ToolResult ExportLatestChartAsDataUri(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic chartObjects = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                chartObjects = sheet.ChartObjects();
                var count = Convert.ToInt32(chartObjects.Count);
                if (count <= 0)
                {
                    return new ToolResult { Ok = true, Data = new { chart_image = (string)null, note = "当前工作表没有图表，无法生成图片。" } };
                }

                // 按 chart_name 或取最后一张；不能凭 count 猜测（单图表且未命名时按名取空串会 1004）
                var chartName = Optional(args, "chart_name");
                chartObject = string.IsNullOrWhiteSpace(chartName) ? sheet.ChartObjects(count) : sheet.ChartObjects(chartName);

                chart = chartObject.Chart;
                var width = Convert.ToInt32(chartObject.Width);
                var height = Convert.ToInt32(chartObject.Height);

                var seriesCount = 0;
                var pointTotal = 0;
                var chartType = 0;
                var hasTitle = false;
                var titleText = (string)null;
                try
                {
                    chartType = Convert.ToInt32(chart.ChartType);
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("chart", "ExcelChartService.cs:ExportLatestChartAsDataUri 静默兜底触发: " + xaiEx.Message);
                }

                try
                {
                    dynamic title = chart.ChartTitle;
                    if (title != null)
                    {
                        hasTitle = true;
                        var text = Convert.ToString(title.Text);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            titleText = text;
                        }
                    }
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("chart", "ExcelChartService.cs:ExportLatestChartAsDataUri 静默兜底触发: " + xaiEx.Message);
                    hasTitle = false;
                }

                try
                {
                    dynamic sc = chart.SeriesCollection();
                    seriesCount = Convert.ToInt32(sc.Count);
                    for (var i = 1; i <= seriesCount; i++)
                    {
                        dynamic s = sc.Item(i);
                        pointTotal += Convert.ToInt32(s.Points.Count);
                    }
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("chart", "ExcelChartService.cs:method? 静默兜底触发: " + xaiEx.Message);
                }

                // Excel 导出 PNG：写临时文件 → 读字节 → base64 data URI → 清理。
                var tempPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "xai_chart_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
                try
                {
                    chart.Export(tempPath, "PNG");
                    if (!System.IO.File.Exists(tempPath))
                    {
                        return new ToolResult { Ok = false, ErrorCode = "CHART_EXPORT_FAILED", Data = "图表导出失败（未生成图片文件）。" };
                    }

                    var bytes = System.IO.File.ReadAllBytes(tempPath);
                    var dataUri = "data:image/png;base64," + Convert.ToBase64String(bytes);
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            chart_image = dataUri,
                            chart_name = Convert.ToString(chartObject.Name),
                            chart_type = chartType,
                            has_data = seriesCount > 0 && pointTotal > 0,
                            series_count = seriesCount,
                            point_count = pointTotal,
                            has_title = hasTitle,
                            title = titleText,
                            width,
                            height,
                            note = "图表图片已生成，供核对标题 / 数据是否与需求一致。"
                        }
                    };
                }
                finally
                {
                    try
                    {
                        if (System.IO.File.Exists(tempPath))
                        {
                            System.IO.File.Delete(tempPath);
                        }
                    }
                    catch (Exception xaiEx) {
                        AppLog.Warn("chart", "ExcelChartService.cs:method? 静默兜底触发: " + xaiEx.Message);
                    }
                }
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)chartObjects);
                // sheet/workbook 是 application 的共享 RCW（ActiveSheet/ActiveWorkbook），
                // FinalRelease 会连带销毁调用方持有的同身份 RCW——只置空交 GC。
                sheet = null;
                workbook = null;
            }
        }

        private dynamic FindChartObject(dynamic sheet, string chartName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(chartName))
                {
                    return sheet.ChartObjects(chartName);
                }

                return sheet.ChartObjects(1);
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:FindChartObject 静默兜底触发: " + xaiEx.Message);
                return null;
            }
        }

        /// <summary>图表自检（AI 生成后的校验闭环）。读图表元数据 + 数据源系列公式 + 导出图片，给出可审计的自检结论。
        /// 判据：① 图表是否真正生成；② 是否有有效数据（系列/点数）；③ 数据源区域（由 series.Formula 解析出来）；
        /// ④ 与预期类型/系列数是否一致；⑤ 渲染图片是否可导出（供 AI “看图”核对）。</summary>
        public ToolResult VerifyChart(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic chartObjects = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                var sheetName = Optional(args, "sheet");
                var chartName = Optional(args, "chart_name");
                var expectedSeries = TryGetOptionalInt(args, "expected_series");
                var expectedChartType = Optional(args, "expected_chart_type");
                var expectedTitle = Optional(args, "expected_title");
                var expectedDataSource = Optional(args, "expected_data_source");
                var expectedSecondaryAxis = args.TryGetValue("expected_secondary_axis", out var esaRaw)
                    && esaRaw != null
                    && int.TryParse(Convert.ToString(esaRaw), out var esaNum)
                    ? esaNum
                    : (int?)null;

                // 防“未解析占位符泄漏”：流程占位 {step2.chart_type} 若因前一步输出缺失而未替换成真实值，
                // 仍带花括号的字面量会被当成“断言值”产生假告警。凡仍含 { } 的一律视为未提供，跳过对应校验。
                if (ContainsBraces(expectedTitle)) expectedTitle = null;
                if (ContainsBraces(expectedDataSource)) expectedDataSource = null;
                if (ContainsBraces(expectedChartType)) expectedChartType = null;

                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[sheetName];
                chartObjects = sheet.ChartObjects();
                var count = Convert.ToInt32(chartObjects.Count);
                if (count <= 0)
                {
                    var r0 = ComputeChartVerdict(false, 0, 0, 0, false,
                        expectedSeries, NormalizeChartFamily(expectedChartType), null);
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            chart_found = false,
                            verdict = r0.verdict,
                            issues = r0.issues,
                            selfcheck = r0.selfcheck,
                            note = "当前工作表没有图表，说明图表并未真正生成。",
                            chart_image = (string)null
                        }
                    };
                }

                chartObject = string.IsNullOrWhiteSpace(chartName) ? sheet.ChartObjects(count) : sheet.ChartObjects(chartName);
                chart = chartObject.Chart;

                // dynamic 形参传染：TryGetChartType/CollectSeriesMeta 接 dynamic，调用点返回值退化为 dynamic，
                // 命名元组元素（SeriesCount/TextSource…）不是运行时成员，须用 Item1..Item6 还原为静态类型，
                // 否则下游 ChartTypeFamily/ComputeChartVerdict 会被冲成动态分派、元组无法析构。
                var chartType = (int)TryGetChartType(chart);
                var hasTitle = false;
                var title = (string)null;
                TryGetTitle(chart, ref hasTitle, ref title);
                var meta = CollectSeriesMeta(chart, workbook, sheet);
                var sCount = meta.SeriesCount;
                var pTotal = meta.PointTotal;
                var emptyS = meta.EmptySeries;
                var textSource = meta.TextSource;
                var seriesInfo = meta.Series;
                var dataRanges = meta.DataRanges;
                var secondaryAxisCount = meta.SecondaryAxisCount;
                var secondaryAxisKnown = meta.SecondaryAxisKnown;

                var actualFamily = ChartTypeFamily(chartType);
                // 显式快照为静态局部（切断 dynamic 实参传染导致的返回元组动态化）
                int staticSCount = sCount;
                int staticPTotal = pTotal;
                int staticEmptyS = emptyS;
                bool staticTextSource = textSource;
                int staticExpectedSeries = expectedSeries ?? 0;
                bool hasExpectedSeries = expectedSeries.HasValue;
                string staticFamily = actualFamily;
                string staticNorm = NormalizeChartFamily(expectedChartType);
                var verdictTuple = ComputeChartVerdict(true,
                    staticSCount, staticPTotal, staticEmptyS, staticTextSource,
                    hasExpectedSeries ? (int?)staticExpectedSeries : null, staticNorm, staticFamily);
                var verdict = (string)verdictTuple.Item1;
                var issues = (List<string>)verdictTuple.Item2;
                var selfcheck = (string)verdictTuple.Item3;

                // 标题正确性（元数据核验）：给出 expected_title 时比对实际标题，断言是否写对了。
                var titleOk = (bool?)null;
                if (!string.IsNullOrWhiteSpace(expectedTitle))
                {
                    titleOk = title != null && TitleMatches(title, expectedTitle);
                }

                // 数据源正确性（series.Formula 核验）：给出 expected_data_source 时，只要任一系列
                // 的数据源区域与预期区域相交即判定数据源命中。
                var dataSourceOk = (bool?)null;
                if (!string.IsNullOrWhiteSpace(expectedDataSource))
                {
                    string ds = expectedDataSource;
                    List<string> ranges = dataRanges;
                    dataSourceOk = ranges.Exists(r => RangesOverlap(r, ds));
                }

                // 导出图片供 AI “看到”渲染效果 + 像素级核验「数据是否真的画出来了」；失败不阻断自检，仅置空。
                string chartImage = null;
                var dataRendered = (bool?)null;
                var nonBlankRatio = (double?)null;
                try
                {
                    if (TryExportChartPngBytes(chartObject, out byte[] pngBytes))
                    {
                        chartImage = "data:image/png;base64," + Convert.ToBase64String(pngBytes);
                        var stats = ComputeRenderStats(pngBytes);
                        dataRendered = stats.dataRendered;
                        nonBlankRatio = stats.nonBlankRatio;
                    }
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("chart", "ExcelChartService.cs:method? 静默兜底触发: " + xaiEx.Message);
                    chartImage = null;
                }

                // 把标题 / 数据源 / 渲染核验的失败项并入 verdict、issues、selfcheck，与环境已有判据叠加。
                var extraIssues = new List<string>();
                if (titleOk == false)
                {
                    AppendVerdictFlag(ref verdict, "title");
                    extraIssues.Add($"图表标题与预期不符：预期「{expectedTitle}」，实际{(string.IsNullOrWhiteSpace(title) ? "无标题" : "「" + title + "」")}。");
                }

                if (dataSourceOk == false)
                {
                    AppendVerdictFlag(ref verdict, "source");
                    extraIssues.Add($"图表数据源与预期不符：预期选用「{expectedDataSource}」，实际系列数据源为 {string.Join("、", dataRanges.Count == 0 ? new List<string> { "(未能解析)" } : dataRanges)}。");
                }

                // M-D1：次轴断言——expected_secondary_axis 给出时，实际次轴系列数必须一致。
                // 仅当元数据可读（secondaryAxisKnown）时判定；读不出视作环境不支持，不误报。
                var secondaryAxisOk = (bool?)null;
                if (expectedSecondaryAxis.HasValue && secondaryAxisKnown)
                {
                    secondaryAxisOk = secondaryAxisCount == expectedSecondaryAxis.Value;
                    if (secondaryAxisOk == false)
                    {
                        AppendVerdictFlag(ref verdict, "secondary_axis");
                        extraIssues.Add($"次轴系列数与预期不符：预期 {expectedSecondaryAxis.Value} 个系列挂次轴，实际 {secondaryAxisCount} 个。");
                    }
                }

                if (dataRendered == false && pTotal > 0)
                {
                    // 已有可绘制的数据点、但渲染主体几乎空白：数据可能未被真正绘制（空图已由“empty”判据覆盖）。
                    AppendVerdictFlag(ref verdict, "blank");
                    extraIssues.Add($"渲染核验发现图表内容几乎为空白（非白像素占比 {nonBlankRatio:0.0%}）：虽存在 {pTotal} 个数据点，但数据可能因值区被当文本、颜色过淡或图表尺寸过小而未真实显示。");
                }

                if (extraIssues.Count > 0)
                {
                    issues.AddRange(extraIssues);
                    selfcheck += "；" + string.Join("；", extraIssues);
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_found = true,
                        chart_name = Convert.ToString(chartObject.Name),
                        chart_type = chartType,
                        chart_type_name = ChartTypeChineseName(chartType),
                        series_count = sCount,
                        empty_series_count = emptyS,
                        point_total = pTotal,
                        has_text_source = textSource,
                        series = seriesInfo,
                        data_ranges = dataRanges,
                        has_title = hasTitle,
                        title,
                        title_ok = titleOk,
                        expected_title = string.IsNullOrWhiteSpace(expectedTitle) ? null : expectedTitle,
                        data_source_ok = dataSourceOk,
                        expected_data_source = string.IsNullOrWhiteSpace(expectedDataSource) ? null : expectedDataSource,
                        secondary_axis_count = secondaryAxisKnown ? (int?)secondaryAxisCount : null,
                        secondary_axis_ok = secondaryAxisOk,
                        expected_secondary_axis = expectedSecondaryAxis,
                        data_rendered = dataRendered,
                        nonblank_ratio = nonBlankRatio,
                        expected_chart_type = string.IsNullOrWhiteSpace(expectedChartType) ? null : expectedChartType,
                        expected_series = expectedSeries,
                        verdict,
                        issues,
                        selfcheck,
                        chart_image = chartImage,
                        note = "图表自检结果：请结合 issues / selfcheck 判断是否达标；实际数据源区域见 data_ranges（来自 series 公式），title_ok / data_source_ok / data_rendered 分别核验标题、数据源与渲染是否有内容。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("VERIFY_CHART_FAILED", "图表自检失败：" + ex.Message);
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)chartObjects);
                // 2026-09-12：不再释放 sheet/workbook——它们是调用方共享的同身份 RCW，
                // 减计数断开后调用方后续操作即炸 InvalidComObjectException（M-D1 纪律），
                // GC 兜底即可。
            }
        }

        private int TryGetChartType(dynamic chart)
        {
            try
            {
                return Convert.ToInt32(chart.ChartType);
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:TryGetChartType 静默兜底触发: " + xaiEx.Message);
                return 0;
            }
        }

        private void TryGetTitle(dynamic chart, ref bool hasTitle, ref string title)
        {
            try
            {
                dynamic t = chart.ChartTitle;
                if (t != null)
                {
                    hasTitle = true;
                    var text = Convert.ToString(t.Text);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        title = text;
                    }
                }
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:TryGetTitle 静默兜底触发: " + xaiEx.Message);
                hasTitle = false;
            }
        }

        /// <summary>系列元数据结果（M-D1：脱离 ValueTuple，规避 dynamic 运行时绑定坑）。</summary>
        private sealed class SeriesMeta
        {
            public int SeriesCount;
            public int PointTotal;
            public int EmptySeries;
            public bool TextSource;
            public List<object> Series = new List<object>();
            public List<string> DataRanges = new List<string>();
            public int SecondaryAxisCount;
            public bool SecondaryAxisKnown;
        }

        private SeriesMeta CollectSeriesMeta(object chartObj, object workbookObj, object sheetObj)
        {
            // 显式 object 入参 + 强类型结果类：彻底切断 dynamic 传染。
            dynamic chart = chartObj;
            dynamic workbook = workbookObj;
            dynamic sheet = sheetObj;
            var meta = new SeriesMeta();
            meta.SeriesCount = 0;
            meta.PointTotal = 0;
            meta.EmptySeries = 0;
            meta.TextSource = false;
            meta.SecondaryAxisCount = 0;
            meta.SecondaryAxisKnown = false;
            var series = new List<object>();
            var dataRanges = new List<string>();
            try
            {
                dynamic sc = chart.SeriesCollection();
                meta.SeriesCount = Convert.ToInt32(sc.Count);
                for (var i = 1; i <= meta.SeriesCount; i++)
                {
                    dynamic s = sc.Item(i);
                    var points = 0;
                    try
                    {
                        points = Convert.ToInt32(s.Points.Count);
                    }
                    catch (Exception xaiEx) {
                        AppLog.Warn("chart", "ExcelChartService.cs:CollectSeriesMeta 静默兜底触发: " + xaiEx.Message);
                    }

                    meta.PointTotal += points;
                    if (points <= 0)
                    {
                        meta.EmptySeries++;
                    }

                    var formula = (string)null;
                    try
                    {
                        formula = Convert.ToString(s.Formula);
                    }
                    catch (Exception xaiEx) {
                        AppLog.Warn("chart", "ExcelChartService.cs:CollectSeriesMeta 静默兜底触发: " + xaiEx.Message);
                    }

                    var ranges = ParseSeriesFormulaRanges(formula);
                    foreach (var r in ranges)
                    {
                        if (!meta.DataRanges.Contains(r))
                        {
                            meta.DataRanges.Add(r);
                        }
                    }

                    // 值区抽样判数值性：scatter/bar 等若值区是文本，Points.Count 仍可能>0（看似 ok 实为“画不出的图”），
                    // 必须显式告警。读失败或无数值区则跳过，不影响 verdict。
                    var dataKind = (string)null;
                    var lastRange = ExtractLastRange(formula);
                    if (points > 0 && lastRange != null)
                    {
                        dataKind = ClassifySourceSamples(ReadSampleRange(workbook, sheet, lastRange, 6));
                        if (dataKind == "text" || dataKind == "mixed")
                        {
                            meta.TextSource = true;
                        }
                    }

                    var name = (string)null;
                    try
                    {
                        var n = Convert.ToString(s.Name);
                        if (!string.IsNullOrWhiteSpace(n))
                        {
                            name = n;
                        }
                    }
                    catch (Exception xaiEx) {
                        AppLog.Warn("chart", "ExcelChartService.cs:method? 静默兜底触发: " + xaiEx.Message);
                    }

                    meta.Series.Add(new
                    {
                        index = i,
                        name,
                        points,
                        formula,
                        data_ranges = ranges,
                        data_kind = dataKind,
                        has_data = points > 0,
                        on_secondary_axis = TryGetSecondaryAxis(s, out bool sec) ? sec : (bool?)null
                    });
                    if (TryGetSecondaryAxis(s, out bool isSecondary) && isSecondary)
                    {
                        meta.SecondaryAxisCount++;
                    }
                    meta.SecondaryAxisKnown = true;
                }
            }
            catch (Exception metaEx)
            {
                Console.WriteLine("[CollectSeriesMeta] 异常: " + metaEx.GetType().Name + " " + metaEx.Message + " @ 行" + metaEx.StackTrace);
                meta.SeriesCount = 0;
                meta.PointTotal = 0;
                meta.EmptySeries = 0;
                meta.TextSource = false;
                meta.SecondaryAxisCount = 0;
                meta.SecondaryAxisKnown = false;
            }

            Console.WriteLine("[CollectSeriesMeta] 完成: s=" + meta.SeriesCount + " sec=" + meta.SecondaryAxisCount + " known=" + meta.SecondaryAxisKnown);
            return meta;
        }

        /// <summary>读取系列的 AxisGroup（1=primary / 2=secondary）；读取失败返回 false 且 out 值为 false。</summary>
        private static bool TryGetSecondaryAxis(dynamic series, out bool onSecondary)
        {
            onSecondary = false;
            try
            {
                var group = Convert.ToInt32(series.AxisGroup);
                onSecondary = group == 2; // xlSecondary
                return true;
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:TryGetSecondaryAxis 静默兜底触发: " + xaiEx.Message);
                return false;
            }
        }

        private static bool TryExportChartPngBytes(dynamic chartObject, out byte[] bytes)
        {
            bytes = null;
            var tempPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "xai_verify_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png");
            try
            {
                dynamic chart = chartObject.Chart;
                chart.Export(tempPath, "PNG");
                if (!System.IO.File.Exists(tempPath))
                {
                    return false;
                }

                bytes = System.IO.File.ReadAllBytes(tempPath);
                return bytes != null && bytes.Length > 0;
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:TryExportChartPngBytes 静默兜底触发: " + xaiEx.Message);
                bytes = null;
                return false;
            }
            finally
            {
                try
                {
                    if (System.IO.File.Exists(tempPath))
                    {
                        System.IO.File.Delete(tempPath);
                    }
                }
                catch (Exception xaiEx) {
                    AppLog.Warn("chart", "ExcelChartService.cs:TryExportChartPngBytes 静默兜底触发: " + xaiEx.Message);
                }
            }
        }

        private static int? TryGetOptionalInt(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                if (int.TryParse(Convert.ToString(raw), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                {
                    return v;
                }
            }

            return null;
        }

        // ---------- 纯静态判据（不依赖 COM，可单元测试） ----------

        /// <summary>字符串是否仍含花括号（未解析的 {name} 占位符），用于让自检工具跳过“断言值缺失”的假告警。</summary>
        public static bool ContainsBraces(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return value.IndexOf('{') >= 0 || value.IndexOf('}') >= 0;
        }

        /// <summary>给 verdict 追加后缀标记（verdict=="ok" 时直接取该标记名，否则追加 +name，不重复）。</summary>
        public static void AppendVerdictFlag(ref string verdict, string name)
        {
            if (string.Equals(verdict, "ok", StringComparison.Ordinal))
            {
                verdict = name;
            }
            else if (verdict.IndexOf("+" + name, StringComparison.Ordinal) < 0)
            {
                verdict += "+" + name;
            }
        }

        /// <summary>规范化后比对实际标题与预期标题（忽略大小写、首尾空白与成对引号）。未给预期返回 true（未断言）。</summary>
        public static bool TitleMatches(string actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            var a = actual.Trim().Trim('"').Trim();
            var e = expected.Trim().Trim('"').Trim();
            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>解析单元格区域（支持 A1、$A$1:$B$8、'表'!C2:D9）为矩形坐标；解析失败返回 null。</summary>
        public static (int C1, int C2, int R1, int R2)? ParseRangeRect(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return null;
            }

            var t = token.Trim();
            var bang = t.LastIndexOf('!');
            if (bang >= 0)
            {
                t = t.Substring(bang + 1);
            }

            t = t.Replace("$", string.Empty);
            var parts = t.Split(':');
            var from = parts[0];
            var to = parts.Length > 1 ? parts[1] : parts[0];
            if (!TryCell(from, out var fc, out var fr) || !TryCell(to, out var tc, out var tr))
            {
                return null;
            }

            return (Math.Min(fc, tc), Math.Max(fc, tc), Math.Min(fr, tr), Math.Max(fr, tr));
        }

        private static bool TryCell(string cell, out int col, out int row)
        {
            col = 0;
            row = 0;
            if (string.IsNullOrWhiteSpace(cell))
            {
                return false;
            }

            var c = 0;
            var i = 0;
            for (; i < cell.Length && char.IsLetter(cell[i]); i++)
            {
                c = c * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1);
            }

            if (c <= 0 || i >= cell.Length || !int.TryParse(cell.Substring(i), NumberStyles.Integer, CultureInfo.InvariantCulture, out var r) || r <= 0)
            {
                return false;
            }

            col = c;
            row = r;
            return true;
        }

        /// <summary>两个单元格区域是否相交（共享任一单元格即算命中）。解析失败视为不相交。</summary>
        public static bool RangesOverlap(string a, string b)
        {
            var ra = ParseRangeRect(a);
            var rb = ParseRangeRect(b);
            if (ra == null || rb == null)
            {
                return false;
            }

            return ra.Value.C1 <= rb.Value.C2 && rb.Value.C1 <= ra.Value.C2
                && ra.Value.R1 <= rb.Value.R2 && rb.Value.R1 <= ra.Value.R2;
        }

        /// <summary>
        /// 像素级核验渲染图是否真的画出了内容（数据/系列是否可见）：
        /// 用 GDI+/System.Drawing 逐像素抽样统计「非白内容像素」占比。背景按近白（三通道都 ≥240）判定，
        /// 其余（深色轴线文字 / 彩色柱条折线）都计入内容。全白或几乎空白的图该占比极低 → 判定数据未显示。
        /// 抽样步长随图片总像素自适应，兼顾速度与稳定性。
        /// </summary>
        public static (bool dataRendered, double nonBlankRatio, int sampledPixels, int contentPixels) ComputeRenderStats(byte[] pngBytes)
        {
            const double threshold = 0.01;
            if (pngBytes == null || pngBytes.Length < 8)
            {
                return (false, 0.0, 0, 0);
            }

            try
            {
                using (var ms = new MemoryStream(pngBytes))
                using (var bmp = new Bitmap(ms))
                {
                    var w = bmp.Width;
                    var h = bmp.Height;
                    if (w <= 0 || h <= 0)
                    {
                        return (false, 0.0, 0, 0);
                    }

                    // 目标约 90k 个采样点，长边尽量覆盖，避免大图逐像素的额外开销。
                    var step = Math.Max(1, (int)Math.Sqrt((double)w * h / 90000.0));
                    var content = 0;
                    var total = 0;
                    for (var y = 0; y < h; y += step)
                    {
                        for (var x = 0; x < w; x += step)
                        {
                            total++;
                            var c = bmp.GetPixel(x, y);
                            if (!(c.R >= 240 && c.G >= 240 && c.B >= 240))
                            {
                                content++;
                            }
                        }
                    }

                    var ratio = total > 0 ? (double)content / total : 0.0;
                    return (ratio >= threshold, ratio, total, content);
                }
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:static 静默兜底触发: " + xaiEx.Message);
                return (false, 0.0, 0, 0);
            }
        }

        /// <summary>把各种书写形式（推荐名/中文名/家族名）统一到图表族，用于 cross-check。未知返回 null。</summary>
        public static string NormalizeChartFamily(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var n = name.Trim();
            if (n.Equals("Column", StringComparison.OrdinalIgnoreCase) || n.Equals("ColumnClustered", StringComparison.OrdinalIgnoreCase)
                || n.Contains("柱形") || n.Contains("Column"))
            {
                return "Column";
            }

            if (n.Equals("Bar", StringComparison.OrdinalIgnoreCase) || n.Contains("条形"))
            {
                return "Bar";
            }

            if (n.Equals("Line", StringComparison.OrdinalIgnoreCase) || n.Contains("折线"))
            {
                return "Line";
            }

            if (n.Equals("Pie", StringComparison.OrdinalIgnoreCase) || n.Contains("饼"))
            {
                return "Pie";
            }

            if (n.Equals("Scatter", StringComparison.OrdinalIgnoreCase) || n.Equals("XYScatter", StringComparison.OrdinalIgnoreCase)
                || n.Contains("散点"))
            {
                return "Scatter";
            }

            if (n.Equals("Area", StringComparison.OrdinalIgnoreCase) || n.Contains("面积"))
            {
                return "Area";
            }

            // ---- 2016+ 原生新图表族（需求②a/②b）----
            if (n.Equals("Waterfall", StringComparison.OrdinalIgnoreCase) || n.Contains("瀑布"))
            {
                return "Waterfall";
            }
            if (n.Equals("Treemap", StringComparison.OrdinalIgnoreCase) || n.Contains("树状"))
            {
                return "Treemap";
            }
            if (n.Equals("Sunburst", StringComparison.OrdinalIgnoreCase) || n.Contains("旭日"))
            {
                return "Sunburst";
            }
            if (n.Equals("Pareto", StringComparison.OrdinalIgnoreCase) || n.Contains("排列"))
            {
                return "Pareto";
            }
            if (n.Equals("Histogram", StringComparison.OrdinalIgnoreCase) || n.Contains("直方"))
            {
                return "Histogram";
            }

            return null;
        }

        public static string ChartTypeFamily(int xlType)
        {
            foreach (var t in ChartTypeMap)
            {
                if (t.XlType == xlType)
                {
                    return t.Family;
                }
            }

            return null;
        }

        public static string ChartTypeChineseName(int xlType)
        {
            foreach (var t in ChartTypeMap)
            {
                if (t.XlType == xlType)
                {
                    return t.Name;
                }
            }

            return null;
        }

        /// <summary>从 series.Formula（A1 风格，形如 =SERIES("名",Sheet1!$A$2:$A$7,Sheet1!$B$2:$B$7,1)）
        /// 提取全部单元格/区域引用，即该系列实际使用的数据源。R1C1 风格无法可靠解析则返回空列表。</summary>
        public static List<string> ParseSeriesFormulaRanges(string formula)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(formula))
            {
                return result;
            }

            // R1C1 / 公式含右方括号（如 R1C1:R1C5、R[1]C[-1]）→ 放弃解析。
            if (formula.IndexOf('[') >= 0
                || System.Text.RegularExpressions.Regex.IsMatch(formula, @"(?<![A-Za-z])R\d+C\d+"))
            {
                return result;
            }

            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
                formula, @"(?<=!|\b)(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                var r = m.Value;
                if (!result.Contains(r))
                {
                    result.Add(r);
                }
            }

            return result;
        }

        /// <summary>取数据源中的“值区”——SERIES 公式里最后一个区域引用（意为 plot order 前的值列）：
        /// 柱/折/散点图为第 3 参，饼图第 2 参，名称若是单元格引用时其本身也排在值区前，故“最后那个区域”可稳定代理值区。</summary>
        public static string ExtractLastRange(string formula)
        {
            var rs = ParseSeriesFormulaRanges(formula);
            return rs.Count > 0 ? rs[rs.Count - 1] : null;
        }

        /// <summary>抽样判数据类型（决定图表是否真能显示数值）。返回 numeric | text | mixed | empty。</summary>
        public static string ClassifySourceSamples(object[] samples)
        {
            var nonEmpty = 0;
            var numeric = 0;
            var text = 0;
            foreach (var s in samples ?? new object[0])
            {
                if (s == null)
                {
                    continue;
                }

                var txt = Convert.ToString(s);
                if (string.IsNullOrWhiteSpace(txt))
                {
                    continue;
                }

                nonEmpty++;
                if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    numeric++;
                }
                else
                {
                    text++;
                }
            }

            if (nonEmpty == 0)
            {
                return "empty";
            }

            if (text == 0)
            {
                return "numeric";
            }

            return numeric == 0 ? "text" : "mixed";
        }

        /// <summary>读 areaToken（形如 Sheet1!$B$2:$B$7 或 $B$2:$B$7）前 max 个单元格值；跨表自动切目标表，失败返回空。</summary>
        private object[] ReadSampleRange(dynamic workbook, dynamic sheet, string rangeToken, int max)
        {
            try
            {
                object v;
                var bang = rangeToken.IndexOf('!');
                if (bang >= 0)
                {
                    var sheetName = rangeToken.Substring(0, bang).Trim('\'').Trim();
                    var addr = rangeToken.Substring(bang + 1);
                    v = string.Equals(sheetName, Convert.ToString(sheet.Name), StringComparison.OrdinalIgnoreCase)
                        ? sheet.Range[addr].Value2
                        : workbook.Worksheets.Item[sheetName].Range[addr].Value2;
                }
                else
                {
                    v = sheet.Range[rangeToken].Value2;
                }

                if (v is Array arr)
                {
                    var list = new List<object>();
                    foreach (object item in arr)
                    {
                        list.Add(item);
                        if (list.Count >= max)
                        {
                            break;
                        }
                    }

                    return list.ToArray();
                }

                return new[] { v };
            }
            catch (Exception xaiEx) {
                AppLog.Warn("chart", "ExcelChartService.cs:ReadSampleRange 静默兜底触发: " + xaiEx.Message);
                return new object[0];
            }
        }

        /// <summary>生成自检结论。verdict：no_chart | empty | partial | ok（可选叠加 +mismatch/+text）。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeChartVerdict(
            bool chartFound,
            int seriesCount,
            int pointTotal,
            int emptySeries,
            bool textSource,
            int? expectedSeries,
            string expectedCharFamily,
            string actualChartFamily)
        {
            var issues = new List<string>();
            string verdict;
            if (!chartFound)
            {
                verdict = "no_chart";
                issues.Add("当前工作表没有任何图表，说明图表并未真正生成，请重新生成并确认数据源范围。");
            }
            else if (seriesCount <= 0)
            {
                verdict = "empty";
                issues.Add("图表没有任何数据系列（series_count=0），属于空图表：数据源未关联、写入失败或图表被清空。");
            }
            else if (pointTotal <= 0)
            {
                verdict = "empty";
                issues.Add($"图表含 {seriesCount} 个系列但总点数为 0：系列未关联到有效数值范围，多为数据源区域选错、区域为空或数值被当作文本。");
            }
            else if (emptySeries > 0)
            {
                verdict = "partial";
                issues.Add($"有 {emptySeries} 个系列没有任何数据点（共 {seriesCount} 个系列）：部分系列数据源缺失或为空，建议核对对应 series.data_ranges。");
            }
            else
            {
                verdict = "ok";
            }

            if (textSource)
            {
                if (verdict == "ok")
                {
                    verdict = "text";
                }
                else if (verdict.IndexOf("+text", StringComparison.Ordinal) < 0)
                {
                    verdict += "+text";
                }

                issues.Add("检测到值区包含文本或混合类型数据：图表可能无法正确显示数值趋势（如散点图X/Y轴文本无法绘图、折线图文本值被当0处理），建议确保数据源为纯数值。");
            }

            if (!string.IsNullOrWhiteSpace(expectedCharFamily) && !string.IsNullOrWhiteSpace(actualChartFamily)
                && !string.Equals(expectedCharFamily, actualChartFamily, StringComparison.OrdinalIgnoreCase))
            {
                if (verdict == "ok")
                {
                    verdict = "mismatch";
                }
                else if (verdict.IndexOf("+mismatch", StringComparison.Ordinal) < 0)
                {
                    verdict += "+mismatch";
                }

                issues.Add($"图表类型与预期不符：预期「{expectedCharFamily}」，实际是「{actualChartFamily}」（图表类型名：{ChartTypeChineseName(ChartFamilyToType(actualChartFamily))}）。请核对图表类型是否真按需求创建。");
            }

            if (expectedSeries.HasValue && seriesCount > 0 && expectedSeries.Value != seriesCount)
            {
                var m = $"系列数量与预期不符：预期 {expectedSeries.Value} 个，实际 {seriesCount} 个。";
                if (!issues.Contains(m))
                {
                    issues.Add(m);
                }
            }

            var sb = new System.Text.StringBuilder();
            if (!chartFound)
            {
                sb.Append("自检未通过：没有生成任何图表。");
            }
            else
            {
                sb.Append("自检")
                  .Append(verdict == "ok" ? "通过" : "未通过")
                  .Append("：图表已生成，类型「")
                  .Append(actualChartFamily ?? "未知")
                  .Append("」，共 ")
                  .Append(seriesCount)
                  .Append(" 个系列 / ")
                  .Append(pointTotal)
                  .Append(" 个数据点");
                if (emptySeries > 0)
                {
                    sb.Append("（其中 ").Append(emptySeries).Append(" 个系列为空）");
                }
                if (textSource)
                {
                    sb.Append("（值区含文本数据，可能无法正确绘图）");
                }
            }

            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        private static int ChartFamilyToType(string family)
        {
            return family switch
            {
                "Column" => 51,
                "Bar" => 57,
                "Line" => 4,
                "Pie" => 5,
                "Scatter" => -4169,
                "Area" => 1,
                _ => 0
            };
        }

        private static readonly (int XlType, string Family, string Name)[] ChartTypeMap =
        {
            (51, "Column", "柱形图"), (52, "Column", "堆积柱形图"), (53, "Column", "百分比堆积柱形图"), (54, "Column", "三维柱形图"),
            (57, "Bar", "条形图"), (58, "Bar", "堆积条形图"), (59, "Bar", "百分比堆积条形图"),
            (4, "Line", "折线图"), (63, "Line", "堆积折线图"), (64, "Line", "百分比堆积折线图"), (65, "Line", "数据点折线图"),
            (-4169, "Scatter", "散点图"), (72, "Scatter", "平滑线散点图"), (73, "Scatter", "无标记平滑线散点图"),
            (74, "Scatter", "折线散点图"), (75, "Scatter", "无标记折线散点图"),
            (5, "Pie", "饼图"), (68, "Pie", "复合饼图"), (69, "Pie", "分离型饼图"),
            (1, "Area", "面积图"), (2, "Area", "堆积面积图"), (3, "Area", "百分比堆积面积图"),
            // ---- P1 扩展（M-D1）：环形族 / 雷达族 ----
            (-4111, "Combo", "组合图"), (-4120, "Doughnut", "环形图"), (-4121, "Doughnut", "分离型环形图"),
            (-4151, "Radar", "雷达图"), (-4152, "Radar", "填充雷达图"), (-4153, "Radar", "数据点雷达图"),
            // ---- 2016+ 原生新图表（需求②a，真机类型库权威值）----
            (119, "Waterfall", "瀑布图"), (117, "Treemap", "树状图"), (120, "Sunburst", "旭日图"), (122, "Pareto", "排列图"),
            (118, "Histogram", "直方图"), (121, "BoxWhisker", "箱形图"), (123, "Funnel", "漏斗图")
        };

        private static long ParseColor(string colorHex)
        {
            var hex = colorHex.TrimStart('#');
            if (hex.Length == 6)
            {
                hex = "FF" + hex;
            }

            if (hex.Length != 8)
            {
                throw new FormatException("颜色必须为 #RRGGBB 或 #AARRGGBB。");
            }

            var value = uint.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var alpha = (value >> 24) & 0xFF;
            var red = (value >> 16) & 0xFF;
            var green = (value >> 8) & 0xFF;
            var blue = value & 0xFF;
            _ = alpha;
            // Excel COM 的 Color 使用 0x00BBGGRR 布局。
            return (blue << 16) | (green << 8) | red;
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
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

        private static string GetString(Dictionary<string, object> args, string key, string defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToString(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

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