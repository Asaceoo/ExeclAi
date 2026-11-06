using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 图表读回诊断与修复（v3.7.1.37 批次，用户诉求：AI 生成图表后能读取截图/诊断/修复/审查）：
    /// - diagnose_chart：无期望全面体检——结构读回（类型/系列/点数/空系列/源区域）+ 渲染像素核验
    ///   + 数据守恒对账 + 根因归类 + 修复建议 + ExcelTrace 记录（诊断有依据/有结论/有记录）。
    /// - repair_chart：按诊断结论修复——空系列/点数错乱 → SetSourceData 重绑（经典类型）；
    ///   117~123 原生新图 → 临时表重建（复用 CreateChartAdvanced 的临时表路径）；标题缺失补标题；
    ///   修复后自动复检并返回前后对比 + chart_image（AI 读图闭环）。
    /// </summary>
    public sealed partial class ExcelChartService
    {
        /// <summary>2016+ 原生新图表类型枚举（与 ExcelPivotChartService 的 nativeTypes 对应）。</summary>
        private static readonly HashSet<int> NativeNewChartTypes = new HashSet<int> { 117, 118, 119, 120, 121, 122, 123 };

        public ToolResult DiagnoseChart(Dictionary<string, object> args)
        {
            // 复用 verify_chart 的读回与渲染核验（同一实例内部调用）
            var verify = VerifyChart(args);
            try
            {
                JObject j = JObject.FromObject(verify.Data ?? new { });
                bool chartFound = j.Value<bool?>("chart_found") ?? false;
                string verdict = j.Value<string>("verdict") ?? "unknown";
                int seriesCount = j.Value<int?>("series_count") ?? 0;
                int pointTotal = j.Value<int?>("point_total") ?? 0;
                int emptySeries = j.Value<int?>("empty_series_count") ?? 0;
                bool textSource = j.Value<bool?>("has_text_source") ?? false;
                bool? dataRendered = j.Value<bool?>("data_rendered");
                double? nonBlank = j.Value<double?>("nonblank_ratio");
                var findings = new JArray();
                foreach (var issue in j.Value<JArray>("issues") ?? new JArray())
                { findings.Add(JObject.FromObject(new { basis = (string)issue, severity = "warn" })); }

                // 根因归类（基于结构证据的组合判定）
                string rootCause;
                var suggestions = new List<string>();
                if (!chartFound)
                {
                    rootCause = "NO_CHART";
                    findings.Add(JObject.FromObject(new { basis = "当前工作表没有任何图表对象", severity = "fail" }));
                    suggestions.Add("重新创建图表：create_excel_chart / create_chart_advanced / create_combo_chart。");
                }
                else if (pointTotal == 0)
                {
                    rootCause = "NO_POINTS";
                    findings.Add(JObject.FromObject(new { basis = "系列存在但数据点总数为 0（数据未绑定或源区域为空）", severity = "fail" }));
                    suggestions.Add("调用 repair_chart（提供 source_address）重新绑定源数据。");
                }
                else if (emptySeries > 0)
                {
                    rootCause = "EMPTY_SERIES";
                    findings.Add(JObject.FromObject(new { basis = emptySeries + " 个系列数据点为空", severity = "fail" }));
                    suggestions.Add("调用 repair_chart（提供 source_address）重绑源数据；若源列含文本型数字先 clean_excel_data。");
                }
                else if (textSource)
                {
                    rootCause = "TEXT_SOURCE";
                    findings.Add(JObject.FromObject(new { basis = "数值区被 Excel 当作文本（源列含文本型数字或格式异常）", severity = "warn" }));
                    suggestions.Add("调用 clean_excel_data 将文本型数字转为数值，再重建图表。");
                }
                else if (dataRendered == false)
                {
                    rootCause = "BLANK_RENDER";
                    findings.Add(JObject.FromObject(new { basis = "渲染核验：非白像素占比 " + (nonBlank.HasValue ? (nonBlank.Value * 100).ToString("0.0", CultureInfo.InvariantCulture) + "%" : "未知") + "，图表主体几乎空白", severity = "fail" }));
                    suggestions.Add("调用 repair_chart 重建/重绑；检查图表尺寸是否过小。");
                }
                else if (verdict != "ok")
                {
                    rootCause = "EXPECT_CHECK";
                    suggestions.Add("按 issues 逐条修复后复检 verify_chart。");
                }
                else
                {
                    rootCause = "OK";
                    suggestions.Add("图表体检通过，无需修复。");
                }

                ExcelTrace.Info("diagnose_chart", "root_cause=" + rootCause + " verdict=" + verdict
                    + " series=" + seriesCount + " pts=" + pointTotal + " empty=" + emptySeries);

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        tool = "diagnose_chart",
                        chart_found = chartFound,
                        chart_type_name = j.Value<string>("chart_type_name"),
                        chart_name = j.Value<string>("chart_name"),
                        series_count = seriesCount,
                        empty_series_count = emptySeries,
                        point_total = pointTotal,
                        has_text_source = textSource,
                        data_ranges = j.Value<JArray>("data_ranges"),
                        data_rendered = dataRendered,
                        nonblank_ratio = nonBlank,
                        has_title = j.Value<bool?>("has_title"),
                        verdict,
                        root_cause = rootCause,
                        findings,   // 诊断依据（逐条：basis=证据, severity=等级）
                        suggestions,// 修复建议（逐条可执行动作，含工具名）
                        recorded = true,
                        chart_image = j.Value<string>("chart_image"),
                        note = "诊断结论=root_cause；findings 为依据（basis/severity）；suggestions 为建议动作；诊断已记录（ExcelTrace）。"
                    }
                };
            }
            catch (Exception ex)
            {
                return Invalid("DIAGNOSE_FAILED", "图表诊断失败：" + ex.Message);
            }
        }

        /// <summary>
        /// repair_chart：按诊断结论修复图表。动作矩阵：
        /// ① type 提供且为 117~123 原生新图（或现图 chart_type∈117~123）→ 重建式修复：删旧图 →
        ///    复用 CreateChartAdvanced 的临时表路径重建（源 address 必填）；② 经典类型 → SetSourceData
        ///    重绑（source_address 缺省时从系列 Formula 解析的原区域重取）；③ title 提供且缺失 → 补标题。
        /// 修复后自动复检（verify_chart）并返回前后对比 + chart_image（AI 读图闭环）。
        /// </summary>
        public ToolResult RepairChart(Dictionary<string, object> args, ExcelPivotChartService pivotChart)
        {
            var sheetName = Optional(args, "sheet");
            var chartName = Optional(args, "chart_name");
            var sourceAddress = Optional(args, "source_address");
            var type = Optional(args, "type");
            var title = Optional(args, "title");

            var before = VerifyChart(args);
            var actions = new List<string>();
            try
            {
                dynamic workbook = application.ActiveWorkbook
                    ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                dynamic sheet = string.IsNullOrWhiteSpace(sheetName)
                    ? workbook.Worksheets.Item[1]
                    : workbook.Worksheets.Item[sheetName];
                dynamic chartObjects = sheet.ChartObjects();
                int count = Convert.ToInt32(chartObjects.Count);
                if (count <= 0)
                {
                    return Invalid("CHART_NOT_FOUND", "当前工作表没有图表对象，无法修复。请先创建图表。");
                }
                dynamic chartObject = string.IsNullOrWhiteSpace(chartName)
                    ? chartObjects.Item(count)
                    : chartObjects.Item(chartName);
                dynamic chart = chartObject.Chart;
                int chartType = 0;
                try { chartType = Convert.ToInt32(chart.ChartType); } catch { }
                string nameBefore = Convert.ToString(chartObject.Name);
                bool isNative = NativeNewChartTypes.Contains(chartType) || (type != null && int.TryParse(type, out _));

                // ① 原生新图：重建式修复（临时表路径在 CreateChartAdvanced 内）
                if ((isNative || type != null && !int.TryParse(type, out _)) && !string.IsNullOrWhiteSpace(sourceAddress) && pivotChart != null)
                {
                    string rebuildType = type;
                    if (string.IsNullOrWhiteSpace(rebuildType))
                    {
                        // chart_type 数字 → 类型名（nativeTypes 的反向映射）
                        rebuildType = chartType switch
                        {
                            117 => "treemap", 118 => "histogram", 119 => "waterfall",
                            120 => "sunburst", 121 => "boxwhisker", 122 => "pareto", _ => null
                        };
                    }
                    if (rebuildType != null)
                    {
                        // 重建分支前重取引用（前置 VerifyChart 的 finally 可能已分离旧 RCW，双保险）
                        workbook = application.ActiveWorkbook;
                        sheet = workbook.Worksheets.Item[string.IsNullOrWhiteSpace(sheetName)
                            ? Convert.ToString(workbook.Worksheets.Item[1].Name)
                            : sheetName];
                        try { chartObject.Delete(); } catch { }
                        var rebuildArgs = new Dictionary<string, object>
                        {
                            ["type"] = rebuildType,
                            ["source_address"] = sourceAddress,
                            ["sheet"] = Convert.ToString(sheet.Name)
                        };
                        var rr = pivotChart.CreateChartAdvanced(rebuildArgs);
                        if (!rr.Ok)
                        {
                            return Invalid("REPAIR_FAILED", "重建式修复失败：" + Convert.ToString(rr.Data));
                        }
                        actions.Add("重建原生图表（type=" + rebuildType + "，源=" + sourceAddress + "）");
                        // 重建后复检 + 截图
                        var afterRebuild = VerifyChart(args);
                        ExcelTrace.Info("repair_chart", "rebuilt type=" + rebuildType + " chart=" + nameBefore);
                        return BuildRepairResult(actions, before, afterRebuild, Convert.ToString(sheet.Name));
                    }
                    return Invalid("REPAIR_NEEDS_TYPE", "原生新图表修复需要提供 type 参数（treemap/histogram/waterfall/sunburst/boxwhisker/pareto）。");
                }

                // ② 经典类型：SetSourceData 重绑
                if (!string.IsNullOrWhiteSpace(sourceAddress))
                {
                    dynamic rng = sheet.Range[sourceAddress];
                    chart.SetSourceData(rng, 2);
                    Release((object)rng);
                    actions.Add("SetSourceData 重绑源数据（" + sourceAddress + "）");
                }
                else
                {
                    // 从系列 Formula 解析的原区域重绑（Resolve 数据源自 ParseSourceFromFormula 同源逻辑）
                    try
                    {
                        dynamic ser = chart.SeriesCollection().Item(1);
                        string formula = Convert.ToString(ser.Formula);
                        Release((object)ser);
                        var m = System.Text.RegularExpressions.Regex.Match(formula, @"([A-Za-z0-9_\u4e00-\u9fa5（）() 0-9]+)!(\$[A-Z]+\$?\d+:\$[A-Z]+\$?\d+)");
                        if (m.Success)
                        {
                            dynamic rng2 = sheet.Range[m.Groups[2].Value.Replace("$", "")];
                            chart.SetSourceData(rng2, 2);
                            Release((object)rng2);
                            actions.Add("从系列公式解析区域重绑（" + m.Groups[2].Value + "）");
                        }
                    }
                    catch { }
                }

                // ③ 标题补全
                if (!string.IsNullOrWhiteSpace(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                    actions.Add("补全标题（" + title + "）");
                }

                if (actions.Count == 0)
                {
                    actions.Add("未发现可自动修复项——请先调用 diagnose_chart 查看根因与建议。");
                }

                var afterVerify = VerifyChart(args);
                ExcelTrace.Info("repair_chart", string.Join("; ", actions) + " | before_verdict="
                    + (JObject.FromObject(before.Data ?? new { })).Value<string>("verdict"));
                return BuildRepairResult(actions, before, afterVerify, sheetName);
            }
            catch (Exception ex)
            {
                ExcelTrace.Error("repair_chart", "RepairChart", ex);
                return Invalid("REPAIR_FAILED", "图表修复失败：" + ex.Message);
            }
        }

        private ToolResult BuildRepairResult(List<string> actions, ToolResult before, ToolResult after, string sheetName)
        {
            JObject jAfter = JObject.FromObject(after.Data ?? new { });
            string verdict = jAfter.Value<string>("verdict") ?? "unknown";
            var repairResult = new ToolResult
            {
                Ok = true,
                Data = new
                {
                    tool = "repair_chart",
                    repaired_actions = actions,
                    before_verdict = (JObject.FromObject(before.Data ?? new { })).Value<string>("verdict"),
                    after = jAfter,
                    chart_image = jAfter.Value<string>("chart_image"),
                    sheet = sheetName,
                    note = "修复动作与前后对比已记录（ExcelTrace）；chart_image 为修复后截图（AI 可直接读图核验）。"
                }
            };
            return repairResult;
        }
    }
}
