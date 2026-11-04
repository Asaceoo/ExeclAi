using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelAiAssistant.Excel
{
    public sealed class ReportSection
    {
        public string Title { get; set; }
        public List<string> Lines { get; set; } = new List<string>();
    }

    public sealed class ReportResult
    {
        public List<ReportSection> Sections { get; set; } = new List<ReportSection>();
        public object[,] ReportGrid { get; set; }
        public string Markdown { get; set; }
        public string Html { get; set; }
        public int WordCount { get; set; }

        /// <summary>P0-2（Q6）：生成后自检结果（数字对账/空段/占位符/覆盖匹配），对齐图表 auto_check 口径。</summary>
        public ReportSelfCheck SelfCheck { get; set; }
    }

    /// <summary>P0-2 报告自检结果。</summary>
    public sealed class ReportSelfCheck
    {
        public string Verdict { get; set; } = "ok";
        public List<string> Issues { get; set; } = new List<string>();
        public int CheckedCount { get; set; }
    }

    /// <summary>P0-1 基线配对。</summary>
    public sealed class BaselinePair
    {
        public int ActualColumn { get; set; }
        public int BaselineColumn { get; set; }
        public string Kind { get; set; }   // target / period
        public string Basis { get; set; }  // 配对依据（列名语义）
    }

    /// <summary>
    /// C2 模板化文字报告（纯计算/纯文本组装，不依赖 COM，可直接单测）。
    /// 从数据矩阵推断 概览/趋势/异常/重点列摘要/相关性/建议 六个章节，并生成工作表网格、Markdown、HTML。
    /// </summary>
    public static class ReportTextBuilder
    {
        private const int MaxColumns = 50;
        private const int MaxAnomalySamples = 5;
        private const int MaxTrendColumns = 5;
        private const int MaxKeyColumns = 3;
        // 相关性默认阈值（可通过 Build 的 correlationThreshold 参数覆盖，工具面见 analyze_column_correlation / generate_analysis_report）
        private const double CorrelationThreshold = 0.7;
        private static readonly string[] CorrelationNote =
        {
            "相关性强并不代表因果关系，请结合业务口径进一步核实。"
        };

        public static ReportResult Build(object[,] matrix, string sheet, string address, int anomalyStd, int maxNumericColumns, double correlationThreshold = 0.7, IEnumerable<string> sectionFilter = null)
        {
            var result = new ReportResult();
            // 章节白名单：null=全部章节；显式传入时未列出的章节不产出（三种输出形态一致生效）。
            var allow = sectionFilter == null
                ? null
                : new HashSet<string>(sectionFilter.Where(s => !string.IsNullOrWhiteSpace(s)));
            var correlationCutoff = Math.Max(0.05, Math.Min(correlationThreshold, 0.95));
            var cap = Math.Max(1, maxNumericColumns);
            var columnsCount = Math.Min(matrix.GetLength(1), Math.Min(MaxColumns, cap));
            var dataRows = Math.Max(matrix.GetLength(0) - 1, 0);
            var k = Math.Max(1, anomalyStd);

            // ---- 逐列归类 ----
            var names = new string[columnsCount];
            var numericColumns = new List<int>();
            var numsByColumn = new Dictionary<int, List<double>>();
            var rowsByColumn = new Dictionary<int, List<int>>();
            var nonBlankByColumn = new int[columnsCount];
            var dataCellCount = 0;
            for (var c = 0; c < columnsCount; c++)
            {
                names[c] = Header(matrix, c);
                var numbers = new List<double>();
                var rowIndexes = new List<int>();
                var nonBlank = 0;
                for (var r = 1; r < matrix.GetLength(0); r++)
                {
                    var cell = matrix[r, c];
                    if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                    {
                        continue;
                    }

                    nonBlank++;
                    if (double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        numbers.Add(number);
                        rowIndexes.Add(r);
                    }
                }

                nonBlankByColumn[c] = nonBlank;
                dataCellCount += nonBlank;
                if (numbers.Count > 0 && numbers.Count >= dataRows * 0.6)
                {
                    numericColumns.Add(c);
                    numsByColumn[c] = numbers;
                    rowsByColumn[c] = rowIndexes;
                }
            }

            var numericColCount = numericColumns.Count;
            var textColumnCount = columnsCount - numericColCount - CountBlankColumns(nonBlankByColumn, dataRows);
            var missingCells = columnsCount * dataRows - dataCellCount;
            var missingRate = columnsCount * dataRows == 0
                ? 0.0
                : (double)missingCells / (columnsCount * dataRows);

            // ---- 概览 ----
            Add(result, allow, "概览", new[]
            {
                "范围：" + (string.IsNullOrWhiteSpace(sheet) ? "" : sheet + "!") + address,
                "维度：" + dataRows + " 行 × " + columnsCount + " 列",
                "数值列 " + numericColCount + " 个，文本列 " + textColumnCount + " 个",
                "缺失单元格 " + missingCells + " 个（占数据区 " + Pct(missingRate) + "）"
            });

            // ---- 趋势 ----
            var trends = new List<KeyValuePair<double, string>>();
            foreach (var c in numericColumns)
            {
                var values = numsByColumn[c];
                if (values.Count < 3)
                {
                    continue;
                }

                var direction = Trend(values, out var magnitude);
                if (direction != "波动")
                {
                    var first = values[0];
                    var last = values[values.Count - 1];
                    var line = names[c] + "：" + direction + "（首 " + R4(first) +
                               " → 末 " + R4(last) + "，变动 " + Pct(magnitude) + "）";
                    trends.Add(new KeyValuePair<double, string>(Math.Abs(magnitude), line));
                }
            }

            if (trends.Count == 0)
            {
                Add(result, allow, "趋势", "未识别到明显的单调上升或下降趋势；如需趋势可转为日期序列后再分析。");
            }
            else
            {
                Add(result, allow, "趋势",
                    trends.OrderByDescending(kv => kv.Key).Take(MaxTrendColumns).Select(kv => kv.Value).ToList());
            }

            // ---- 异常 ----
            var anomalyLines = new List<string>();
            var anomalyCountTotal = 0;
            foreach (var c in numericColumns)
            {
                var values = numsByColumn[c];
                if (values.Count < 3)
                {
                    continue;
                }

                var mean = values.Average();
                var variance = values.Select(x => (x - mean) * (x - mean)).Sum() / Math.Max(values.Count - 1, 1);
                var stdev = Math.Sqrt(variance);
                if (stdev <= 0)
                {
                    continue;
                }

                var low = mean - k * stdev;
                var high = mean + k * stdev;
                var outliers = new List<string>();
                var count = 0;
                for (var i = 0; i < values.Count; i++)
                {
                    if (values[i] < low || values[i] > high)
                    {
                        count++;
                        if (outliers.Count < MaxAnomalySamples)
                        {
                            var label = Label(matrix, rowsByColumn[c][i]);
                            outliers.Add(label + " = " + R4(values[i]));
                        }
                    }
                }

                anomalyCountTotal += count;
                var sample = outliers.Count > 0 ? "，如 " + string.Join("、", outliers) : "";
                anomalyLines.Add(names[c] + "：" + count + " 个离群（±" + k + "σ）" + sample);
            }

            if (anomalyLines.Count == 0)
            {
                Add(result, allow, "异常", "未发现超过 mean±" + k + "σ 的明显离群点。");
            }
            else
            {
                Add(result, allow, "异常", anomalyLines);
            }

            // ---- 重点列摘要 ----
            var keyLines = new List<string>();
            var byRange = numericColumns
                .Select(c => new { c, range = numsByColumn[c].Max() - numsByColumn[c].Min() })
                .Where(x => !double.IsNaN(x.range))
                .OrderByDescending(x => x.range)
                .Take(MaxKeyColumns)
                .ToList();
            foreach (var item in byRange)
            {
                var values = numsByColumn[item.c];
                keyLines.Add(names[item.c] + "：范围 " + R4(values.Min()) + "~" + R4(values.Max()) +
                             "，均值 " + R4(values.Average()) + "，标准差 " + R4(Math.Sqrt(
                                 values.Select(x => (x - values.Average()) * (x - values.Average())).Sum() / Math.Max(values.Count - 1, 1))));
            }

            if (keyLines.Count > 0)
            {
                Add(result, allow, "重点列摘要", keyLines);
            }

            // ---- 相关性 ----
            var corrLines = new List<KeyValuePair<double, string>>();
            for (var i = 0; i < numericColumns.Count; i++)
            {
                for (var j = i + 1; j < numericColumns.Count; j++)
                {
                    var ci = numericColumns[i];
                    var cj = numericColumns[j];
                    var r = CorrelationAligned(rowsByColumn[ci], numsByColumn[ci], rowsByColumn[cj], numsByColumn[cj], 1, dataRows + 1);
                    if (double.IsNaN(r) || Math.Abs(r) < correlationCutoff)
                    {
                        continue;
                    }

                    corrLines.Add(new KeyValuePair<double, string>(Math.Abs(r),
                        names[ci] + " 与 " + names[cj] + "：r=" + r.ToString("0.00", CultureInfo.InvariantCulture) +
                        "（" + (r >= 0 ? "正" : "负") + "相关）"));
                }
            }

            if (corrLines.Count == 0)
            {
                Add(result, allow, "相关性", "数值列间未发现 |r|≥" + Pct(correlationCutoff) + " 的明显相关。");
            }
            else
            {
                var lines = corrLines.OrderByDescending(kv => kv.Key).Take(MaxKeyColumns).Select(kv => kv.Value).ToList();
                lines.AddRange(CorrelationNote);
                Add(result, allow, "相关性", lines);
            }

            Add(result, allow, "建议", Suggest());

            // ---- P0-1（Q4 拍板）：基线探测 + 对比洞察 ----
            var baselines = DetectBaselines(names, columnsCount, numericColumns);
            if (baselines.Count > 0)
            {
                var baselineLines = new List<string>();
                foreach (var bl in baselines)
                {
                    var actual = numsByColumn[bl.ActualColumn];
                    var baseline = numsByColumn[bl.BaselineColumn];
                    var n = Math.Min(actual.Count, baseline.Count);
                    if (bl.Kind == "target")
                    {
                        // 目标/预算：逐行求达成率（实际合计 / 目标合计），0 目标防除零
                        var actualSum = actual.Take(n).Sum();
                        var targetSum = baseline.Take(n).Sum();
                        if (Math.Abs(targetSum) < 1e-12)
                        {
                            baselineLines.Add(names[bl.ActualColumn] + " vs " + names[bl.BaselineColumn] +
                                "：目标合计为 0，无法计算达成率。");
                        }
                        else
                        {
                            var rate = actualSum / targetSum;
                            baselineLines.Add(names[bl.ActualColumn] + " vs " + names[bl.BaselineColumn] +
                                "：达成率 " + Pct(rate) + "（实际合计 " + R4(actualSum) + " / 目标合计 " + R4(targetSum) + "，配对依据：" + bl.Basis + "）");
                        }
                    }
                    else
                    {
                        // 同期/上期：按行同环比（本期 vs 同期，行对齐平均变动）
                        var changes = new List<double>();
                        for (var i = 0; i < n; i++)
                        {
                            if (Math.Abs(baseline[i]) > 1e-12)
                            {
                                changes.Add((actual[i] - baseline[i]) / Math.Abs(baseline[i]));
                            }
                        }

                        if (changes.Count == 0)
                        {
                            baselineLines.Add(names[bl.ActualColumn] + " vs " + names[bl.BaselineColumn] +
                                "：同期值全为 0，无法计算同环比。");
                        }
                        else
                        {
                            var avg = changes.Average();
                            var up = changes.Count(x => x > 0);
                            var down = changes.Count(x => x < 0);
                            baselineLines.Add(names[bl.ActualColumn] + " vs " + names[bl.BaselineColumn] +
                                "：平均同环比 " + Pct(avg) + "（上升 " + up + " 行 / 下降 " + down + " 行，配对依据：" + bl.Basis + "）");
                        }
                    }
                }

                Add(result, allow, "同环比对比", baselineLines);
            }
            else
            {
                // P0-1 显式声明（验收口径：无基线数据注入 → 声明必须出现）——
                // 并入已有概览节（同一标题 Add 两次会产出两个概览节，属组装缺陷）。
                var declaration = "⚠ 未发现对比基线列（同期/目标/预算语义）——本报告结论仅基于绝对值。提供上月/目标数据可升级为对比洞察。";
                var overviewSection = result.Sections.FirstOrDefault(s => s.Title == "概览");
                if (overviewSection != null)
                {
                    overviewSection.Lines.Add(declaration);
                }
                else
                {
                    Add(result, allow, "概览", declaration);
                }
            }

            // ---- P0-2（Q6 拍板）：生成后自检回显（对齐图表 auto_check 治理口径） ----
            result.SelfCheck = SelfCheck(result, allow, matrix, names);

            // ---- 组装三模式输出 ----
            var wordCount = result.Sections.Sum(s => s.Lines.Sum(l => l.Length));
            result.WordCount = wordCount;
            result.ReportGrid = BuildGrid(sheet, address, result.Sections);
            result.Markdown = BuildMarkdown(sheet, address, result.Sections);
            result.Html = BuildHtml(sheet, address, result.Sections);
            return result;
        }

        private static void Add(ReportResult result, HashSet<string> allow, string title, IEnumerable<string> lines)
        {
            if (allow != null && !allow.Contains(title))
            {
                return;
            }

            result.Sections.Add(new ReportSection { Title = title, Lines = lines.ToList() });
        }

        private static void Add(ReportResult result, HashSet<string> allow, string title, string line)
        {
            Add(result, allow, title, new[] { line });
        }

        /// <summary>
        /// P0-1（Q4）：基线列探测（纯逻辑确定性）——列名语义启发式：
        /// 目标类（目标/预算/计划/target/budget/goal）与同期类（同期/上期/上月/去年/prior/last）。
        /// 配对：基线列名去掉语义词后的剩余文本与其他数值列名匹配（如「销售额目标」→「销售额」）；
        /// 无配对不硬猜（该基线列跳过）。每个数值列至多配一个目标基线+一个同期基线。
        /// </summary>
        internal static List<BaselinePair> DetectBaselines(string[] names, int columnsCount, List<int> numericColumns)
        {
            var result = new List<BaselinePair>();
            var targetKeywords = new[] { "目标", "预算", "计划", "target", "budget", "goal", "plan" };
            var periodKeywords = new[] { "同期", "上期", "上月", "去年", "prior", "last" };

            foreach (var c in numericColumns)
            {
                var colName = names[c] ?? string.Empty;
                var lower = colName.ToLowerInvariant();

                // 目标类
                var kw = targetKeywords.FirstOrDefault(k => lower.Contains(k) || colName.Contains(k));
                if (kw != null)
                {
                    var stripped = colName.Replace(kw, string.Empty).Trim();
                    if (stripped.Length > 0)
                    {
                        // 剩余文本与其他数值列名匹配
                        var match = numericColumns.FirstOrDefault(o => o != c &&
                            ((names[o] ?? string.Empty).IndexOf(stripped, StringComparison.OrdinalIgnoreCase) >= 0));
                        if (match >= 0)
                        {
                            result.Add(new BaselinePair
                            {
                                ActualColumn = match,
                                BaselineColumn = c,
                                Kind = "target",
                                Basis = "「" + colName + "」含目标语义词「" + kw + "」，剩余「" + stripped + "」配对「" + names[match] + "」"
                            });
                            continue;
                        }
                    }
                }

                // 同期类
                kw = periodKeywords.FirstOrDefault(k => lower.Contains(k) || colName.Contains(k));
                if (kw != null)
                {
                    var stripped = colName.Replace(kw, string.Empty).Trim();
                    if (stripped.Length > 0)
                    {
                        var match = numericColumns.FirstOrDefault(o => o != c &&
                            ((names[o] ?? string.Empty).IndexOf(stripped, StringComparison.OrdinalIgnoreCase) >= 0));
                        if (match >= 0)
                        {
                            result.Add(new BaselinePair
                            {
                                ActualColumn = match,
                                BaselineColumn = c,
                                Kind = "period",
                                Basis = "「" + colName + "」含同期语义词「" + kw + "」，剩余「" + stripped + "」配对「" + names[match] + "」"
                            });
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// P0-2（Q6）：生成后自检（纯逻辑）——空段/占位符/覆盖匹配/关键数字格式四类。
        /// 负例必须被抓到（对齐 verify_dashboard_deep U29 篡改必抓先例）。
        /// </summary>
        internal static ReportSelfCheck SelfCheck(ReportResult result, HashSet<string> allow, object[,] matrix, string[] names)
        {
            var check = new ReportSelfCheck();
            foreach (var section in result.Sections)
            {
                check.CheckedCount++;
                // ① 空段
                if (section.Lines == null || section.Lines.Count == 0)
                {
                    check.Issues.Add("空段：「" + section.Title + "」无内容行。");
                    continue;
                }

                foreach (var line in section.Lines)
                {
                    // ② 占位符/TODO
                    if (line.Contains("TODO") || line.Contains("占位") || line.Contains("待补") || line.Trim().Length == 0)
                    {
                        check.Issues.Add("占位符：「" + section.Title + "」含「" + line.Trim() + "」。");
                    }
                }
            }

            // ③ 覆盖匹配：白名单里声明但未产出的章节
            if (allow != null)
            {
                foreach (var allowed in allow)
                {
                    if (!result.Sections.Any(s => s.Title == allowed))
                    {
                        check.Issues.Add("覆盖缺口：sections 白名单声明了「" + allowed + "」但未产出（数据形态不支持）。");
                    }
                }
            }

            // ④ 概览数字与矩阵对账（维度行数）
            var overview = result.Sections.FirstOrDefault(s => s.Title == "概览");
            if (overview != null && overview.Lines.Count > 0)
            {
                var dimLine = overview.Lines.FirstOrDefault(l => l.StartsWith("维度："));
                if (dimLine != null)
                {
                    var actualRows = Math.Max(matrix.GetLength(0) - 1, 0);
                    var claimedRows = 0;
                    var m = System.Text.RegularExpressions.Regex.Match(dimLine, "维度：([0-9]+) 行");
                    if (m.Success && int.TryParse(m.Groups[1].Value, out claimedRows) && claimedRows != actualRows)
                    {
                        check.Issues.Add("数字对账：概览维度行数 " + claimedRows + " 与实际数据行 " + actualRows + " 不一致。");
                    }
                }
            }

            check.Verdict = check.Issues.Count == 0 ? "ok" : "issues";
            return check;
        }

        private static int CountBlankColumns(int[] nonBlank, int dataRows)
        {
            if (dataRows <= 0)
            {
                return 0;
            }

            return nonBlank.Count(value => value == 0);
        }

        private static string Trend(List<double> values, out double magnitude)
        {
            var increasing = true;
            var decreasing = true;
            magnitude = 0.0;
            for (var i = 1; i < values.Count; i++)
            {
                if (values[i] < values[i - 1]) { increasing = false; }
                if (values[i] > values[i - 1]) { decreasing = false; }
            }

            if (increasing && !decreasing)
            {
                magnitude = RelativeChange(values[0], values[values.Count - 1]);
                return "上升";
            }

            if (decreasing && !increasing)
            {
                magnitude = RelativeChange(values[0], values[values.Count - 1]);
                return "下降";
            }

            if (increasing && decreasing)
            {
                return "平稳";
            }

            return "波动";
        }

        private static double RelativeChange(double first, double last)
        {
            if (Math.Abs(first) < 1e-12)
            {
                return 0.0;
            }

            return (last - first) / Math.Abs(first);
        }

        private static double CorrelationAligned(
            List<int> rowsA, List<double> a,
            List<int> rowsB, List<double> b,
            int startRow, int endRowExclusive)
        {
            var x = new List<double>(64);
            var y = new List<double>(64);
            var i = 0;
            var j = 0;
            while (i < rowsA.Count && j < rowsB.Count)
            {
                if (rowsA[i] == rowsB[j] && rowsA[i] >= startRow && rowsA[i] < endRowExclusive)
                {
                    x.Add(a[i]);
                    y.Add(b[j]);
                    i++;
                    j++;
                }
                else if (rowsA[i] < rowsB[j])
                {
                    i++;
                }
                else
                {
                    j++;
                }
            }

            if (x.Count < 3)
            {
                return double.NaN;
            }

            var xMean = x.Average();
            var yMean = y.Average();
            var numerator = 0.0;
            var denomX = 0.0;
            var denomY = 0.0;
            for (var n = 0; n < x.Count; n++)
            {
                var dx = x[n] - xMean;
                var dy = y[n] - yMean;
                numerator += dx * dy;
                denomX += dx * dx;
                denomY += dy * dy;
            }

            var denom = Math.Sqrt(denomX * denomY);
            return denom == 0 ? double.NaN : numerator / denom;
        }

        private static string[] Suggest()
        {
            // 模板化建议；具体触发条件由调用方在行内展开。
            return new[]
            {
                "建议先补充缺失单元格，再以完整样本复核关键统计。",
                "对离群点所在记录逐条复核，确认是否为录入或口径问题。",
                "对高相关的数值列，深入业务解释其联动关系，而非直接视为因果。",
                "可结合“推荐图表旁路”：将重点列生成图表便于直观审阅。"
            };
        }

        private static object[,] BuildGrid(string sheet, string address, List<ReportSection> sections)
        {
            var lines = new List<string>
            {
                "XAI 分析报告",
                "来源：" + (string.IsNullOrWhiteSpace(sheet) ? "" : sheet + "!") + address,
                "生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
            };
            foreach (var section in sections)
            {
                lines.Add("◆ " + section.Title);
                foreach (var line in section.Lines)
                {
                    lines.Add("- " + line);
                }
            }

            var grid = new object[lines.Count, 1];
            for (var i = 0; i < lines.Count; i++)
            {
                grid[i, 0] = lines[i];
            }

            return grid;
        }

        private static string BuildMarkdown(string sheet, string address, List<ReportSection> sections)
        {
            var builder = new StringBuilder();
            builder.Append("# XAI 分析报告\n\n");
            builder.Append("**来源**：" + (string.IsNullOrWhiteSpace(sheet) ? "" : sheet + "!") + address + "  \n");
            builder.Append("**生成时间**：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "\n\n");
            foreach (var section in sections)
            {
                builder.Append("## ").Append(section.Title).AppendLine();
                foreach (var line in section.Lines)
                {
                    builder.Append("- ").Append(line).AppendLine();
                }
                builder.AppendLine();
            }
            return builder.ToString();
        }

        private static string BuildHtml(string sheet, string address, List<ReportSection> sections)
        {
            var builder = new StringBuilder();
            builder.Append("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">");
            builder.Append("<title>XAI 分析报告</title></head><body>");
            builder.Append("<h1>XAI 分析报告</h1>");
            builder.Append("<p>来源：").Append(sheet).Append("!").Append(address).Append("<br>");
            builder.Append("生成时间：").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).Append("</p>");
            foreach (var section in sections)
            {
                builder.Append("<h2>").Append(section.Title).Append("</h2><ul>");
                foreach (var line in section.Lines)
                {
                    builder.Append("<li>").Append(HtmlEscape(line)).Append("</li>");
                }
                builder.Append("</ul>");
            }
            builder.Append("</body></html>");
            return builder.ToString();
        }

        private static string HtmlEscape(string text)
        {
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static string Header(object[,] matrix, int column)
        {
            if (matrix.GetLength(0) == 0)
            {
                return "COL" + (column + 1);
            }

            var header = matrix[0, column];
            var text = header == null ? string.Empty : Convert.ToString(header).Trim();
            return string.IsNullOrWhiteSpace(text) ? "COL" + (column + 1) : text;
        }

        private static string Label(object[,] matrix, int rowIndex)
        {
            if (rowIndex < matrix.GetLength(0))
            {
                var first = matrix[rowIndex, 0];
                var text = first == null ? string.Empty : Convert.ToString(first).Trim();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text;
                }
            }
            return "第" + (rowIndex + 1) + "行";
        }

        private static double R4(double value)
        {
            return Math.Round(value, 4);
        }

        private static string Pct(double ratio)
        {
            return (ratio * 100).ToString("0.0", CultureInfo.InvariantCulture) + "%";
        }
    }
}