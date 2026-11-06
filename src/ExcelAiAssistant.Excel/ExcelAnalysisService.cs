using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第三期-数据分析：区域统计摘要、相关性分析、列值分布、图表推荐。只读 + 纯 .NET 计算。
    /// </summary>
    public sealed class ExcelAnalysisService
    {
        private const int MaxAnalyzedCells = 100000;

        private readonly dynamic application;
        private readonly OperationRecorder recorder;

        public ExcelAnalysisService(object application)
            : this(application, null)
        {
        }

        public ExcelAnalysisService(object application, OperationRecorder recorder)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.recorder = recorder;
        }

        public ToolResult AnalyzeRangeStatistics(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            var colStats = new List<object>();
            var numericColumnCount = 0;
            var numericCols = new List<int>();
            for (var column = 0; column < columns; column++)
            {
                var name = Header(matrix, column);
                var numbers = new List<double>();
                var texts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var nonNull = 0;
                var nullCount = 0;
                for (var row = 1; row < matrix.GetLength(0); row++)
                {
                    // CVErr 错误单元格归一为错误文本，防止被 TryParse 当数字污染统计（M16）。
                    var cell = ExcelErrorValues.Normalize(matrix[row, column]);
                    if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                    {
                        nullCount++;
                        continue;
                    }

                    nonNull++;
                    var text = Convert.ToString(cell).Trim();
                    if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        numbers.Add(number);
                    }
                    else
                    {
                        texts.TryGetValue(text, out var count);
                        texts[text] = count + 1;
                    }
                }

                var isNumeric = numbers.Count > 0;
                if (isNumeric && numbers.Count >= nonNull * 0.8)
                {
                    numericColumnCount++;
                    numericCols.Add(column + 1);
                    // 离群点检测须在 Sort 之前做（DetectOutliers 返回的索引对应原始顺序，即数据行序）。
                    var outlierIndexes = OutlierDetector.DetectOutliers(numbers.ToArray(), "iqr");
                    var outlierValues = outlierIndexes.Select(i => R4(numbers[i - 1])).ToList();
                    numbers.Sort();
                    var count = numbers.Count;
                    var mean = numbers.Average();
                    var median = count % 2 == 0
                        ? (numbers[count / 2 - 1] + numbers[count / 2]) / 2.0
                        : numbers[count / 2];
                    var variance = numbers.Select(x => (x - mean) * (x - mean)).Sum() / Math.Max(count - 1, 1);
                    colStats.Add(new
                    {
                        column = column + 1,
                        name,
                        type = "number",
                        non_null = nonNull,
                        null_count = nullCount,
                        min = R4(numbers.Min()),
                        max = R4(numbers.Max()),
                        mean = R4(mean),
                        median = R4(median),
                        stdev = R4(Math.Sqrt(variance)),
                        sum = R4(numbers.Sum()),
                        unique_count = numbers.Distinct().Count()
                    });
                }
                else
                {
                    colStats.Add(new
                    {
                        column = column + 1,
                        name,
                        type = nonNull == 0 ? "blank" : "text",
                        non_null = nonNull,
                        null_count = nullCount,
                        unique_count = texts.Count,
                        top_values = texts.OrderByDescending(kv => kv.Value).Take(5)
                            .Select(kv => new { value = kv.Key, count = kv.Value }).ToArray()
                    });
                }
            }

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    rows = matrix.GetLength(0),
                    columns,
                    numeric_column_count = numericColumnCount,
                    text_column_count = columns - numericColumnCount,
                    columns_summary = colStats.ToArray()
                }
            };
        }

        public ToolResult CorrelationAnalysis(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var method = GetString(args, "method", "pearson");
            var threshold = Math.Max(0.05, Math.Min(GetDouble(args, "threshold", 0.7), 0.95));
            // 归一化 method：pearson→pearson；rank/spearman 均为秩相关(Spearman=对秩算Pearson)；
            // 未知值（如 kendall）明确报错，避免静默按 Pearson 计算却返回误导性 method 名。
            if (method.Equals("rank", StringComparison.OrdinalIgnoreCase) ||
                method.Equals("spearman", StringComparison.OrdinalIgnoreCase))
            {
                method = "rank";
            }
            else if (!method.Equals("pearson", StringComparison.OrdinalIgnoreCase))
            {
                return Invalid("INVALID_METHOD", $"method 仅支持 pearson/spearman，当前为 {method}。");
            }

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            var numericColumns = new List<int>();

            // 找出数值列（表头之后至少有 2 个可解析数值）。
            var columnData = new List<List<double>>();
            for (var column = 0; column < columns; column++)
            {
                var numbers = new List<double>();
                for (var row = 1; row < matrix.GetLength(0); row++)
                {
                    // CVErr 错误单元格归一为错误文本，防止被 TryParse 当数字参与相关系数（M16）。
                    var cell = ExcelErrorValues.Normalize(matrix[row, column]);
                    if (cell == null)
                    {
                        continue;
                    }

                    if (double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        numbers.Add(number);
                    }
                }

                if (numbers.Count >= 2)
                {
                    numericColumns.Add(column);
                    columnData.Add(numbers);
                }
            }

            var n = numericColumns.Count;
            var pairs = new List<object>();
            var coefficients = new double[n, n];
            for (var i = 0; i < n; i++)
            {
                coefficients[i, i] = 1.0;
            }

            for (var i = 0; i < n; i++)
            {
                for (var j = i + 1; j < n; j++)
                {
                    var coefficient = ComputeCorrelation(columnData[i], columnData[j], method);
                    coefficients[i, j] = coefficient;
                    coefficients[j, i] = coefficient;
                    pairs.Add(new
                    {
                        col_i = numericColumns[i] + 1,
                        col_j = numericColumns[j] + 1,
                        name_i = Header(matrix, numericColumns[i]),
                        name_j = Header(matrix, numericColumns[j]),
                        coefficient = R4(coefficient),
                        strong = !double.IsNaN(coefficient) && Math.Abs(coefficient) >= threshold
                    });
                }
            }

            var ordered = pairs.Cast<object>()
                .Select(p => (dynamic)p)
                .OrderByDescending(p => Math.Abs((double)p.coefficient))
                .Select(p => (object)p)
                .ToList();

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    method,
                    threshold = R4(threshold),
                    strong_pair_count = ordered.Cast<object>().Select(p => (dynamic)p).Count(p => (bool)p.strong),
                    numeric_column_count = n,
                    columns = numericColumns.Select(c => new { column = c + 1, name = Header(matrix, c) }).ToArray(),
                    top_pairs = ordered.Take(10).ToArray()
                }
            };
        }

        public ToolResult ColumnValueDistribution(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var column = GetInt(args, "column", 1);
            var top = Math.Max(1, Math.Min(GetInt(args, "top", 10), 100));
            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            if (column < 1 || column > columns)
            {
                return Invalid("INVALID_COLUMN", $"列号必须介于 1 到 {columns}。");
            }

            var distribution = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var nullCount = 0;
            for (var row = 1; row < matrix.GetLength(0); row++)
            {
                // CVErr 错误单元格归一为错误文本，防止以 "-2146826246" 字面值混入分布（M16）。
                var cell = ExcelErrorValues.Normalize(matrix[row, column - 1]);
                if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                {
                    nullCount++;
                    continue;
                }

                var text = Convert.ToString(cell).Trim();
                distribution.TryGetValue(text, out var count);
                distribution[text] = count + 1;
            }

            var topValues = distribution.OrderByDescending(kv => kv.Value).Take(top)
                .Select(kv => new { value = kv.Key, count = kv.Value }).ToArray();
            var countSpan = AnalysisBasis.ColumnSpan(read.Sheet, read.Address, column);
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    sheet = read.Sheet,
                    column,
                    name = Header(matrix, column - 1),
                    distinct_count = distribution.Count,
                    null_count = nullCount,
                    total_rows = Math.Max(matrix.GetLength(0) - 1, 0),
                    distribution = distribution.OrderByDescending(kv => kv.Value).Take(top)
                        .Select(kv => new { value = kv.Key, count = kv.Value }).ToArray()
                }
            };
        }

        /// <summary>
        /// 列质量画像（M14，对标 Power Query 列质量/分布/剖析单面板）：对指定列一次给出
        /// 质量（空值率/错误率/唯一值/类型占比）、分布（Top N 高频值）与剖析（数值列统计/文本长度统计），
        /// 并附质量提示（blank_heavy/has_errors/mixed_types/padded_text/constant）。只读。
        /// </summary>
        public ToolResult ColumnProfile(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var column = GetInt(args, "column", 1);
            var top = Math.Max(1, Math.Min(GetInt(args, "top", 10), 100));
            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            if (column < 1 || column > columns)
            {
                return Invalid("INVALID_COLUMN", $"列号必须介于 1 到 {columns}。");
            }

            var profile = ColumnProfileEngine.Compute(matrix, column, top);
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    column,
                    name = Header(matrix, column - 1),
                    total_rows = profile.TotalRows,
                    quality = new
                    {
                        blank_count = profile.BlankCount,
                        blank_rate = profile.BlankRate,
                        error_count = profile.ErrorCount,
                        error_rate = profile.ErrorRate,
                        distinct_count = profile.DistinctCount,
                        distinct_rate = profile.DistinctRate,
                        type_counts = profile.TypeCounts
                    },
                    numbers = profile.Numbers == null
                        ? null
                        : new
                        {
                            count = profile.Numbers.Count,
                            min = profile.Numbers.Min,
                            max = profile.Numbers.Max,
                            mean = profile.Numbers.Mean,
                            median = profile.Numbers.Median,
                            std_dev = profile.Numbers.StdDev
                        },
                    text = profile.Text == null
                        ? null
                        : new
                        {
                            count = profile.Text.Count,
                            min_length = profile.Text.MinLength,
                            max_length = profile.Text.MaxLength,
                            avg_length = profile.Text.AvgLength,
                            whitespace_only_count = profile.Text.WhitespaceOnlyCount,
                            padded_count = profile.Text.PaddedCount
                        },
                    top_values = profile.TopValues.Select(v => new { value = v.Value, count = v.Count }).ToArray(),
                    quality_flags = profile.QualityFlags
                }
            };
        }

        /// <summary>
        /// FR-K1 业务口径规则文件：读取活动工作簿同目录的《_口径.md》（每行一条「- 术语: 定义」）。
        /// 只读这一个固定文件名（安全红线：不做任意文件读取）；隐私红线：文件仅本地读取，绝不外发上传。
        /// 限幅：最多 200 行 / 8KB（超出截断并在返回值标记）；条目上限 100。找不到文件时返回创建模板。
        /// </summary>
        public ToolResult ReadBusinessGlossary(Dictionary<string, object> args)
        {
            try
            {
                dynamic workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    return Invalid("NO_WORKBOOK", "当前没有打开的工作簿，无法定位口径文件。");
                }

                var fullName = Convert.ToString(workbook.FullName);
                if (string.IsNullOrWhiteSpace(fullName) || !System.IO.Path.IsPathRooted(fullName))
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { found = false, reason = "workbook_unsaved", hint = GlossaryHint }
                    };
                }

                var directory = System.IO.Path.GetDirectoryName(fullName);
                var path = System.IO.Path.Combine(directory ?? string.Empty, GlossaryFileName);
                if (!System.IO.File.Exists(path))
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { found = false, reason = "file_missing", expected_path = path, hint = GlossaryHint }
                    };
                }

                var text = ReadGlossaryText(path);
                var allLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                var totalLines = allLines.Length;
                var lineCapHit = totalLines > GlossaryMaxLines;
                var sizeCapHit = false;
                var kept = new List<string>();
                var size = 0;
                foreach (var line in allLines)
                {
                    if (kept.Count >= GlossaryMaxLines)
                    {
                        lineCapHit = true;
                        break;
                    }

                    var cost = System.Text.Encoding.UTF8.GetByteCount(line + "\n");
                    if (size + cost > GlossaryMaxBytes)
                    {
                        sizeCapHit = true;
                        break;
                    }

                    size += cost;
                    kept.Add(line);
                }

                var entries = Core.Diagnostics.GlossaryParser.Parse(kept, GlossaryMaxEntries);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        found = true,
                        file = System.IO.Path.GetFileName(path),
                        directory,
                        total_lines = totalLines,
                        line_cap = GlossaryMaxLines,
                        truncated = lineCapHit || sizeCapHit,
                        truncated_by = lineCapHit ? "line_cap_200" : (sizeCapHit ? "size_cap_8kb" : (string)null),
                        entry_count = entries.Count,
                        entries = entries.Select(e => new { term = e.Term, definition = e.Definition }).ToArray()
                    }
                };
            }
            catch (System.Runtime.InteropServices.COMException ce)
            {
                return Invalid("GLOSSARY_READ_FAILED", "读取口径文件失败：" + ce.Message);
            }
            catch (Exception ex)
            {
                return Invalid("GLOSSARY_READ_FAILED", "读取口径文件失败：" + ex.Message);
            }
        }

        private const string GlossaryFileName = "_口径.md";
        private const int GlossaryMaxLines = 200;
        private const int GlossaryMaxBytes = 8 * 1024;
        private const int GlossaryMaxEntries = 100;

        private const string GlossaryHint =
            "未找到 _口径.md。可在工作簿同目录创建（UTF-8 编码），格式每行一条「- 术语: 定义」，例如：\n" +
            "- 毛利: 收入-成本（不含税，不含退货）\n" +
            "- 活跃用户: 当日有登录且会话时长≥5分钟\n" +
            "- 大区: 仅华东/华北/华南/西南四区";

        /// <summary>口径文件解码：BOM 优先 → UTF-8 严格校验 → GBK 兜底（中文系统记事本常见编码）。</summary>
        private static string ReadGlossaryText(string path)
        {
            var bytes = System.IO.File.ReadAllBytes(path);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                return System.Text.Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            }

            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                return System.Text.Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            }

            try
            {
                return new System.Text.UTF8Encoding(false, true).GetString(bytes);
            }
            catch (System.Text.DecoderFallbackException)
            {
                return System.Text.Encoding.GetEncoding(936).GetString(bytes);
        }
        }

        // ---- 趋势预测（轻量：线性回归/移动平均外推；季节 Holt/ARIMA 走 python-advanced-analysis 技能，分工口径不变）----
        public ToolResult TrendForecast(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var timeColumn = Math.Max(1, GetInt(args, "time_column", 1));
            var valueColumn = Math.Max(1, GetInt(args, "value_column", 2));
            var periods = Math.Max(1, Math.Min(GetInt(args, "periods", 3), 12));
            var window = Math.Max(2, Math.Min(GetInt(args, "window", 3), 12));
            var method = GetString(args, "method", "linear").ToLowerInvariant();
            if (method != "linear" && method != "moving_average")
            {
                return Invalid("INVALID_METHOD", "method 仅支持 linear/moving_average；季节性 Holt-Winters/ARIMA 属高阶预测，请走 python-advanced-analysis 技能（run_python_excel + statsmodels）。");
            }

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            if (timeColumn > columns || valueColumn > columns)
            {
                return Invalid("INVALID_COLUMN", $"time_column/value_column 必须介于 1 到 {columns}。");
            }
            if (timeColumn == valueColumn)
            {
                return Invalid("INVALID_COLUMN", "time_column 与 value_column 不能为同一列。");
            }

            var labels = new List<string>();
            var values = new List<double>();
            for (var row = 1; row < matrix.GetLength(0); row++)
            {
                // CVErr 错误单元格归一（M16 同款），非数值行跳过不进样本。
                var cell = ExcelErrorValues.Normalize(matrix[row, valueColumn - 1]);
                if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                {
                    continue;
                }
                if (!double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                {
                    continue;
                }
                var timeCell = ExcelErrorValues.Normalize(matrix[row, timeColumn - 1]);
                labels.Add(timeCell == null ? string.Empty : Convert.ToString(timeCell).Trim());
                values.Add(number);
            }

            // 预测是外推，样本过少会严重误导，fail-closed 拒算（与“拒绝生成空图表”同哲学）。
            if (values.Count < 6)
            {
                return Invalid("INSUFFICIENT_POINTS", $"有效数值点仅 {values.Count} 个，趋势预测至少需要 6 个连续数值点；请补充历史数据，或对短序列改用移动平均口径并在报告注明局限。");
            }

            var forecasts = new List<object>();
            TrendForecaster.ForecastResult computed;
            try
            {
                computed = TrendForecaster.Compute(values, method, window, periods);
            }
            catch (ArgumentException)
            {
                return Invalid("INSUFFICIENT_POINTS", $"有效数值点仅 {values.Count} 个，趋势预测至少需要 {TrendForecaster.MinPoints} 个连续数值点；请补充历史数据，或对短序列改用移动平均口径并在报告注明局限。");
            }

            var trendText = computed.Trend;
            var timeName = Header(matrix, timeColumn - 1);
            var valueName = Header(matrix, valueColumn - 1);
            var valueSpan = AnalysisBasis.ColumnSpan(read.Sheet, read.Address, valueColumn);
            var timeSpan = AnalysisBasis.ColumnSpan(read.Sheet, read.Address, timeColumn);
            for (var k = 1; k <= periods; k++)
            {
                forecasts.Add(new { period = TrendForecaster.NextLabel(labels.Count > 0 ? labels[labels.Count - 1] : null, k), value = R4(computed.ForecastValues[k - 1]) });
            }

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    method,
                    time_column = timeColumn,
                    value_column = valueColumn,
                    time_column_name = timeName,
                    value_column_name = valueName,
                    points = values.Count,
                    periods,
                    window = method == "moving_average" ? window : 0,
                    trend = trendText,
                    slope = R4(computed.Slope),
                    r2 = R4(computed.R2),
                    forecasts = forecasts.ToArray(),
                    basis = new Dictionary<string, object>
                    {
                        ["source"] = AnalysisBasis.Source(read.Sheet, read.Address),
                        ["sample_rule"] = $"时间列「{timeName}」+ 数值列「{valueName}」共 {values.Count} 个有效数值点（非数值/空值/错误值行剔除）",
                        ["method_text"] = method == "linear"
                            ? "最小二乘线性回归外推 y=a+b·x（x=1..n），r2 为拟合优度；r2 低于 0.5 时预测仅供参考"
                            : $"最近 {window} 期移动平均基准水平外推（不含趋势项）",
                        ["boundary_text"] = "线性外推不覆盖季节性与突变，季节性/复杂序列请走 python-advanced-analysis 技能（statsmodels）",
                        ["formula_examples"] = method == "linear"
                            ? new[]
                            {
                                $"SLOPE({valueSpan},{timeSpan})",
                                $"INTERCEPT({valueSpan},{timeSpan})",
                                $"FORECAST.LINEAR(新期数,{valueSpan},{timeSpan})"
                            }
                            : new[]
                            {
                                $"AVERAGE(最近{window}期数值区间)"
                            }
                    }
                }
            };
        }

        // ---- 假设检验标准工具（Welch t/配对 t/ANOVA/卡方；纯逻辑在 Core/Tools/HypothesisTester）----
        public ToolResult HypothesisTest(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var test = GetString(args, "test", "two_sample_t").ToLowerInvariant();
            var alpha = Math.Max(0.001, Math.Min(GetDouble(args, "alpha", 0.05), 0.2));

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            var validTests = new[] { "two_sample_t", "paired_t", "anova", "chi_square" };
            if (Array.IndexOf(validTests, test) < 0)
            {
                return Invalid("INVALID_TEST", "test 仅支持 two_sample_t / paired_t / anova / chi_square；非参检验（Mann-Whitney 等）与复杂设计请走 python-advanced-analysis 技能。");
            }

            HypothesisTester.TestResult r;
            string sampleRule;
            string[] formulaExamples;
            try
            {
                if (test == "two_sample_t" || test == "anova")
                {
                    var valueColumn = Math.Max(1, GetInt(args, "value_column", 2));
                    var groupColumn = Math.Max(1, GetInt(args, "group_column", 1));
                    if (valueColumn > columns || groupColumn > columns || valueColumn == groupColumn)
                    {
                        return Invalid("INVALID_COLUMN", $"value_column/group_column 须介于 1 到 {columns} 且不能为同一列。");
                    }

                    var groups = new List<(string Label, List<double> Values)>();
                    var byLabel = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
                    for (var row = 1; row < matrix.GetLength(0); row++)
                    {
                        var cell = ExcelErrorValues.Normalize(matrix[row, valueColumn - 1]);
                        if (cell == null || !double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                        {
                            continue;
                        }

                        var gCell = ExcelErrorValues.Normalize(matrix[row, groupColumn - 1]);
                        var label = gCell == null || string.IsNullOrWhiteSpace(Convert.ToString(gCell)) ? "未分组" : Convert.ToString(gCell).Trim();
                        if (!byLabel.TryGetValue(label, out var list))
                        {
                            list = new List<double>();
                            byLabel[label] = list;
                        }

                        list.Add(number);
                    }

                    groups.AddRange(byLabel.Select(kv => (kv.Key, kv.Value)));
                    if (test == "two_sample_t")
                    {
                        if (groups.Count != 2)
                        {
                            return Invalid("INVALID_GROUP_COUNT", $"two_sample_t 恰好需要 2 组，当前 {groups.Count} 组（组标签来自 group_column）；≥3 组请用 test=anova。");
                        }

                        r = HypothesisTester.WelchTTest(groups[0].Values, groups[1].Values, groups[0].Label, groups[1].Label);
                        sampleRule = $"数值列「{Header(matrix, valueColumn - 1)}」按分组列「{Header(matrix, groupColumn - 1)}」分成 2 组（成对完整观测，非数值行剔除）";
                        formulaExamples = new[]
                        {
                            $"T.TEST({AnalysisBasis.ColumnSpan(read.Sheet, read.Address, groups[1].Label == null ? 1 : valueColumn)},{AnalysisBasis.ColumnSpan(read.Sheet, read.Address, valueColumn)},2,3)（Excel 双样本异方差 t 检验）"
                        };
                    }
                    else
                    {
                        if (groups.Count < 2)
                        {
                            return Invalid("INVALID_GROUP_COUNT", $"anova 至少需要 2 组，当前 {groups.Count} 组。");
                        }

                        r = HypothesisTester.AnovaFTest(groups);
                        sampleRule = $"数值列「{Header(matrix, valueColumn - 1)}」按分组列「{Header(matrix, groupColumn - 1)}」分为 {groups.Count} 组";
                        formulaExamples = new[]
                        {
                            $"F.DIST.RT(F值,{r.Df.ToString("0", CultureInfo.InvariantCulture)},{r.Df2.ToString("0", CultureInfo.InvariantCulture)})",
                            "单因素方差分析：组间方差 / 组内方差"
                        };
                    }
                }
                else if (test == "paired_t")
                {
                    var valueColumn = Math.Max(1, GetInt(args, "value_column", 2));
                    var valueColumn2 = Math.Max(1, GetInt(args, "value_column2", 3));
                    if (valueColumn > columns || valueColumn2 > columns || valueColumn == valueColumn2)
                    {
                        return Invalid("INVALID_COLUMN", $"value_column/value_column2 须介于 1 到 {columns} 且不能为同一列。");
                    }

                    var a = new List<double>();
                    var b = new List<double>();
                    for (var row = 1; row < matrix.GetLength(0); row++)
                    {
                        var c1 = ExcelErrorValues.Normalize(matrix[row, valueColumn - 1]);
                        var c2 = ExcelErrorValues.Normalize(matrix[row, valueColumn2 - 1]);
                        if (c1 == null || c2 == null) continue;
                        if (!double.TryParse(Convert.ToString(c1).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var n1)) continue;
                        if (!double.TryParse(Convert.ToString(c2).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var n2)) continue;
                        a.Add(n1);
                        b.Add(n2);
                    }

                    r = HypothesisTester.PairedTTest(a, b, Header(matrix, valueColumn - 1), Header(matrix, valueColumn2 - 1));
                    sampleRule = $"配对列「{Header(matrix, valueColumn - 1)}」与「{Header(matrix, valueColumn2 - 1)}」逐行成对完整观测，共 {r.TotalN} 对";
                    formulaExamples = new[] { $"T.TEST({AnalysisBasis.ColumnSpan(read.Sheet, read.Address, valueColumn)},{AnalysisBasis.ColumnSpan(read.Sheet, read.Address, valueColumn2)},2,1)（Excel 成对双样本 t 检验）" };
                }
                else
                {
                    var columnA = Math.Max(1, GetInt(args, "column_a", 1));
                    var columnB = Math.Max(1, GetInt(args, "column_b", 2));
                    if (columnA > columns || columnB > columns || columnA == columnB)
                    {
                        return Invalid("INVALID_COLUMN", $"column_a/column_b 须介于 1 到 {columns} 且不能为同一列。");
                    }

                    var pairs = new List<(string, string)>();
                    for (var row = 1; row < matrix.GetLength(0); row++)
                    {
                        var c1 = ExcelErrorValues.Normalize(matrix[row, columnA - 1]);
                        var c2 = ExcelErrorValues.Normalize(matrix[row, columnB - 1]);
                        if (c1 == null || c2 == null) continue;
                        var s1 = Convert.ToString(c1).Trim();
                        var s2 = Convert.ToString(c2).Trim();
                        if (s1.Length == 0 || s2.Length == 0) continue;
                        pairs.Add((s1, s2));
                    }

                    r = HypothesisTester.ChiSquareIndependence(pairs);
                    sampleRule = $"类别列「{Header(matrix, columnA - 1)}」×「{Header(matrix, columnB - 1)}」交叉列联表，共 {r.TotalN} 条有效配对";
                    formulaExamples = new[] { "CHISQ.TEST(实际区域, 期望区域)（期望=行合计×列合计/总计）" };
                }
            }
            catch (ArgumentException ex)
            {
                return Invalid("HYPOTHESIS_INPUT_INVALID", ex.Message);
            }

            HypothesisTester.MarkSignificance(r, alpha);
            var conclusion = r.Significant
                ? $"在 α={alpha.ToString("0.###", CultureInfo.InvariantCulture)} 水平上显著（p={r.PValue.ToString("0.####", CultureInfo.InvariantCulture)} < α）：拒绝原假设"
                : $"在 α={alpha.ToString("0.###", CultureInfo.InvariantCulture)} 水平上不显著（p={r.PValue.ToString("0.####", CultureInfo.InvariantCulture)} ≥ α）：不能拒绝原假设";

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    method = r.Method,
                    statistic = R4(r.Statistic),
                    df = R4(r.Df),
                    df2 = r.Df2 > 0 ? R4(r.Df2) : 0,
                    p_value = r.PValue,
                    alpha,
                    significant = r.Significant,
                    conclusion,
                    effect_name = r.EffectName,
                    effect_size = r.EffectSize.HasValue ? R4(r.EffectSize.Value) : 0,
                    total_n = r.TotalN,
                    groups = r.Groups.Select(g => new { label = g.Label, n = g.N, mean = R4(g.Mean), variance = R4(g.Variance) }).ToArray(),
                    premise_notes = r.PremiseNotes.ToArray(),
                    basis = new Dictionary<string, object>
                    {
                        ["source"] = AnalysisBasis.Source(read.Sheet, read.Address),
                        ["sample_rule"] = sampleRule,
                        ["method_text"] = r.Method == "welch_t"
                            ? "Welch 双样本 t 检验（不假设方差齐性，Welch–Satterthwaite 自由度）"
                            : r.Method == "paired_t"
                                ? "配对 t 检验（逐行差值）"
                                : r.Method == "anova_f"
                                    ? "单因素方差分析 F 检验"
                                    : "卡方独立性检验（列联表）",
                        ["alpha_text"] = $"显著性水平 α={alpha.ToString("0.###", CultureInfo.InvariantCulture)}，p<α 判显著",
                        ["formula_examples"] = formulaExamples
                    }
                }
            };
        }

        // ---- 聚类分群标准工具（KMeans++/Lloyd + 轮廓系数定 k；纯逻辑在 Core/Tools/ClusterAnalyzer）----
        public ToolResult ClusterAnalysis(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var standardize = !GetString(args, "standardize", "on").Equals("off", StringComparison.OrdinalIgnoreCase);
            var kMaxRaw = GetInt(args, "k_max", 6);
            var kMax = Math.Max(ClusterAnalyzer.MinK, Math.Min(ClusterAnalyzer.MaxK, kMaxRaw));
            int? explicitK = null;
            if (args.TryGetValue("k", out var kObj) && kObj != null)
            {
                explicitK = Math.Max(ClusterAnalyzer.MinK, Math.Min(ClusterAnalyzer.MaxK, Convert.ToInt32(kObj is Newtonsoft.Json.Linq.JValue j ? j.Value : kObj)));
            }

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(0) > 0 ? matrix.GetLength(1) : 0;

            // 特征列解析：显式 columns（如 "2,3,4"）优先；缺省=数值占比≥60% 的列（与报告引擎同口径）。
            var featureCols = new List<int>();
            var columnsRaw = GetString(args, "columns", null);
            if (!string.IsNullOrWhiteSpace(columnsRaw))
            {
                foreach (var part in columnsRaw.Split(new[] { ',', '，', '、' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (int.TryParse(part.Trim(), out var col) && col >= 1 && col <= columns)
                    {
                        featureCols.Add(col - 1);
                    }
                }

                if (featureCols.Count == 0)
                {
                    return Invalid("INVALID_COLUMNS", $"columns 未解析出有效特征列；请传列号如 2,3,4（介于 1 到 {columns}）。");
                }
            }
            else
            {
                var dataRows = Math.Max(matrix.GetLength(0) - 1, 0);
                for (var c = 0; c < columns; c++)
                {
                    var numeric = 0;
                    for (var r = 1; r < matrix.GetLength(0); r++)
                    {
                        var cell = ExcelErrorValues.Normalize(matrix[r, c]);
                        if (cell != null && double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                        {
                            numeric++;
                        }
                    }

                    if (dataRows > 0 && numeric >= dataRows * 0.6)
                    {
                        featureCols.Add(c);
                    }
                }

                if (featureCols.Count == 0)
                {
                    return Invalid("NO_NUMERIC_FEATURES", "区域内未识别出数值特征列；请用 columns 参数显式指定特征列号（如 2,3,4）。");
                }
            }

            var featureNames = featureCols.Select(c => Header(matrix, c)).ToArray();

            // 完整观测行：所有特征列均为数值的行才参与聚类（缺失行剔除并计数）。
            var features = new List<double[]>();
            var skipped = 0;
            for (var r = 1; r < matrix.GetLength(0); r++)
            {
                var vec = new double[featureCols.Count];
                var ok = true;
                for (var i = 0; i < featureCols.Count; i++)
                {
                    var cell = ExcelErrorValues.Normalize(matrix[r, featureCols[i]]);
                    if (cell == null || !double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        ok = false;
                        break;
                    }

                    vec[i] = number;
                }

                if (ok) features.Add(vec); else skipped++;
            }

            ClusterAnalyzer.ClusterResult r2;
            try
            {
                r2 = ClusterAnalyzer.Compute(features, featureNames, explicitK, kMax, standardize);
            }
            catch (ArgumentException ex)
            {
                return Invalid("INSUFFICIENT_POINTS", ex.Message);
            }

            // 分派明细上限：≤1000 行全量返回，超出建议按簇画像/走 Python（防返回体爆炸）。
            var assignments = Enumerable.Range(0, r2.TotalPoints)
                .Select(i => new { data_row = i + 1, cluster = r2.Labels[i] })
                .Take(1000)
                .ToArray();

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    method = "kmeans++_lloyd",
                    k = r2.K,
                    best_silhouette = R4(r2.BestSilhouette),
                    standardized = r2.Standardized,
                    total_points = r2.TotalPoints,
                    skipped_rows = skipped,
                    k_trace = r2.KTrace.Select(t => new { k = t.K, silhouette = R4(t.Silhouette) }).ToArray(),
                    feature_columns = featureCols.Select((c, i) => new { column = c + 1, name = featureNames[i] }).ToArray(),
                    profiles = r2.Profiles.Select(p => new
                    {
                        cluster = p.Cluster,
                        size = p.Size,
                        means = featureCols.Select((c, i) => new { name = featureNames[i], mean = R4(p.MeanOriginal[i]) }).ToArray()
                    }).ToArray(),
                    assignments = assignments,
                    assignments_omitted = r2.TotalPoints > assignments.Length,
                    basis = new Dictionary<string, object>
                    {
                        ["source"] = AnalysisBasis.Source(read.Sheet, read.Address),
                        ["sample_rule"] = $"特征列「{string.Join("/", featureNames)}」完整观测行 {r2.TotalPoints} 行（含缺失/非数值行 {skipped} 行剔除）",
                        ["method_text"] = (r2.Standardized ? "Z-score 标准化后 " : "") + $"K-Means++ 初始化 + Lloyd 迭代（5 次重启取最优惯性）",
                        ["k_selection_text"] = explicitK.HasValue
                            ? $"k 由用户显式指定为 {explicitK.Value}（该 k 的平均轮廓系数见 best_silhouette）"
                            : $"k 在 2~{kMax} 内按平均轮廓系数最大自动选定（k_trace 可查各 k 得分）",
                        ["boundary_text"] = "KMeans 假设球形簇；轮廓系数仅作参考，分群业务解读请结合画像；DBSCAN/层次聚类等走 python-advanced-analysis 技能"
                    }
                }
            };
        }

        // ---- 多元线性回归标准工具（OLS + 系数表 + 残差诊断；纯逻辑在 Core/Tools/RegressionAnalyzer）----
        public ToolResult RegressionAnalysis(Dictionary<string, object> args)
        {
            var address = Required(args, "address");

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(0) > 0 ? matrix.GetLength(1) : 0;

            // 因变量列解析：显式 target_column 优先；缺省取区域最后一列。
            int targetCol;
            if (args.TryGetValue("target_column", out var tObj) && tObj != null)
            {
                targetCol = Convert.ToInt32(tObj is Newtonsoft.Json.Linq.JValue jv ? jv.Value : tObj);
                if (targetCol < 1 || targetCol > columns)
                {
                    return Invalid("INVALID_COLUMNS", $"target_column={targetCol} 超出范围（1~{columns}）。");
                }

                targetCol -= 1;
            }
            else
            {
                targetCol = columns - 1;
            }

            // 特征列解析：显式 columns（如 "2,3"）优先；缺省=除因变量外数值占比≥60% 的列。
            var featureCols = new List<int>();
            var columnsRaw = GetString(args, "columns", null);
            if (!string.IsNullOrWhiteSpace(columnsRaw))
            {
                foreach (var part in columnsRaw.Split(new[] { ',', '，', '、' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (int.TryParse(part.Trim(), out var col) && col >= 1 && col <= columns && col - 1 != targetCol)
                    {
                        featureCols.Add(col - 1);
                    }
                }

                if (featureCols.Count == 0)
                {
                    return Invalid("INVALID_COLUMNS", $"columns 未解析出有效特征列；请传列号如 2,3（介于 1 到 {columns} 且不等于因变量列 {targetCol + 1}）。");
                }
            }
            else
            {
                var dataRows = Math.Max(matrix.GetLength(0) - 1, 0);
                for (var c = 0; c < columns; c++)
                {
                    if (c == targetCol) continue;
                    var numeric = 0;
                    for (var r = 1; r < matrix.GetLength(0); r++)
                    {
                        var cell = ExcelErrorValues.Normalize(matrix[r, c]);
                        if (cell != null && double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                        {
                            numeric++;
                        }
                    }

                    if (dataRows > 0 && numeric >= dataRows * 0.6)
                    {
                        featureCols.Add(c);
                    }
                }

                if (featureCols.Count == 0)
                {
                    return Invalid("NO_NUMERIC_FEATURES", "区域内未识别出数值特征列；请用 columns 参数显式指定特征列号（如 2,3）。");
                }
            }

            var targetName = Header(matrix, targetCol);
            var featureNames = featureCols.Select(c => Header(matrix, c)).ToArray();

            // 完整观测行：因变量与所有特征列均为数值的行才参与回归（缺失行剔除并计数）。
            var yValues = new List<double>();
            var features = new List<double[]>();
            var skipped = 0;
            for (var r = 1; r < matrix.GetLength(0); r++)
            {
                var yCell = ExcelErrorValues.Normalize(matrix[r, targetCol]);
                if (yCell == null || !double.TryParse(Convert.ToString(yCell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var yNum))
                {
                    skipped++;
                    continue;
                }

                var vec = new double[featureCols.Count];
                var ok = true;
                for (var i = 0; i < featureCols.Count; i++)
                {
                    var cell = ExcelErrorValues.Normalize(matrix[r, featureCols[i]]);
                    if (cell == null || !double.TryParse(Convert.ToString(cell).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
                    {
                        ok = false;
                        break;
                    }

                    vec[i] = number;
                }

                if (ok)
                {
                    yValues.Add(yNum);
                    features.Add(vec);
                }
                else
                {
                    skipped++;
                }
            }

            RegressionAnalyzer.RegressionResult rr;
            try
            {
                rr = RegressionAnalyzer.Compute(yValues, features, featureNames);
            }
            catch (ArgumentException ex)
            {
                return Invalid("REGRESSION_FAILED", ex.Message);
            }

            // 诊断明细上限：≤1000 行全量返回，超出省略（防返回体爆炸）。
            var diagnostics = rr.Observations
                .Select(o => new { data_row = o.Row, actual = R4(o.Actual), fitted = R4(o.Fitted), residual = R4(o.Residual), cooks_d = R6(o.CooksD) })
                .Take(1000)
                .ToArray();

            var ySpan = AnalysisBasis.ColumnSpan(read.Sheet, read.Address, targetCol + 1);
            var xSpans = string.Join(",", featureCols.Select(c => AnalysisBasis.ColumnSpan(read.Sheet, read.Address, c + 1)));

            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    method = "ols",
                    n = rr.TotalObs,
                    df_model = rr.DfModel,
                    df_residual = rr.DfResidual,
                    r2 = R4(rr.R2),
                    adj_r2 = R4(rr.AdjR2),
                    f_stat = R4(rr.FStat),
                    f_p = R6(rr.FPValue),
                    residual_std_error = R4(rr.ResidualStdError),
                    durbin_watson = double.IsNaN(rr.DurbinWatson) ? null : (object)R4(rr.DurbinWatson),
                    target = new { column = targetCol + 1, name = targetName },
                    feature_columns = featureCols.Select((c, i) => new { column = c + 1, name = featureNames[i] }).ToArray(),
                    coefficients = rr.Coefficients.Select(c => new
                    {
                        name = c.Name,
                        estimate = R4(c.Estimate),
                        std_error = R4(c.StdError),
                        t_stat = R4(c.TStat),
                        p_value = R6(c.PValue)
                    }).ToArray(),
                    vif = rr.Vifs.Select(v => new { name = v.Name, vif = R4(v.Vif) }).ToArray(),
                    skipped_rows = skipped,
                    diagnostics = diagnostics,
                    diagnostics_omitted = rr.Observations.Count > diagnostics.Length,
                    basis = new Dictionary<string, object>
                    {
                        ["source"] = AnalysisBasis.Source(read.Sheet, read.Address),
                        ["sample_rule"] = $"因变量「{targetName}」，特征列「{string.Join("/", featureNames)}」，完整观测 {rr.TotalObs} 行（含缺失/非数值行 {skipped} 行剔除）",
                        ["method_text"] = "普通最小二乘（OLS）正规方程求解；系数显著性为双侧 t 检验，模型整体显著性为 F 检验（统计口径纯 C# 可复算）",
                        ["formula_examples"] = new[]
                        {
                            $"LINEST({ySpan},({xSpans}),TRUE,TRUE)（数组公式，返回回归统计量）",
                            $"TREND({ySpan},({xSpans}),新X区域)（按回归模型预测）",
                            "RSQ(" + ySpan + "," + xSpans.Split(',')[0] + ")（单特征 R²）"
                        },
                        ["boundary_text"] = "R² 高不代表因果关系；Durbin-Watson≈2 提示残差无自相关，VIF>10 提示多重共线性，Cook's D 大（如 >4/n）的行对系数影响大；非线性/逻辑回归/正则化走 python-advanced-analysis 技能"
                    }
                }
            };
        }

        // ---- 文本列变体聚类建议标准工具（FR-K2 探索式清洗；纯逻辑在 Core/Tools/TextVariantClusterer）----
        public ToolResult TextVariants(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var column = GetInt(args, "column", 1);
            var maxDistance = Math.Max(1, Math.Min(GetInt(args, "max_distance", 2), 4));

            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(0) > 0 ? matrix.GetLength(1) : 0;
            if (column < 1 || column > columns)
            {
                return Invalid("INVALID_COLUMN", $"column={column} 超出范围（1~{columns}）。");
            }

            // 逐行取文本列原值（跳过表头行；错误值归一后计入，空白计入 empty 不参与聚类）。
            var values = new List<string>();
            for (var r = 1; r < matrix.GetLength(0); r++)
            {
                var cell = ExcelErrorValues.Normalize(matrix[r, column - 1]);
                values.Add(cell == null ? null : Convert.ToString(cell));
            }

            TextVariantClusterer.ClusterResult cr;
            try
            {
                cr = TextVariantClusterer.Compute(values, maxDistance);
            }
            catch (ArgumentException ex)
            {
                return Invalid("TOO_MANY_UNIQUE_VALUES", ex.Message);
            }

            var colName = Header(matrix, column - 1);
            return new ToolResult
            {
                Ok = true,
                Truncated = read.Truncated,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    column,
                    name = colName,
                    total_rows = cr.TotalRows,
                    unique_values = cr.UniqueValues,
                    empty_count = cr.EmptyCount,
                    max_distance = maxDistance,
                    group_count = cr.Groups.Count,
                    variant_value_count = cr.Groups.Sum(g => g.Members.Count),
                    groups = cr.Groups.Take(50).Select(g => new
                    {
                        representative = g.Representative,
                        total_count = g.TotalCount,
                        members = g.Members.Select(m => new { value = m.Value, count = m.Count, match = m.MatchKind }).ToArray(),
                        merge_options = new { longest = g.Longest, shortest = g.Shortest, most_frequent = g.MostFrequent }
                    }).ToArray(),
                    groups_omitted = cr.Groups.Count > 50,
                    basis = new Dictionary<string, object>
                    {
                        ["source"] = AnalysisBasis.Source(read.Sheet, read.Address),
                        ["sample_rule"] = $"文本列「{colName}」共 {cr.TotalRows} 行（空白 {cr.EmptyCount} 行不计入），唯一值 {cr.UniqueValues} 个",
                        ["method_text"] = $"归一化（去空白/全角转半角/大小写折叠）后贪心聚类：归一化相等=exact，编辑距离≤{maxDistance} 且共享首字符=fuzzy（LCP≥1 防短中文串误合，算法经沙箱真值组召回验证）",
                        ["boundary_text"] = "本工具只建议不合并；确认后的合并动作须用 write_excel_cells_batch 将变体单元格改写为代表值（改值语义，不得用 dedup_range——那是删行），走确认门可回滚；指纹算法仅对拉丁字符有效不做中文承诺；唯一值超 1500 或需模糊匹配进阶（rapidfuzz）走 python-advanced-analysis 技能"
                    }
                }
            };
        }

        // ---- 看板洞察卡标准工具（FR 蓝图批7：AI 结论可视化进看板；布局纯逻辑在 Core/Tools/InsightCardLayout）----
        public ToolResult AddInsightCard(Dictionary<string, object> args)
        {
            var title = Required(args, "title");
            var body = Required(args, "body");
            var cellAnchor = GetString(args, "cell", null);
            var sheetName = GetString(args, "sheet", null);
            var colorName = GetString(args, "color", "blue");
            var widthPt = Convert.ToSingle(GetInt(args, "width_pt", (int)InsightCardLayout.DefaultWidthPt));
            var replace = args.TryGetValue("replace_existing", out var rObj) &&
                          Convert.ToString(rObj is Newtonsoft.Json.Linq.JValue jv ? jv.Value : rObj).Equals("true", StringComparison.OrdinalIgnoreCase);

            // 布局纯逻辑：自动高度/截断
            var layout = InsightCardLayout.Estimate(title, body, widthPt);
            var bodyText = InsightCardLayout.TruncateBody(body, widthPt, InsightCardLayout.MaxBodyLines);

            dynamic sheet;
            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? application.ActiveSheet
                    : application.ActiveWorkbook.Worksheets.Item(sheetName);
            }
            catch (Exception xaiEx)
            {
                return Invalid("INVALID_SHEET", $"定位目标工作表失败（{(sheetName ?? "活动表")}）：{xaiEx.Message}");
            }

            // 幂等替换：清掉本表既有洞察卡（xai_insight_ 前缀）防反复调用堆积
            if (replace)
            {
                var doomed = new List<dynamic>();
                foreach (var shp in sheet.Shapes)
                {
                    if (Convert.ToString(shp.Name).StartsWith("xai_insight_", StringComparison.Ordinal))
                    {
                        doomed.Add(shp);
                    }
                }

                foreach (var shp in doomed)
                {
                    shp.Delete();
                }
            }

            try
            {
                dynamic anchor = string.IsNullOrEmpty(cellAnchor)
                    ? sheet.Range("A1")
                    : sheet.Range(cellAnchor);
                var left = Convert.ToSingle(anchor.Left);
                var top = Convert.ToSingle(anchor.Top);
                var rgb = InsightCardLayout.ColorBgr(colorName);
                var shapeName = "xai_insight_" + DateTime.Now.ToString("HHmmssfff");

                dynamic shape = sheet.Shapes.AddShape(5, left, top, layout.WidthPt, layout.HeightPt); // 5 = msoShapeRoundedRectangle
                shape.Name = shapeName;
                shape.Fill.ForeColor.RGB = rgb;
                shape.Line.Visible = 0;
                shape.Shadow.Visible = 0;
                dynamic tf = shape.TextFrame2;
                tf.WordWrap = -1; // msoTrue
                tf.MarginLeft = 8f;
                tf.MarginRight = 8f;
                tf.MarginTop = 6f;
                tf.MarginBottom = 6f;
                tf.TextRange.Text = title + "\n" + bodyText;
                tf.TextRange.Font.Size = InsightCardLayout.BodyFontSize;
                tf.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF; // 白字
                try
                {
                    tf.TextRange.Paragraphs(1).Font.Size = InsightCardLayout.TitleFontSize;
                    tf.TextRange.Paragraphs(1).Font.Bold = -1; // msoTrue
                }
                catch (Exception xaiEx)
                {
                    AppLog.Warn("insight_card", "标题段加粗降级（整卡同字号不影响交付）：" + xaiEx.Message);
                }

                // 内嵌读回自检：文本真实落形状才算成功
                var readBack = Convert.ToString(tf.TextRange.Text) ?? string.Empty;
                var ok = readBack.Contains(title) && shape.Name.StartsWith("xai_insight_", StringComparison.Ordinal);
                if (!ok)
                {
                    shape.Delete();
                    return Invalid("INSIGHT_CARD_FAILED", "洞察卡写入后读回校验未命中标题，已删除半成品；请重试或缩小文字量。");
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = string.IsNullOrEmpty(sheetName) ? Convert.ToString(sheet.Name) : sheetName,
                        shape_name = shapeName,
                        cell = string.IsNullOrEmpty(cellAnchor) ? "A1" : cellAnchor,
                        verdict = "ok",
                        width_pt = (int)layout.WidthPt,
                        height_pt = (int)layout.HeightPt,
                        body_lines = layout.BodyLines,
                        body_truncated = layout.Truncated,
                        basis = new Dictionary<string, object>
                        {
                            ["source"] = $"洞察卡形状 {shapeName}（{Convert.ToString(sheet.Name)}!{cellAnchor ?? "A1"} 锚点）",
                            ["method_text"] = $"圆角矩形卡片：标题 {InsightCardLayout.TitleFontSize}pt 加粗 + 正文 {InsightCardLayout.BodyFontSize}pt 白字，卡高按文字量自动估算（正文超 {InsightCardLayout.MaxBodyLines} 行截断省略），写入后读回文本校验命中标题",
                            ["boundary_text"] = "洞察卡是展示形状，不参与单元格快照回滚；删除可重跑本工具带 replace_existing=true 清空重写，或随 delete_dashboard 删表移除。卡片文字必须来自工具真实分析结论，禁止编造数值"
                        }
                    }
                };
            }
            catch (Exception xaiEx)
            {
                AppLog.Warn("insight_card", "洞察卡创建失败：" + xaiEx.Message);
                return Invalid("INSIGHT_CARD_FAILED", $"洞察卡创建失败：{xaiEx.Message}（目标表可能处于保护态或选中了不可用锚点）");
            }
        }

        public ToolResult RecommendChart(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var matrix = read.Matrix;
            var columns = matrix.GetLength(1);
            var rows = Math.Max(matrix.GetLength(0) - 1, 0);
            var recommendation = Classify(matrix);

            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    sheet = read.Sheet,
                    address = read.Address,
                    rows,
                    columns,
                    numeric_column_count = recommendation.Numeric.Count,
                    text_column_count = recommendation.Text.Count,
                    date_column_count = recommendation.Date.Count,
                    recommended_chart_type = recommendation.ChartType,
                    x_column = recommendation.ChartType == "null" ? (int?)null : recommendation.XColumn + 1,
                    series_columns = recommendation.SeriesFeature().Select(c => c + 1).ToArray(),
                    rationale = recommendation.Rationale
                }
            };
        }

        /// <summary>
        /// C1：推荐图表一键插入。内部复用 Classify 的选型，自动定坐标列与锚点，插入中文标题图表。
        /// </summary>
        public ToolResult InsertRecommendedChart(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var read = ReadMatrix(args, address);
            if (!string.IsNullOrEmpty(read.Error))
            {
                return Invalid("TOO_MANY_CELLS", read.Error);
            }

            var recommendation = Classify(read.Matrix);
            if (recommendation.ChartType == "null")
            {
                return Invalid("NO_CHART", recommendation.Rationale);
            }

            var title = Optional(args, "title");
            if (string.IsNullOrWhiteSpace(title))
            {
                title = read.Sheet + " 分析";
            }

            var anchor = Optional(args, "anchor");
            var width = GetInt(args, "width", 420);
            var height = GetInt(args, "height", 260);

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic source = null;
            dynamic anchorRange = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                // v3.7.1.44 RCW 纪律：workbook/sheet 为共享身份 RCW，只赋值使用，不 FinalRelease（见 finally）
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];

                source = sheet.Range[address];
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

                var chartType = ChartTypeMapper.Resolve(recommendation.ChartType);
                chartObject = sheet.ChartObjects().Add(left, top, width, height);
                chart = chartObject.Chart;
                chart.ChartType = ChartTypeMapper.ResolveXlChartType(recommendation.ChartType);
                chart.SetSourceData(source);
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;

                var sheetNameForRecord = Convert.ToString(sheet.Name);
                var chartNameForRecord = Convert.ToString(chartObject.Name);
                var sourceAddressForRecord = Convert.ToString(source.Address);
                recorder?.Invoke(
                    "insert_recommended_chart",
                    args,
                    "chart-insert",
                    new
                    {
                        sheet = sheetNameForRecord,
                        chart_name = chartNameForRecord,
                        source_address = sourceAddressForRecord,
                        chart_type = recommendation.ChartType
                    },
                    new ChartInsertPayload { Sheet = sheetNameForRecord, ChartName = chartNameForRecord });

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        chart_name = chartNameForRecord,
                        chart_type = recommendation.ChartType,
                        chart_type_resolved = chartType,
                        anchor = string.IsNullOrWhiteSpace(anchor) ? "source-right" : anchor,
                        source_address = sourceAddressForRecord,
                        series_columns = recommendation.SeriesFeature().Select(c => c + 1).ToArray(),
                        category_column = recommendation.XColumn + 1,
                        title
                    }
                };
            }
            finally
            {
                // v3.7.1.44 RCW 纪律修复：chart/chartObject/anchorRange/source 为本方法自建 RCW，照常释放；
                // sheet/workbook 是共享身份（ActiveWorkbook/ActiveSheet/Worksheets.Item），FinalRelease 会减断
                // 共享 RCW，后续工具访问同一 sheet 即抛 InvalidComObjectException（真机 U3b 引爆 U8 实证）。
                Release((object)chart);
                Release((object)chartObject);
                Release((object)anchorRange);
                Release((object)source);
            }
        }

        /// <summary>
        /// 依据数据形状给出图表选型。供 recommend_excel_chart 与 insert_recommended_chart 复用。
        /// </summary>
        private static ChartRecommendation Classify(object[,] matrix)
        {
            var columns = matrix.GetLength(1);
            var numeric = new List<int>();
            var text = new List<int>();
            var date = new List<int>();
            for (var column = 0; column < columns; column++)
            {
                var numbers = 0;
                var dates = 0;
                var texts = 0;
                var nonNull = 0;
                for (var row = 1; row < matrix.GetLength(0); row++)
                {
                    var cell = matrix[row, column];
                    if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell)))
                    {
                        continue;
                    }

                    nonNull++;
                    var value = Convert.ToString(cell).Trim();
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    {
                        numbers++;
                    }
                    else if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                    {
                        dates++;
                    }
                    else
                    {
                        texts++;
                    }
                }

                if (nonNull == 0)
                {
                    continue;
                }

                if (numbers >= nonNull * 0.8)
                {
                    numeric.Add(column);
                }
                else if (dates >= nonNull * 0.8)
                {
                    date.Add(column);
                }
                else
                {
                    text.Add(column);
                }
            }

            int? xColumn;
            string chartType;
            string rationale;
            if (numeric.Count == 0)
            {
                chartType = "null";
                rationale = "没有数值列，无法直接生成图表；建议先汇总或转数值。";
                xColumn = null;
            }
            else if (text.Count > 0 && numeric.Count <= 1)
            {
                chartType = "Pie";
                rationale = "一个分类列 + 单个数值列，适合占比分析（饼图）。";
                xColumn = text[0];
            }
            else if (text.Count > 0)
            {
                chartType = "ColumnClustered";
                rationale = "分类列 + 多个数值列，适合分类对比（簇状柱形图）。";
                xColumn = text[0];
            }
            else if (date.Count > 0)
            {
                chartType = "Line";
                rationale = "含日期列，适合趋势分析（折线图）。";
                xColumn = date[0];
            }
            else
            {
                chartType = "Line";
                rationale = "多数值列且无明确分类，建议先用行号作为 X 轴查看趋势（折线图）。";
                xColumn = null;
            }

            return new ChartRecommendation
            {
                Numeric = numeric,
                Text = text,
                Date = date,
                ChartType = chartType,
                XColumn = xColumn ?? (numeric.Count > 0 ? numeric[0] : 0),
                Rationale = rationale
            };
        }

        private sealed class ChartRecommendation
        {
            public List<int> Numeric { get; set; } = new List<int>();
            public List<int> Text { get; set; } = new List<int>();
            public List<int> Date { get; set; } = new List<int>();
            public string ChartType { get; set; }
            public int XColumn { get; set; }
            public string Rationale { get; set; }

            public IEnumerable<int> SeriesFeature()
            {
                return Numeric.Count > 0 ? Numeric.Take(6) : new int[0];
            }
        }

        private static double ComputeCorrelation(List<double> a, List<double> b, string method)
        {
            var pairs = new List<Tuple<double, double>>();
            var count = Math.Min(a.Count, b.Count);
            for (var i = 0; i < count; i++)
            {
                pairs.Add(Tuple.Create(a[i], b[i]));
            }

            if (pairs.Count < 2)
            {
                return 0.0;
            }

            var x = pairs.Select(p => p.Item1).ToArray();
            var y = pairs.Select(p => p.Item2).ToArray();
            if (string.Equals(method, "rank", StringComparison.OrdinalIgnoreCase))
            {
                x = Rank(x);
                y = Rank(y);
            }

            var xMean = x.Average();
            var yMean = y.Average();
            var numerator = 0.0;
            var denomX = 0.0;
            var denomY = 0.0;
            for (var i = 0; i < x.Length; i++)
            {
                var dx = x[i] - xMean;
                var dy = y[i] - yMean;
                numerator += dx * dy;
                denomX += dx * dx;
                denomY += dy * dy;
            }

            var denom = Math.Sqrt(denomX * denomY);
            return denom == 0 ? 0.0 : numerator / denom;
        }

        private static double[] Rank(double[] values)
        {
            var indexed = values.Select((value, index) => new { value, index }).ToArray();
            var sorted = indexed.OrderBy(v => v.value).ToArray();
            var ranks = new double[values.Length];
            var i = 0;
            while (i < sorted.Length)
            {
                var j = i;
                while (j + 1 < sorted.Length && Math.Abs(sorted[j].value - sorted[j + 1].value) < 1e-12)
                {
                    j++;
                }

                var average = (i + j) / 2.0 + 1.0;
                for (var k = i; k <= j; k++)
                {
                    ranks[sorted[k].index] = average;
                }

                i = j + 1;
            }

            return ranks;
        }

        private ReadResult ReadMatrix(Dictionary<string, object> args, string address)
        {
            // v3.7.1.44：workbook/sheet 为共享身份 RCW，本方法不再持有/释放（见 finally 注释）
            dynamic range = null;
            dynamic value = null;
            try
            {
                var workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                range = sheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                // 全量读入类工具护栏：按 rows×cols 总单元格显式拒绝，勿只标 truncated 而仍拉取整表——
                // 否则整张 UsedRange 会被拉成 17 亿元素矩阵 OOM。int 相乘会溢出（1048576×16384 溢出为负/0
                // 令 truncated 误判 false），必须用 long 乘。
                var cellCount = (long)rows * (long)columns;
                if (cellCount > MaxAnalyzedCells)
                {
                    return new ReadResult
                    {
                        Sheet = Convert.ToString(sheet.Name),
                        Address = Convert.ToString(range.Address),
                        Error = $"数据区域过大（{rows} 行 × {columns} 列，{cellCount} 个单元），超过可分析上限 {MaxAnalyzedCells}，请缩小区域或分批处理。"
                    };
                }

                value = range.Value2;
                var matrix = FlattenValues(value, rows, columns);
                return new ReadResult
                {
                    Sheet = Convert.ToString(sheet.Name),
                    Address = Convert.ToString(range.Address),
                    Matrix = matrix,
                    Truncated = false
                };
            }
            finally
            {
                // v3.7.1.44 RCW 纪律修复：sheet/workbook 是共享身份 RCW（ActiveWorkbook/ActiveSheet/
                // Worksheets.Item 与宿主及后续工具共享同一实例），FinalRelease 会把它减断，
                // 后续任何工具访问同一 sheet 即抛 InvalidComObjectException（真机 U3b 引爆 U8 实证）。
                // 只释放本方法自建的 range/value（每次新 RCW，安全）；与 PivotChartService 枚举纪律同源。
                Release((object)value);
                Release((object)range);
            }
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

        private static double R4(double value)
        {
            return Math.Round(value, 4);
        }

        private static double R6(double value)
        {
            return Math.Round(value, 6);
        }

        private static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source == null || rows == 0 || columns == 0)
            {
                return result;
            }

            if (source is object[,] array)
            {
                var rBound = Math.Min(rows, array.GetLength(0));
                var cBound = Math.Min(columns, array.GetLength(1));
                for (var row = 0; row < rBound; row++)
                {
                    for (var column = 0; column < cBound; column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }
            }
            else if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static double GetDouble(Dictionary<string, object> args, string key, double defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToDouble(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value, CultureInfo.InvariantCulture);
        }

        private static string GetString(Dictionary<string, object> args, string key, string defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToString(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
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

        private sealed class ReadResult
        {
            public string Sheet { get; set; }
            public string Address { get; set; }
            public object[,] Matrix { get; set; }
            public bool Truncated { get; set; }
            public string Error { get; set; }
        }
    }
}