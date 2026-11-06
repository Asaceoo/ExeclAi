using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 列质量画像引擎（M14，对标 Power Query 列质量/分布/剖析单面板）。
    /// 纯静态计算、无 COM 依赖，输入为已读取的二维矩阵（含标题行），可单元测试。
    /// 三联输出：质量（空值率/错误率/唯一值/类型占比）+ 分布（Top N 高频值）+ 剖析（数值统计或文本长度统计）。
    /// </summary>
    public static class ColumnProfileEngine
    {
        private static readonly string[] KnownErrors =
        {
            "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#SPILL!", "#CALC!", "#GETTING_DATA"
        };

        /// <summary>CVErr 映射统一取自 ExcelErrorValues（M16 全量审查：单一来源，全项目共用）。</summary>
        private static IReadOnlyDictionary<int, string> CvErrCodes => ExcelErrorValues.Codes;

        /// <summary>列画像结果。质量/分布/剖析三段字段均以 JSON 友好的标量与数组表达。</summary>
        public sealed class ColumnProfile
        {
            public int TotalRows { get; set; }

            public int BlankCount { get; set; }

            public double BlankRate { get; set; }

            public int ErrorCount { get; set; }

            public double ErrorRate { get; set; }

            public int DistinctCount { get; set; }

            public double DistinctRate { get; set; }

            /// <summary>类型占比（键：number/text/date/bool/error/blank）。</summary>
            public Dictionary<string, int> TypeCounts { get; set; }

            /// <summary>数值剖析（列内存在数值时非 null）。</summary>
            public NumericStats Numbers { get; set; }

            /// <summary>文本剖析（列内存在文本时非 null）。</summary>
            public TextStats Text { get; set; }

            /// <summary>高频值分布（已按次数降序、值升序稳定排序）。</summary>
            public List<ValueCount> TopValues { get; set; }

            /// <summary>质量提示：blank_heavy / has_errors / mixed_types / padded_text / constant。</summary>
            public List<string> QualityFlags { get; set; }
        }

        public sealed class NumericStats
        {
            public int Count { get; set; }

            public double Min { get; set; }

            public double Max { get; set; }

            public double Mean { get; set; }

            /// <summary>中位数（偶数个取中间两数均值）。</summary>
            public double Median { get; set; }

            /// <summary>样本标准差（n-1；样本数 &lt; 2 时为 0）。</summary>
            public double StdDev { get; set; }
        }

        public sealed class TextStats
        {
            public int Count { get; set; }

            public int MinLength { get; set; }

            public int MaxLength { get; set; }

            public double AvgLength { get; set; }

            /// <summary>纯空白（非 null 但全空格）单元格数。</summary>
            public int WhitespaceOnlyCount { get; set; }

            /// <summary>首尾带空格的文本单元格数（清洗候选）。</summary>
            public int PaddedCount { get; set; }
        }

        public sealed class ValueCount
        {
            public string Value { get; set; }

            public int Count { get; set; }
        }

        /// <summary>
        /// 对矩阵第 columnIndex 列（1 基，跳过第 0 行标题）计算画像。
        /// 类型判定顺序：blank → error → bool → date → number → text（bool 先于 number，防 Convert 误吞）。
        /// </summary>
        public static ColumnProfile Compute(object[,] matrix, int columnIndex, int top)
        {
            var profile = new ColumnProfile
            {
                TypeCounts = new Dictionary<string, int>(StringComparer.Ordinal),
                TopValues = new List<ValueCount>(),
                QualityFlags = new List<string>()
            };

            if (matrix == null || matrix.GetLength(0) < 2 ||
                columnIndex < 1 || columnIndex > matrix.GetLength(1))
            {
                return profile;
            }

            var rows = matrix.GetLength(0) - 1;
            profile.TotalRows = rows;

            var numbers = new List<double>();
            var textLengths = new List<int>();
            var distribution = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var paddedCount = 0;
            var whitespaceOnlyCount = 0;
            var typeCounts = profile.TypeCounts;

            for (var row = 1; row < matrix.GetLength(0); row++)
            {
                var cell = matrix[row, columnIndex - 1];
                var kind = Classify(cell, out var text, out var numeric);
                Increment(typeCounts, kind);

                switch (kind)
                {
                    case "blank":
                        break;
                    case "error":
                        profile.ErrorCount++;
                        // 错误码也是列的一个取值（Power Query 同口径计入唯一值与分布），如 "#N/A"×2。
                        distribution.TryGetValue(text, out var errCount);
                        distribution[text] = errCount + 1;
                        break;
                    case "bool":
                        distribution.TryGetValue(text, out var boolCount);
                        distribution[text] = boolCount + 1;
                        break;
                    case "date":
                        distribution.TryGetValue(text, out var dateCount);
                        distribution[text] = dateCount + 1;
                        break;
                    case "number":
                        numbers.Add(numeric);
                        distribution.TryGetValue(numeric.ToString(CultureInfo.InvariantCulture), out var numCount);
                        distribution[numeric.ToString(CultureInfo.InvariantCulture)] = numCount + 1;
                        break;
                    default:
                        textLengths.Add(text.Length);
                        // 填充检测必须基于原始字符串（Classify 内已 Trim，trimmed 后首尾空格信息丢失）。
                        var raw = cell as string;
                        if (raw != null && raw.Trim().Length != raw.Length)
                        {
                            paddedCount++;
                        }

                        distribution.TryGetValue(text, out var textCount);
                        distribution[text] = textCount + 1;
                        break;
                }

                if (kind == "blank" && cell != null)
                {
                    whitespaceOnlyCount++;
                }
            }

            profile.BlankCount = typeCounts.TryGetValue("blank", out var blank) ? blank : 0;
            profile.BlankRate = rows == 0 ? 0 : Math.Round(profile.BlankCount / (double)rows, 4);
            profile.ErrorRate = rows == 0 ? 0 : Math.Round(profile.ErrorCount / (double)rows, 4);

            var distinctValues = distribution.Keys.Count;
            profile.DistinctCount = distinctValues;
            profile.DistinctRate = rows == 0 ? 0 : Math.Round(distinctValues / (double)rows, 4);

            if (numbers.Count > 0)
            {
                profile.Numbers = ComputeNumericStats(numbers);
            }

            if (textLengths.Count > 0)
            {
                profile.Text = new TextStats
                {
                    Count = textLengths.Count,
                    MinLength = textLengths.Min(),
                    MaxLength = textLengths.Max(),
                    AvgLength = Math.Round(textLengths.Average(), 2),
                    WhitespaceOnlyCount = whitespaceOnlyCount,
                    PaddedCount = paddedCount
                };
            }

            profile.TopValues = distribution
                .OrderByDescending(kv => kv.Value).ThenBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, top))
                .Select(kv => new ValueCount { Value = kv.Key, Count = kv.Value })
                .ToList();

            profile.QualityFlags = BuildFlags(profile, numbers.Count, textLengths.Count);
            return profile;
        }

        private static string Classify(object cell, out string text, out double numeric)
        {
            text = null;
            numeric = 0d;
            if (cell == null)
            {
                return "blank";
            }

            if (cell is string s)
            {
                text = s.Trim();
                if (text.Length == 0)
                {
                    return "blank";
                }

                if (KnownErrors.Contains(text, StringComparer.OrdinalIgnoreCase))
                {
                    return "error";
                }

                return "text";
            }

            if (cell is bool)
            {
                text = Convert.ToString(cell);
                return "bool";
            }

            if (cell is DateTime dt)
            {
                text = dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                return "date";
            }

            if (cell is double || cell is float || cell is decimal || cell is int || cell is long || cell is short)
            {
                // CVErr 整数（错误单元格）先于数值判定，映射回错误码文本进分布。
                if (cell is int code && CvErrCodes.TryGetValue(code, out var errorCode))
                {
                    text = errorCode;
                    return "error";
                }

                numeric = Convert.ToDouble(cell, CultureInfo.InvariantCulture);
                return "number";
            }

            text = Convert.ToString(cell, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(text) ? "blank" : "text";
        }

        private static NumericStats ComputeNumericStats(List<double> values)
        {
            var sorted = values.OrderBy(v => v).ToList();
            var count = sorted.Count;
            var mean = sorted.Average();
            double median = count % 2 == 1
                ? sorted[count / 2]
                : (sorted[count / 2 - 1] + sorted[count / 2]) / 2.0;
            var stdDev = 0d;
            if (count > 1)
            {
                var variance = sorted.Sum(v => (v - mean) * (v - mean)) / (count - 1);
                stdDev = Math.Round(Math.Sqrt(variance), 4);
            }

            return new NumericStats
            {
                Count = count,
                Min = sorted[0],
                Max = sorted[count - 1],
                Mean = Math.Round(mean, 4),
                Median = Math.Round(median, 4),
                StdDev = stdDev
            };
        }

        private static List<string> BuildFlags(ColumnProfile profile, int numberCount, int textCount)
        {
            var flags = new List<string>();
            if (profile.TotalRows == 0)
            {
                return flags;
            }

            if (profile.BlankRate >= 0.5)
            {
                flags.Add("blank_heavy");
            }

            if (profile.ErrorCount > 0)
            {
                flags.Add("has_errors");
            }

            if (numberCount > 0 && textCount > 0)
            {
                flags.Add("mixed_types");
            }

            if (profile.Text != null && profile.Text.PaddedCount > 0)
            {
                flags.Add("padded_text");
            }

            if (profile.DistinctCount == 1 && profile.BlankCount == 0)
            {
                flags.Add("constant");
            }

            return flags;
        }

        private static void Increment(Dictionary<string, int> map, string key)
        {
            map.TryGetValue(key, out var count);
            map[key] = count + 1;
        }
    }
}
