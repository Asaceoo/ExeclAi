using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>单一异常标记：Row/Col 均从 1 计；Type ∈ outlier/format/empty。</summary>
    public sealed class CellFlag
    {
        public int Row;
        public int Col;
        public string Type;
        public string Value;
        public string Reason;
        public string Suggestion;
    }

    /// <summary>纯逻辑异常/格式/空值检测，不接触 COM，供标记式清洗阶段计划用。</summary>
    public static class OutlierDetector
    {
        private const double IqrFactor = 1.5;
        private const double ZThreshold = 3.0;

        public static List<int> DetectOutliers(double[] values) => DetectOutliers(values, "iqr");

        /// <summary>对数值列返回异常值的行索引（1 起）。method: iqr|zscore|auto。</summary>
        public static List<int> DetectOutliers(double[] values, string method)
        {
            var result = new List<int>();
            if (values == null || values.Length < 3) return result;

            var valid = new List<int>();
            for (int i = 0; i < values.Length; i++)
            {
                double v = values[i];
                if (!double.IsNaN(v) && !double.IsInfinity(v)) valid.Add(i);
            }
            if (valid.Count < 3) return result;

            string m = (method ?? "auto").ToLowerInvariant();
            if (m == "auto") m = valid.Count < 30 ? "iqr" : "zscore";

            if (m == "zscore")
            {
                double[] sv = valid.Select(i => values[i]).ToArray();
                double mean = sv.Average();
                double variance = sv.Sum(v => (v - mean) * (v - mean)) / sv.Length;
                double sd = Math.Sqrt(variance);
                if (sd < 1e-12) return result;
                foreach (int i in valid)
                    if (Math.Abs(values[i] - mean) / sd > ZThreshold) result.Add(i + 1);
                return result;
            }

            // IQR（默认）：分位数基于有序样本，极端值不污染中心位置
            double[] sorted = valid.Select(i => values[i]).OrderBy(v => v).ToArray();
            double q1 = Quantile(sorted, 0.25);
            double q3 = Quantile(sorted, 0.75);
            double iqr = q3 - q1;
            if (Math.Abs(iqr) < 1e-12) return result;
            double lo = q1 - IqrFactor * iqr;
            double hi = q3 + IqrFactor * iqr;
            foreach (int i in valid)
                if (values[i] < lo || values[i] > hi) result.Add(i + 1);
            return result;
        }

        /// <summary>文本列中疑似「文本型数字」的行（1 起）：字符串但可被解析为数字。</summary>
        public static List<int> DetectTextAsNumber(object[,] block, int col)
        {
            var result = new List<int>();
            if (block == null || col < 1 || col > block.GetLength(1)) return result;
            for (int r = 0; r < block.GetLength(0); r++)
            {
                object v = block[r, col - 1];
                string s = v as string;
                if (s != null && string.IsNullOrWhiteSpace(s) == false)
                {
                    double d;
                    if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                        result.Add(r + 1);
                }
            }
            return result;
        }

        /// <summary>日期列中不可解析为日期的行（1 起）。</summary>
        public static List<int> DetectBadDates(object[,] block, int col)
        {
            var result = new List<int>();
            if (block == null || col < 1 || col > block.GetLength(1)) return result;
            for (int r = 0; r < block.GetLength(0); r++)
            {
                object v = block[r, col - 1];
                string s = v as string;
                if (s == null || string.IsNullOrWhiteSpace(s)) continue; // 空值交给 empty 处理
                DateTime dt;
                if (!DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    result.Add(r + 1);
            }
            return result;
        }

        /// <summary>生成按类型分组的标记计划 plan[type]→[CellFlag]。flagTypes 缺省时全部检测。</summary>
        public static Dictionary<string, List<CellFlag>> BuildPlan(object[,] block, string[] flagTypes)
        {
            var plan = new Dictionary<string, List<CellFlag>>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> wanted = flagTypes == null
                ? null
                : new HashSet<string>(flagTypes, StringComparer.OrdinalIgnoreCase);
            if (block == null) return plan;

            int rows = block.GetLength(0), cols = block.GetLength(1);

            if (IsWanted(wanted, "empty") && rows > 0 && cols > 0)
            {
                var empties = new List<CellFlag>();
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        if (IsEmpty(block[r, c]))
                            empties.Add(new CellFlag
                            {
                                Row = r + 1, Col = c + 1, Type = "empty",
                                Value = "", Reason = "空值/空字符串", Suggestion = "核对或补充数据"
                            });
                if (empties.Count > 0) plan["empty"] = empties;
            }

            for (int c = 1; c <= cols; c++)
            {
                string kind = ColumnKind(block, c, rows);
                int[] colRows = Enumerable.Range(0, rows).ToArray();

                if (kind == "numeric" && IsWanted(wanted, "outlier"))
                {
                    var vals = new double[rows];
                    for (int r = 0; r < rows; r++)
                    {
                        double d;
                        vals[r] = TryNumber(block[r, c - 1], out d) ? d : double.NaN;
                    }
                    var outliers = new List<CellFlag>();
                    foreach (int i in DetectOutliers(vals, "iqr"))
                        outliers.Add(new CellFlag { Row = i, Col = c, Type = "outlier", Value = FormatValue(block[i - 1, c - 1]), Reason = "数值偏离邻域分布(IQR)", Suggestion = "核对该值是否笔误" });
                    if (outliers.Count > 0) plan["outlier"] = outliers;
                }
                else if (kind == "date" && IsWanted(wanted, "format"))
                {
                    var bads = new List<CellFlag>();
                    foreach (int i in DetectBadDates(block, c))
                        bads.Add(new CellFlag { Row = i, Col = c, Type = "format", Value = FormatValue(block[i - 1, c - 1]), Reason = "无法解析为日期", Suggestion = "修正为统一日期格式" });
                    if (bads.Count > 0) plan["format"] = bads;
                }
                else if (kind == "text" && IsWanted(wanted, "format"))
                {
                    var tn = new List<CellFlag>();
                    foreach (int i in DetectTextAsNumber(block, c))
                        tn.Add(new CellFlag { Row = i, Col = c, Type = "format", Value = FormatValue(block[i - 1, c - 1]), Reason = "疑似文本型数字", Suggestion = "转成数值以便运算" });
                    if (tn.Count > 0) plan["format"] = tn;
                }
            }
            return plan;
        }

        private static bool IsWanted(HashSet<string> wanted, string type)
        {
            return wanted == null || wanted.Contains(type);
        }

        private static bool IsEmpty(object v)
        {
            if (v == null) return true;
            string s = v as string;
            if (s != null) return s.Length == 0;
            return false;
        }

        private static string ColumnKind(object[,] block, int col, int rows)
        {
            int numeric = 0, date = 0, text = 0, nonEmpty = 0;
            for (int r = 0; r < rows; r++)
            {
                object v = block[r, col - 1];
                if (IsEmpty(v)) continue;
                nonEmpty++;
                string s = v as string;
                if (s == null) { numeric++; continue; }
                DateTime dt;
                double d;
                if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) date++;
                else if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) numeric++;
                else text++;
            }
            if (nonEmpty == 0) return "text";
            double numRatio = (double)numeric / nonEmpty;
            double dateRatio = (double)date / nonEmpty;
            if (numRatio >= 0.7) return "numeric";
            if (dateRatio >= 0.6) return "date";
            return "text";
        }

        private static bool TryNumber(object v, out double d)
        {
            d = 0;
            if (v == null) return false;
            if (v is double || v is float || v is int || v is long || v is decimal)
            {
                d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                return true;
            }
            return double.TryParse(v.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }

        private static string FormatValue(object v)
        {
            return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        private static double Quantile(double[] sorted, double p)
        {
            int n = sorted.Length;
            double pos = (n - 1) * p;
            int lo = (int)Math.Floor(pos);
            int hi = (int)Math.Ceiling(pos);
            if (lo == hi) return sorted[lo];
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
        }
    }
}