using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// FR-K3 分析结论附计算依据：把分析族工具的数据源/列区间拼成 Excel 公式文本，
    /// 让返回值 basis 里的每个数字都可回表核对。纯逻辑、internal 供单测。
    /// 注意：ReadMatrix 首行恒为表头，数据行从地址起始行 +1 开始（公式示例区间据此跳过表头）。
    /// </summary>
    internal static class AnalysisBasis
    {
        private static readonly Regex SpanRx =
            new Regex(@"^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$", RegexOptions.Compiled);

        /// <summary>人话化数据源位置：'Sheet1'!B2:H21（去 $ 便于阅读）。</summary>
        internal static string Source(string sheet, string comAddress)
        {
            return Quote(sheet) + "!" + (comAddress ?? string.Empty).Replace("$", string.Empty);
        }

        internal static string Quote(string sheet)
        {
            return "'" + (sheet ?? string.Empty).Replace("'", "''") + "'";
        }

        internal static string ColumnLetter(int column1Based)
        {
            var n = column1Based;
            var letters = string.Empty;
            while (n > 0)
            {
                var rem = (n - 1) % 26;
                letters = (char)('A' + rem) + letters;
                n = (n - 1) / 26;
            }

            return letters.Length > 0 ? letters : "A";
        }

        internal static int FromLetters(string letters)
        {
            var n = 0;
            foreach (var ch in letters.ToUpperInvariant())
            {
                n = n * 26 + (ch - 'A' + 1);
            }

            return n;
        }

        internal static bool TryParseSpan(
            string comAddress, out int colStart, out int rowStart, out int colEnd, out int rowEnd)
        {
            colStart = rowStart = colEnd = rowEnd = 0;
            if (string.IsNullOrWhiteSpace(comAddress))
            {
                return false;
            }

            var m = SpanRx.Match(comAddress.Trim());
            if (!m.Success)
            {
                return false;
            }

            colStart = FromLetters(m.Groups[1].Value);
            rowStart = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            if (m.Groups[3].Success)
            {
                colEnd = FromLetters(m.Groups[3].Value);
                rowEnd = int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
            }
            else
            {
                colEnd = colStart;
                rowEnd = rowStart;
            }

            return true;
        }

        /// <summary>单列**数据**区间引用（跳过表头行）：'Sheet1'!$C$3:$C$21。列越界或地址不可解析返回空串。</summary>
        internal static string ColumnSpan(string sheet, string comAddress, int column1Based)
        {
            if (!TryParseSpan(comAddress, out var c1, out var r1, out var c2, out var r2))
            {
                return string.Empty;
            }

            if (column1Based < c1 || column1Based > c2 || r2 <= r1)
            {
                return string.Empty;
            }

            var letter = ColumnLetter(column1Based);
            return Quote(sheet) + "!$" + letter + "$" + (r1 + 1) + ":$" + letter + "$" + r2;
        }

        /// <summary>逐数值列给公式示例（默认 AVERAGE），最多 cap 条；不可解析的列跳过。</summary>
        internal static string[] FormulaExamples(
            string sheet, string comAddress, string func, IEnumerable<int> columns1Based, int cap)
        {
            var list = new List<string>();
            foreach (var col in columns1Based)
            {
                if (list.Count >= Math.Max(cap, 0))
                {
                    break;
                }

                var span = ColumnSpan(sheet, comAddress, col);
                if (span.Length > 0)
                {
                    list.Add(func + "(" + span + ")");
                }
            }

            return list.ToArray();
        }

        /// <summary>明细行号区间（人话）：数据行 = 地址起始行 +1 ～ 结束行。</summary>
        internal static string DetailRows(string comAddress, int totalRows)
        {
            if (totalRows <= 1)
            {
                return "无明细行";
            }

            if (!TryParseSpan(comAddress, out var c1, out var r1, out var c2, out var r2))
            {
                return "共 " + (totalRows - 1) + " 行明细（不含表头）";
            }

            return "第 " + (r1 + 1) + " 行～第 " + (r1 + totalRows - 1) + " 行（共 " + (totalRows - 1) + " 行明细）";
        }

        internal static Dictionary<string, object> StatisticsBasis(
            string sheet, string comAddress, int totalRows, List<int> numericCols1Based)
        {
            return new Dictionary<string, object>
            {
                ["source"] = Source(sheet, comAddress),
                ["header_row"] = "第 1 行为表头，不参与统计",
                ["detail_rows"] = DetailRows(comAddress, totalRows),
                ["numeric_rule"] = "列内非空值 ≥80% 可解析为数字才判为数值列；min/max/mean/median/stdev/sum 只计入非空数值",
                ["error_rule"] = "错误单元格（#N/A 等）归一为错误文本，不参与数值统计",
                ["formula_examples"] = FormulaExamples(sheet, comAddress, "AVERAGE", numericCols1Based, 6)
            };
        }

        /// <summary>成对完整观测：仅取两字典同一行键都存在的观测。修复旧实现按各自非空数组
        /// 下标错位配对的缺陷（两列空值分布不同时相关系数会算错——FR-K3 落地实证）。</summary>
        internal static (List<double> a, List<double> b) PairAligned(
            Dictionary<int, double> a, Dictionary<int, double> b)
        {
            var xs = new List<double>();
            var ys = new List<double>();
            foreach (var kv in a)
            {
                if (b.TryGetValue(kv.Key, out var v))
                {
                    xs.Add(kv.Value);
                    ys.Add(v);
                }
            }

            return (xs, ys);
        }
    }
}
