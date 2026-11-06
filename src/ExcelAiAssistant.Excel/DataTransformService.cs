using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    public static class DataTransformService
    {
        public static TransformOutcome RemoveBlankRows(object[,] values)
        {
            var keptRows = new List<int>();
            for (var row = 0; row < values.GetLength(0); row++)
            {
                if (Enumerable.Range(0, values.GetLength(1)).Any(column => !IsBlank(values[row, column])))
                {
                    keptRows.Add(row);
                }
            }

            return new TransformOutcome(CopySelectedRows(values, keptRows), values.GetLength(0) - keptRows.Count);
        }

        public static TransformOutcome FixDates(object[,] values, int columnIndex, string format = null)
        {
            ValidateColumn(values, columnIndex);
            var changed = 0;
            for (var row = 0; row < values.GetLength(0); row++)
            {
                var parsed = ParseDate(values[row, columnIndex - 1], format);
                if (!Equals(parsed, values[row, columnIndex - 1]))
                {
                    changed++;
                }

                values[row, columnIndex - 1] = parsed;
            }

            return new TransformOutcome(values, changed);
        }

        public static TransformOutcome FillBlanks(object[,] values, int? columnIndex, object fillValue)
        {
            var changed = 0;
            var start = columnIndex.HasValue ? columnIndex.Value - 1 : 0;
            var end = columnIndex.HasValue ? columnIndex.Value : values.GetLength(1);
            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = start; column < end; column++)
                {
                    if (IsBlank(values[row, column]))
                    {
                        values[row, column] = fillValue;
                        changed++;
                    }
                }
            }

            return new TransformOutcome(values, changed);
        }

        /// <summary>去首尾空格：对矩阵全部字符串单元格执行 Trim()，仅统计实际发生变化的格数。</summary>
        public static TransformOutcome TrimWhitespace(object[,] values)
        {
            var changed = 0;
            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    var v = values[row, column];
                    if (v == null || v is DBNull)
                    {
                        continue;
                    }

                    var s = Convert.ToString(v, CultureInfo.InvariantCulture);
                    var trimmed = s.Trim();
                    if (!string.Equals(s, trimmed))
                    {
                        values[row, column] = trimmed;
                        changed++;
                    }
                }
            }

            return new TransformOutcome(values, changed);
        }

        /// <summary>去控制字符：移除单元格文本中全部不可见控制字符（Char.IsControl，不含空格等可见空白）。</summary>
        public static TransformOutcome RemoveNonPrintable(object[,] values)
        {
            var changed = 0;
            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    var v = values[row, column];
                    if (v == null || v is DBNull)
                    {
                        continue;
                    }

                    var s = Convert.ToString(v, CultureInfo.InvariantCulture);
                    var cleaned = new string(s.Where(ch => !char.IsControl(ch)).ToArray());
                    if (!string.Equals(s, cleaned))
                    {
                        values[row, column] = cleaned;
                        changed++;
                    }
                }
            }

            return new TransformOutcome(values, changed);
        }

        public static FilterResult ApplyFilter(
            object[,] values,
            int columnIndex,
            string @operator,
            string value)
        {
            ValidateColumn(values, columnIndex);
            var rows = new List<object[]>();
            for (var row = 0; row < values.GetLength(0); row++)
            {
                var record = Enumerable.Range(0, values.GetLength(1)).Select(column => values[row, column]).ToArray();
                if (Matches(record[columnIndex - 1], @operator, value))
                {
                    rows.Add(record);
                }
            }

            var result = new object[rows.Count, values.GetLength(1)];
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    result[rowIndex, column] = rows[rowIndex][column];
                }
            }

            return new FilterResult(result, rows.Count);
        }

        public static bool Matches(object cell, string @operator, string value)
        {
            var text = Convert.ToString(cell) ?? string.Empty;
            switch (@operator)
            {
                case "equals": return string.Equals(text, value, StringComparison.OrdinalIgnoreCase);
                case "not_equals": return !string.Equals(text, value, StringComparison.OrdinalIgnoreCase);
                case "contains": return text.IndexOf(value ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0;
                case "empty": return IsBlank(cell);
                case "not_empty": return !IsBlank(cell);
                case "gt": case "gte": case "lt": case "lte": return CompareNumbers(cell, value, @operator);
                default: throw new ArgumentException($"不支持的筛选运算符：{@operator}");
            }
        }

        private static bool CompareNumbers(object left, string right, string @operator)
        {
            if (!double.TryParse(right, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var rightNumber))
            {
                return false;
            }

            double leftNumber;
            if (left is IConvertible convertible && !(left is string))
            {
                leftNumber = convertible.ToDouble(System.Globalization.CultureInfo.InvariantCulture);
            }
            else if (!double.TryParse(Convert.ToString(left), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out leftNumber))
            {
                return false;
            }

            return @operator switch
            {
                "gt" => leftNumber > rightNumber,
                "gte" => leftNumber >= rightNumber,
                "lt" => leftNumber < rightNumber,
                "lte" => leftNumber <= rightNumber,
                _ => false
            };
        }

        private static object ParseDate(object value, string format)
        {
            if (value == null || value is DateTime)
            {
                return value;
            }

            var text = Convert.ToString(value)?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return value;
            }

            if (!string.IsNullOrEmpty(format) && DateTime.TryParseExact(
                    text,
                    format,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out var exact))
            {
                return exact;
            }

            if (DateTime.TryParse(text, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out var parsed))
            {
                return parsed;
            }

            return DateTime.TryParse(text, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var invariant)
                ? invariant
                : value;
        }

        private static void ValidateColumn(object[,] values, int columnIndex)
        {
            if (columnIndex < 1 || columnIndex > values.GetLength(1))
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), $"列号必须在 1 到 {values.GetLength(1)} 之间。");
            }
        }

        private static bool IsBlank(object value)
        {
            return value == null || Convert.ToString(value).Trim().Length == 0;
        }

        private static object[,] CopySelectedRows(object[,] values, IReadOnlyCollection<int> rows)
        {
            var result = new object[rows.Count, values.GetLength(1)];
            var outputRow = 0;
            foreach (var sourceRow in rows)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    result[outputRow, column] = values[sourceRow, column];
                }

                outputRow++;
            }

            return result;
        }
    }

    public class TransformOutcome
    {
        public object[,] Values { get; }

        public int ChangedCount { get; }

        public TransformOutcome(object[,] values, int changedCount)
        {
            Values = values;
            ChangedCount = changedCount;
        }
    }

    public sealed class FilterResult : TransformOutcome
    {
        public int MatchedRows { get; }

        public FilterResult(object[,] values, int matchedRows)
            : base(values, 0)
        {
            MatchedRows = matchedRows;
        }
    }
}

