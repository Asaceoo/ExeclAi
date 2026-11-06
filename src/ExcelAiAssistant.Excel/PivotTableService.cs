using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    public sealed class PivotSummary
    {
        public string[] RowLabels { get; }

        public string[] ColumnHeaders { get; }

        public object[,] Values { get; }

        public int SourceRows { get; }

        public PivotSummary(string[] rowLabels, string[] columnHeaders, object[,] values, int sourceRows)
        {
            RowLabels = rowLabels;
            ColumnHeaders = columnHeaders;
            Values = values;
            SourceRows = sourceRows;
        }
    }

    public static class PivotTableService
    {
        public static PivotSummary SumByRow(
            object[,] values,
            int labelColumn,
            int valueColumn)
        {
            Validate(values, labelColumn, nameof(labelColumn));
            Validate(values, valueColumn, nameof(valueColumn));

            var labels = new List<string>();
            var totals = new List<double>();
            var count = 0;

            for (var row = 1; row < values.GetLength(0); row++)
            {
                if (IsBlank(values[row, labelColumn - 1]))
                {
                    continue;
                }

                var number = ToDouble(values[row, valueColumn - 1]);
                if (number == null)
                {
                    continue;
                }

                var label = Convert.ToString(values[row, labelColumn - 1]);
                var index = labels.FindIndex(item => string.Equals(item, label, StringComparison.OrdinalIgnoreCase));
                if (index < 0)
                {
                    labels.Add(label);
                    totals.Add(number.Value);
                    index = labels.Count - 1;
                }
                else
                {
                    totals[index] += number.Value;
                }

                count++;
            }

            var result = new object[labels.Count + 1, 2];
            result[0, 0] = "Row Labels";
            result[0, 1] = "Total";
            for (var index = 0; index < labels.Count; index++)
            {
                result[index + 1, 0] = labels[index];
                result[index + 1, 1] = totals[index];
            }

            return new PivotSummary(labels.ToArray(), new[] { "Total" }, result, count);
        }

        private static void Validate(object[,] values, int columnIndex, string parameterName)
        {
            if (columnIndex < 1 || columnIndex > values.GetLength(1))
            {
                throw new ArgumentOutOfRangeException(parameterName);
            }
        }

        private static bool IsBlank(object value)
        {
            return value == null || Convert.ToString(value).Trim().Length == 0;
        }

        private static double? ToDouble(object value)
        {
            try
            {
                return Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
            {
                return null;
            }
        }
    }
}
