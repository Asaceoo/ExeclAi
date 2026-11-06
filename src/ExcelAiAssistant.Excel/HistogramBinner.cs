using System;
using System.Globalization;
using System.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 直方图分箱（纯计算，可单测），供原生 xlHistogram 不可用时的回退柱状图。
    /// </summary>
    public static class HistogramBinner
    {
        public static object[,] BuildBins(double[] values, int binCount)
        {
            if (values == null || values.Length == 0)
            {
                return new object[,] { { "区间", "频数" } };
            }

            var bins = Math.Max(1, binCount);
            var floor = values.Min();
            var ceil = values.Max();
            var span = ceil - floor;
            if (span <= 0)
            {
                return new object[,] { { "区间", "频数" }, { R(floor, 2) + "（单值）", values.Length } };
            }

            var width = span / bins;
            var counts = new int[bins];
            foreach (var value in values)
            {
                var index = (int)((value - floor) / width);
                if (index >= bins) { index = bins - 1; }
                if (index < 0) { index = 0; }
                counts[index]++;
            }

            var result = new object[bins + 1, 2];
            result[0, 0] = "区间";
            result[0, 1] = "频数";
            for (var i = 0; i < bins; i++)
            {
                result[i + 1, 0] = "[" + R(floor + i * width, 2) + ", " + R(floor + (i + 1) * width, 2) + ")";
                result[i + 1, 1] = counts[i];
            }

            return result;
        }

        private static string R(double value, int precision)
        {
            return Math.Round(value, precision).ToString(CultureInfo.InvariantCulture);
        }
    }
}