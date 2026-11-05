using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 轻量趋势预测纯逻辑（analyze_trend_forecast，2026-09-14）：
    /// 线性回归（最小二乘）/移动平均基准水平外推；不接触 COM，供服务层与单测使用。
    /// 季节性 Holt-Winters/ARIMA 属高阶预测，走 python-advanced-analysis 技能（statsmodels），本类不做。
    /// </summary>
    public static class TrendForecaster
    {
        /// <summary>预测是外推，少于 6 个样本点会严重误导，调用方须 fail-closed 拒算。</summary>
        public const int MinPoints = 6;

        public sealed class ForecastResult
        {
            public double Slope { get; set; }
            public double Intercept { get; set; }
            public double R2 { get; set; }
            public string Trend { get; set; }
            public List<double> ForecastValues { get; set; } = new List<double>();
        }

        /// <summary>method: linear | moving_average；periods 为向前预测期数；window 仅 moving_average 使用。</summary>
        public static ForecastResult Compute(IReadOnlyList<double> values, string method, int window, int periods)
        {
            if (values == null || values.Count < MinPoints)
            {
                throw new ArgumentException($"有效数值点不足，趋势预测至少需要 {MinPoints} 个。");
            }

            var result = new ForecastResult();
            if (string.Equals(method, "linear", StringComparison.OrdinalIgnoreCase))
            {
                var n = values.Count;
                var xMean = (n + 1) / 2.0;
                var yMean = values.Average();
                double sxx = 0, sxy = 0, syy = 0;
                for (var i = 0; i < n; i++)
                {
                    sxx += (i + 1 - xMean) * (i + 1 - xMean);
                    sxy += (i + 1 - xMean) * (values[i] - yMean);
                    syy += (values[i] - yMean) * (values[i] - yMean);
                }

                result.Slope = sxx == 0 ? 0 : sxy / sxx;
                result.Intercept = yMean - result.Slope * xMean;
                result.R2 = syy == 0 ? 1.0 : (sxy * sxy) / (sxx * syy);
                result.Trend = result.Slope > 0 ? "上升" : result.Slope < 0 ? "下降" : "平稳";
                for (var k = 1; k <= periods; k++)
                {
                    result.ForecastValues.Add(result.Intercept + result.Slope * (n + k));
                }
            }
            else
            {
                // 移动平均：以最近 window 期均值为基准水平外推（不含趋势项）。
                var w = Math.Max(2, Math.Min(window, values.Count));
                var baseLevel = values.Skip(values.Count - w).Take(w).Average();
                result.Slope = 0;
                result.Intercept = baseLevel;
                result.R2 = 0;
                result.Trend = "水平（移动平均不含趋势项）";
                for (var k = 1; k <= periods; k++)
                {
                    result.ForecastValues.Add(baseLevel);
                }
            }

            return result;
        }

        /// <summary>期数标签延伸：末尾为数字（如 2026-03、第12月、12）则递增；原数字带前导零时保留位宽（2026-03 → 2026-04）。</summary>
        public static string NextLabel(string lastLabel, int k)
        {
            if (string.IsNullOrEmpty(lastLabel))
            {
                return "T+" + k;
            }

            var match = System.Text.RegularExpressions.Regex.Match(lastLabel, "^(.*?)(\\d+)$");
            if (match.Success && int.TryParse(match.Groups[2].Value, out var num))
            {
                var digits = match.Groups[2].Value;
                var next = num + k;
                var text = digits.Length > 1 && digits[0] == '0'
                    ? next.ToString().PadLeft(digits.Length, '0')
                    : next.ToString();
                return match.Groups[1].Value + text;
            }

            return "T+" + k;
        }
    }
}
