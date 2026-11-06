using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// Excel 错误值（CVErr）识别与归一——全项目单一来源（M16 全量审查提炼）。
    /// Excel 错误单元格经 COM Value2 读回是 System.Int32 错误码（M14 真机探针实证：
    /// "#N/A" 写入后宿主解析为错误值，读回 Int32 -2146826246）。不识别会被误判为
    /// number 污染统计、排序、凑数、写回（错误值被静默改写成数字）。
    /// ColumnProfileEngine 及各分析/加工/校验服务统一从此处取映射。
    /// </summary>
    public static class ExcelErrorValues
    {
        /// <summary>CVErr 整数 → 错误文本映射（7 码）。</summary>
        public static readonly IReadOnlyDictionary<int, string> Codes = new Dictionary<int, string>
        {
            [-2146826288] = "#NULL!",
            [-2146826281] = "#DIV/0!",
            [-2146826275] = "#VALUE!",
            [-2146826265] = "#REF!",
            [-2146826259] = "#NAME?",
            [-2146826253] = "#NUM!",
            [-2146826246] = "#N/A"
        };

        /// <summary>是否为 CVErr 整数形态的错误单元格。</summary>
        public static bool IsCveErr(object value)
        {
            return value is int code && Codes.ContainsKey(code);
        }

        /// <summary>归一：CVErr 整数映射为错误文本（如 "#N/A"），其它值原样返回。</summary>
        public static object Normalize(object value)
        {
            if (value is int code && Codes.TryGetValue(code, out var text))
            {
                return text;
            }

            return value;
        }
    }
}
