using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// v3.8.1.76 写入文本保护（审查批次：真机自检实证 18 位身份证被 COM Value2 写入重解析为
    /// 5.787E+17（精度丢失）、日期串被重解析为日期序列号）。纯逻辑判定哪些字符串值必须先锁
    /// 文本格式（NumberFormat="@"）再写，防止 Excel 按"手工输入"语义重解析造成数据损坏。
    ///
    /// 规则（宁窄勿宽，只拦「被解析即损坏」的形态）：
    ///   1) 纯数字字符串长度 ≥ 11——手机号 11 位 / 身份证 18 位 / 长编码；≥ 16 位 double 必丢
    ///      精度（真机实证），≥ 11 位会被科学计数法展示且语义上几乎总是"编码"而非"数值"；
    ///   2) 前导 0 的纯数字字符串（长度 ≥ 2）——前导零是编码语义，解析为数字即丢失。
    ///
    /// 日期串（如 1987-07-04）不拦：Excel 将其存为原生日期与用户手工输入行为一致，属正常
    /// 语义化而非损坏；重解析情况经写入结果 coerced_cells 上报给 AI 复核（ShouldReportCoercion）。
    /// </summary>
    public static class TextPreserveAdvisor
    {
        /// <summary>纯数字字符串触发文本锁的最小长度（手机号位数）。</summary>
        public const int MinDigitLength = 11;

        /// <summary>coerced_cells 只上报长度 ≥ 8 的输入字符串（日期串/长编码），避免普通数字串淹没信噪比。</summary>
        public const int MinCoercionReportLength = 8;

        /// <summary>判定该单元格值是否必须先锁文本格式再写入。</summary>
        public static bool ShouldForceText(object cellValue)
        {
            var s = cellValue as string;
            if (s == null && cellValue is JValue jv)
                s = jv.Value as string;
            if (string.IsNullOrEmpty(s) || s.Length < 2)
                return false;
            foreach (var c in s)
            {
                if (c < '0' || c > '9')
                    return false;
            }
            return s.Length >= MinDigitLength || s[0] == '0';
        }

        /// <summary>
        /// 返回需要锁文本格式的单元格坐标（1 基，相对于 values 数组），供写入层逐格设
        /// NumberFormat="@" 后再整块写值。
        /// </summary>
        public static List<KeyValuePair<int, int>> FindTextLockCells(object[,] values)
        {
            var cells = new List<KeyValuePair<int, int>>();
            if (values == null)
                return cells;
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            for (var r = 1; r <= rows; r++)
            {
                for (var c = 1; c <= cols; c++)
                {
                    if (ShouldForceText(values[r - 1, c - 1]))
                        cells.Add(new KeyValuePair<int, int>(r, c));
                }
            }
            return cells;
        }

        /// <summary>
        /// 写入后回读发现"输入字符串被存为数值/日期"时，判定是否值得上报 coerced_cells。
        /// 只上报长字符串（日期串、长编码形态），普通短数字串被数值化属无歧义等价，不上报防噪声。
        /// </summary>
        public static bool ShouldReportCoercion(object inputValue)
        {
            var s = inputValue as string;
            if (s == null && inputValue is JValue jv)
                s = jv.Value as string;
            return !string.IsNullOrEmpty(s) && s.Length >= MinCoercionReportLength;
        }
    }
}
