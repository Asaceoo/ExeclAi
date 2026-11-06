using System;
using System.Globalization;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 文本/序列日期容错解析（纯逻辑，确定性，可单测；XLL UDF M1 提公来源：
    /// ExcelAiAssistant.Excel.ExcelUtilityToolsService 的 internal 静态方法，原实现保持转发以单一来源）。
    /// 无 COM / 无 IO / 无状态——UDF 运行红线（设计-函数库UDF.md §4.2）可安全复用。
    /// </summary>
    public static class TextDateParser
    {
        /// <summary>
        /// 容错解析日期：
        /// - DateTime 直接返回；
        /// - double 序列值仅在「单元格是日期格式或显式 inputFormat」时按 OADate 转换
        ///   （防 ID/年份等纯数字列被误写为乱日期；上限用 OADate 合法最大 9999-12-31≈2958465）；
        /// - 文本按 inputFormat 精确解析，未给格式走 CurrentCulture 宽松解析。
        /// 解析失败返回 null（fail-closed）。
        /// </summary>
        public static DateTime? TryParseDate(object v, string inputFormat, string numberFormat)
        {
            if (v == null)
            {
                return null;
            }

            if (v is DateTime dt)
            {
                return dt;
            }

            if (v is double d && d > 0 && d < 2958466 &&
                (IsDateLikeFormat(numberFormat) || !string.IsNullOrWhiteSpace(inputFormat)))
            {
                // Excel 序列日期：默认只在单元格是「日期格式」时当序列日期转换（避免 ID/年份等纯数字列被误写为
                // 乱日期）；用户显式传 input_format 视为明确意图，允许整列按序列日期转换。
                // 上限用 OADate 合法最大 9999-12-31（约 2958465），不能只到 30000——否则 1982 年后的
                // 日期（OADate 约 4 万+，如 2024≈45298）会被误判为无法解析而拒绝。
                try
                {
                    return DateTime.FromOADate(d);
                }
                catch
                {
                    return null;
                }
            }

            var s = Convert.ToString(v, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(s))
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(inputFormat))
            {
                if (DateTime.TryParseExact(s, inputFormat, CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces, out var exact))
                {
                    return exact;
                }

                return null;
            }

            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed))
            {
                return parsed;
            }

            return null;
        }

        /// <summary>判断某自定义数字格式是否属于「日期/时间」类（含年/月/日/时/分/秒记号）。纯数字/文本格式判定为非日期。</summary>
        public static bool IsDateLikeFormat(string fmt)
        {
            if (string.IsNullOrWhiteSpace(fmt)) return false;
            for (var i = 0; i < fmt.Length; i++)
            {
                var c = fmt[i];
                if (c == 'y' || c == 'm' || c == 'd' || c == 'h' || c == 's' ||
                    c == 'Y' || c == 'M' || c == 'D' || c == 'H' || c == 'S')
                {
                    return true;
                }
            }

            return false;
        }
    }
}
