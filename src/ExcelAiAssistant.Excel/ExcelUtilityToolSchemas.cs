using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>通用数据加工工具（第二批：文本/日期/编号/公式/多表合并）参数 Schema。</summary>
    internal static class ExcelUtilityToolSchemas
    {
        public static Dictionary<string, object> TextClean()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要清洗的区域地址，如 A1:C40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("trim", "boolean", "去除单元格文本首尾空格，默认 true。"),
                Property("collapse_whitespace", "boolean", "把连续空白合并为单个空格，默认 false。"),
                Property("remove_non_printable", "boolean", "去除控制字符，默认 false。"),
                EnumProperty("case", new[] { "none", "upper", "lower" }),
                Property("remove_chars", "string", "要删除的字符集合（逐个删除），如 \"-（）\"。"),
                Property("find", "string", "要查找替换的文本（可选）。"),
                Property("replace", "string", "替换成的文本，缺省删除匹配文本。"));
        }

        public static Dictionary<string, object> TextSplit()
        {
            return Object(
                Names("address", "column"),
                Property("address", "string", "源区域地址，如 A1:A40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("column", "integer", "要拆分的源列号（从 1 计）。"),
                Property("delimiter", "string", "分隔符，默认英文逗号。"),
                Property("target_column", "integer", "拆分结果起始列号，缺省为源列+1。"),
                Property("keep_source", "boolean", "是否保留源列，默认 true。"));
        }

        public static Dictionary<string, object> TextMerge()
        {
            return Object(
                Names("address", "columns", "target_column"),
                Property("address", "string", "源区域地址，如 A1:C40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("columns", "array", "要合并的列号数组（从 1 计），如 [1,2,3]。"),
                Property("target_column", "integer", "合并结果写入的列号（从 1 计）。"),
                Property("delimiter", "string", "列间分隔符，默认空格。"),
                Property("skip_empty", "boolean", "是否跳过空单元格，默认 true。"));
        }

        public static Dictionary<string, object> DateNormalize()
        {
            return Object(
                Names("address", "column"),
                Property("address", "string", "源区域地址，如 A1:A40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("column", "integer", "要标准化的日期列号（从 1 计）。"),
                Property("input_format", "string", "输入格式（如 yyyy/M/d），缺省自动识别。"),
                Property("output_format", "string", "输出格式，默认 yyyy-MM-dd。"),
                Property("invalid_as", "string", "无法解析时的替代文本，缺省保留原值。"));
        }

        public static Dictionary<string, object> NumberPad()
        {
            return Object(
                Names("address", "column", "width"),
                Property("address", "string", "源区域地址，如 A1:A40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("column", "integer", "要格式化的编号列号（从 1 计）。"),
                Property("width", "integer", "编号总宽度（含补位），如 6 表示 000123。"),
                Property("prefix", "string", "编号前缀，如 \"NO-\"。"),
                Property("suffix", "string", "编号后缀，如 \"-A\"。"),
                Property("pad_char", "string", "补位字符，默认 0。"));
        }

        public static Dictionary<string, object> FillFormula()
        {
            return Object(
                Names("address", "formula"),
                Property("address", "string", "要填充公式的区域地址，如 B2:B40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("formula", "string", "公式模板，用 {row} 表示当前行号，如 \"=A{row}*B{row}\"。"),
                Property("start_row", "integer", "起始行号，缺省按区域首行自动计算。"));
        }

        public static Dictionary<string, object> MergeSheets()
        {
            return Object(
                Names("source_sheets", "target_sheet"),
                Property("source_sheets", "array", "要合并的源工作表名数组，如 [\"1月\",\"2月\"]。"),
                Property("target_sheet", "string", "结果写入的工作表名（不存在则报错）。"),
                Property("include_headers", "boolean", "是否带表头（只取第一个源表的表头），默认 true。"),
                Property("start_row", "integer", "源数据起始行（1 表示含表头），默认 1。"));
        }

        private static string[] Names(params string[] names)
        {
            return names;
        }

        private static Dictionary<string, object> Object(string[] required, params Dictionary<string, object>[] properties)
        {
            var map = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                map[Convert.ToString(property["name"])] = property;
            }

            var schema = new Dictionary<string, object> { ["type"] = "object", ["properties"] = map };
            if (required != null && required.Length > 0)
            {
                schema["required"] = required;
            }

            return schema;
        }

        private static Dictionary<string, object> Property(string name, string type, string description)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }

        private static Dictionary<string, object> EnumProperty(string name, string[] values)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
        }
    }
}
