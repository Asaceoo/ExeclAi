using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>数据加工类新工具（排序/去重/查找替换/转置/条件格式）参数 Schema。</summary>
    internal static class ExcelDataToolSchemas
    {
        public static Dictionary<string, object> SortRange()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要排序的区域地址，如 A1:C40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                IntProperty("sort_column", "主排序列号（从 1 计），默认 1。"),
                EnumProperty("direction", new[] { "asc", "desc" }),
                IntProperty("sort_column2", "次排序列号（可选，从 1 计）。"),
                EnumProperty("direction2", new[] { "asc", "desc" }),
                Property("header_rows", "integer", "视为表头的行数，排序时置顶保持不动，默认 1，无表头传 0。"));
        }

        public static Dictionary<string, object> DedupRange()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要去重的区域地址，如 A1:C50。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("columns", "array", "判重依据列号数组（从 1 计），缺省用整行所有列。"),
                EnumProperty("keep", new[] { "first", "last" }),
                Property("header_rows", "integer", "视为表头的行数，默认 1，无表头传 0。"),
                Property("case_sensitive", "boolean", "文本判重是否区分大小写，默认 false。"));
        }

        public static Dictionary<string, object> FindReplaceRange()
        {
            return Object(
                Names("address", "find"),
                Property("address", "string", "要查找替换的区域地址，如 A1:C50。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("find", "string", "要查找的文本。"),
                Property("replace", "string", "替换成的文本，缺省表示删除匹配文本。"),
                Property("match_case", "boolean", "是否区分大小写，默认 false。"),
                Property("whole_cell", "boolean", "是否仅整格完全相等才替换，默认 false（部分替换）。"));
        }

        public static Dictionary<string, object> TransposeRange()
        {
            return Object(
                Names("address"),
                Property("address", "string", "源区域地址（其左上角作为转置结果写入起点），如 A1:C40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("target_sheet", "string", "结果写入的工作表，缺省与源同表。"),
                Property("target_address", "string", "结果写入的左上角地址，缺省覆盖写回源区域左上角。"));
        }

        public static Dictionary<string, object> ConditionalFormat()
        {
            return Object(
                Names("address", "rule"),
                Property("address", "string", "要设置/清除条件格式的区域地址，如 A1:A50。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                EnumProperty("rule", new[] { "greater", "less", "between", "duplicates", "data_bar", "clear" }),
                Property("value", "number", "数值边界（greater/less 使用）。"),
                Property("value2", "number", "数值上限（between 使用的第二个边界）。"),
                Property("fill_color", "string", "高亮填充色，如 \"#ffd9d9\"（greater/less/between/duplicates 使用）。"),
                Property("text_color", "string", "字体颜色，如 \"#9c0006\"。"),
                Property("bar_color", "string", "数据条颜色，如 \"#63be7b\"（data_bar 使用）。"));
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

        private static Dictionary<string, object> IntProperty(string name, string description)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["minimum"] = 0 };
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