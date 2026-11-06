using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>标记式清洗工具（P6-B）参数 Schema。</summary>
    internal static class ExcelFlagToolSchemas
    {
        public static Dictionary<string, object> FlagCells()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要扫描标记的区域地址，如 A1:D200。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("flag_types", "array", "要检测的类型数组：outlier(异常值)/format(格式异常-文本数字与坏日期)/empty(空值)。缺省全部检测。"),
                Property("fill_colors", "object", "各类型高亮色映射（type→\"#RRGGBB\"），缺省用内置色板。"));
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
    }
}