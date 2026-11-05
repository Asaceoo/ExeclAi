using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>公式解释工具（P6-A）参数 Schema。只读工具。</summary>
    internal static class ExcelExplainToolSchemas
    {
        public static Dictionary<string, object> ExplainFormula()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要解释的公式单元格地址，如 E2。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("max_depth", "integer", "关联公式下钻深度上限（1-6，默认 3）。解释结果默认只给一层关联公式提示，模型可追问后再次调用本工具逐层展开。"));
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