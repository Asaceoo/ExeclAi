using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>联网搜索工具 JSON Schema（v3.8.1.83）。</summary>
    public static class SearchSchemas
    {
        public static Dictionary<string, object> WebSearch()
        {
            return Object(
                Names("query"),
                Property("query", "string", "搜索关键词。只传关键词（如「制造业 平均良率 2026」），禁止包含工作簿原始数据或敏感信息。"),
                IntProperty("limit", "返回结果条数，默认 5，上限 8。"));
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

            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = map
            };

            if (required != null && required.Length > 0)
            {
                schema["required"] = required;
            }

            return schema;
        }

        private static Dictionary<string, object> Property(string name, string type, string description = null)
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
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "integer",
                ["description"] = description
            };
        }
    }
}
