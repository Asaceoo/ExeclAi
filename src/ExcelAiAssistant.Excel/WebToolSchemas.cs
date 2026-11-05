using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>网页抓取工具（fetch_url）参数 Schema。只读工具。</summary>
    internal static class WebToolSchemas
    {
        public static Dictionary<string, object> FetchUrl()
        {
            return Object(
                Names("url"),
                Property("url", "string", "要抓取的网页 URL，仅支持 http/https，例如 https://example.com/page。"),
                Property("timeout_ms", "integer", "请求超时毫秒，默认 30000，范围 1000-60000。"),
                Property("max_bytes", "integer", "最大读取字节数，默认 262144（256KB），范围 1024-1048576(1MB)，超出部分截断。"),
                Property("sanitize", "boolean", "是否对返回文本脱敏（隐藏邮箱/手机号/身份证），默认 true。"),
                Property("extract", "string", "内容提取方式：html（转为纯文本，默认）或 raw（保留原始文本）。"));
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