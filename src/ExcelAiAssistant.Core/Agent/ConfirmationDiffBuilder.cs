using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Agent
{
    public static class ConfirmationDiffBuilder
    {
        public static string Build(
            string tool,
            System.Collections.Generic.Dictionary<string, object> args,
            Core.Models.ToolResult preview)
        {
            var address = args != null && args.TryGetValue("address", out var value)
                ? Convert.ToString(value)
                : "当前选区";
            var rows = (preview?.Data as Newtonsoft.Json.Linq.JObject)?["rows"]?.ToObject<int?>();
            var columns = (preview?.Data as Newtonsoft.Json.Linq.JObject)?["columns"]?.ToObject<int?>();

            if (rows.HasValue && columns.HasValue)
            {
                return $"工具 {tool} 将修改 {address}（{rows} 行 × {columns} 列）。";
            }

            return $"工具 {tool} 将修改 {address}。";
        }
    }
}
