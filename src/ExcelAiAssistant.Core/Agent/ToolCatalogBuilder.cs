using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Agent
{
    public static class ToolCatalogBuilder
    {
        public static List<Dictionary<string, object>> Build(
            IReadOnlyCollection<Core.Tools.ToolDescriptor> tools)
        {
            var catalog = new List<Dictionary<string, object>>();
            foreach (var tool in tools)
            {
                catalog.Add(Definition(tool.Name, tool.Description, tool.Parameters));
            }

            return catalog;
        }

        public static Dictionary<string, object> Definition(
            string name,
            string description,
            Dictionary<string, object> parameters)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "function",
                ["function"] = new Dictionary<string, object>
                {
                    ["name"] = name,
                    ["description"] = description ?? string.Empty,
                    ["parameters"] = parameters ?? new Dictionary<string, object>
                    {
                        ["type"] = "object",
                        ["properties"] = new Dictionary<string, object>()
                    }
                }
            };
        }
    }
}
