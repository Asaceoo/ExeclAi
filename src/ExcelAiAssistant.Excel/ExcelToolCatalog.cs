using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelToolCatalog
    {
        public static ToolDescriptor Create(
            string name,
            string description,
            RiskLevel riskLevel,
            Func<Dictionary<string, object>, ToolResult> execute,
            Dictionary<string, object> parameters = null)
        {
            return new ToolDescriptor(name, description, riskLevel, execute, parameters);
        }

        public static Dictionary<string, object> ArraySchema(string description)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "array",
                ["description"] = description,
                ["items"] = new Dictionary<string, object>
                {
                    ["type"] = "object",
                    ["properties"] = new Dictionary<string, object>
                    {
                        ["tool"] = new Dictionary<string, object>
                        {
                            ["type"] = "string",
                            ["enum"] = new[] { "remove_blank_rows", "fix_dates", "split_column", "fill_blanks" }
                        },
                        ["sheet"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["address"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["column"] = new Dictionary<string, object> { ["type"] = "integer", ["minimum"] = 1 },
                        ["date_format"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["delimiter"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["target_address"] = new Dictionary<string, object> { ["type"] = "string" },
                        ["value"] = new Dictionary<string, object> { ["type"] = "string" }
                    }
                }
            };
        }

        public static Dictionary<string, object> FilterSchema()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["sheet"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["address"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["column"] = new Dictionary<string, object> { ["type"] = "integer", ["minimum"] = 1 },
                    ["operator"] = new Dictionary<string, object>
                    {
                        ["type"] = "string",
                        ["enum"] = new[] { "equals", "not_equals", "contains", "gt", "gte", "lt", "lte", "empty", "not_empty" }
                    },
                    ["value"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["target_address"] = new Dictionary<string, object> { ["type"] = "string" }
                },
                ["required"] = new[] { "address", "column", "operator" }
            };
        }

        public static Dictionary<string, object> FormatSchema()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["sheet"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["address"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["number_format"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["bold"] = new Dictionary<string, object> { ["type"] = "boolean" },
                    ["italic"] = new Dictionary<string, object> { ["type"] = "boolean" },
                    ["font_color"] = new Dictionary<string, object> { ["type"] = "integer" },
                    ["back_color"] = new Dictionary<string, object> { ["type"] = "integer" },
                    ["auto_fit_columns"] = new Dictionary<string, object> { ["type"] = "boolean" }
                },
                ["required"] = new[] { "address" }
            };
        }

        public static Dictionary<string, object> SheetSchema(params string[] required)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["name"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["sheet"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["before_sheet"] = new Dictionary<string, object> { ["type"] = "string" },
                    ["after_sheet"] = new Dictionary<string, object> { ["type"] = "string" }
                },
                ["required"] = required
            };
        }
    }
}


