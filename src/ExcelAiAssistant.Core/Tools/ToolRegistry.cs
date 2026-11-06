using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Core.Tools
{
    public sealed class ToolDescriptor
    {
        public string Name { get; }

        public string Description { get; }

        public RiskLevel RiskLevel { get; }

        public Func<Dictionary<string, object>, ToolResult> Execute { get; }

        public Dictionary<string, object> Parameters { get; }

        public ToolDescriptor(
            string name,
            string description,
            RiskLevel riskLevel,
            Func<Dictionary<string, object>, ToolResult> execute,
            Dictionary<string, object> parameters = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("工具名称不能为空", nameof(name));
            }

            Name = name.Trim();
            Description = description ?? string.Empty;
            RiskLevel = riskLevel;
            Execute = execute ?? throw new ArgumentNullException(nameof(execute));
            Parameters = parameters;
        }

        public static object[,] NormalizeValues(object input)
        {
            switch (input)
            {
                case object[,] array:
                    return array;
                case IEnumerable<object> sequence:
                    var items = sequence.Cast<object>().ToList();
                    var scalarOnly = items.All(item => !(item is IEnumerable<object>) || item is string);
                    if (scalarOnly)
                    {
                        var result = new object[items.Count, 1];
                        for (var index = 0; index < items.Count; index++)
                        {
                            result[index, 0] = UnwrapScalar(items[index]);
                        }

                        return result;
                    }

                    var width = items.Count == 0 ? 0 : items.Max(item => ((IEnumerable<object>)item).Count());
                    var matrix = new object[items.Count, width];
                    for (var rowIndex = 0; rowIndex < items.Count; rowIndex++)
                    {
                        var row = ((IEnumerable<object>)items[rowIndex]).ToList();
                        for (var columnIndex = 0; columnIndex < width; columnIndex++)
                        {
                            matrix[rowIndex, columnIndex] =
                                columnIndex < row.Count ? UnwrapScalar(row[columnIndex]) : null;
                        }
                    }

                    return matrix;
                default:
                    var single = new object[1, 1];
                    single[0, 0] = UnwrapScalar(input);
                    return single;
            }
        }

        private static object UnwrapScalar(object value)
        {
            if (value is Newtonsoft.Json.Linq.JValue jValue)
            {
                return jValue.Value;
            }

            if (value is IEnumerable<object> list && !(value is string))
            {
                var first = list.FirstOrDefault();
                return UnwrapScalar(first);
            }

            return value;
        }
    }
}
