using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class ExcelToolHostTests
    {
        [Fact]
        public void Registry_includes_context_preview_write_and_undo_tools()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            Assert.Contains(registry.Tools, tool => tool.Name == "get_current_excel_context");
            Assert.Contains(registry.Tools, tool => tool.Name == "preview_excel_range");
            Assert.Contains(registry.Tools, tool => tool.Name == "write_excel_cell");
            Assert.Contains(registry.Tools, tool => tool.Name == "write_excel_range");
            Assert.Contains(registry.Tools, tool => tool.Name == "undo_excel_operation");
            Assert.Equal(RiskLevel.NormalWrite, registry.GetRiskLevel("undo_excel_operation"));
        }

        [Fact]
        public void Registry_includes_second_batch_utility_tools()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            foreach (var name in new[]
            {
                "text_clean", "text_split", "text_merge",
                "date_normalize", "number_pad", "fill_formula", "merge_sheets"
            })
            {
                Assert.True(registry.TryGet(name, out var descriptor), $"缺少工具：{name}");
                Assert.Equal(RiskLevel.NormalWrite, descriptor.RiskLevel);
                Assert.NotNull(descriptor.Parameters);
            }
        }

        [Fact]
        public void Workbook_hash_is_stable_and_covers_full_name()
        {
            var first = ExcelToolHost.CreateWorkbookHash(@"C:\book.xlsx");
            var second = ExcelToolHost.CreateWorkbookHash(@"C:\book.xlsx");

            Assert.Equal(first, second);
            Assert.NotEqual(ExcelToolHost.CreateWorkbookHash(@"C:\other.xlsx"), first);
        }

        [Fact]
        public void Normalizes_json_matrix_to_two_dimensional_values()
        {
            var method = typeof(ExcelToolHost).GetMethod(
                "NormalizeValues",
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Static);
            Assert.NotNull(method);

            var input = new JArray
            {
                new JArray { "Region", 123 },
                new JArray { "East", 45 }
            };

            var result = Assert.IsType<object[,]>(method.Invoke(null, new[] { (object)input }));
            Assert.Equal(2, result.GetLength(0));
            Assert.Equal(2, result.GetLength(1));
            Assert.Equal("Region", result[0, 0]);
            Assert.Equal(123L, result[0, 1]);
            Assert.Equal("East", result[1, 0]);
            Assert.Equal(45L, result[1, 1]);
        }

        private static Dictionary<string, object> Args(params (string Key, object Value)[] values)
        {
            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, value) in values)
            {
                args[key] = value;
            }

            return args;
        }
    }
}
