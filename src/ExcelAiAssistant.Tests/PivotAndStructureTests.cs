using System;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class PivotTableServiceTests
    {
        [Fact]
        public void Sums_values_by_row_label()
        {
            var values = new object[,]
            {
                { "Region", "Amount" },
                { "East", 10D },
                { "West", 5D },
                { "East", 7D },
                { "", 99D }
            };

            var result = PivotTableService.SumByRow(values, 1, 2);

            Assert.Equal(3, result.SourceRows);
            Assert.Equal(3, result.Values.GetLength(0));
            Assert.Equal("East", result.Values[1, 0]);
            Assert.Equal(17D, result.Values[1, 1]);
        }
    }

    public sealed class ExcelToolRegistryTests
    {
        [Fact]
        public void Registry_includes_m3_phase_two_tools()
        {
            var host = new ExcelToolHost(new object());
            var registry = host.CreateRegistry();

            Assert.Contains(registry.Tools, tool => tool.Name == "clean_excel_data");
            Assert.Contains(registry.Tools, tool => tool.Name == "filter_excel_data");
            Assert.Contains(registry.Tools, tool => tool.Name == "add_excel_worksheet");
            Assert.Contains(registry.Tools, tool => tool.Name == "rename_excel_worksheet");
            Assert.Contains(registry.Tools, tool => tool.Name == "delete_excel_worksheet");
            Assert.Contains(registry.Tools, tool => tool.Name == "create_excel_pivot_summary");
            Assert.Contains(registry.Tools, tool => tool.Name == "create_excel_chart");

            Assert.Equal(Core.Models.RiskLevel.StructuralChange, registry.GetRiskLevel("add_excel_worksheet"));
            Assert.Equal(Core.Models.RiskLevel.Dangerous, registry.GetRiskLevel("delete_excel_worksheet"));
            Assert.Equal(Core.Models.RiskLevel.StructuralChange, registry.GetRiskLevel("create_excel_chart"));
        }
    }
}

