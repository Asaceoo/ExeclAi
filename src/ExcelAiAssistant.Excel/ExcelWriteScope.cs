using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

namespace ExcelAiAssistant.Excel
{
    public sealed class ExcelWriteScope
    {
        private readonly ExcelToolHost host;

        public ExcelWriteScope(ExcelToolHost host)
        {
            this.host = host;
        }

        public ToolResult Write(
            dynamic workbook,
            dynamic sheet,
            dynamic range,
            object[,] values,
            string valueMode,
            Dictionary<string, object> args,
            string tool)
        {
            return host.WriteRangeCore(workbook, sheet, range, values, valueMode, args, tool);
        }
    }
}
