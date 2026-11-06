using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>标记式清洗（P6-B）：识别异常/格式/空值并高亮 + 生成可审计诊断 sheet。</summary>
    public static class ExcelFlagToolsRegistration
    {
        public static void RegisterFlagTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer,
            Action<string, object, string, object, object> recordStructuralUndo)
        {
            var service = new ExcelFlagCellsService(application, writer, recordStructuralUndo);
            registry.Register(new ToolDescriptor(
                "flag_cells",
                "对区域内数据做标记式清洗识别：异常值(IQR/Z-score)、格式异常(文本型数字/坏日期)、空值；在源区用异色高亮，并生成「异常标记诊断」表（含行/列/类型/原因/建议），诊断表可整体删除并恢复。只标记不修改数据。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.FlagCells(args)),
                ExcelFlagToolSchemas.FlagCells()));
        }
    }
}