using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelAdvancedToolRegistration
    {
        public static void RegisterAdvancedTools(
            ToolRegistry registry,
            ExcelReadService read,
            ExcelEditService edit,
            ExcelScratchpadService scratchpad,
            Func<Dictionary<string, object>, ToolResult> batchWrite,
            OperationRecorder recorder = null)
        {
            edit.SetRecorder(recorder);
            Register(registry, "get_excel_sheet_list", "列出工作簿中所有工作表的名称、可见性与 UsedRange 概览。", RiskLevel.ReadOnly, () => read.GetSheetList, AdvancedToolSchemas.SheetList());
            Register(registry, "preview_excel_sheet", "预览工作表的 UsedRange 维度、前几行、尾部几行与列类型推断。", RiskLevel.ReadOnly, () => read.PreviewSheet, AdvancedToolSchemas.PreviewSheet());
            Register(registry, "get_excel_range_values_and_formulas", "读取区域的数值、公式与数字格式；超过阈值仅返回摘要。", RiskLevel.ReadOnly, () => read.ReadRange, AdvancedToolSchemas.ReadRange());
            Register(registry, "get_current_selection_values_and_formulas", "读取当前选区的数值、公式与数字格式。", RiskLevel.ReadOnly, () => read.ReadSelection, AdvancedToolSchemas.ReadSelection());
            Register(registry, "select_excel_range", "选中并激活指定区域。", RiskLevel.ReadOnly, () => read.SelectRange, AdvancedToolSchemas.SelectRange());
            Register(registry, "find_excel_cells", "在指定区域查找匹配的值或公式单元格。", RiskLevel.ReadOnly, () => read.FindCells, AdvancedToolSchemas.FindCells());
            Register(registry, "manage_excel_filter", "按列条件过滤（隐藏不匹配行）或清除该过滤；基于行隐藏模拟，不产生原生 AutoFilter。", RiskLevel.NormalWrite, () => read.ManageFilter, AdvancedToolSchemas.Filter());
            Register(registry, "manage_excel_named_range", "查询、创建、更新或删除命名区域。", RiskLevel.NormalWrite, () => edit.ManageNamedRange, AdvancedToolSchemas.NamedRange());
            Register(registry, "write_excel_cells_batch", "向多个不连续单元格批量写入值或公式。", RiskLevel.NormalWrite, () => batchWrite, AdvancedToolSchemas.BatchWrite());
            Register(registry, "append_calculation_scratchpad", "向 AI 计算草稿追加输入、公式与期望输出并回填结果。", RiskLevel.NormalWrite, () => scratchpad.Append, AdvancedToolSchemas.AppendScratchpad());
            Register(registry, "clear_calculation_scratchpad", "清空 AI 计算草稿内容。", RiskLevel.StructuralChange, () => scratchpad.Clear, AdvancedToolSchemas.ClearScratchpad());
            Register(registry, "insert_excel_dimension", "插入行或列。", RiskLevel.StructuralChange, () => edit.InsertDimension, AdvancedToolSchemas.Dimension());
            Register(registry, "delete_excel_dimension", "删除行或列，涉及数据破坏请谨慎使用。", RiskLevel.Dangerous, () => edit.DeleteDimension, AdvancedToolSchemas.Dimension());
            Register(registry, "copy_paste_excel_range", "在工作簿内复制并粘贴区域。", RiskLevel.StructuralChange, () => edit.CopyPasteRange, AdvancedToolSchemas.CopyPaste());
            Register(registry, "create_toc_sheet", "生成包含各工作表链接的工作簿目录页。", RiskLevel.NormalWrite, () => edit.CreateToc, AdvancedToolSchemas.Toc());
        }

        private static void Register(
            ToolRegistry registry,
            string name,
            string description,
            RiskLevel risk,
            Func<Func<Dictionary<string, object>, ToolResult>> target,
            Dictionary<string, object> parameters)
        {
            var execute = target();
            registry.Register(new ToolDescriptor(name, description, risk, args => Execute(() => execute(args)), parameters));
        }

        private static ToolResult Execute(Func<ToolResult> action)
        {
            return ExcelToolHost.Execute(action);
        }
    }
}