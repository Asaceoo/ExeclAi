using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>批量落地一批高价值数据加工内置工具（排序/去重/查找替换/转置/条件格式）。</summary>
    public static class ExcelDataToolsRegistration
    {
        public static void RegisterDataTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer)
        {
            var data = new ExcelDataToolsService(application, writer);
            Register(registry, "sort_range",
                "对区域内数据按一列或多列排序（数值/文本智能比较，升/降序），表头保持不动；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.SortRange, ExcelDataToolSchemas.SortRange());
            Register(registry, "dedup_range",
                "对区域内数据去重：按整行或指定列判重，保留首次/末次出现，其余重复行删除并向上压缩，尾部补空；整块可撤销。",
                RiskLevel.NormalWrite, () => data.DedupRange, ExcelDataToolSchemas.DedupRange());
            Register(registry, "find_replace_range",
                "在区域内批量查找并替换文本（可区分大小写、可整格匹配），替换单元格数可控并记录撤销。",
                RiskLevel.NormalWrite, () => data.FindReplaceRange, ExcelDataToolSchemas.FindReplaceRange());
            Register(registry, "transpose_range",
                "把区域行列互换（转置），可选择写入当前选区左上角或指定目标位置。",
                RiskLevel.NormalWrite, () => data.TransposeRange, ExcelDataToolSchemas.TransposeRange());
            Register(registry, "format_excel_conditional",
                "对区域应用条件格式：大于/小于/介于/重复值高亮或数据条，可用 rule=clear 一键清除规则。",
                RiskLevel.NormalWrite, () => data.ConditionalFormat, ExcelDataToolSchemas.ConditionalFormat());
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