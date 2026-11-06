using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>通用数据加工工具（第二批）：文本清洗/拆分/合并、日期标准化、编号补零、批量填充公式、多表合并。</summary>
    public static class ExcelUtilityToolsRegistration
    {
        public static void RegisterUtilityTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer)
        {
            var data = new ExcelUtilityToolsService(application, writer);
            Register(registry, "text_clean",
                "对区域内文本做清洗：去首尾空格、合并连续空白、去控制字符、统一大小写、删除指定字符、查找替换；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.TextClean, ExcelUtilityToolSchemas.TextClean());
            Register(registry, "text_split",
                "把一列文本按分隔符拆分为多列（如姓名拆姓/名、地址拆省市区），可保留或清空源列；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.TextSplit, ExcelUtilityToolSchemas.TextSplit());
            Register(registry, "text_merge",
                "把多列文本按分隔符合并到一列（如姓+名拼全名、多列地址拼完整地址），可跳过空单元格；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.TextMerge, ExcelUtilityToolSchemas.TextMerge());
            Register(registry, "date_normalize",
                "把一列日期统一为标准格式（自动识别或指定输入格式，支持 Excel 序列日期），无法解析的可替换为指定文本；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.DateNormalize, ExcelUtilityToolSchemas.DateNormalize());
            Register(registry, "number_pad",
                "把一列编号补零/加前后缀格式化为定长文本（如 1 → NO-000001-A），适合规范台账编号；写入结果并记录撤销。",
                RiskLevel.NormalWrite, () => data.NumberPad, ExcelUtilityToolSchemas.NumberPad());
            Register(registry, "fill_formula",
                "向区域内批量填充公式模板（用 {row} 表示当前行号自动替换），整块可撤销；适合按行批量套用公式。",
                RiskLevel.NormalWrite, () => data.FillFormula, ExcelUtilityToolSchemas.FillFormula());
            Register(registry, "merge_sheets",
                "把多个工作表的数据纵向合并到目标工作表（可选带表头，只取第一个源表的表头），适合月度/分部门台账汇总。",
                RiskLevel.NormalWrite, () => data.MergeSheets, ExcelUtilityToolSchemas.MergeSheets());
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
