using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>数据智能填充（P6-C'）：ai_fill_blank_cells 只读分析 + apply_fill_blank_cells 受控写入。</summary>
    public static class ExcelAiFillToolsRegistration
    {
        public static void RegisterAiFillTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer)
        {
            var service = new ExcelAiFillService(
                (object)application,
                writer,
                isControlledEnabled: () => new ControlledExecutionStore().Load().Enabled);

            registry.Register(new ToolDescriptor(
                "ai_fill_blank_cells",
                "只读分析一个区域（含表头+数据）里的空白单元格：返回空缺清单、按列分组的表头与非空样本、以及可参考的其他数据样本，不写入任何内容。供模型做语义推理后逐格填出建议值。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.AnalyzeBlanks(args)),
                ExcelAiFillToolSchemas.AnalyzeBlanks()));

            registry.Register(new ToolDescriptor(
                "apply_fill_blank_cells",
                "把确认后的建议值就地回填到指定区域的空白单元格（非空白格一律不动，会被拒绝），整块写回并记录可撤销。需开启受控执行总开关，否则拒绝。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.ApplyFills(args)),
                ExcelAiFillToolSchemas.ApplyFills()));
        }
    }
}