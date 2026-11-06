using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>公式解释（P6-A）：真机结构化拆解公式（函数→参数→引用值），只读不产生写操作。</summary>
    public static class ExcelExplainToolsRegistration
    {
        public static void RegisterExplainTools(ToolRegistry registry, dynamic application)
        {
            var service = new ExcelExplainService(application);
            registry.Register(new ToolDescriptor(
                "explain_excel_formula",
                "只读解释一个公式：真机拆解其函数/参数/引用单元格的实际值，返回结构化信息供转述。无法真机求值的引用（跨表/UDF/数组等）会显式标记为降级，不会编造解释。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.ExplainFormula(args)),
                ExcelExplainToolSchemas.ExplainFormula()));
        }
    }
}