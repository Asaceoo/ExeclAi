using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>增强工具注册：环比分析（读+可选写）、批注管理、数据有效性。</summary>
    public static class ExcelEnhanceToolsRegistration
    {
        public static void RegisterEnhanceTools(
            ToolRegistry registry,
            ExcelEnhanceService service)
        {
            registry.Register(new ToolDescriptor(
                "rate_change_analysis",
                "读取 label+value 两列，逐行计算绝对变化与百分比变化（环比），可选将结果写回；适合时间/周期序列分析。只读分析，写回需显式 output_address。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.RateChangeAnalysis(args)),
                ExcelEnhanceSchemas.RateChangeAnalysis()));

            registry.Register(new ToolDescriptor(
                "cell_comment",
                "管理单元格批注：get 读取、set 写入、clear 清除一个区域的批注。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.CellComment(args)),
                ExcelEnhanceSchemas.CellComment()));

            registry.Register(new ToolDescriptor(
                "data_validation",
                "设置数据有效性：下拉列表（list，source 为单元格范围或逗号列表）或整数区间校验（interval）。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.DataValidation(args)),
                ExcelEnhanceSchemas.DataValidation()));
        }
    }
}