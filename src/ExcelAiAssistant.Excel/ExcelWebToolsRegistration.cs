using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>网页抓取工具注册：只读，不写工作簿、不执行脚本。</summary>
    public static class ExcelWebToolsRegistration
    {
        public static void RegisterWebTools(ToolRegistry registry)
        {
            registry.Register(new ToolDescriptor(
                "fetch_url",
                "只读抓取一个 http/https 网页或文本 URL，返回其内容文本（自动转纯文本，默认脱敏隐藏邮箱/手机/身份证）供分析；不写工作簿、不执行任何代码、不访问本地文件。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => WebFetchService.Fetch(args)),
                WebToolSchemas.FetchUrl()));
        }
    }
}