// 临时诊断：列出 WPS 宿主视角（dashboard=null）的注册工具清单与总数
using System;
using System.Linq;
using ExcelAiAssistant.Excel;

namespace ToolCountDiag
{
    internal static class Diag
    {
        private static int Main()
        {
            dynamic excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) { Console.WriteLine("no-excel"); return 2; }
            dynamic app = Activator.CreateInstance(excelType);
            app.Visible = false;
            app.DisplayAlerts = false;
            try
            {
                var host = new ExcelToolHost((object)app);
                var registry = host.CreateRegistry();
                var tools = registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase).ToList();
                Console.WriteLine("TOOL_COUNT=" + tools.Count);
                foreach (var t in tools)
                {
                    Console.WriteLine(t.Name + "\t" + t.RiskLevel);
                }
            }
            finally
            {
                foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    try { p.Kill(); } catch { }
                    p.Dispose();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return 0;
        }
    }
}
