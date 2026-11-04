// 临时诊断：列出 WPS 宿主视角（dashboard=null）的注册工具清单与总数
using ExcelAiAssistant.Core.Diagnostics;
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
                // FIX（M27R3 审查）：原实现无条件 Kill 全机 EXCEL —— 会连累用户正在编辑的表格，
                // 且会让 Excel 把全部加载项取消勾选（LoadBehavior 3→2，实测事故）。
                // 现仅强杀「无主窗口」的 automation 孤儿（有窗口的一律优雅关闭），并在强杀后复位加载项勾选。
                ExcelProcessHygiene.CleanupOrphansOnly();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return 0;
        }
    }
}
