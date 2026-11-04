// 诊断工具：列出当前宿主可见的注册工具清单与总数。
// M29 审查修复：原实现任何异常都会冒泡成 Windows "应用程序错误 0xe0434352" 弹窗
// （CLR 未处理异常）；现统一 try/catch，异常时打印可读错误并以非 0 退出码结束，绝不再弹窗。
using System;
using System.Linq;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Excel;

namespace ToolCountDiag
{
    internal static class Diag
    {
        private static int Main()
        {
            object app = null;
            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Console.WriteLine("no-excel");
                    return 2;
                }

                dynamic created = Activator.CreateInstance(excelType);
                app = (object)created;
                created.Visible = false;
                created.DisplayAlerts = false;

                var host = new ExcelToolHost(app);
                var registry = host.CreateRegistry();
                var tools = registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase).ToList();
                Console.WriteLine("TOOL_COUNT=" + tools.Count);
                foreach (var t in tools)
                {
                    Console.WriteLine(t.Name + "\t" + t.RiskLevel);
                }

                return 0;
            }
            catch (Exception ex)
            {
                // 诊断工具不应把宿主拖进 WER 弹窗：异常 → 单行可读输出 + 退出码 1
                var msg = ex.Message;
                if (ex is System.Runtime.InteropServices.COMException com)
                {
                    msg += " (HRESULT:0x" + unchecked((uint)com.HResult).ToString("X8") + ")";
                }

                Console.Error.WriteLine("TOOL_COUNT_ERROR: " + ex.GetType().Name + ": " + msg);
                return 1;
            }
            finally
            {
                // 只回收无主 automation 实例；绝不触碰用户正在编辑的 Excel
                try { ExcelProcessHygiene.CleanupOrphansOnly(); } catch { }
                try
                {
                    if (app != null && System.Runtime.InteropServices.Marshal.IsComObject(app))
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
                    }
                }
                catch { }

                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch { }
            }
        }
    }
}