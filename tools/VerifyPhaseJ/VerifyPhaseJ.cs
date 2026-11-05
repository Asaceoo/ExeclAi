using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyPhaseJ
{
    /// <summary>
    /// 自动真机冒烟（第 J 批）：P5-C1 跨工作簿只读引用端到端。
    /// 启动真实 Excel：
    ///  1) open_excel_external_book 以只读句柄打开外部工作簿，返回句柄与工作表清单；
    ///  2) read_external_range 读回该句柄指定区域，数值与外部文件一致；
    ///  3) 打开外部句柄后当前活动工作簿仍是写入目标（安全边界：仅读不受影响）；
    ///  4) 重复打开同一外部文件复用句柄（不泄漏）；
    ///  5) close_external_book 显式关闭句柄，后续读取返回 HANDLE_CLOSED。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures)
                {
                    Console.WriteLine("FAILED: " + f);
                }

                return 1;
            }

            Console.WriteLine("ALL PHASE-J CROSS-WORKBOOK CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            dynamic app = null;
            dynamic wbActive = null;
            var activeName = string.Empty;
            string extPath = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;

                wbActive = app.Workbooks.Add();
                activeName = Convert.ToString(wbActive.Name);

                // 外部工作簿：写入 4 行，保存为临时文件。
                var ext = app.Workbooks.Add();
                dynamic ws = ext.Worksheets.Item[1];
                ws.Name = "ExtData";
                ws.Range["A1:C4"].Value2 = new object[,]
                {
                    { "id", "stat", "score" },
                    { 1D, "open", 91D },
                    { 2D, "read", 82D },
                    { 3D, "probe", 74D }
                };
                extPath = Path.Combine(Path.GetTempPath(), "xai-phasej-ext-" + Guid.NewGuid().ToString("N") + ".xlsx");
                ext.SaveAs(extPath);
                Release((object)ws);
                ext.Close(false);
                Release((object)ext);

                var registry = new ExternalBookRegistry((object)app);
                var toolReg = new ToolRegistry();
                ExcelExternalBookTools.Register(toolReg, registry);

                // 1) open 外部只读句柄
                var info1 = registry.Open(extPath);
                Assert("open-handle", info1 != null && info1.Token.StartsWith("ext_"), "open 返回 ext_ 句柄");
                Assert("open-sheets", info1.Sheets.Contains("ExtData"), "句柄工作表含 ExtData");

                // 2) 通过工具读回指定区域
                var read = ReadRange(toolReg, info1.Token, "ExtData", "A1:C4");
                var resultJson = read.Data as string ?? string.Empty;
                Assert("read-eq", read.Ok && resultJson.Contains("82") && resultJson.Contains("probe")
                    && resultJson.Contains("score"), "read_external_range 读回外部数据");

                // 3) 打开外部句柄后活动工作簿未被改动
                var activeAfter = Convert.ToString(app.ActiveWorkbook.Name);
                Assert("active-unchanged", activeAfter == activeName, "读外部后活动工作簿仍是写入目标（=" + activeAfter + "）");

                // 4) 重复打开复用句柄
                var info2 = registry.Open(extPath);
                Assert("open-dedup", info2.Token == info1.Token, "重复打开复用同一句柄（不泄漏）");

                // 5) close 后读取失败
                Assert("close", registry.Close(info1.Token), "显式关闭句柄");
                var closed = ReadRange(toolReg, info1.Token, "ExtData", "A1:C2");
                Assert("read-after-close", !closed.Ok && closed.ErrorCode == "HANDLE_CLOSED",
                    "关闭后读取返回句柄失效（error=" + closed.ErrorCode + "）");
            }
            finally
            {
                try { if (wbActive != null) { wbActive.Close(false); Release((object)wbActive); } }
                catch (Exception ex) { Console.WriteLine("NOTE  active-close: " + Safe(ex)); }
                try { if (app != null) { app.ScreenUpdating = true; app.Quit(); Release((object)app); } }
                catch (Exception ex) { Console.WriteLine("NOTE  app-quit: " + Safe(ex)); }
                try { if (extPath != null && File.Exists(extPath)) { File.Delete(extPath); } }
                catch { }
            }
        }

        private static ToolResult ReadRange(ToolRegistry reg, string handle, string sheet, string range)
        {
            if (reg.TryGet("read_external_range", out var tool))
            {
                return tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["handle"] = handle,
                    ["sheet"] = sheet,
                    ["range"] = range
                });
            }

            return new ToolResult { Ok = false, ErrorCode = "NOT_REGISTERED", Data = "read_external_range 未注册" };
        }

        private static void Assert(string name, bool condition, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("PASS " + name + " — " + detail);
            }
            else
            {
                _failed++;
                Failures.Add(name + ": " + detail);
                Console.WriteLine("FAIL " + name + " — " + detail);
            }
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + ": " + detail);
            Console.WriteLine("FAIL " + name + " — " + detail);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        private static string Safe(Exception ex)
        {
            return ex == null ? string.Empty : ex.Message;
        }
    }
}
