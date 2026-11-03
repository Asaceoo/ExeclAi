using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace WpsToolHarness
{
    /// <summary>
    /// M3 审计 harness：连 KET.Application，构建全量 ToolRegistry。
    /// 用法：WpsToolHarness.exe [list|smoke|battery]
    ///   list    = 输出全部工具名 + RiskLevel
    ///   smoke   = 只读冒烟（上下文/预览）
    ///   battery = 按品类代表工具执行真机验证（WPS 兼容审计核心）
    /// </summary>
    internal static class Program
    {
        private static Dictionary<string, object> Args(params object[] kv)
        {
            var dict = new Dictionary<string, object>();
            for (var i = 0; i + 1 < kv.Length; i += 2)
            {
                dict[(string)kv[i]] = kv[i + 1];
            }
            return dict;
        }

        private static int Main(string[] args)
        {
            var mode = args.Length > 0 ? args[0] : "list";

            // 【时机关键】共享/独占判定必须在 CreateInstance 之前：WPS 单实例模型下
            // 起完实例再检测必然命中自己刚起的 et.exe，导致独占模式下 DisplayAlerts=false
            // 永不生效 → delete_excel_worksheet 弹模态确认框永久挂起（M6 实测回归，
            // M3 能跑通正因当时尚无此判定）。battery 模式再叠加独占守卫（下方）。
            bool etWasRunning = System.Diagnostics.Process.GetProcessesByName("et").Length > 0;
            // 自起实例 PID 快照：退出时仅兑底清理新增的 et.exe（不碰用户已开的进程）。
            var preEtPids = new HashSet<int>(
                System.Diagnostics.Process.GetProcessesByName("et").Select(p => p.Id));

            // battery 独占守卫：电池是风险工具真机门禁（NormalWrite 逐项真跑），
            // 若已有 et.exe 在跑则拒绝执行——绝不附着用户实例：① Workbooks.Add 会在用户
            // 会话里弹新工作簿污染工作区；② 共享模式下不能设 DisplayAlerts=false，
            // delete_excel_worksheet 等操作会弹模态确认框永久挂起（实测两次卡死同点）；
            // ③ 写类工具副作用外溢到用户工作簿（实测残留 T1/T2 工作表）。
            // 门禁语义：电池只接受独占 WPS 环境，提示用户关闭后重跑。
            if ((string.Equals(mode, "battery", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(mode, "full", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(mode, "m7", StringComparison.OrdinalIgnoreCase)) && etWasRunning)
            {
                Console.WriteLine("BATTERY_REQUIRES_EXCLUSIVE: 检测到正在运行的 WPS 表格（et.exe），电池门禁拒绝在共享实例上执行。请关闭全部 WPS 窗口后重跑。");
                return 3;
            }

            dynamic app;
            try
            {
                var type = Type.GetTypeFromProgID("KET.Application");
                if (type == null)
                {
                    Console.WriteLine("PROGID_NOT_FOUND: KET.Application 未注册");
                    return 2;
                }

                app = Activator.CreateInstance(type);
            }
            catch (Exception ex)
            {
                Console.WriteLine("KET_CONNECT_FAIL: " + ex.Message);
                return 2;
            }

            // 宿主感知保护：共享模式（启动前已有 et.exe）不设 Visible/DisplayAlerts、不 Quit，
            // 只关闭自建工作簿，避免把用户开着的 WPS 连带隐藏/退出造成工作丢失。
            // （etWasRunning 已在 Main 入口、CreateInstance 之前采样，见上方注释。）
            if (etWasRunning) { Console.WriteLine("HOST_SHARED=1 (existing et.exe detected; will not Quit)"); }

            try
            {
                if (!etWasRunning)
                {
                    app.Visible = false;
                    app.DisplayAlerts = false;
                }
                dynamic workbook = app.Workbooks.Add();
                var host = new ExcelAiAssistant.Excel.ExcelToolHost((object)app);
                var registry = host.CreateRegistry();
                var tools = registry.Tools.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase).ToList();

                Console.WriteLine("TOOL_COUNT=" + tools.Count);
                foreach (var tool in tools)
                {
                    Console.WriteLine(tool.Name + "\t" + tool.RiskLevel);
                }

                if (string.Equals(mode, "smoke", StringComparison.OrdinalIgnoreCase))
                {
                    RunOne(registry, "get_current_excel_context", new Dictionary<string, object>());
                    RunOne(registry, "preview_excel_range", Args("address", "A1:B2"));
                }
                else if (string.Equals(mode, "battery", StringComparison.OrdinalIgnoreCase))
                {
                    RunOne(registry, "write_excel_range", Args("address", "A1:B2", "values", new[] { new[] { 1, 2 }, new[] { 3, 4 } }));
                    RunOne(registry, "write_excel_cells_batch", Args("cells", new object[] { new Dictionary<string, object> { ["address"] = "D1", ["value"] = "x" } }));
                    RunOne(registry, "append_excel_rows", Args("start_cell", "F1", "values", new[] { new[] { "a", "b" } }));
                    RunOne(registry, "sort_range", Args("address", "A1:B2", "sort_column", 1, "direction", "asc"));
                    RunOne(registry, "dedup_range", Args("address", "A1:B2"));
                    RunOne(registry, "transpose_range", Args("address", "A1:B2"));
                    RunOne(registry, "find_replace_range", Args("address", "A1:B2", "find", "x", "replace", "y"));
                    RunOne(registry, "format_excel_range", Args("address", "A1:B2", "bold", true));
                    RunOne(registry, "format_excel_conditional", Args("address", "A1:B2", "rule", "greater", "value", 1));
                    RunOne(registry, "data_validation", Args("address", "E1:E5", "type", "list", "source", "甲,乙,丙"));
                    RunOne(registry, "flag_cells", Args("address", "A1:B2"));
                    RunOne(registry, "add_excel_worksheet", Args("name", "T1"));
                    RunOne(registry, "rename_excel_worksheet", Args("sheet", "T1", "name", "T2"));
                    RunOne(registry, "delete_excel_worksheet", Args("sheet", "T2"));
                    RunOne(registry, "create_excel_chart", Args("source_address", "A1:B2", "chart_type", "ColumnClustered"));
                    RunOne(registry, "create_excel_pivot_summary", Args("source_address", "A1:B9", "label_column", 1, "value_column", 2, "target_address", "H1"));
                    RunOne(registry, "create_excel_sparkline", Args("data_range", "A1:B1", "dest_range", "J1"));
                    RunOne(registry, "list_external_links", new Dictionary<string, object>());
                    RunOne(registry, "get_excel_sheet_list", new Dictionary<string, object>());
                    RunOne(registry, "undo_excel_operation", new Dictionary<string, object>());
                    RunOne(registry, "export_sheets_to_pdf", Args("file_path", System.IO.Path.Combine(System.IO.Path.GetTempPath(), "wps_harness_export.pdf")));
                    RunOne(registry, "run_vba_snippet", Args("snippet", "Beep"));
                    // 负向用例：坏参数必须返回结构化错误而非崩溃/挂起（工具闭环拿真实报错）
                    RunOne(registry, "write_excel_range", Args("address", "A1:B2")); // 缺 values 参数
                }
                else if (string.Equals(mode, "full", StringComparison.OrdinalIgnoreCase))
                {
                    FullScenarios.Run(registry, app);
                }
                else if (string.Equals(mode, "m7", StringComparison.OrdinalIgnoreCase))
                {
                    // M6 复审专项：schema 对象项契约 + 静默成功防线真机验证
                    FullScenarios.RunM7(registry, app);
                }
                else if (string.Equals(mode, "m17", StringComparison.OrdinalIgnoreCase))
                {
                    // M17 专项：写路径工具保护态集中守卫（SHEET_PROTECTED 结构化错误）
                    var m17Failed = FullScenarios.RunM17(registry, app);
                    Console.WriteLine("M17_FAILED=" + m17Failed);
                    if (m17Failed > 0)
                    {
                        try { workbook.Close(false); } catch { }
                        if (!etWasRunning) { try { app.Quit(); } catch { } }
                        return 4;
                    }
                }

                try { workbook.Close(false); }
                catch (Exception teardown) { Console.WriteLine("TEARDOWN_WARN(close): " + teardown.Message); }
                if (!etWasRunning)
                {
                    try { app.Quit(); }
                    catch (Exception teardown) { Console.WriteLine("TEARDOWN_WARN(quit): " + teardown.Message); }
                    // 兑底：RCW 分离等异常可能让 Quit() 未能送达，独占模式下残留的 et.exe
                    // 会令下一次门禁 battery 被独占守卫拦截（TOOL_COUNT=0 → throw）。
                    // 只 kill 本次运行新增的 et 进程，不碰用户已开的实例。
                    foreach (var p in System.Diagnostics.Process.GetProcessesByName("et"))
                    {
                        if (!preEtPids.Contains(p.Id))
                        {
                            try { p.Kill(); Console.WriteLine("TEARDOWN_KILLED_ET pid=" + p.Id); } catch { }
                        }
                    }
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("HARNESS_FAIL: " + ex);
                try { if (!etWasRunning) { app.Quit(); } } catch { }
                if (!etWasRunning)
                {
                    foreach (var p in System.Diagnostics.Process.GetProcessesByName("et"))
                    {
                        if (!preEtPids.Contains(p.Id))
                        {
                            try { p.Kill(); } catch { }
                        }
                    }
                }
                return 3;
            }
        }

        private static void RunOne(ExcelAiAssistant.Core.Tools.ToolRegistry registry, string name, Dictionary<string, object> args)
        {
            if (!registry.TryGet(name, out var tool))
            {
                Console.WriteLine("RESULT\t" + name + "\tNOT_REGISTERED\t-\t-");
                return;
            }

            try
            {
                var result = tool.Execute(args);
                var payload = result?.Data == null ? "-" : JsonConvert.SerializeObject(result.Data);
                if (payload.Length > 220)
                {
                    payload = payload.Substring(0, 220) + "…";
                }

                Console.WriteLine("RESULT\t" + name + "\t" + (result?.Ok == true ? "OK" : "FAIL") + "\t" + (result?.ErrorCode ?? "-") + "\t" + payload);
            }
            catch (Exception ex)
            {
                var message = ex.Message;
                if (message.Length > 200)
                {
                    message = message.Substring(0, 200) + "…";
                }

                Console.WriteLine("RESULT\t" + name + "\tEXCEPTION\t" + ex.GetType().Name + "\t" + message);
            }
        }
    }
}
