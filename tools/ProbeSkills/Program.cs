using System;
using System.Reflection;
using System.Threading;

internal static class ProbeSkills
{
    private static int Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        dynamic excel = null;
        try
        {
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            excel.Visible = false;
            excel.DisplayAlerts = false;
            // 故意不建工作簿：复现「Excel 启动初期无 ActiveWorkbook」场景，
            // 验证 Initialize 不再因 GetCurrentContext 抛异常而整体失败、技能仍能加载。

            dynamic addins = excel.COMAddIns;
            object addinObj = null;
            foreach (var item in addins)
            {
                var progId = Convert.ToString(item.ProgID ?? string.Empty);
                Console.WriteLine("  addin progid=" + progId + " connect=" + item.Connect);
                if (string.Equals(progId, "ExcelAiAssistant.AddIn", StringComparison.OrdinalIgnoreCase))
                {
                    for (var attempt = 0; attempt < 5 && addinObj == null; attempt++)
                    {
                        try
                        {
                            addinObj = item.Object;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("  -> Object access failed: " + ex.Message);
                            addinObj = null;
                        }

                        if (addinObj == null)
                        {
                            Console.WriteLine("  -> attempt " + attempt + " null, waiting for load...");
                            Thread.Sleep(2000);
                        }
                    }

                    Console.WriteLine("  -> Object resolved: " + (addinObj != null ? addinObj.GetType().FullName : "null"));
                    break;
                }
            }

            if (addinObj == null)
            {
                Console.WriteLine("ADDIN NOT FOUND (Object null) — 保持 Excel 运行让前端 Initialize 执行 LoadSkills...");
                Thread.Sleep(15000);
                Console.WriteLine("WAIT DONE");
                return 1;
            }

            Console.WriteLine("ADDIN LOADED: " + addinObj.GetType().FullName);
            var addinType = addinObj.GetType();

            var fTaskPane = addinType.GetField("taskPane", BindingFlags.NonPublic | BindingFlags.Instance);
            object taskPane = fTaskPane?.GetValue(addinObj);
            if (taskPane == null)
            {
                var mShow = addinType.GetMethod("ShowTaskPane", BindingFlags.NonPublic | BindingFlags.Instance);
                mShow?.Invoke(addinObj, null);
                for (var i = 0; i < 50 && taskPane == null; i++)
                {
                    Thread.Sleep(200);
                    taskPane = fTaskPane?.GetValue(addinObj);
                }
            }

            if (taskPane == null)
            {
                Console.WriteLine("TASKPANE NOT CREATED");
                return 2;
            }

            Console.WriteLine("TASKPANE: " + taskPane.GetType().FullName);
            var pContent = taskPane.GetType().GetProperty("ContentControl");
            object host = pContent?.GetValue(taskPane);
            if (host == null)
            {
                Console.WriteLine("CONTENT NULL");
                return 3;
            }

            Console.WriteLine("HOST: " + host.GetType().FullName);
            var fBridge = host.GetType().GetField("sidebarBridge", BindingFlags.NonPublic | BindingFlags.Instance);
            object bridge = null;
            for (var i = 0; i < 60 && bridge == null; i++)
            {
                bridge = fBridge?.GetValue(host);
                if (bridge == null)
                {
                    Thread.Sleep(200);
                }
            }

            if (bridge == null)
            {
                Console.WriteLine("BRIDGE NOT CREATED");
                return 4;
            }

            Console.WriteLine("BRIDGE: " + bridge.GetType().FullName);

            // 场景一：无工作簿时 Initialize 必须成功且带技能（本次修复的核心）
            var mInit = bridge.GetType().GetMethod("Initialize");
            var initJson = (string)mInit.Invoke(bridge, null);
            Console.WriteLine("INIT JSON: " + initJson);
            if (initJson == null || initJson.IndexOf("\"ok\":true", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Console.WriteLine("INIT FAILED (ok!=true) — 无工作簿时初始化仍失败");
                return 6;
            }

            if (initJson.IndexOf("\"skills\":[", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Console.WriteLine("INIT SKILLS EMPTY — 无工作簿时技能列表为空");
                return 7;
            }

            Console.WriteLine("INIT OK (no workbook, skills loaded)");

            // 场景二：GetSkills 独立接口
            var mGet = bridge.GetType().GetMethod("GetSkills");
            var task = (System.Threading.Tasks.Task<string>)mGet.Invoke(bridge, null);
            var json = task.GetAwaiter().GetResult();
            Console.WriteLine("SKILLS JSON: " + json);
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex);
            return 5;
        }
        finally
        {
            try
            {
                if (excel != null)
                {
                    excel.ScreenUpdating = true;
                    excel.Quit();
                }
            }
            catch
            {
            }
        }
    }
}
