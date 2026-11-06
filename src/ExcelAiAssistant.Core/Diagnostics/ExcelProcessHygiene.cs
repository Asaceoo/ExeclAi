using System;
using System.Diagnostics;
using Microsoft.Win32;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// M27R3 审查收口：进程与加载项卫生。
    /// 背景：真机验证脚本原先无条件 Kill 全机 EXCEL / et.exe，导致 ①误杀用户正在编辑的会话；
    /// ②Excel 判定加载项「启动崩溃」并把全部加载项取消勾选（LoadBehavior 3→2，实测事故，用户报「插件全没了」）。
    /// 约定：只优雅关闭有主窗口的实例；仅无主窗口的 automation 孤儿按 PID 强杀；强杀后复位加载项勾选。
    /// </summary>
    public static class ExcelProcessHygiene
    {
        /// <summary>只清理无主窗口的 automation 孤儿，并在清理后复位 Excel 加载项勾选（LoadBehavior=3）。</summary>
        public static void CleanupOrphansOnly()
        {
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (p.MainWindowHandle == IntPtr.Zero) { p.Kill(); } // 无窗口 = automation 孤儿，不参与加载项加载
                    }
                    catch { }
                    finally { try { p.Dispose(); } catch { } }
                }
            }
            catch { }

            RestoreAddinLoadBehavior();
        }

        /// <summary>把被 Excel 取消勾选的加载项复位为「随 Excel 启动并加载」（LoadBehavior=3）。</summary>
        public static void RestoreAddinLoadBehavior()
        {
            foreach (var hive in new[] { RegistryHive.CurrentUser })
            {
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                    using (var addins = baseKey.OpenSubKey(@"Software\Microsoft\Office\Excel\Addins", writable: true))
                    {
                        if (addins == null) { continue; }
                        foreach (var sub in addins.GetSubKeyNames())
                        {
                            try
                            {
                                using (var key = addins.OpenSubKey(sub, writable: true))
                                {
                                    var current = key?.GetValue("LoadBehavior");
                                    if (current != null && Convert.ToInt32(current) == 2) { key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
