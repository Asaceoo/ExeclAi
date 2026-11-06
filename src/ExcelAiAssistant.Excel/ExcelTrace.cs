using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// Excel 层轻量结构化日志门面（2026-09-12 日志全覆盖批次）：
    /// Addin 层 DiagnosticsLog 因分层不可反向引用，Excel 层服务（图表/看板/截图等）的
    /// 关键路径与异常此前只有返回值没有日志，远程调试/事后定位困难——本门面补齐。
    /// 设计（最成熟方案的最小子集）：
    /// - 落盘 %TEMP%\ExcelAiAssistant.excel.log，与 Addin 层主日志分离互不干扰；
    /// - 单文件 5MB 轮转，保留最近 3 份备份；
    /// - 级别 Info/Warn/Error；写入全程持锁 + 容错（日志失败绝不影响业务）；
    /// - 全部方法吞异常：日志是旁路，不是业务依赖。
    /// </summary>
    internal static class ExcelTrace
    {
        private const long RotateThresholdBytes = 5L * 1024 * 1024;
        private const int RotateBackupsToKeep = 3;
        private static readonly object SyncRoot = new object();

        private static string LogPathLazy;
        private static string LogPath => LogPathLazy ?? (LogPathLazy = Path.Combine(Path.GetTempPath(), "ExcelAiAssistant.excel.log"));

        public static void Info(string tool, string message) => Write("INFO", tool, message);
        public static void Warn(string tool, string message) => Write("WARN", tool, message);
        public static void Error(string tool, string message) => Write("ERROR", tool, message);

        /// <summary>记录工具异常：类型 + 消息 + 堆栈首行（定位用）。</summary>
        public static void Error(string tool, string where, Exception ex)
        {
            if (ex == null) { return; }
            var stack = ex.StackTrace ?? string.Empty;
            int nl = stack.IndexOf('\n');
            var firstLine = nl > 0 ? stack.Substring(0, nl).Trim() : stack.Trim();
            Write("ERROR", tool, where + " | " + ex.GetType().Name + ": " + ex.Message + " @ " + firstLine);
        }

        private static void Write(string level, string tool, string message)
        {
            try
            {
                lock (SyncRoot)
                {
                    try { RotateIfNeeded(); } catch { }
                    var line = new StringBuilder();
                    line.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                        .Append(" [").Append(level).Append("] [").Append(tool ?? "excel").Append("] ")
                        .Append(message ?? string.Empty)
                        .AppendLine();
                    File.AppendAllText(LogPath, line.ToString(), Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // 日志是旁路：磁盘满/权限/AV 拦截等一律吞掉，不影响 Excel 操作
            }
        }

        private static void RotateIfNeeded()
        {
            if (!File.Exists(LogPath)) { return; }
            if (new FileInfo(LogPath).Length < RotateThresholdBytes) { return; }
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backup = LogPath.Replace(".log", "." + stamp + ".log");
            File.Move(LogPath, backup);
            // 清理超量备份
            var dir = Path.GetDirectoryName(LogPath);
            var name = Path.GetFileNameWithoutExtension(LogPath);
            var backups = Directory.GetFiles(dir ?? ".", name + ".*.log");
            if (backups.Length > RotateBackupsToKeep)
            {
                Array.Sort(backups);
                for (int i = 0; i < backups.Length - RotateBackupsToKeep; i++)
                {
                    File.Delete(backups[i]);
                }
            }
        }
    }
}
