using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAiAssistant.Addin
{
    /// <summary>日志级别，用于设置面板按严重程度过滤。</summary>
    public enum LogLevel
    {
        Info,
        Warn,
        Error
    }

    /// <summary>
    /// 结构化运行日志：既落盘到 %TEMP%\ExcelAiAssistant.log，也维护一份进程内环形缓冲，
    /// 供设置页「运行日志」实时查询（含级别过滤、清空、自检输出）。
    /// </summary>
    internal static class DiagnosticsLog
    {
        private static readonly object SyncRoot = new object();
        private const int RingCapacity = 4000;

        /// <summary>单文件轮转阈值：超过 5MB 滚出备份，防止 %TEMP% 日志无限膨胀。</summary>
        private const long RotateThresholdBytes = 5L * 1024 * 1024;

        /// <summary>轮转备份保留个数：ExcelAiAssistant.log → ExcelAiAssistant.<时间戳>.log，仅保留最近 5 份。</summary>
        private const int RotateBackupsToKeep = 5;

        private static readonly LinkedList<LogEntry> Ring = new LinkedList<LogEntry>();

        private static string LogPathLazy;
        private static string LogPath => LogPathLazy ?? (LogPathLazy = Path.Combine(Path.GetTempPath(), "ExcelAiAssistant.log"));

        private static readonly object RingTrimLock = new object();

        public static void Write(string message) => Write(LogLevel.Info, message);

        public static void Info(string message) => Write(LogLevel.Info, message);

        public static void Warn(string message) => Write(LogLevel.Warn, message);

        public static void Error(string message) => Write(LogLevel.Error, message);

        public static void Write(LogLevel level, string message)
        {
            Append(level, message);
        }

        private static void Append(LogLevel level, string message)
        {
            var now = DateTime.Now;
            var line = $"{now:yyyy-MM-dd HH:mm:ss.fff} [{level.ToString().ToUpperInvariant()}] {message}";
            try
            {
                lock (SyncRoot)
                {
                    RotateIfNeeded();
                    File.AppendAllText(LogPath, line + Environment.NewLine);
                }
            }
            catch
            {
                // 磁盘不可写时不能影响主流程（自引用打点会递归——审查 R1 修正：此处天然经 Ring 环缓存留痕）

            }

            lock (RingTrimLock)
            {
                Ring.AddLast(new LogEntry { Ticks = now.Ticks, Time = now, Level = level, Message = message, Line = line });
                while (Ring.Count > RingCapacity)
                {
                    Ring.RemoveFirst();
                }
            }
        }

        /// <summary>必须在持有 SyncRoot 时调用：主日志超阈值则滚出带时间戳备份并裁剪历史备份数量。
        /// 轮转失败（文件被占用等）静默跳过——本轮仍追加写入，下轮重试。</summary>
        private static void RotateIfNeeded()
        {
            try
            {
                var file = new FileInfo(LogPath);
                if (!file.Exists || file.Length < RotateThresholdBytes)
                {
                    return;
                }

                var backup = LogPath.Replace(".log", $"-{DateTime.Now:yyyyMMdd-HHmmss}.log");
                File.Move(LogPath, backup);
                PruneRotateBackups();
            }
            catch
            {
                // 轮转失败不阻断写日志主流程
            }
        }

        /// <summary>仅保留最近 RotateBackupsToKeep 份 ExcelAiAssistant-*.log 备份，更早的删除。</summary>
        private static void PruneRotateBackups()
        {
            try
            {
                var dir = Path.GetDirectoryName(LogPath);
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                {
                    return;
                }

                var backups = Directory.GetFiles(dir, "ExcelAiAssistant-*.log")
                    .OrderByDescending(File.GetLastWriteTime)
                    .ToList();
                for (var i = RotateBackupsToKeep; i < backups.Count; i++)
                {
                    try { File.Delete(backups[i]); }
                    catch
                    {
                        // 单个备份删除失败不影响其余裁剪
                    }
                }
            }
            catch
            {
                // 裁剪失败不影响主流程
            }
        }

        /// <summary>按级别与尾部条数返回最近日志（新→旧）。</summary>
        public static IReadOnlyList<LogEntry> Query(int tail, string level)
        {
            lock (RingTrimLock)
            {
                var filtered = level == null
                    ? Ring.ToList()
                    : Ring.Where(e => string.Equals(e.Level.ToString(), level, StringComparison.OrdinalIgnoreCase)).ToList();
                if (tail > 0 && filtered.Count > tail)
                {
                    filtered.RemoveRange(0, filtered.Count - tail);
                }

                filtered.Reverse(); // 新→旧
                return filtered;
            }
        }

        /// <summary>清空进程内日志；同时把已落盘文件改名为 ExcelAiAssistant.<时间戳>.log 备份，再从空文件继续。</summary>
        public static string Clear()
        {
            lock (RingTrimLock)
            {
                Ring.Clear();
            }

            string backup = null;
            try
            {
                lock (SyncRoot)
                {
                    if (File.Exists(LogPath))
                    {
                        backup = LogPath.Replace(".log", $"-{DateTime.Now:yyyyMMdd-HHmmss}.log");
                        File.Move(LogPath, backup);
                    }
                }
            }
            catch
            {
                backup = null;
            }

            return backup;
        }
    }

    /// <summary>一条结构化日志。</summary>
    internal sealed class LogEntry
    {
        public long Ticks { get; set; }
        public DateTime Time { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; }
        public string Line { get; set; }
    }
}