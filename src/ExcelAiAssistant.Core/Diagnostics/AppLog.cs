using System;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// 跨程序集日志门面：Core/Excel 层写运行日志的唯一入口，宿主（Addin）启动时把真实落盘实现
    /// 注入为 <see cref="Sink"/>。默认 no-op，保证纯逻辑层可无副作用地单测；未注入时静默丢弃，
    /// 生产由 OnConnection 注入 DiagnosticsLog。
    /// 目的：让「覆盖全部功能的运行日志」不再受程序集边界限制——所有工具/服务层都能统一记录，
    /// 便于调试定位任一子系统的错误，而不是各自吞掉异常。
    /// </summary>
    public static class AppLog
    {
        /// <summary>注入的日志接收器；线程安全由接收方保证，本方每次读取取最新快照。</summary>
        public static Action<LogEntry> Sink;

        /// <summary>宿主（Addin）在启动时调用，把真实落盘日志接入。</summary>
        public static void SetSink(Action<LogEntry> sink) => Sink = sink;

        public static void Info(string message) => Write(CoreLogLevel.Info, message);

        public static void Warn(string message) => Write(CoreLogLevel.Warn, message);

        public static void Error(string message) => Write(CoreLogLevel.Error, message);

        /// <summary>带分类标签的日志（如 "python" / "vba" / "tool" / "verify"），便于前端按类别过滤。</summary>
        public static void Info(string category, string message) => Write(CoreLogLevel.Info, category, message);

        public static void Warn(string category, string message) => Write(CoreLogLevel.Warn, category, message);

        public static void Error(string category, string message) => Write(CoreLogLevel.Error, category, message);

        private static void Write(CoreLogLevel level, string message) => Write(level, null, message);

        private static void Write(CoreLogLevel level, string category, string message)
        {
            try
            {
                Sink?.Invoke(new LogEntry
                {
                    Time = DateTime.Now,
                    Level = level,
                    Category = category,
                    Message = message
                });
            }
            catch (Exception)
            {
                // 日志接收方异常不得影响业务主流程。
            }
        }
    }

    /// <summary>与宿主日志一致的三级级别。</summary>
    public enum CoreLogLevel
    {
        Info = 0,
        Warn = 1,
        Error = 2
    }

    /// <summary>统一的结构化日志条目：带分类与时间，不限程序集。</summary>
    public sealed class LogEntry
    {
        public DateTime Time { get; set; }

        public CoreLogLevel Level { get; set; }

        /// <summary>业务分类：python/json/vba/tool/verify 等，可空。</summary>
        public string Category { get; set; }

        public string Message { get; set; }
    }
}