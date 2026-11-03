using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// 一键自检日志异常提取（降噪版）：
    /// - 跳过「回合开始」心跳行（含「空闲超时=」配置回显，曾被误报为异常——真机实证）。
    /// - 泛化的「超时」签名收紧为「执行超时」（只匹配真实超时失败，不匹配配置回显）。
    /// - 其余签名保持：Info 级工具日志也可能带真实错误码（如 EXCEL_COM_ERROR），故签名匹配不限定 [Error] 级。
    /// - 事件签名去重（v3.7.1.58）：同一次失败会以三种行文落多份日志（[python] worker / 工具事件 / 回合摘要），
    ///   且主日志与轮转副本同文；旧整行精确去重挡不住，一键自检把 10 个真实故障虚报 31 处（真机实证）。
    ///   现按「归一化错误码 + 10 秒时间桶」签名去重：同桶同码只留首条；不同时间点的独立重试仍各自保留。
    /// </summary>
    public static class LogIssueExtractor
    {
        private static readonly string[] Patterns =
        {
            "EXCEL_COM_ERROR", "TOOL_EXCEPTION", "ok=False", "ok=false",
            "PYTHON_RUN_ERROR", "WRITE_RANGE_INVALID", "PYTHON_DEPS_INSTALL_FAILED",
            "输入字符串的格式不正确", "cells 必须是对象数组", "无效的窗口句柄",
            "任务撤销", "执行超时", "Shutdown", "InvalidComObjectException", "RuntimeBinderException"
        };

        private static readonly Regex TimestampRx =
            new Regex(@"^\s*(\d{1,2}:\d{2}:\d{2})(?:\.\d+)?", RegexOptions.Compiled);

        /// <summary>从格式化日志行（"HH:mm:ss.fff [Level] message"）提取异常条目，按事件签名去重后最多 30 条。</summary>
        public static List<string> Extract(List<string> lines)
        {
            var issues = new List<string>();
            var seen = new HashSet<string>();
            foreach (var line in lines ?? new List<string>())
            {
                // 心跳行降噪：「回合开始 工具上限=… 空闲超时=…」是每回合固定 Info 回显，不是异常。
                if (line.IndexOf("回合开始", StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                if (line.IndexOf("[Error]", StringComparison.Ordinal) >= 0 || Patterns.Any(line.Contains))
                {
                    var trimmed = line.Length > 260 ? line.Substring(0, 260) + "…" : line;
                    if (seen.Add(Signature(line)))
                    {
                        issues.Add(trimmed);
                    }
                }
            }

            return issues.Take(30).ToList();
        }

        /// <summary>事件签名 = 归一化错误码 + 10 秒时间桶。同桶同码视为同一故障的多层/多副本落笔。</summary>
        private static string Signature(string line)
        {
            var bucket = "";
            var tm = TimestampRx.Match(line);
            if (tm.Success)
            {
                var parts = tm.Groups[1].Value.Split(':');
                if (parts.Length == 3 && int.TryParse(parts[2], out var sec))
                {
                    bucket = parts[0] + ":" + parts[1] + ":" + (sec / 10 * 10).ToString("00");
                }
            }

            return ClassifyCode(line) + "|" + bucket;
        }

        /// <summary>归一化错误码：worker 层失败行（"执行失败 exit=1: Traceback…"）不带错误码 token，
        /// 归并到 PYTHON_RUN_ERROR，使同一次 Python 失败的三层行文共享同一签名。</summary>
        private static string ClassifyCode(string line)
        {
            if (line.IndexOf("Traceback", StringComparison.Ordinal) >= 0 &&
                (line.IndexOf("执行失败", StringComparison.Ordinal) >= 0 ||
                 line.IndexOf("PYTHON_RUN_ERROR", StringComparison.Ordinal) >= 0))
            {
                return "PYTHON_RUN_ERROR";
            }

            var best = -1;
            var code = "[Error]";
            foreach (var p in Patterns)
            {
                var i = line.IndexOf(p, StringComparison.Ordinal);
                if (i >= 0 && (best < 0 || i < best))
                {
                    best = i;
                    code = p;
                }
            }

            return code;
        }
    }
}
