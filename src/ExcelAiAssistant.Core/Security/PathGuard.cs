using System;
using System.IO;
using System.Linq;

namespace ExcelAiAssistant.Core.Security
{
    /// <summary>
    /// 模型可控输出路径的写盘守门：拒绝含 NUL、空串或路径穿越（`..` 段）的路径，
    /// 防止被提示注入的模型用「..\..」把报表/图表写到任意系统目录或覆盖用户文件。
    /// 仅做防御性校验，不做目录白名单，以便保留"导出到任意绝对路径"的合法用法。
    /// </summary>
    public static class PathGuard
    {
        /// <summary>校验模型可控的输出路径，返回解析后的完整路径；不合法抛 ArgumentException。</summary>
        public static string GuardOutputPath(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                throw new ArgumentException("输出路径无效。", nameof(raw));
            }

            var trimmed = raw.Trim().Replace('/', '\\');
            if (trimmed.IndexOf('\\') >= 0 &&
                trimmed.Split('\\').Contains(".."))
            {
                throw new ArgumentException("输出路径不允许包含 '..' 路径成分（防止穿越逃逸）。", nameof(raw));
            }

            if (raw.IndexOf('\0') >= 0)
            {
                throw new ArgumentException("输出路径含非法字符（NUL）。", nameof(raw));
            }

            return trimmed;
        }

        /// <summary>校验模型可控的输出路径并补全缺失的扩展名（大小写不敏感）。</summary>
        public static string GuardAndComplete(string raw, string extension)
        {
            var path = GuardOutputPath(raw);
            if (!path.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
            {
                path = path + extension;
            }

            return path;
        }
    }
}