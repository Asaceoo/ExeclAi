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
            // 按段判断 '..'（不限"含反斜杠"才检查）：单段 ".." 归一化后会落到父目录，
            // 旧实现只在含 '\' 时检查，GuardOutputPath("..") 可直接穿越。
            if (Array.IndexOf(trimmed.Split('\\'), "..") >= 0)
            {
                throw new ArgumentException("输出路径不允许包含 '..' 路径成分（防止穿越逃逸）。", nameof(raw));
            }

            if (raw.IndexOf('\0') >= 0)
            {
                throw new ArgumentException("输出路径含非法字符（NUL）。", nameof(raw));
            }

            // M28 审查：必须返回绝对规范化路径——SaveAs/Export 由 Excel 进程按自身 CWD 落盘，
            // File.Exists/Directory.CreateDirectory 由 server 进程按自身 CWD 判断；相对路径会让
            // 覆盖门与自检落在不同目录（误拒/静默覆盖）。统一在此解析，调用方一律用绝对值。
            var full = Path.GetFullPath(trimmed);

            // M28 审查（第三轮）：拒绝 UNC/设备路径与网络驱动器——File.Exists/导出会向远端发起
            // SMB/NTLM 外连（映射盘 Z: 形如本地盘符，必须按卷类型判定，不能只看前缀）。
            if (full.StartsWith("\\\\", StringComparison.Ordinal))
            {
                throw new ArgumentException("输出路径不允许 UNC/设备路径（防凭据外连）：" + full, nameof(raw));
            }

            try
            {
                var root = Path.GetPathRoot(full);
                if (string.IsNullOrWhiteSpace(root) ||
                    new DriveInfo(root).DriveType != DriveType.Fixed)
                {
                    throw new ArgumentException("输出路径仅支持本地固定磁盘（拒绝网络盘/光驱/内存盘）：" + full, nameof(raw));
                }
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new ArgumentException("输出路径卷类型无法判定，已拒绝：" + full + "（" + ex.Message + "）", nameof(raw));
            }

            return full;
        }

        /// <summary>
        /// M28（第三轮续）：覆盖写之前备份原文件（同目录 &lt;name&gt;.bak-&lt;时间戳&gt;），
        /// 让 overwrite=true 至少可人工恢复；备份失败由调用方中止写入（fail-closed）。
        /// </summary>
        public static string CreateBackup(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            var stamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var backup = filePath + ".bak-" + stamp;
            var suffix = 1;
            while (File.Exists(backup))
            {
                backup = filePath + ".bak-" + stamp + "-" + suffix;
                suffix++;
            }

            File.Copy(filePath, backup, false);
            // File.Copy 会把目标写入时间设成源文件的旧时间——不刷新排序与保留都会失真
            //（陈旧源文件的"刚建备份"会被当最旧/超期误删）。备份时间以创建时刻为准。
            try { File.SetLastWriteTimeUtc(backup, DateTime.UtcNow); } catch { }

            // M28（第三轮续）：备份保留策略——同一文件只保留最近 10 份，且清理 >30 天的旧备份，
            // 防止反复覆盖导致 .bak 无限堆积（清理失败不影响本次备份结果）。
            try
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    // M28（第四轮自查）：GetFiles 的 8.3 短名匹配会让 "a.txt.bak-*" 误配 "a.txtx.bak-*"——
                    // 改用精确前缀 + EnumerateFiles。
                    var backupPrefix = Path.GetFileName(filePath) + ".bak-";
                    var stale = Directory.EnumerateFiles(directory)
                        .Where(candidate =>
                        {
                            var name = Path.GetFileName(candidate);
                            if (!name.StartsWith(backupPrefix, StringComparison.OrdinalIgnoreCase))
                            {
                                return false;
                            }

                            // 只认本工具生成的 <yyyyMMddHHmmss>(-n)? 形态——用户自建的
                            // “报告.xlsx.bak-notes.txt” 不得被纳入清理范围。
                            var suffix = name.Substring(backupPrefix.Length);
                            return System.Text.RegularExpressions.Regex.IsMatch(suffix, "^[0-9]{14}(-[0-9]+)?$");
                        })
                        .OrderByDescending(File.GetLastWriteTimeUtc)
                        .ToList();
                    var cutoff = DateTime.UtcNow.AddDays(-30);
                    for (var index = 0; index < stale.Count; index++)
                    {
                        // 注释口径是“保留最近 10 份”：index 0..9 保留，第 11 份起才删。
                        if (index >= 10 || File.GetLastWriteTimeUtc(stale[index]) < cutoff)
                        {
                            try { File.Delete(stale[index]); } catch { }
                        }
                    }
                }
            }
            catch { }

            return backup;
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