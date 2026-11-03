using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// workbuddy 记忆桥（M26）：让外部 AI 智能体（经 MCP/CLI）读写本仓库的
    /// .workbuddy/memory 长期记忆与 docs 层文档，实现「Excel 助手 ⇄ 项目记忆」双向连接。
    /// 纯文件逻辑，无 COM 依赖。
    /// </summary>
    public static class AgentMemoryBridge
    {
        /// <summary>默认 workbuddy 记忆根（仓库相对，可由调用方覆盖）。</summary>
        public static string ResolveMemoryRoot(string overrideRoot)
        {
            if (!string.IsNullOrWhiteSpace(overrideRoot) && Directory.Exists(overrideRoot)) return overrideRoot;
            // 从程序集位置向上找 .workbuddy
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            for (var i = 0; i < 6 && !string.IsNullOrEmpty(dir); i++)
            {
                var candidate = Path.Combine(dir, ".workbuddy", "memory");
                if (Directory.Exists(candidate)) return candidate;
                dir = Path.GetDirectoryName(dir);
            }
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".workbuddy", "memory");
        }

        /// <summary>读记忆：默认读 MEMORY.md；指定 date 读该日日志；指定 name 读任意文件。</summary>
        public static object Read(string memoryRoot, string name, string date, int maxChars)
        {
            var root = ResolveMemoryRoot(memoryRoot);
            var file = ResolveFile(root, name, date);
            if (file == null) return new { ok = false, error = "记忆文件不存在", root = root };
            var text = File.ReadAllText(file);
            var limit = maxChars > 0 ? maxChars : 20000;
            var truncated = text.Length > limit;
            return new
            {
                ok = true,
                file = Path.GetFileName(file),
                chars = text.Length,
                truncated = truncated,
                content = truncated ? text.Substring(0, limit) : text
            };
        }

        /// <summary>追加记忆条目（不删改既有内容）。</summary>
        public static object Append(string memoryRoot, string text, string name, string date)
        {
            if (string.IsNullOrWhiteSpace(text)) return new { ok = false, error = "text 不能为空" };
            var root = ResolveMemoryRoot(memoryRoot);
            Directory.CreateDirectory(root);
            var file = ResolveFile(root, name, date) ?? Path.Combine(root, string.IsNullOrWhiteSpace(name) ? (string.IsNullOrWhiteSpace(date) ? DateTime.Now.ToString("yyyy-MM-dd") : date) + ".md" : name);
            var entry = "\n- [" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + "] " + text.Trim() + "\n";
            File.AppendAllText(file, entry);
            return new { ok = true, file = Path.GetFileName(file), appended_chars = entry.Length };
        }

        /// <summary>搜索记忆：跨全部记忆文件做子串匹配，返回带文件名与行号的命中。</summary>
        public static object Search(string memoryRoot, string query, int limit)
        {
            if (string.IsNullOrWhiteSpace(query)) return new { ok = false, error = "query 不能为空" };
            var root = ResolveMemoryRoot(memoryRoot);
            if (!Directory.Exists(root)) return new { ok = false, error = "记忆目录不存在", root = root };
            var hits = new List<object>();
            var cap = limit > 0 ? limit : 30;
            foreach (var f in Directory.GetFiles(root, "*.md").OrderByDescending(x => File.GetLastWriteTimeUtc(x)))
            {
                var lines = File.ReadAllLines(f);
                for (var i = 0; i < lines.Length && hits.Count < cap; i++)
                {
                    if (lines[i].IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        hits.Add(new { file = Path.GetFileName(f), line = i + 1, text = lines[i].Trim().Length > 200 ? lines[i].Trim().Substring(0, 200) : lines[i].Trim() });
                    }
                }
                if (hits.Count >= cap) break;
            }
            return new { ok = true, query = query, hit_count = hits.Count, hits = hits };
        }

        private static string ResolveFile(string root, string name, string date)
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                var p = Path.Combine(root, name);
                return File.Exists(p) ? p : null;
            }
            if (!string.IsNullOrWhiteSpace(date))
            {
                var p = Path.Combine(root, date + ".md");
                return File.Exists(p) ? p : null;
            }
            var mem = Path.Combine(root, "MEMORY.md");
            return File.Exists(mem) ? mem : null;
        }
    }
}
