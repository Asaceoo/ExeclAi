using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// FR-K1 业务口径规则文件解析：工作簿同目录《_口径.md》，每行一条「- 术语: 定义」
    /// （全角/半角冒号均可，* 列表符同样接受）。纯逻辑可单测。
    /// 规则：非列表行（标题/代码块/空行）不解析为条目；术语与定义去首尾空白。
    /// </summary>
    public static class GlossaryParser
    {
        private static readonly Regex EntryRx =
            new Regex(@"^\s*[-*]\s*(?<term>[^:：]+?)\s*[:：]\s*(?<definition>.+?)\s*$", RegexOptions.Compiled);

        public sealed class GlossaryEntry
        {
            public string Term { get; set; }
            public string Definition { get; set; }
        }

        /// <summary>解析口径条目；maxEntries 上限保护上下文（默认 100）。</summary>
        public static List<GlossaryEntry> Parse(IEnumerable<string> lines, int maxEntries = 100)
        {
            var entries = new List<GlossaryEntry>();
            if (lines == null)
            {
                return entries;
            }

            foreach (var line in lines)
            {
                if (entries.Count >= Math.Max(maxEntries, 0))
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var m = EntryRx.Match(line);
                if (!m.Success)
                {
                    continue;
                }

                var term = m.Groups["term"].Value.Trim();
                var definition = m.Groups["definition"].Value.Trim();
                if (term.Length == 0 || definition.Length == 0)
                {
                    continue;
                }

                entries.Add(new GlossaryEntry { Term = term, Definition = definition });
            }

            return entries;
        }
    }
}
