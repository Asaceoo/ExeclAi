using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// Sidebar 内联 JS 静态审计（对应踩坑手册 §4.1：一个内联 JS 语法错误废掉整个 WebView2 侧边栏）。
    /// 从 index.html 提取 &lt;script&gt; 块，做字符串/注释感知的括号配平与引号闭合扫描。
    /// 这是启发式护栏（不能证明 JS 完全合法，但能确定性拦住「C# 拼接括号失衡」这类高发错误）。
    /// </summary>
    public static class SidebarInlineScriptAuditor
    {
        /// <summary>单个脚本块的扫描结果。</summary>
        public sealed class ScanResult
        {
            public int ScriptIndex { get; set; }

            public bool Balanced { get; set; }

            /// <summary>首个失衡位置的行号（相对 index.html 全文，1 起）。</summary>
            public int Line { get; set; }

            public string Problem { get; set; }

            /// <summary>失衡所在行的原文（截断），便于人工定位。</summary>
            public string LineText { get; set; }
        }

        /// <summary>定位仓库内的 Sidebar/index.html（从测试输出目录向上找 src 目录）。</summary>
        public static string LocateSidebarHtml()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                var candidate = Path.Combine(
                    dir.FullName, "src", "ExcelAiAssistant.Addin", "Sidebar", "index.html");
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                dir = dir.Parent;
            }

            return null;
        }

        /// <summary>提取全部内联 script 块（不含外部 src 引用）及其在全文中的起始行号。</summary>
        public static List<(string code, int startLine)> ExtractInlineScripts(string html)
        {
            var results = new List<(string, int)>();
            var matches = Regex.Matches(
                html,
                "<script\\b[^>]*>(.*?)</script>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                if (match.Groups[1].Success)
                {
                    var startLine = 1 + html.Substring(0, match.Groups[1].Index)
                        .Split('\n').Length - 1;
                    results.Add((match.Groups[1].Value, startLine));
                }
            }

            return results;
        }

        private static readonly HashSet<string> RegexPrecedingKeywords = new HashSet<string>
        {
            "return", "typeof", "instanceof", "in", "of", "new", "delete", "void", "do", "else", "case"
        };

        /// <summary>启发式判定：'/' 是否应视为正则字面量的开始（而非除号）。
        /// 向前找最近的有效字符；是 (,=:[!&amp;|?{; 等标点或行首 → 正则；
        /// 是标识符结尾时，若该标识符是 return/typeof 等关键字 → 正则，否则 → 除号。</summary>
        private static bool RegexCanStart(string code, int slashIndex)
        {
            var i = slashIndex - 1;
            while (i >= 0 && char.IsWhiteSpace(code[i]))
            {
                i--;
            }

            if (i < 0)
            {
                return true;
            }

            var ch = code[i];
            if (char.IsLetterOrDigit(ch) || ch == '_' || ch == '$')
            {
                var end = i;
                while (i >= 0 && (char.IsLetterOrDigit(code[i]) || code[i] == '_' || code[i] == '$'))
                {
                    i--;
                }

                var word = code.Substring(i + 1, end - i);
                return RegexPrecedingKeywords.Contains(word);
            }

            return ch == '(' || ch == ',' || ch == '=' || ch == ':' || ch == '[' || ch == '!'
                || ch == '&' || ch == '|' || ch == '?' || ch == '{' || ch == ';';
        }

        /// <summary>字符串/注释/正则字面量感知的括号配平与引号闭合扫描。
        /// 正则字面量按业界经典启发式识别：'/' 的前一个有效字符是 (,=:[!&amp;|?{; 或行首时视为正则开始，
        /// 否则视为除号。这是启发式护栏（不能证明 JS 完全合法，但能确定性拦住「C# 拼接括号失衡」）。</summary>
        public static ScanResult Scan(string code, int scriptIndex, int startLine)
        {
            var stack = new Stack<(char bracket, int line)>();
            var line = startLine;
            var state = "code"; // code / lineComment / blockComment / singleQuote / doubleQuote / template / regex
            var regexCharClass = false;
            var pairs = new Dictionary<char, char> { [')'] = '(', [']'] = '[', ['}'] = '{' };
            var openers = new HashSet<char> { '(', '[', '{' };

            for (var i = 0; i < code.Length; i++)
            {
                var ch = code[i];
                if (ch == '\n')
                {
                    line++;
                    if (state == "lineComment")
                    {
                        state = "code";
                    }

                    continue;
                }

                switch (state)
                {
                    case "lineComment":
                        break;
                    case "blockComment":
                        if (ch == '*' && i + 1 < code.Length && code[i + 1] == '/')
                        {
                            state = "code";
                            i++;
                        }

                        break;
                    case "singleQuote":
                    case "doubleQuote":
                    case "template":
                        if (ch == '\\')
                        {
                            i++;
                        }
                        else if ((state == "singleQuote" && ch == '\'') ||
                                 (state == "doubleQuote" && ch == '"') ||
                                 (state == "template" && ch == '`'))
                        {
                            state = "code";
                        }

                        break;
                    case "regex":
                        if (ch == '\\')
                        {
                            i++;
                        }
                        else if (ch == '[')
                        {
                            regexCharClass = true;
                        }
                        else if (ch == ']')
                        {
                            regexCharClass = false;
                        }
                        else if (ch == '/' && !regexCharClass)
                        {
                            state = "code";
                        }

                        break;
                    default:
                        if (ch == '/' && i + 1 < code.Length && code[i + 1] == '/')
                        {
                            state = "lineComment";
                            i++;
                        }
                        else if (ch == '/' && i + 1 < code.Length && code[i + 1] == '*')
                        {
                            state = "blockComment";
                            i++;
                        }
                        else if (ch == '/' && RegexCanStart(code, i))
                        {
                            state = "regex";
                            regexCharClass = false;
                        }
                        else if (ch == '\'' || ch == '"' || ch == '`')
                        {
                            state = ch == '\'' ? "singleQuote" : ch == '"' ? "doubleQuote" : "template";
                        }
                        else if (openers.Contains(ch))
                        {
                            stack.Push((ch, line));
                        }
                        else if (pairs.ContainsKey(ch))
                        {
                            if (stack.Count == 0 || stack.Peek().bracket != pairs[ch])
                            {
                                return new ScanResult
                                {
                                    ScriptIndex = scriptIndex,
                                    Balanced = false,
                                    Line = line,
                                    Problem = "多余的闭合括号 '" + ch + "'"
                                };
                            }

                            stack.Pop();
                        }

                        break;
                }
            }

            if (state == "singleQuote" || state == "doubleQuote" || state == "template")
            {
                return new ScanResult
                {
                    ScriptIndex = scriptIndex,
                    Balanced = false,
                    Line = line,
                    Problem = "字符串字面量未闭合（" + state + "）"
                };
            }

            if (state == "blockComment")
            {
                return new ScanResult
                {
                    ScriptIndex = scriptIndex,
                    Balanced = false,
                    Line = line,
                    Problem = "块注释未闭合"
                };
            }

            if (stack.Count > 0)
            {
                var top = stack.Peek();
                return new ScanResult
                {
                    ScriptIndex = scriptIndex,
                    Balanced = false,
                    Line = top.line,
                    Problem = "括号 '" + top.bracket + "' 未闭合"
                };
            }

            return new ScanResult { ScriptIndex = scriptIndex, Balanced = true, Line = startLine };
        }

        /// <summary>整份 HTML 的全部内联脚本扫描；返回全部失衡块（为空 = 通过）。</summary>
        public static List<ScanResult> Audit(string html)
        {
            var problems = new List<ScanResult>();
            var scripts = ExtractInlineScripts(html);
            for (var index = 0; index < scripts.Count; index++)
            {
                var scan = Scan(scripts[index].code, index, scripts[index].startLine);
                if (!scan.Balanced)
                {
                    scan.LineText = LineTextOf(scripts[index].code, scripts[index].startLine, scan.Line);
                    problems.Add(scan);
                }
            }

            return problems;
        }

        private static string LineTextOf(string code, int startLine, int line)
        {
            var lines = code.Split('\n');
            var index = line - startLine;
            if (index < 0 || index >= lines.Length)
            {
                return null;
            }

            var text = lines[index].Trim();
            return text.Length <= 160 ? text : text.Substring(0, 160) + "…";
        }

        /// <summary>生成给人看的问题描述（供测试失败消息与构建日志）。</summary>
        public static string DescribeProblems(List<ScanResult> problems)
        {
            var builder = new StringBuilder();
            foreach (var problem in problems)
            {
                builder.Append("script#").Append(problem.ScriptIndex)
                    .Append(" 行 ").Append(problem.Line)
                    .Append("：").Append(problem.Problem);
                if (!string.IsNullOrEmpty(problem.LineText))
                {
                    builder.Append("　← ").Append(problem.LineText);
                }

                builder.AppendLine();
            }

            return builder.ToString();
        }
    }
}
