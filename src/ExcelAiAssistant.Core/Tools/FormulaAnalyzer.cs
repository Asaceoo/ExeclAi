using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>参数来源类别。表达式为复合运算，供解释层逐层核对。</summary>
    public enum ArgKind
    {
        Constant,        // 字符串/数字/布尔/错误字面量
        CellRef,         // A1 / $A$1
        RangeRef,        // A1:C3 / $A$1:$C$3
        NamedRange,      // Sales
        NestedFunc,      // SUM(B2:B5)
        Expression,      // D2*E2 / E2=""
        ExternalRef,     // Sheet2!A1 / [Book1]Sheet1!A1
        Unparseable
    }

    public sealed class ArgToken
    {
        public ArgKind Kind;
        public string Text;
    }

    public sealed class FuncNode
    {
        public string Name;
        public List<ArgToken> Args = new List<ArgToken>();
    }

    /// <summary>纯逻辑公式解析器：拆分参数 + 参数来源分类，不接触 COM，可喂给解释层下钻。</summary>
    public static class FormulaAnalyzer
    {
        private static readonly Regex CellRefRegex = new Regex(
            @"^(\$?[A-Za-z]{1,3}\$?\d{1,7})(:(\$?[A-Za-z]{1,3}\$?\d{1,7}))?$", RegexOptions.Compiled);
        private static readonly Regex NumberRegex = new Regex(
            @"^[+-]?(\d+(\.\d+)?([eE][+-]?\d+)?)%?$", RegexOptions.Compiled);
        private static readonly Regex ErrorRegex = new Regex(
            @"^#(N/A|VALUE!|REF!|DIV/0!|NUM!|NAME\?|NULL!)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex IdentifierRegex = new Regex(
            @"^[A-Za-z_][A-Za-z0-9_.]*$", RegexOptions.Compiled);
        private static readonly Regex R1C1StyleRegex = new Regex(
            @"\bR\[?-?\d*\]?C\[?-?\d*\]?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>在函数体上按深度 1 逗号切分参数，忽略字符串字面量与嵌套括号。</summary>
        public static List<string> SplitArgs(string body)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(body)) return result;
            int depth = 0;
            bool inString = false;
            var sb = new StringBuilder();
            for (int i = 0; i < body.Length; i++)
            {
                char c = body[i];
                if (inString)
                {
                    sb.Append(c);
                    if (c == '"')
                    {
                        // 转义引号 ""
                        if (i + 1 < body.Length && body[i + 1] == '"') { sb.Append('"'); i++; continue; }
                        inString = false;
                    }
                    continue;
                }
                if (c == '"') { inString = true; sb.Append(c); continue; }
                if (c == '(') { depth++; sb.Append(c); continue; }
                if (c == ')') { if (depth > 0) depth--; sb.Append(c); continue; }
                if (c == ',' && depth == 0) { result.Add(sb.ToString().Trim()); sb.Clear(); continue; }
                sb.Append(c);
            }
            if (sb.Length > 0) result.Add(sb.ToString().Trim());
            return result;
        }

        /// <summary>将单个参数文本分类为来源类别。</summary>
        public static ArgKind Classify(string text)
        {
            string t = (text ?? string.Empty).Trim();
            if (t.Length == 0) return ArgKind.Unparseable;

            if (t[0] == '"') return t.EndsWith("\"") ? ArgKind.Constant : ArgKind.Unparseable;
            if (IsLiteralConstant(t)) return ArgKind.Constant;

            // 外部/跨表引用（含 [ 书签 或 带 !）：跨表、命名区域引用一并视为外部降级项
            if (t.IndexOf('[') >= 0 || t.IndexOf('!') >= 0) return ArgKind.ExternalRef;

            Match m = CellRefRegex.Match(t);
            if (m.Success) return m.Groups[2].Success ? ArgKind.RangeRef : ArgKind.CellRef;

            int p = t.IndexOf('(');
            if (p > 0 && IdentifierRegex.IsMatch(t.Substring(0, p)))
            {
                // 内置隐藏函数 _xlfn. 前缀无法解释，降级
                if (t.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) return ArgKind.Unparseable;
                return ArgKind.NestedFunc;
            }

            if (t.IndexOf('+') >= 0 || t.IndexOf('-') >= 0 || t.IndexOf('*') >= 0 ||
                t.IndexOf('/') >= 0 || t.IndexOf('^') >= 0 || t.IndexOf('&') >= 0 ||
                t.IndexOf('=') >= 0 || t.IndexOf('<') >= 0 || t.IndexOf('>') >= 0)
                return ArgKind.Expression;

            if (IdentifierRegex.IsMatch(t)) return ArgKind.NamedRange;
            return ArgKind.Unparseable;
        }

        /// <summary>解析顶层函数；无函数（裸表达式/常量）返回 Name=null 的单参数根。</summary>
        public static FuncNode ParseTopLevel(string formula)
        {
            string t = (formula ?? string.Empty).Trim();
            if (t.StartsWith("=", StringComparison.Ordinal)) t = t.Substring(1).Trim();
            // 整体为无函数的值引用或常量，不再尝试函数
            if (t.Length == 0 || IsLiteralConstant(t) || CellRefRegex.IsMatch(t) || isBareOperand(t))
            {
                var root = new FuncNode();
                root.Args.Add(new ArgToken { Kind = Classify(t), Text = t });
                return root;
            }

            int p = t.IndexOf('(');
            if (p > 0 && IdentifierRegex.IsMatch(t.Substring(0, p)) && t.EndsWith(")"))
            {
                var node = new FuncNode { Name = t.Substring(0, p).ToUpperInvariant() };
                string body = t.Substring(p + 1, t.Length - p - 2);
                foreach (string arg in SplitArgs(body))
                    node.Args.Add(new ArgToken { Kind = Classify(arg), Text = arg });
                return node;
            }

            var fallback = new FuncNode();
            fallback.Args.Add(new ArgToken { Kind = Classify(t), Text = t });
            return fallback;
        }

        private static bool isBareOperand(string t)
        {
            // 不含前导括号的纯表达式（如 A1+B1）不作为函数处理
            int p = t.IndexOf('(');
            return p <= 0;
        }

        private static bool IsLiteralConstant(string t)
        {
            if (NumberRegex.IsMatch(t)) return true;
            string cap = t.ToUpperInvariant();
            if (cap == "TRUE" || cap == "FALSE") return true;
            if (ErrorRegex.IsMatch(t)) return true;
            return false;
        }

        /// <summary>
        /// 从公式文本中提取所有引用地址（E6 真实变体）：
        /// A1 / $A$1 / A1:C3 / $A$1:$C$3 / C:C 整列 / '带空格表'!A1 / Sheet2!A1 / [Book]Sheet!A1。
        /// 忽略字符串字面量。遇 R1C1 样式返回 "UNPARSED:&lt;原公式&gt;" 哨兵，不静默丢弃。
        /// </summary>
        public static List<string> ExtractCellRefs(string formula)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(formula)) return result;

            // R1C1 样式与 A1 样式互斥；.Formula 恒为 A1，此分支只在外部喂入 R1C1 文本时触发
            if (R1C1StyleRegex.IsMatch(formula))
            {
                result.Add("UNPARSED:" + formula.Trim());
                return result;
            }

            var t = formula;
            int i = 0;
            while (i < t.Length)
            {
                // 跳过字符串字面量，避免把引号内的 " A1" 文本误当引用
                if (t[i] == '"')
                {
                    i = SkipStringLiteral(t, i);
                    continue;
                }

                // 单引号引住的表名：'...'!
                if (t[i] == '\'')
                {
                    int close = t.IndexOf('\'', i + 1);
                    if (close > i)
                    {
                        int after = close + 1;
                        if (after < t.Length && t[after] == '!')
                        {
                            string prefix = t.Substring(i, after - i + 1);
                            int refEnd;
                            string refText;
                            if (TryParseA1Ref(t, after + 1, out refEnd, out refText))
                            {
                                result.Add(prefix + refText);
                                i = refEnd;
                                continue;
                            }
                        }
                    }

                    i++;
                    continue;
                }

                // 跨簿：[Book]Sheet! 或 [Book]'Sheet Name'!
                if (t[i] == '[')
                {
                    int close = t.IndexOf(']', i + 1);
                    if (close > i)
                    {
                        int j = close + 1;
                        if (j < t.Length && t[j] == '!')
                        {
                            // 直接 [Book]!A1 形式少见；仍然尝试
                            string prefix = t.Substring(i, j - i + 1);
                            int refEnd;
                            string refText;
                            if (TryParseA1Ref(t, j + 1, out refEnd, out refText))
                            {
                                result.Add(prefix + refText);
                                i = refEnd;
                                continue;
                            }
                        }

                        // [Book]Sheet!A1 或 [Book]'Name'!A1
                        string sheetPrefix = TryParseSheetAfterBook(t, close + 1);
                        if (sheetPrefix != null)
                        {
                            int refStart = close + 1 + sheetPrefix.Length;
                            int refEnd;
                            string refText;
                            if (TryParseA1Ref(t, refStart, out refEnd, out refText))
                            {
                                result.Add(t.Substring(i, close + 1 - i) + sheetPrefix + refText);
                                i = refEnd;
                                continue;
                            }
                        }
                    }

                    i++;
                    continue;
                }

                // 无引号表名/定义名前缀：Name!
                if (IsColumnLetter(t[i]) || t[i] == '_')
                {
                    string name = TryParseNamePrefix(t, i, out int prefixEnd);
                    if (name != null)
                    {
                        int refEnd;
                        string refText;
                        if (TryParseA1Ref(t, prefixEnd + 1, out refEnd, out refText))
                        {
                            result.Add(name + refText);
                            i = refEnd;
                            continue;
                        }

                        i = prefixEnd + 1;
                        continue;
                    }
                }

                // 普通 A1 / 区域 / 整列引用（含 $）
                if (IsColumnLetter(t[i]) || t[i] == '$')
                {
                    int refEnd;
                    string refText;
                    if (TryParseA1Ref(t, i, out refEnd, out refText))
                    {
                        result.Add(refText);
                        i = refEnd;
                        continue;
                    }
                }

                i++;
            }

            return result;
        }

        /// <summary>从下标 i 起吞入一个标识符（A-Za-z _ 开头，后接字母/数字/._）；其后续紧跟 ! 则视为表名前缀并原样返回，否则返回 null。</summary>
        private static string TryParseNamePrefix(string s, int i, out int prefixEnd)
        {
            prefixEnd = i;
            int j = i;
            if (j < s.Length && (IsColumnLetter(s[j]) || s[j] == '_')) j++;
            while (j < s.Length && (IsColumnLetter(s[j]) || char.IsDigit(s[j]) || s[j] == '_' || s[j] == '.'))
                j++;
            // 至少一个字符 + 紧跟 '!' 才成表前缀，避免把函数名/裸标识当前缀
            if (j > i && j < s.Length && s[j] == '!')
            {
                prefixEnd = j;
                return s.Substring(i, j - i + 1);
            }

            return null;
        }

        /// <summary>在 [Book] 之后解析可选的 'Sheet Name'! 或 SheetName! 前缀。</summary>
        private static string TryParseSheetAfterBook(string s, int start)
        {
            if (start >= s.Length) return null;
            if (s[start] == '\'')
            {
                int close = s.IndexOf('\'', start + 1);
                if (close > start && close + 1 < s.Length && s[close + 1] == '!')
                {
                    return s.Substring(start, close + 2 - start);
                }

                return null;
            }

            if (IsColumnLetter(s[start]) || s[start] == '_')
            {
                int j = start;
                if (IsColumnLetter(s[j]) || s[j] == '_') j++;
                while (j < s.Length && (IsColumnLetter(s[j]) || char.IsDigit(s[j]) || s[j] == '_' || s[j] == '.'))
                    j++;
                if (j < s.Length && s[j] == '!')
                {
                    return s.Substring(start, j - start + 1);
                }
            }

            return null;
        }

        /// <summary>
        /// 自下标 start 起尝试解析一个 A1 引用：单元格、区域、整列。可带 $。返回 true 时 refText/refEnd 有效，否则 refEnd=start。
        /// </summary>
        private static bool TryParseA1Ref(string s, int start, out int refEnd, out string refText)
        {
            refEnd = start;
            refText = null;
            int i = start;
            bool absCol1 = false;
            if (i < s.Length && s[i] == '$') { absCol1 = true; i++; }
            int letterBegin = i;
            while (i < s.Length && IsColumnLetter(s[i])) i++;
            string col1 = s.Substring(letterBegin, i - letterBegin);
            if (col1.Length == 0 || col1.Length > 3) return false;

            // 可选行前 $，再跟数字 → 单元格
            int digitCol1 = i;
            bool absRow1 = false;
            if (i < s.Length && s[i] == '$') { absRow1 = true; i++; }
            if (i < s.Length && char.IsDigit(s[i]))
            {
                int rowBegin = i;
                while (i < s.Length && char.IsDigit(s[i])) i++;
                string row1 = s.Substring(rowBegin, i - rowBegin);
                string addr = CellToken(absCol1, col1, absRow1, row1);

                // 单元格区域 :$C$3
                if (i < s.Length && s[i] == ':')
                {
                    int j = i + 1;
                    bool absCol2 = false;
                    if (j < s.Length && s[j] == '$') { absCol2 = true; j++; }
                    int lb2 = j;
                    while (j < s.Length && IsColumnLetter(s[j])) j++;
                    string col2 = s.Substring(lb2, j - lb2);
                    if (col2.Length > 0 && col2.Length <= 3)
                    {
                        bool absRow2 = false;
                        if (j < s.Length && s[j] == '$') { absRow2 = true; j++; }
                        if (j < s.Length && char.IsDigit(s[j]))
                        {
                            int rb2 = j;
                            while (j < s.Length && char.IsDigit(s[j])) j++;
                            string row2 = s.Substring(rb2, j - rb2);
                            refEnd = j;
                            refText = addr + ":" + CellToken(absCol2, col2, absRow2, row2);
                            return true;
                        }
                    }
                }

                refEnd = i;
                refText = addr;
                return true;
            }

            // 意外多吞了行前 $：回退到数字列处再判定（$A$1 场景 absRow1 已在上面命中 digit，不会到此）
            if (absRow1) i = digitCol1;

            // 整列引用 C:D（列字母后紧跟冒号，第二段仅字母非数字）
            if (i < s.Length && s[i] == ':')
            {
                int j = i + 1;
                bool absCol2 = false;
                if (j < s.Length && s[j] == '$') { absCol2 = true; j++; }
                int lb2 = j;
                while (j < s.Length && IsColumnLetter(s[j])) j++;
                string col2 = s.Substring(lb2, j - lb2);
                // 第二段后不能再跟数字（否则是 C:D5 形式的非法混合）
                if (col2.Length > 0 && col2.Length <= 3 && !(j < s.Length && char.IsDigit(s[j])))
                {
                    refEnd = j;
                    refText = ColToken(absCol1, col1) + ":" + ColToken(absCol2, col2);
                    return true;
                }
            }

            return false;
        }

        private static string CellToken(bool absCol, string col, bool absRow, string row)
        {
            return (absCol ? "$" : "") + col + (absRow ? "$" : "") + row;
        }

        private static string ColToken(bool absCol, string col)
        {
            return (absCol ? "$" : "") + col;
        }

        /// <summary>判断给定文本是否为单个单元格引用（A1 / $A$1，不含区域/工作表前缀）。</summary>
        public static bool IsCellRef(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var t = text.Trim();
            return CellRefRegex.IsMatch(t) && !Regex.IsMatch(t, @":");
        }

        /// <summary>判断是否为单元格区域引用（A1:C3 / $A$1:$C$3）。</summary>
        public static bool IsRangeRef(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var t = text.Trim();
            return CellRefRegex.IsMatch(t) && t.IndexOf(':') >= 0;
        }

        /// <summary>把相对/绝对地址统一为规范键（去除 $），用于循环引用检测去重。返回 null 表示非单元格地址。</summary>
        public static string CanonicalKey(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var m = CellRefRegex.Match(text.Trim());
            if (!m.Success) return null;
            var first = m.Groups[1].Value.Replace("$", "").ToUpperInvariant();
            var third = m.Groups[3].Success ? m.Groups[3].Value.Replace("$", "").ToUpperInvariant() : null;
            return third == null ? first : first + ":" + third;
        }

        /// <summary>跳过字符串字面量（含 "" 转义），返回越过后首个下标。</summary>
        private static int SkipStringLiteral(string s, int i)
        {
            i++;
            while (i < s.Length)
            {
                if (s[i] == '"')
                {
                    if (i + 1 < s.Length && s[i + 1] == '"') { i += 2; continue; }
                    return i + 1;
                }

                i++;
            }

            return s.Length;
        }

        private static bool IsColumnLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        private static bool IsAllLetters(string s)
        {
            foreach (var c in s)
            {
                if (!IsColumnLetter(c)) return false;
            }

            return true;
        }
    }
}