using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Workflow
{
    /// <summary>
    /// 流程"去作用区域"：把步骤参数里固化的 A1 区域字面量改写为 ``{selection}`` 占位符，
    /// 使流程在执行时取"当前选区"动态解析，从而复用任意数据区域（纯逻辑，可单测）。
    /// 只处理自定义（非内置）流程；改写保守，仅命中"确实像区域字面量"之处。
    /// </summary>
    public static class FlowScopeNormalizer
    {
        public const string SelectionMarker = "{selection}";
        public const string SelectionSheetMarker = "{selection_sheet}";
        /// <summary>「插入选区」作用区域标记：`'#XAI_SCOPE: <完全限定并集地址>#`，解析后一次性应用并移除。</summary>
        public const string ScopeMarkerTag = "#XAI_SCOPE:";

        // 独立 A1 区域字面量，如 $C$5:$F$22 / C5:F22 / '表'!$C$5 等。要求两侧为字母/数字/表的拼接边界，
        // 防止误伤代码片段里的随机字母数字。这里只匹配"带冒号的区域"，单格地址交给 Range 上下文规则处理。
        private static readonly Regex RangeAddressLiteral = new Regex(
            @"(?<![A-Za-z0-9_])('(?:[^']|'')+'!)?\$?[A-Z]{1,3}\$?[0-9]+\s*:\s*\$?[A-Z]{1,3}\$?[0-9]+(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        // VBA Range("...") / .Range("...") 上下文：整体把区域替换为 {selection}，保留 .Range( 前缀形式。
        private static readonly Regex VbaRangeLiteral = new Regex(
            @"(?i)(\.?Range\s*\(\s*)""((?:[^""]*!)\s*)?\$?[A-Z]{1,3}\$?[0-9]+(?:\s*:\s*\$?[A-Z]{1,3}\$?[0-9]+)?""\s*\)",
            RegexOptions.Compiled);

        // 值形如"纯 A1 区域地址"（可带 $、可带表前缀、单格或区域）整体替换；
        // 仅当值整串就是一个地址时命中，避免误伤段落/代码里的随机字母数字。
        private static readonly Regex PureAddressValue = new Regex(
            @"^\s*(?:'(?:[^']|'')+'!\s*)?\$?[A-Z]{1,3}\$?[0-9]+(?:\s*:\s*\$?[A-Z]{1,3}\$?[0-9]+)?\s*$",
            RegexOptions.Compiled);

        // 「插入选区」作用区域解析用途：
        // 捕获任意 Range("...")，内部地址允许带工作表前缀 + 单个/多面积（逗号分隔），$ 与全限定皆可。
        private static readonly Regex VbaRangeAny = new Regex(
            @"(?i)(\.?Range\s*\(\s*)""(?<addr>[^""\r\n]*)""\s*\)",
            RegexOptions.Compiled);

        // 提取 #XAI_SCOPE: 后的地址串（到行尾；去掉末尾的收尾 #）。
        private static readonly Regex ScopeMarkerExtract = new Regex(
            @"#XAI_SCOPE:(?<addr>[^\r\n]*)",
            RegexOptions.Compiled);

        // 移除整条含 #XAI_SCOPE: 的标记行（含其换行）。
        private static readonly Regex ScopeMarkerLineRemove = new Regex(
            @"(?m)^[^\r\n]*#XAI_SCOPE:[^\r\n]*(\r?\n|$)",
            RegexOptions.Compiled);

        /// <summary>就地改写模板步骤参数中的区域字面量；仅自定义流程生效。返回改写次数。</summary>
        public static int NormalizeTemplate(WorkflowTemplate template)
        {
            if (template == null || template.Steps == null || template.Steps.Count == 0) return 0;
            if (template.IsBuiltIn) return 0;

            int total = 0;
            foreach (var step in template.Steps)
            {
                if (step == null) continue;
                if (step.Arguments != null)
                {
                    foreach (var key in new List<string>(step.Arguments.Keys))
                    {
                        var value = step.Arguments[key];
                        var rewritten = RewriteValue(value, key);
                        if (rewritten != null)
                        {
                            step.Arguments[key] = rewritten.NewValue;
                            total += rewritten.Count;
                        }
                    }
                }

                if (step.Verify != null && step.Verify.Arguments != null)
                {
                    foreach (var key in new List<string>(step.Verify.Arguments.Keys))
                    {
                        var value = step.Verify.Arguments[key];
                        var rewritten = RewriteValue(value, key);
                        if (rewritten != null)
                        {
                            step.Verify.Arguments[key] = rewritten.NewValue;
                            total += rewritten.Count;
                        }
                    }
                }
            }

            return total;
        }

        private sealed class RewriteResult
        {
            public object NewValue;
            public int Count;
        }

        /// <summary>对单个参数值进行"去作用域"改写；无改写返回 null。</summary>
        private static RewriteResult RewriteValue(object value, string key)
        {
            switch (value)
            {
                case string text:
                    var rewritten = RewriteString(text, key);
                    if (rewritten == null || rewritten.Count == 0) return null;
                    return new RewriteResult { NewValue = rewritten.NewValue, Count = rewritten.Count };

                case IDictionary<string, object> map:
                    int mapCount = 0;
                    var newMap = new Dictionary<string, object>(map, StringComparer.OrdinalIgnoreCase);
                    foreach (var k in new List<string>(newMap.Keys))
                    {
                        var inner = RewriteValue(newMap[k], k);
                        if (inner != null) { newMap[k] = inner.NewValue; mapCount += inner.Count; }
                    }
                    if (mapCount == 0) return null;
                    return new RewriteResult { NewValue = newMap, Count = mapCount };

                case System.Collections.IEnumerable sequence when !(value is string):
                    int seqCount = 0;
                    var newList = new List<object>();
                    foreach (var item in sequence)
                    {
                        var inner = RewriteValue(item, key);
                        newList.Add(inner != null ? inner.NewValue : item);
                        seqCount += inner != null ? inner.Count : 0;
                    }
                    if (seqCount == 0) return null;
                    return new RewriteResult { NewValue = newList, Count = seqCount };

                default:
                    return null;
            }
        }

        private static RewriteResult RewriteString(string text, string key)
        {
            if (string.IsNullOrEmpty(text)) return null;

            // 1) 整串就是纯 A1 区域地址（可带表前缀/美元$）→ 整体替换为 {selection}
            if (PureAddressValue.IsMatch(text))
            {
                return new RewriteResult { NewValue = SelectionMarker, Count = 1 };
            }

            // 2) VBA Range("...") / .Range("...") 上下文整体替换为 Range("{selection}")
            if (!VbaRangeLiteral.IsMatch(text)) return null;
            int count = 0;
            var output = VbaRangeLiteral.Replace(text, m =>
            {
                count++;
                return m.Groups[1].Value + "\"" + SelectionMarker + "\")";
            });

            return count > 0 && !string.Equals(output, text, StringComparison.Ordinal)
                ? new RewriteResult { NewValue = output, Count = count }
                : null;
        }

        /// <summary>对自由 VBA 代码去作用域：把固化区域改写为 {selection} 占位符。返回新代码；无改写时返回原串。</summary>
        public static string DescopeVbaCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return code ?? string.Empty;
            var result = RewriteString(code, "snippet");
            return result != null ? Convert.ToString(result.NewValue) : code;
        }

        /// <summary>判断某流程步骤参数字符串是否引用选区标记（供执行期预算选区用）。</summary>
        public static bool ReferencesSelection(object value)
        {
            switch (value)
            {
                case string s:
                    return s != null &&
                           (s.IndexOf(SelectionMarker, StringComparison.Ordinal) >= 0 ||
                            s.IndexOf(SelectionSheetMarker, StringComparison.Ordinal) >= 0);
                case IDictionary<string, object> map:
                    foreach (var v in map.Values) if (ReferencesSelection(v)) return true;
                    return false;
                case System.Collections.IEnumerable seq when !(value is string):
                    foreach (var item in seq) if (ReferencesSelection(item)) return true;
                    return false;
                default:
                    return false;
            }
        }

        /// <summary>判断模板是否有任何步骤（含注入自检）真正引用选区标记。</summary>
        public static bool TemplateReferencesSelection(WorkflowTemplate template)
        {
            if (template == null || template.Steps == null) return false;
            foreach (var step in template.Steps)
            {
                if (step == null) continue;
                if (step.Arguments != null)
                {
                    foreach (var value in step.Arguments.Values)
                    {
                        if (ReferencesSelection(value)) return true;
                    }
                }

                if (step.Verify != null && step.Verify.Arguments != null)
                {
                    foreach (var value in step.Verify.Arguments.Values)
                    {
                        if (ReferencesSelection(value)) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>「插入选区」解析结果。</summary>
        public sealed class ScopeResolveOutcome
        {
            public bool HasMarker;
            public string UnionAddress;
            /// <summary>有标记但找不到可替换的数据块 Range（仅单格/无 Range）→ 调用方应报错中止。</summary>
            public bool NoDataBlock;
            public bool Applied => HasMarker && !NoDataBlock;
            public string Code;
        }

        /// <summary>代码是否含「插入选区」作用区域标记。</summary>
        public static bool HasScopeMarker(string code)
        {
            return !string.IsNullOrWhiteSpace(code)
                && code.IndexOf(ScopeMarkerTag, StringComparison.Ordinal) >= 0;
        }

        /// <summary>提取 #XAI_SCOPE: 后的完全限定并集地址。无标记或解析不出返回 false。</summary>
        public static bool TryExtractScopeAddress(string code, out string union)
        {
            union = null;
            if (!HasScopeMarker(code)) return false;
            var m = ScopeMarkerExtract.Match(code ?? string.Empty);
            if (!m.Success) return false;
            union = m.Groups["addr"].Value.Trim();
            if (union.EndsWith("#", StringComparison.Ordinal))
            {
                union = union.Substring(0, union.Length - 1).TrimEnd();
            }
            union = union.Trim();
            return union.Length > 0;
        }

        /// <summary>解析并应用「插入选区」标记：把代码里面积最大的数据块 Range 替换为并集区域，
        /// 同时顶掉 {selection}/{selection_sheet}，最后移除标记行（一次性）。无标记时原样返回。</summary>
        public static ScopeResolveOutcome ResolveScope(string code)
        {
            if (string.IsNullOrWhiteSpace(code) || !HasScopeMarker(code))
            {
                return new ScopeResolveOutcome { HasMarker = false, Code = code ?? string.Empty };
            }

            var outcome = new ScopeResolveOutcome { HasMarker = true, Code = code };
            if (!TryExtractScopeAddress(code, out var union))
            {
                outcome.NoDataBlock = true;
                return outcome;
            }
            outcome.UnionAddress = union;
            // 工作表名可能含双引号（Excel 允许），注入 VBA 双引号字符串前须按 VBA 规则转义为 ""，否则破坏字符串边界。
            var vbaSafe = union.Replace("\"", "\"\"");

            long bestCells = -1;
            Match best = null;
            foreach (Match m in VbaRangeAny.Matches(code))
            {
                if (IsInVbaCommentRegion(code, m.Index)) continue;
                var addr = m.Groups["addr"].Value;

                long cells;
                // `{selection}` 占位即"此处就是作用区域"，作为首选目标。
                if (addr.IndexOf(SelectionMarker, StringComparison.Ordinal) >= 0
                    || addr.IndexOf(SelectionSheetMarker, StringComparison.Ordinal) >= 0)
                {
                    cells = long.MaxValue / 2;
                }
                else if (!TryGetCellCount(addr, out cells))
                {
                    continue;
                }

                if (cells > 1 && cells > bestCells)
                {
                    bestCells = cells;
                    best = m;
                }
            }

            if (best == null)
            {
                outcome.NoDataBlock = true;
                return outcome;
            }

            var replacement = best.Groups[1].Value + "\"" + vbaSafe + "\")";
            var rewritten = code.Remove(best.Index, best.Length).Insert(best.Index, replacement);
            rewritten = rewritten.Replace(SelectionMarker, vbaSafe);

            var singleSheet = FirstQualifiedSheet(union);
            if (singleSheet != null)
            {
                rewritten = rewritten.Replace(SelectionSheetMarker, singleSheet.Replace("\"", "\"\""));
            }
            else
            {
                rewritten = rewritten.Replace(SelectionSheetMarker, string.Empty);
            }

            rewritten = ScopeMarkerLineRemove.Replace(rewritten, m2 => m2.Groups[1].Success ? m2.Groups[1].Value : "");
            rewritten = rewritten.TrimStart();
            outcome.Code = rewritten;
            return outcome;
        }

        /// <summary>index 所在逻辑行里、index 之前是否已有 VBA 注释符（'）→ 该 Range 属注释，不应参与替换。</summary>
        private static bool IsInVbaCommentRegion(string code, int index)
        {
            int lineStart = code.LastIndexOf('\n', index);
            lineStart = lineStart < 0 ? 0 : lineStart + 1;
            for (int i = lineStart; i < index; i++)
            {
                if (code[i] == '\'') return true;
            }
            return false;
        }

        /// <summary>计算区域地址（含多面积）的总单元格数；无法解析返回 false。</summary>
        private static bool TryGetCellCount(string addr, out long cells)
        {
            cells = 0;
            if (string.IsNullOrWhiteSpace(addr)) return false;
            foreach (var area in addr.Split(','))
            {
                if (!TryAreaCells(StripWorksheetPrefix(area), out long a)) return false;
                cells += a;
            }
            return cells > 0;
        }

        /// <summary>剥离 'Sheet Name'! 或 SheetName! 这类工作表前缀。</summary>
        private static string StripWorksheetPrefix(string addr)
        {
            var t = addr.Trim();
            var m = Regex.Match(t, @"^(?:'[^']*'|[A-Za-z_][\w.]*)!\s*(.*)$");
            return m.Success ? m.Groups[1].Value.Trim() : t.Trim();
        }

        /// <summary>计算单一面积（C5 / C5:F22）的单元格数。单格=1。</summary>
        private static bool TryAreaCells(string core, out long cells)
        {
            cells = 0;
            if (string.IsNullOrWhiteSpace(core)) return false;
            var parts = core.Split(':');
            if (parts.Length == 1)
            {
                if (!TryCell(parts[0], out _, out _)) return false;
                cells = 1;
                return true;
            }
            if (parts.Length == 2)
            {
                if (!TryCell(parts[0], out int c1, out int r1)) return false;
                if (!TryCell(parts[1], out int c2, out int r2)) return false;
                cells = (Math.Abs(r2 - r1) + 1L) * (Math.Abs(c2 - c1) + 1L);
                return cells > 0;
            }
            return false;
        }

        /// <summary>解析单一单元格地址 C5 / $C$5 → 列号与行号。</summary>
        private static bool TryCell(string s, out int col, out int row)
        {
            col = 0;
            row = 0;
            var t = s.Trim().TrimStart('$');
            if (t.Length == 0) return false;
            int i = 0;
            while (i < t.Length && char.IsLetter(t[i])) i++;
            if (i == 0 || i >= t.Length) return false;
            if (!int.TryParse(t.Substring(i).TrimStart('$'), out row) || row <= 0) return false;
            long cv = 0;
            foreach (var ch in t.Substring(0, i).TrimStart('$'))
            {
                int d = ch >= 'a' ? ch - 'a' + 1 : ch - 'A' + 1;
                cv = cv * 26 + d;
                if (cv > int.MaxValue) return false;
            }
            col = (int)cv;
            return col >= 1;
        }

        /// <summary>取并集地址中首个工作表限定名的表名（去引号），供 {selection_sheet} 替换用；无则 null。</summary>
        private static string FirstQualifiedSheet(string union)
        {
            foreach (var area in union.Split(','))
            {
                var m = Regex.Match(area.Trim(), @"^(?:'([^']*)'|([A-Za-z_][\w.]*))!");
                if (m.Success)
                {
                    return m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                }
            }
            return null;
        }
    }
}