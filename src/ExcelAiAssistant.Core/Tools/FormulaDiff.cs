using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 公式变更对比（纯逻辑，确定性，可单测）：
    /// 对比改前/改后两条公式，输出人话变更标签（D2 公式历史 diff）。
    /// 判定优先级：未变化 &gt; 新增/清除 &gt; IFERROR 增删 &gt; 仅引用方式调整 &gt; 函数增删 &gt; 参数/条件调整 &gt; 公式重写。
    /// </summary>
    public static class FormulaDiff
    {
        /// <summary>
        /// 对比两条公式（可带 = 前缀，空串表示原为空/被清除）。
        /// 返回人话标签（单个主标签，detail 已并入标签文本）。
        /// </summary>
        public static string DescribeChange(string oldFormula, string newFormula)
        {
            var oldBody = Normalize(oldFormula);
            var newBody = Normalize(newFormula);

            if (oldBody.Length == 0 && newBody.Length == 0) return "未变化";
            if (oldBody.Length == 0) return "新增公式";
            if (newBody.Length == 0) return "清除公式";
            if (string.Equals(oldBody, newBody, StringComparison.OrdinalIgnoreCase)) return "未变化";

            // IFERROR 包裹/解除（壳内公式一致才算，避免误判）
            if (!StartsWithIfError(oldBody) && StartsWithIfError(newBody))
            {
                var inner = UnwrapIfError(newBody);
                if (string.Equals(inner, oldBody, StringComparison.OrdinalIgnoreCase))
                {
                    return "增加 IFERROR 错误兜底";
                }
            }
            if (StartsWithIfError(oldBody) && !StartsWithIfError(newBody))
            {
                var inner = UnwrapIfError(oldBody);
                if (string.Equals(inner, newBody, StringComparison.OrdinalIgnoreCase))
                {
                    return "移除 IFERROR 错误兜底";
                }
            }

            // 仅引用方式调整（绝对/相对）：抹平 $ 后一致
            var oldAbs = StripAnchor(oldBody);
            var newAbs = StripAnchor(newBody);
            if (string.Equals(oldAbs, newAbs, StringComparison.OrdinalIgnoreCase))
            {
                return "仅调整引用方式（相对/绝对）";
            }

            // 函数名集合对比
            var oldFns = ExtractFunctions(oldBody);
            var newFns = ExtractFunctions(newBody);
            var added = new List<string>();
            var removed = new List<string>();
            foreach (var fn in newFns)
            {
                if (!oldFns.Contains(fn)) added.Add(fn);
            }
            foreach (var fn in oldFns)
            {
                if (!newFns.Contains(fn)) removed.Add(fn);
            }
            if (added.Count > 0 || removed.Count > 0)
            {
                var parts = new List<string>();
                if (added.Count > 0) parts.Add("新增函数 " + string.Join("/", added));
                if (removed.Count > 0) parts.Add("移除函数 " + string.Join("/", removed));
                return "函数变化：" + string.Join("，", parts);
            }

            // 函数集相同但内容不同 → 参数/条件调整
            return "参数/条件调整";
        }

        /// <summary>
        /// 归一化：去 = 前缀、去首尾空白与内部空白。
        /// 字符串字面量（"..."）内的空白是内容差异，原样保留（审查者视角：防「华东 区」vs「华东区」误判未变化）。
        /// </summary>
        private static string Normalize(string formula)
        {
            var text = (formula ?? string.Empty).Trim();
            if (text.StartsWith("=", StringComparison.Ordinal)) text = text.Substring(1);

            var sb = new System.Text.StringBuilder(text.Length);
            var inString = false;
            foreach (var ch in text)
            {
                if (ch == '"')
                {
                    inString = !inString;
                    sb.Append(ch);
                    continue;
                }
                if (!inString && (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n'))
                {
                    continue;
                }
                sb.Append(ch);
            }
            return sb.ToString();
        }

        private static bool StartsWithIfError(string body)
        {
            return body.StartsWith("IFERROR(", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>剥掉最外层 IFERROR( x, fallback ) 壳，返回 x（按括号配对找第一个顶层逗号）。</summary>
        private static string UnwrapIfError(string body)
        {
            // body 以 IFERROR( 开头
            var inner = body.Substring("IFERROR(".Length);
            var depth = 1;
            var inString = false;
            for (var i = 0; i < inner.Length; i++)
            {
                var ch = inner[i];
                if (ch == '"')
                {
                    inString = !inString;
                    continue;
                }
                if (inString) continue;
                if (ch == '(') depth++;
                else if (ch == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        // 找到配对右括号，剥壳成功；内部再找顶层逗号切掉 fallback
                        var content = inner.Substring(0, i);
                        var d2 = 0;
                        var s2 = false;
                        for (var j = 0; j < content.Length; j++)
                        {
                            var cj = content[j];
                            if (cj == '"') { s2 = !s2; continue; }
                            if (s2) continue;
                            if (cj == '(') d2++;
                            else if (cj == ')') d2--;
                            else if (cj == ',' && d2 == 0)
                            {
                                return content.Substring(0, j).Trim();
                            }
                        }
                        return content.Trim();
                    }
                }
            }
            return body; // 括号不配对（异常形态）→ 原样返回，判定落到后续分支
        }

        /// <summary>抹平绝对引用记号 $（用于「仅引用方式调整」判定）。</summary>
        private static string StripAnchor(string body)
        {
            return body.Replace("$", string.Empty);
        }

        /// <summary>提取公式中的函数名集合（大写）：标识符后紧跟 ( 且前面不是引用记号。</summary>
        private static HashSet<string> ExtractFunctions(string body)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(body, "([A-Za-z_][A-Za-z0-9_.]*)\\s*\\("))
            {
                // 排除单元格引用形态（如 A1( 不存在，但范围名后跟括号的定义名称调用也算函数，保留）
                result.Add(m.Groups[1].Value.ToUpperInvariant());
            }
            return result;
        }
    }
}
