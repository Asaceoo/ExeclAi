using System;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 把 AI 生成/沉淀的 VBA 里「写死的硬编码作用域」（如 ActiveSheet.Range("B25:I36")）
    /// 自动改写为预置模板同款的可复用形态：提示用户现场用鼠标「选择要处理的数据区域」。
    /// 这样沉淀出来的工具不再绑定某一次运行的固定单元格，变成通用可复用模板。
    /// </summary>
    public static class VbaReusabilizer
    {
        // 覆盖 ActiveSheet.Range / ws.Range / .Range / Range 的硬编码单元格/矩形地址字面量。
        // 仅匹配纯单元格地址（字母+数字，可带 $，可带 : 表示矩形），不匹配 Range("A" & r) 这类拼接。
        private static readonly Regex HardcodedRange =
            new Regex(
                @"(?<obj>[A-Za-z_][\w.]*\.\s*)?Range\(\s*""(?<addr>(?:\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(?:\$?[A-Za-z]{1,3}\$?\d{1,7}))?)""\s*\)",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>尝试改写；无硬编码范围时原样返回。返回是否发生了改写。</summary>
        public static string MakeReusable(string code, out bool transformed)
        {
            transformed = false;
            if (string.IsNullOrWhiteSpace(code))
            {
                return code;
            }

            var addrMatch = HardcodedRange.Match(code);
            if (!addrMatch.Success)
            {
                return code;
            }

            var varName = "rngScope";
            var guard = 0;
            while (Regex.IsMatch(code, @"\b" + varName + @"\b"))
            {
                guard++;
                if (guard > 20)
                {
                    return code;
                }

                varName = "rngScope" + (guard + 1);
            }

            var injection =
                "  Dim " + varName + " As Range\r\n" +
                "  Set " + varName + " = Application.InputBox(\"选择要处理的数据区域\", Type:=8)";

            var rewritten = HardcodedRange.Replace(code, varName);
            rewritten = InsertAfterSub(rewritten, injection);
            if (string.Equals(rewritten, code, StringComparison.Ordinal))
            {
                return code;
            }

            transformed = true;
            return rewritten;
        }

        private static string InsertAfterSub(string code, string injection)
        {
            var lines = Regex.Split(code, "\r\n|\r|\n");
            for (var i = 0; i < lines.Length; i++)
            {
                if (Regex.IsMatch(lines[i], @"^\s*Sub\s+\w+"))
                {
                    return InsertAt(lines, i + 1, injection);
                }
            }

            // 没有 Sub 声明：插到顶部（跳过可选的 Option 指令）。
            var insertAt = 0;
            while (insertAt < lines.Length &&
                   Regex.IsMatch(lines[insertAt], @"^\s*Option\s+\w+"))
            {
                insertAt++;
            }

            return InsertAt(lines, insertAt, injection);
        }

        private static string InsertAt(string[] lines, int index, string injection)
        {
            var result = new System.Collections.Generic.List<string>(lines);
            var injectLines = Regex.Split(injection, "\r\n|\r|\n");
            result.InsertRange(index, injectLines);
            return string.Join("\r\n", result);
        }
    }
}