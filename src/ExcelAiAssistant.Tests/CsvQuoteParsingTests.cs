using System;
using System.Collections.Generic;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M27R2 回归锁：CSV 引号感知切分必须在 ImportCsv 真正接上（此前重放时退回裸 Split，
    /// 导致 "苹果,红" 被切成两列）。反射调用私有实现，锁住行为。
    /// </summary>
    public class CsvQuoteParsingTests
    {
        private static string[] Split(string line, string delimiter)
        {
            var method = typeof(ExcelAiAssistant.Excel.ExcelElementOps).GetMethod("SplitCsvLine",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            Assert.NotNull(method);
            return (string[])method.Invoke(null, new object[] { line, delimiter });
        }

        private static System.Collections.Generic.List<string> SplitRecords(string text)
        {
            return InvokeRecords(text, out _);
        }

        /// <summary>M28：一次取得记录与「引号是否未闭合」状态（新签名为 (text, out bool)）。</summary>
        private static System.Collections.Generic.List<string> InvokeRecords(string text, out bool unbalancedQuotes)
        {
            var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
            System.Reflection.MethodInfo method = null;
            foreach (var candidate in typeof(ExcelAiAssistant.Excel.ExcelElementOps).GetMethods(flags))
            {
                if (candidate.Name == "SplitCsvRecords" && candidate.GetParameters().Length == 2)
                {
                    method = candidate;
                    break;
                }
            }

            Assert.NotNull(method);
            var args = new object[] { text, null };
            var records = (System.Collections.Generic.List<string>)method.Invoke(null, args);
            unbalancedQuotes = (bool)args[1];
            return records;
        }

        [Fact]
        public void Unbalanced_quote_is_reported_not_silently_swallowed()
        {
            var records = InvokeRecords("\"a,b\n1,2", out var unbalanced);
            Assert.True(unbalanced, "引号未闭合必须被标记（否则剩余全文被吞成一条记录）");
            Assert.True(records.Count >= 1);
        }

        [Fact]
        public void Mid_field_bare_quote_is_literal_and_does_not_merge_rows()
        {
            // M28 审查：游离引号（英寸号等）不开启引号段——两行各自成记录
            var records = InvokeRecords("a\"b,1\nc\"d,2\n", out var unbalanced);
            Assert.False(unbalanced);
            Assert.Equal(2, records.Count);
            Assert.Equal("a\"b,1", records[0]);
            var single = InvokeRecords("5\"屏,2", out var unbalanced2);
            Assert.False(unbalanced2);
            Assert.Single(single);
        }

        [Fact]
        public void Balanced_quotes_are_not_flagged()
        {
            InvokeRecords("\"a,b\",c", out var unbalanced);
            Assert.False(unbalanced);
        }

        [Fact]
        public void Quoted_newline_field_stays_single_record()
        {
            // M28/R3 修复锁：跨行引号字段 "a\nb" 必须留在同一条记录里（原按 \n 文本 Split 会拆成两条残缺记录）
            var records = SplitRecords("\"a\nb\",x\nplain");
            Assert.Equal(2, records.Count);
            Assert.Equal("\"a\nb\",x", records[0]);
            Assert.Equal("plain", records[1]);
            // 记录再走字段级解析：第一字段值含换行符
            var cells = Split(records[0], ",");
            Assert.Equal(2, cells.Length);
            Assert.Equal("a\nb", cells[0]);
        }

        [Fact]
        public void Quoted_field_with_CRLF_and_escaped_quotes_survives()
        {
            var records = SplitRecords("h1,h2\r\n\"line1\r\nline2\"\"\",z\r\n");
            Assert.Equal(2, records.Count);
            Assert.Equal("h1,h2", records[0]);
            Assert.Equal("\"line1\r\nline2\"\"\",z", records[1]);
            var cells = Split(records[1], ",");
            Assert.Equal(2, cells.Length);
            Assert.Equal("line1\r\nline2\"", cells[0]);
            Assert.Equal("z", cells[1]);
        }

        [Fact]
        public void Blank_interior_line_is_kept_and_trailing_newline_adds_no_record()
        {
            var records = SplitRecords("a,b\n\nc,d\n");
            Assert.Equal(3, records.Count); // 中间空行保留为空记录；尾部换行不追加空记录
            Assert.Equal("a,b", records[0]);
            Assert.Equal("", records[1]);
            Assert.Equal("c,d", records[2]);
        }

        [Fact]
        public void Empty_text_yields_no_records()
        {
            Assert.Empty(SplitRecords(""));
        }

        [Fact]
        public void Quoted_field_with_delimiter_stays_single_cell()
        {
            var cells = Split("\"苹果,红\",1200,正常", ",");
            Assert.Equal(3, cells.Length);
            Assert.Equal("苹果,红", cells[0]);
            Assert.Equal("1200", cells[1]);
            Assert.Equal("正常", cells[2]);
        }

        [Fact]
        public void Escaped_double_quote_is_unescaped()
        {
            // RFC4180：字段内的双引号用两个双引号表示。值 说"重点" 的编码是 "说""重点"""
            var cells = Split("\"说\"\"重点\"\"\",普通", ",");
            Assert.Equal(2, cells.Length);
            Assert.Equal("说\"重点\"", cells[0]);
            Assert.Equal("普通", cells[1]);
        }

        [Fact]
        public void Plain_row_columns_unchanged()
        {
            var cells = Split("a,b,c", ",");
            Assert.Equal(new[] { "a", "b", "c" }, cells);
        }

        [Fact]
        public void ImportCsv_source_uses_quote_aware_splitting()
        {
            // 源码级锁：ImportCsv 的列统计与填充都必须走 SplitCsvLine，禁止退回裸 Split
            var path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "src", "ExcelAiAssistant.Excel", "ExcelElementOps.cs");
            if (!System.IO.File.Exists(path)) { return; } // 非源码树环境跳过
            var src = System.IO.File.ReadAllText(path);
            var importStart = src.IndexOf("public ToolResult ImportCsv(");
            var importBody = src.Substring(importStart, Math.Min(5000, src.Length - importStart));
            // M28/R3：记录切分必须走 SplitCsvRecords（跨行引号字段），字段切分走 SplitCsvLine
            Assert.Contains("SplitCsvRecords(text, delimiter, out var unbalancedQuotes)", importBody);
            Assert.Contains("if (unbalancedQuotes)", importBody);
            Assert.Contains("SplitCsvLine(records[r], delimiter)", importBody);
            Assert.DoesNotContain("text.Split(new[] { \"\\r\\n\", \"\\n\" }", importBody);
            // 多字符 delimiter 必须显式拒绝，禁止静默取首字符
            Assert.Contains("delimiter.Length != 1", importBody);
            // 内存放大上限必须在建表之前判定（超限早退不残留空 sheet）
            Assert.Contains("MATRIX_TOO_LARGE", importBody);
        }
    }
}
