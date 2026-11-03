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
            var importBody = src.Substring(importStart, Math.Min(4000, src.Length - importStart));
            Assert.Contains("SplitCsvLine(rows[r], delimiter)", importBody);
            Assert.DoesNotContain("rows[r].Split(new[] { delimiter }, StringSplitOptions.None)", importBody);
        }
    }
}
