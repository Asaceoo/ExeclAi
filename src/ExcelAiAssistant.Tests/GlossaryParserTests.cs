using System.Collections.Generic;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>FR-K1 业务口径规则文件解析（GlossaryParser）：每行一条「- 术语: 定义」。</summary>
    public class GlossaryParserTests
    {
        [Fact]
        public void Parses_standard_entries()
        {
            var lines = new[]
            {
                "# 口径",
                "- 毛利: 收入-成本（不含税，不含退货）",
                "- 活跃用户：当日有登录且会话时长≥5分钟"
            };
            var entries = GlossaryParser.Parse(lines);
            Assert.Equal(2, entries.Count);
            Assert.Equal("毛利", entries[0].Term);
            Assert.Equal("收入-成本（不含税，不含退货）", entries[0].Definition);
            Assert.Equal("活跃用户", entries[1].Term);
        }

        [Fact]
        public void Accepts_asterisk_bullet_and_strips_whitespace()
        {
            var entries = GlossaryParser.Parse(new[] { "*  大区 :  仅华东/华北/华南/西南  " });
            Assert.Single(entries);
            Assert.Equal("大区", entries[0].Term);
            Assert.Equal("仅华东/华北/华南/西南", entries[0].Definition);
        }

        [Fact]
        public void Skips_headings_code_and_empty_lines()
        {
            var entries = GlossaryParser.Parse(new[]
            {
                "# 标题",
                "",
                "```",
                "- 代码块里的条目: 简化解析不识别代码块围栏，这行会被误收——文档约定口径文件不要放代码块",
                "```",
                "普通段落文字"
            });
            // 简化解析：仅列表行成条目（有冒号即算）；「```」「普通段落文字」不是列表行 → 跳过。
            // 已知局限：代码块内带冒号的列表行会被误收——FR-K1 格式规范明确口径文件不放代码块。
            Assert.Single(entries);
            Assert.Equal("代码块里的条目", entries[0].Term);
        }

        [Fact]
        public void Max_entries_cap_is_enforced()
        {
            var lines = new List<string>();
            for (var i = 0; i < 150; i++)
            {
                lines.Add("- 术语" + i + ": 定义" + i);
            }

            Assert.Equal(100, GlossaryParser.Parse(lines).Count);
            Assert.Equal(10, GlossaryParser.Parse(lines, 10).Count);
        }

        [Fact]
        public void Empty_or_malformed_lines_are_skipped()
        {
            var entries = GlossaryParser.Parse(new[] { "", "   ", "- 缺定义", "没有列表符: 不是条目", null });
            Assert.Empty(entries);
        }

        [Fact]
        public void Null_lines_is_safe()
        {
            Assert.Empty(GlossaryParser.Parse(null));
        }
    }
}
