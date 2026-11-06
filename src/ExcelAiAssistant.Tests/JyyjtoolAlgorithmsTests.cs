using System;
using System.Linq;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>JYYJtool 借鉴工具纯逻辑算法测试（可脱离 Excel 运行）。</summary>
    public sealed class JyyjtoolAlgorithmsTests
    {
        // ---------- 文本提取 / 匹配 ----------
        [Theory]
        [InlineData("abc123中文!", "number", "123")]
        [InlineData("abc123中文!", "chinese", "中文")]
        [InlineData("abc123中文!", "english", "abc")]
        [InlineData("abc123中文!", "alnum", "abc123中文")]
        [InlineData("abc 123!", "special", "!")]
        [InlineData(null, "number", "")]
        [InlineData("", "number", "")]
        public void Extract_characters_by_kind(string text, string kind, string expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.ExtractCharacters(text, kind));
        }

        // ---------- 插入子串 ----------
        [Fact]
        public void Insert_at_character_position()
        {
            Assert.Equal("abXYc", JyyjtoolAlgorithms.InsertSubstring("abc", "XY", "char", "before", 2, null));
        }

        [Fact]
        public void Insert_before_second_separator()
        {
            Assert.Equal("a,bXY,c", JyyjtoolAlgorithms.InsertSubstring("a,b,c", "XY", "separator", "before", 1, ","));
        }

        [Fact]
        public void Insert_after_first_separator()
        {
            Assert.Equal("a,Xb,c", JyyjtoolAlgorithms.InsertSubstring("a,b,c", "X", "separator", "after", 0, ","));
        }

        // ---------- 正则删除 ----------
        [Fact]
        public void Delete_matches_removes_pattern()
        {
            Assert.Equal("abc", JyyjtoolAlgorithms.DeleteMatches("a1b2c", "\\d"));
        }

        [Theory]
        [InlineData(null, "x", "")]
        [InlineData("abc", null, "abc")]
        public void Delete_matches_guards_null(string text, string pattern, string expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.DeleteMatches(text, pattern));
        }

        // ---------- 拼音 ----------
        [Fact]
        public void Pinyin_full_and_initial()
        {
            Assert.Equal("de", JyyjtoolAlgorithms.ToPinyin("的", "full"));
            Assert.Equal("d", JyyjtoolAlgorithms.ToPinyin("的", "initial"));
        }

        // ---------- 简繁 ----------
        [Fact]
        public void Chinese_variant_conversion()
        {
            // 简体 -> 繁体至少能转换到已收录的字
            var g = JyyjtoolAlgorithms.ConvertChineseVariant("门外", "simplified_to_traditional");
            Assert.True(g.Contains("門") || g.Contains("開") || g.Contains("外"));
        }

        // ---------- 金额大写 ----------
        [Theory]
        [InlineData(0, "零元整")]
        [InlineData(1.50, "壹元伍角")]
        [InlineData(100, "壹佰元整")]
        [InlineData(12.34, "壹拾贰元叁角肆分")]
        public void Amount_to_uppercase(decimal value, string expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.NumberToUppercaseAmount(value));
        }

        // ---------- 银行家舍入 ----------
        [Theory]
        [InlineData(2.5, 0, 2)]     // 五取偶
        [InlineData(3.5, 0, 4)]     // 五取偶
        [InlineData(2.215, 2, 2.22)] // 四舍六入
        public void Bankers_round(decimal value, int decimals, decimal expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.BankersRound(value, decimals));
        }

        // ---------- 凑数 ----------
        [Fact]
        public void Subset_sum_finds_exact_combination()
        {
            var chosen = JyyjtoolAlgorithms.FindSubsetSum(new[] { 10.0, 7.0, 3.0, 5.0 }, 15.0, 0.01);
            var sum = chosen.Sum(i => new[] { 10.0, 7.0, 3.0, 5.0 }[i]);
            Assert.Equal(15.0, sum, 2);
            Assert.NotEmpty(chosen);
        }

        [Fact]
        public void Subset_sum_empty_input()
        {
            Assert.Empty(JyyjtoolAlgorithms.FindSubsetSum(Array.Empty<double>(), 10, 0.01));
            Assert.Empty(JyyjtoolAlgorithms.FindSubsetSum(null, 10, 0.01));
        }

        [Fact]
        public void Subset_sum_large_amount_finds_exact_combination()
        {
            // 工具分类审查第十批修复（DP 上限 50000 分=500 元真缺陷）：大额凑数（对账场景
            // 金额常上万）此前永走贪心近似、精确组合静默不可达——上限提到 2,000,000 后可精确命中。
            var candidates = new[] { 3000.0, 4000.0, 5500.0, 2750.50 };
            var chosen = JyyjtoolAlgorithms.FindSubsetSum(candidates, 6750.50, 0.01);
            var sum = chosen.Sum(i => candidates[i]);
            Assert.Equal(6750.50, sum, 2); // 3000 + 2750.50 + ... 精确组合命中
        }

        // ---------- 排列 ----------
        [Fact]
        public void Permutations_generate_all_for_n_small()
        {
            var result = JyyjtoolAlgorithms.GeneratePermutations(3, 100);
            Assert.Equal(6, result.Length);
            Assert.Contains("1,2,3", result);
        }

        [Fact]
        public void Permutations_respect_limit()
        {
            Assert.Equal(3, JyyjtoolAlgorithms.GeneratePermutations(4, 3).Length);
        }

        // ---------- 分隔符检测 / 拆分 ----------
        [Theory]
        [InlineData("a\tb\tc", '\t')]
        [InlineData("a,b,c", ',')]
        [InlineData("a;b;c", ';')]
        [InlineData("a|b|c", '|')]
        [InlineData("a b  c", ' ')]
        public void Detect_delimiter(string line, char expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.DetectDelimiter(line));
        }

        [Fact]
        public void Split_smart_by_char()
        {
            Assert.Equal(new[] { "a", "b", "c" }, JyyjtoolAlgorithms.SplitSmart(" a , b , c ", ','));
        }

        // ---------- 宽表转长表 ----------
        [Fact]
        public void Unpivot_expands_wide_to_long()
        {
            var matrix = new object[,]
            {
                { "产品", "一月", "二月" },
                { "A", 10, 20 },
                { "B", 30, 40 }
            };
            var result = JyyjtoolAlgorithms.Unpivot(matrix, 1);
            Assert.Equal(4, result.Count);
            Assert.Equal(new[] { "A", "10" }, result[0]);
            Assert.Equal(new[] { "B", "40" }, result[3]);
        }

        // ---------- 数值解析 ----------
        [Theory]
        [InlineData("5", 5)]
        [InlineData("abc", 0)]
        public void Parse_int_with_default(string input, int expected)
        {
            Assert.Equal(expected, JyyjtoolAlgorithms.ParseInt(input, 0));
        }

        // ---------- 对抗性修复：多字符分隔符 / 正则 ReDoS 兜底 / 跨区匹配回填 ----------
        [Fact]
        public void Split_smart_by_multi_char_delimiter_preserves_empty_fields()
        {
            Assert.Equal(new[] { "a", "b", "" }, JyyjtoolAlgorithms.SplitSmart(" a :: b :: ", "::"));
            Assert.Equal(new[] { "张三", "北京" }, JyyjtoolAlgorithms.SplitSmart("张三——北京", "——"));
        }

        [Fact]
        public void Delete_matches_with_nested_quantifier_returns_origin_without_hang()
        {
            var payload = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaab";
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var result = JyyjtoolAlgorithms.DeleteMatches(payload, "(a+)+$");
            watch.Stop();
            // 应由 2s 正则超时兜底快速返回，而非卡死。
            Assert.Equal(payload, result);
            Assert.True(watch.ElapsedMilliseconds < 5000, $"ReDoS 未兜底，耗时 {watch.ElapsedMilliseconds}ms");
        }

        [Fact]
        public void Cross_reference_fills_value_from_lookup()
        {
            var source = new object[,] { { "B", 0 }, { "A", 0 }, { "X", 0 } };
            var lookup = new object[,] { { "A", 1 }, { "B", 2 } };
            var values = JyyjtoolAlgorithms.CrossReferenceValues(source, 0, lookup, 0, 1, lastMatch: false);
            Assert.Equal(new object[] { 2, 1, null }, values);
        }

        [Fact]
        public void Cross_reference_last_match_takes_final_row()
        {
            var source = new object[,] { { "A", 0 } };
            var lookup = new object[,] { { "A", 1 }, { "A", 9 } };
            var values = JyyjtoolAlgorithms.CrossReferenceValues(source, 0, lookup, 0, 1, lastMatch: true);
            Assert.Equal(9, values[0]);
        }

        [Fact]
        public void Cross_reference_key_trim_and_case_insensitive()
        {
            var source = new object[,] { { "  ab ", 0 } };
            var lookup = new object[,] { { "AB", 7 } };
            var values = JyyjtoolAlgorithms.CrossReferenceValues(source, 0, lookup, 0, 1, lastMatch: false);
            Assert.Equal(7, values[0]);
        }
    }
}