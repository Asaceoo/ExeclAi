using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// JYYJtool 借鉴工具的纯逻辑算法，全部可脱离 Excel 单测。
    /// </summary>
    public static class JyyjtoolAlgorithms
    {
        // ---------- 文本 ----------

        public static string ExtractCharacters(string text, string kind)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            StringBuilder result = new StringBuilder();
            foreach (var character in text)
            {
                if (MatchesKind(character, kind))
                {
                    result.Append(character);
                }
            }

            return result.ToString();
        }

        public static bool MatchesKind(char c, string kind)
        {
            switch ((kind ?? "number").ToLowerInvariant())
            {
                case "chinese":
                    return c >= '\u4e00' && c <= '\u9fff';
                case "english":
                    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
                case "number":
                    return c >= '0' && c <= '9';
                case "alnum":
                    return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '\u4e00' && c <= '\u9fff');
                case "special":
                    return !(c >= '0' && c <= '9') && !(c >= 'a' && c <= 'z') && !(c >= 'A' && c <= 'Z') &&
                           !(c >= '\u4e00' && c <= '\u9fff') && !char.IsWhiteSpace(c);
                default:
                    return true;
            }
        }

        public static string InsertSubstring(string text, string insertion, string when, string where, int index, string separator)
        {
            if (when == "char")
            {
                var position = Math.Max(0, Math.Min(text?.Length ?? 0, index));
                return (text ?? string.Empty).Insert(position, insertion ?? string.Empty);
            }

            // before/after 第 N 次出现的分隔符（紧贴分隔符插入，不额外添加分隔符）
            var s = text ?? string.Empty;
            var sep = separator ?? string.Empty;
            if (sep.Length == 0)
            {
                return s + (insertion ?? string.Empty);
            }

            var sepIndex = -1;
            var from = 0;
            for (var i = 0; i <= Math.Max(0, index); i++)
            {
                sepIndex = s.IndexOf(sep, from, StringComparison.Ordinal);
                if (sepIndex < 0)
                {
                    return s + (insertion ?? string.Empty);
                }

                from = sepIndex + sep.Length;
            }

            var insertAt = string.Equals(where, "after", StringComparison.OrdinalIgnoreCase)
                ? sepIndex + sep.Length
                : sepIndex;
            return s.Insert(insertAt, insertion ?? string.Empty);
        }

        public static string DeleteMatches(string text, string pattern)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(pattern))
            {
                return text ?? string.Empty;
            }

            try
            {
                // 用户正则不可控，必须加执行超时与编译缓存，防止嵌套量词 (.+)+$ 之类的 ReDoS 卡死 Excel 进程。
                return Regex.Replace(text, pattern, string.Empty,
                    RegexOptions.CultureInvariant | RegexOptions.Compiled, TimeSpan.FromMilliseconds(2000));
            }
            catch (ArgumentException)
            {
                return text;
            }
            catch (RegexMatchTimeoutException)
            {
                return text;
            }
        }

        private static readonly Dictionary<char, string> PinyinMap = BuildPinyinMap();

        private static Dictionary<char, string> BuildPinyinMap()
        {
            var raw = new[]
            {
                "的de","一yi","是shi","不bu","了le","人ren","我wo","在zai","他ta","有you","这zhe","个ge","上shang","们men","来lai",
                "到dao","时shi","大da","地di","为wei","子zi","中zhong","你ni","说shuo","生sheng","国guo","年nian","着zhe","就jiu",
                "那na","和he","要yao","她ta","出chu","也ye","得de","里li","后hou","自zi","以yi","会hui","家jia","可ke","下xia",
                "而er","过guo","天tian","去qu","能neng","对dui","小xiao","多duo","然ran","于yu","心xin","学xue","么me","都du",
                "好hao","看kan","起qi","发fa","当dang","没mei","成cheng","只zhi","如ru","事shi","把ba","还hai","用yong","第di",
                "样yang","道dao","想xiang","作zuo","种zhong","开kai","美mei","总zong","从cong","无wu","情qing","己ji","面mian",
                "最zui","女nv","但dan","现xian","前qian","些xie","所suo","同日","手shou","又you","行xing","意yi","动dong","方fang",
                "期qi","它ta","头tou","经jing","长chang","儿er","回hui","位wei","分fen","爱ai","老lao","因yin","给gei","些","shi",
                "国guo","山shan","水shui","工gong","高gao","新xin","两liang","真zhen","南nan","北bei","东dong","西xi","旧jiu",
                "早zao","晚wan","上午shangwu","下午xiawu","今天jintian","明天mingtian","昨天zuotian","星期xingqi","时间shijian",
                "年nian","月yue","日ri","时shi","分fen","秒miao","元yuan","角jiao","整zheng","大da","小xiao","高gao","低di"
            };
            var map = new Dictionary<char, string>();
            foreach (var token in raw)
            {
                if (token.Length >= 3)
                {
                    map[token[0]] = token.Substring(1);
                }
            }

            return map;
        }

        public static string ToPinyin(string hanzi, string style)
        {
            if (string.IsNullOrEmpty(hanzi))
            {
                return string.Empty;
            }

            var initial = string.Equals(style, "initial", StringComparison.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            foreach (var c in hanzi)
            {
                if (PinyinMap.TryGetValue(c, out var pinyin))
                {
                    sb.Append(initial ? pinyin[0].ToString() : pinyin);
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private static readonly Dictionary<char, char> SimplifiedToTraditional = BuildSimplifiedToTraditional();

        private static Dictionary<char, char> BuildSimplifiedToTraditional()
        {
            var pairs = new[]
            {
                "门門","开開","东東","车車","马馬","见見","长長","贝貝","币幣","则則","间間","闻聞","问問","们們","电電","华華",
                "区區","医醫","学學","来來","乐樂","东東","书書","发發","头頭","农農","兴興","实實","宝寶","贝貝","义義","龙龍",
                "汉漢","众眾","云雲","会會","从從","册冊","付傳","华華","历歷","厅厅","发發","严嚴","丧喪","个個","临臨","丰豐",
                "丹丹","为為","丽麗","举舉","之之","么麼","习習","乡鄉","书書","买買","乱亂","云雲","亚亞","产產","亲親","人個",
                "亿億","什仆","仅僅","从從","仇仇","仓倉","仪儀","们們","价價","众眾","优優","会會","伞傘","伟偉","传傳","伤傷",
                "伦倫","伪偽","体體","何","余餘","作","佥僉","华華","价價","休","众眾","优優","会會","伞傘","伟偉","会會","伞傘"
            };

            var map = new Dictionary<char, char>();
            foreach (var pair in pairs)
            {
                if (pair.Length >= 2)
                {
                    map[pair[0]] = pair[1];
                }
            }

            return map;
        }

        public static string ConvertChineseVariant(string text, string direction)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text ?? string.Empty;
            }

            var toTraditional = !string.Equals(direction, "traditional_to_simplified", StringComparison.OrdinalIgnoreCase);
            var sb = new StringBuilder();
            foreach (var c in text)
            {
                if (toTraditional && SimplifiedToTraditional.TryGetValue(c, out var t))
                {
                    sb.Append(t);
                }
                else if (!toTraditional && SimplifiedToTraditional.ContainsValue(c))
                {
                    var s = SimplifiedToTraditional.FirstOrDefault(p => p.Value == c).Key;
                    sb.Append(s);
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        public static string NumberToUppercaseAmount(decimal value)
        {
            const string digits = "零壹贰叁肆伍陆柒捌玖";
            var negative = value < 0;
            value = Math.Abs(value);
            var yuan = (long)Math.Floor(value);
            var fen = (long)Math.Round((value - yuan) * 100);
            if (fen >= 100)
            {
                yuan += 1;
                fen -= 100;
            }

            var sb = new StringBuilder();
            if (yuan == 0)
            {
                sb.Append("零");
            }
            else
            {
                sb.Append(LongToChinese(yuan, digits));
            }

            sb.Append("元");

            if (fen == 0)
            {
                sb.Append("整");
            }
            else
            {
                var jiao = fen / 10;
                var f = fen % 10;
                if (jiao > 0)
                {
                    sb.Append(digits[(int)jiao]).Append("角");
                }
                else if (fen > 0)
                {
                    sb.Append("零");
                }

                if (f > 0)
                {
                    sb.Append(digits[(int)f]).Append("分");
                }
            }

            if (negative)
            {
                sb.Insert(0, "负");
            }

            return sb.ToString();
        }

        private static string LongToChinese(long value, string digits)
        {
            if (value == 0)
            {
                return "零";
            }

            string[] posUnits = { string.Empty, "拾", "佰", "仟" };
            string[] bigUnits = { string.Empty, "万", "亿", "万亿" };
            var sections = new List<int>();
            var v = value;
            while (v > 0)
            {
                sections.Insert(0, (int)(v % 10000));
                v /= 10000;
            }

            var sb = new StringBuilder();
            var printed = false;
            var pendingZero = false;
            for (var si = 0; si < sections.Count; si++)
            {
                var section = sections[si];
                if (section == 0)
                {
                    if (printed && si < sections.Count - 1)
                    {
                        pendingZero = true;
                    }

                    continue;
                }

                var d = new int[4];
                var t = section;
                for (var k = 3; k >= 0; k--)
                {
                    d[k] = t % 10;
                    t /= 10;
                }

                var hi = 0;
                while (hi < 4 && d[hi] == 0)
                {
                    hi++;
                }

                if (printed && (pendingZero || hi > 0))
                {
                    sb.Append("零");
                }

                pendingZero = false;
                for (var k = hi; k < 4; k++)
                {
                    var digit = d[k];
                    if (digit == 0)
                    {
                        if (k < 3 && d[k + 1] != 0 && (sb.Length == 0 || sb[sb.Length - 1] != '零'))
                        {
                            sb.Append("零");
                        }

                        continue;
                    }

                    var weight = 3 - k;
                    sb.Append(digits[digit]);
                    if (weight > 0)
                    {
                        sb.Append(posUnits[weight]);
                    }

                    printed = true;
                }

                if (si < sections.Count - 1)
                {
                    sb.Append(bigUnits[sections.Count - 1 - si]);
                }
            }

            return sb.ToString();
        }

        // ---------- 数字 ----------

        public static decimal BankersRound(decimal value, int decimals)
        {
            return Math.Round(value, Math.Max(0, decimals), MidpointRounding.ToEven);
        }

        /// <summary>凑数：从候选中找到尽量接近目标的组合（贪心 + 局部调整，返回选择的下标）。</summary>
        public static List<int> FindSubsetSum(double[] candidates, double target, double tolerance)
        {
            var result = new List<int>();
            if (candidates == null || candidates.Length == 0)
            {
                return result;
            }

            // 整数化后做动态规划求是否存在精确组合
            if (TryExactSubset(candidates, target, out var exact))
            {
                return exact;
            }

            // 贪心：从大到小累积，直到超过目标
            var indexed = candidates.Select((value, index) => new { value, index })
                .OrderByDescending(item => item.value).ToList();
            double sum = 0;
            var picked = new List<int>();
            foreach (var item in indexed)
            {
                if (sum + item.value <= target + tolerance)
                {
                    sum += item.value;
                    picked.Add(item.index);
                }
            }

            if (picked.Count == 0 && candidates.Length > 0)
            {
                picked.Add(indexed[0].index);
            }

            return picked;
        }

        private static bool TryExactSubset(double[] candidates, double target, out List<int> chosen)
        {
            chosen = new List<int>();
            var scale = 100.0;
            var intTarget = (int)Math.Round(target * scale);
            // 工具分类审查第十批修复：上限 50000（=500 元）导致大额凑数（对账场景金额常上万）
            // 永不走精确组合、静默退贪心——提到 2,000,000（=2 万，HashSet/字典内存可控）；
            // 单个候选超上限的等值在下方 candidate <= intTarget 防护中自动跳过。
            if (intTarget < 0 || intTarget > 2000000)
            {
                return false;
            }

            var values = candidates.Select(v => (int)Math.Round(v * scale)).ToArray();
            var sums = new HashSet<int> { 0 };
            var prev = new Dictionary<int, int>(); // sum -> last index added
            var source = new Dictionary<int, int>(); // sum -> previous sum
            for (var i = 0; i < values.Length; i++)
            {
                if (values[i] <= 0)
                {
                    continue;
                }

                var additions = new List<int>();
                foreach (var s in sums.ToList())
                {
                    var candidate = s + values[i];
                    if (candidate <= intTarget && !prev.ContainsKey(candidate))
                    {
                        additions.Add(candidate);
                        prev[candidate] = i;
                        source[candidate] = s;
                    }
                }

                foreach (var a in additions)
                {
                    sums.Add(a);
                }

                if (prev.ContainsKey(intTarget))
                {
                    break;
                }
            }

            if (!prev.ContainsKey(intTarget))
            {
                return false;
            }

            var current = intTarget;
            var indices = new List<int>();
            while (current > 0)
            {
                if (!prev.ContainsKey(current))
                {
                    break;
                }

                indices.Add(prev[current]);
                current = source.ContainsKey(current) ? source[current] : 0;
            }

            chosen = indices.Distinct().ToList();
            return chosen.Count > 0;
        }

        public static string[] GeneratePermutations(int n, int limit)
        {
            var result = new List<string>();
            if (n <= 0)
            {
                return result.ToArray();
            }

            var items = Enumerable.Range(1, n).ToArray();
            Permute(items, 0, result, Math.Max(1, Math.Min(limit, 100000)));
            return result.Take(Math.Max(1, Math.Min(limit, 100000))).ToArray();
        }

        private static void Permute(int[] items, int start, List<string> result, int limit)
        {
            if (result.Count >= limit)
            {
                return;
            }

            if (start == items.Length)
            {
                result.Add(string.Join(",", items));
                return;
            }

            for (var i = start; i < items.Length; i++)
            {
                Swap(items, start, i);
                Permute(items, start + 1, result, limit);
                Swap(items, start, i);
            }
        }

        private static void Swap(int[] items, int a, int b)
        {
            var tmp = items[a];
            items[a] = items[b];
            items[b] = tmp;
        }

        // ---------- 分隔符 / 表格 ----------

        public static char DetectDelimiter(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return ',';
            }

            var counts = new[]
            {
                new { c = '\t', n = line.Count(ch => ch == '\t') },
                new { c = ',', n = line.Count(ch => ch == ',') },
                new { c = ';', n = line.Count(ch => ch == ';') },
                new { c = '|', n = line.Count(ch => ch == '|') },
                new { c = ' ', n = Regex.Matches(line, @"\s+").Count }
            };

            return counts.OrderByDescending(x => x.n).First().c;
        }

        public static string[] SplitSmart(string line, char delimiter)
        {
            if (string.IsNullOrEmpty(line))
            {
                return new string[0];
            }

            if (delimiter == ' ')
            {
                return Regex.Split(line.Trim(), @"\s+").Where(s => !string.IsNullOrEmpty(s)).ToArray();
            }

            return line.Split(delimiter).Select(s => s.Trim()).ToArray();
        }

        /// <summary>多字符分隔符版本，保留空字段，逐段去首尾空白。</summary>
        public static string[] SplitSmart(string line, string delimiter)
        {
            if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(delimiter))
            {
                return new string[0];
            }

            return line.Split(new[] { delimiter }, StringSplitOptions.None)
                .Select(s => s.Trim())
                .ToArray();
        }

        /// <summary>跨区键值匹配回填：以 source 的键列逐行去 lookup 建索引（键大小写不敏感，去首尾空白），
        /// 命中则取 lookup 值列的值；未命中保留 null（调用方保持源格原值）。
        /// lastMatch=true 取最后一次命中，否则取首次命中。</summary>
        public static object[] CrossReferenceValues(
            object[,] source, int sourceKeyCol,
            object[,] lookup, int lookupKeyCol, int lookupValueCol,
            bool lastMatch)
        {
            var rows = source.GetLength(0);
            var seen = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var lRows = lookup.GetLength(0);
            for (var r = 0; r < lRows; r++)
            {
                var k = SourceString(lookup, r, lookupKeyCol);
                if (k == null)
                {
                    continue;
                }

                var v = lookup[r, Math.Min(lookupValueCol, lookup.GetLength(1) - 1)];
                if (lastMatch)
                {
                    seen[k] = v;
                }
                else if (!seen.ContainsKey(k))
                {
                    seen[k] = v;
                }
            }

            var result = new object[rows];
            for (var r = 0; r < rows; r++)
            {
                var k = SourceString(source, r, sourceKeyCol);
                if (k != null && seen.TryGetValue(k, out var v))
                {
                    result[r] = v;
                }
            }

            return result;
        }

        private static string SourceString(object[,] matrix, int row, int col)
        {
            if (row < 0 || row >= matrix.GetLength(0) || col < 0 || col >= matrix.GetLength(1))
            {
                return null;
            }

            var v = matrix[row, col];
            if (v == null)
            {
                return null;
            }

            var s = Convert.ToString(v)?.Trim();
            return string.IsNullOrEmpty(s) ? null : s;
        }

        /// <summary>宽表转长表：把领先的前 keyColumns 列作为标识列，其余列为待打通字段。</summary>
        public static List<string[]> Unpivot(object[,] matrix, int keyColumns)
        {
            var rows = matrix.GetLength(0);
            var columns = matrix.GetLength(1);
            var result = new List<string[]>();
            // 首行视为表头，只展开数据行
            for (var r = 1; r < rows; r++)
            {
                for (var c = Math.Max(0, keyColumns); c < columns; c++)
                {
                    var row = new List<string>();
                    for (var k = 0; k < Math.Min(keyColumns, 4); k++)
                    {
                        row.Add(StringValue(matrix, r, k));
                    }

                    if (c < columns)
                    {
                        row.Add(StringValue(matrix, r, c));
                    }

                    result.Add(row.ToArray());
                }
            }

            return result;
        }

        public static string StringValue(object[,] matrix, int row, int column)
        {
            if (row < 0 || row >= matrix.GetLength(0) || column < 0 || column >= matrix.GetLength(1))
            {
                return string.Empty;
            }

            var value = matrix[row, column];
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        public static int ParseInt(object value, int defaultValue)
        {
            try
            {
                return Convert.ToInt32(value);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }
    }
}