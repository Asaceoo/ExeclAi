using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>
    /// 公式构造器（纯逻辑，确定性，可单测）：
    /// 结构化入参 → 生成成品公式。内化三条历史教训进构造器：
    /// 1) 表头行自动钳制（header_rows，默认 1）——根除「区域含表头文本」事故家族；
    /// 2) 同表裸地址不加限定——根除「单限定引用被 Excel 规范化到第一张 sheet」事故；
    /// 3) 跨表引用统一 &#39;Sheet Name&#39;!$A$1 规范形态。
    /// </summary>
    public static class FormulaBuilder
    {
        private static readonly Regex AddrRx =
            new Regex(@"^(?:(?:'(?<qsheet>[^']+)')!|(?:(?<sheet>[^'!:]+)!))?\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d+)(?::\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+))?$", RegexOptions.Compiled);

        /// <summary>解析 A1 地址（可带工作表限定）。返回 false 表示形态非法。</summary>
        public static bool TryParseAddress(string address, out string sheet, out int col1, out int row1, out int col2, out int row2)
        {
            sheet = null; col1 = col2 = row1 = row2 = 0;
            if (string.IsNullOrWhiteSpace(address)) return false;
            var m = AddrRx.Match(address.Trim());
            if (!m.Success) return false;
            if (m.Groups["qsheet"].Success) sheet = m.Groups["qsheet"].Value;
            else if (m.Groups["sheet"].Success) sheet = m.Groups["sheet"].Value;
            col1 = ColToIndex(m.Groups["c1"].Value);
            row1 = int.Parse(m.Groups["r1"].Value, CultureInfo.InvariantCulture);
            if (m.Groups["c2"].Success)
            {
                col2 = ColToIndex(m.Groups["c2"].Value);
                row2 = int.Parse(m.Groups["r2"].Value, CultureInfo.InvariantCulture);
            }
            else
            {
                col2 = col1; row2 = row1;
            }
            return true;
        }

        /// <summary>1 基列号 → 字母（1=A, 27=AA）。</summary>
        public static string ColumnLetter(int col1Based)
        {
            var n = col1Based;
            var sb = new System.Text.StringBuilder();
            while (n > 0)
            {
                var rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        /// <summary>字母 → 1 基列号。</summary>
        public static int ColToIndex(string letters)
        {
            var n = 0;
            foreach (var ch in letters.ToUpperInvariant())
            {
                n = n * 26 + (ch - 'A' + 1);
            }
            return n;
        }

        /// <summary>构造单元格引用（可选工作表限定与绝对化）。</summary>
        public static string CellRef(int col, int row, string sheet, bool abs, bool qualifySheet)
        {
            var letter = ColumnLetter(col);
            var core = abs ? "$" + letter + "$" + row : letter + row;
            if (!qualifySheet || string.IsNullOrEmpty(sheet)) return core;
            return QuoteSheet(sheet) + "!" + core;
        }

        /// <summary>混合引用：列绝对 + 行相对（$X2），逐行模板（rank/cumulative/percent/同环比）填充联动的核心形态。</summary>
        public static string CellRefMixed(int col, int row, string sheet, bool qualifySheet)
        {
            var core = "$" + ColumnLetter(col) + row;
            if (!qualifySheet || string.IsNullOrEmpty(sheet)) return core;
            return QuoteSheet(sheet) + "!" + core;
        }

        /// <summary>构造区域引用（首末格，同上规则）。</summary>
        public static string RangeRef(int col1, int row1, int col2, int row2, string sheet, bool abs, bool qualifySheet)
        {
            // 工作表限定只放首端点（'S'!A1:B2）；尾端点带限定是非法形态
            var first = CellRef(col1, row1, sheet, abs, qualifySheet);
            var last = CellRef(col2, row2, null, abs, false);
            return first + ":" + last;
        }

        /// <summary>工作表名含非字母数字/下划线/中文以外字符时加单引号；至少统一形态。</summary>
        public static string QuoteSheet(string sheet)
        {
            if (string.IsNullOrEmpty(sheet)) return string.Empty;
            var needsQuote = false;
            foreach (var ch in sheet)
            {
                if (!char.IsLetterOrDigit(ch) && ch != '_' && ch > 127)
                {
                    continue; // 中文等高位字符视为安全
                }
                if (!char.IsLetterOrDigit(ch) && ch != '_')
                {
                    needsQuote = true;
                    break;
                }
            }
            return needsQuote ? "'" + sheet.Replace("'", "''") + "'" : sheet;
        }

        /// <summary>表头钳制：把区域首行推到 headerRows+1（数据区从表头之后开始）。</summary>
        public static string ClampDataRows(string address, int headerRows)
        {
            if (!TryParseAddress(address, out var sheet, out var c1, out var r1, out var c2, out var r2))
            {
                return address;
            }
            var firstDataRow = r1 + Math.Max(0, headerRows);
            if (firstDataRow > r2) firstDataRow = r2;
            return RangeRef(c1, firstDataRow, c2, r2, sheet, true, !string.IsNullOrEmpty(sheet));
        }

        /// <summary>
        /// A1 引用相对→绝对（只转引用记号，不动函数名/文本）。
        /// 字符串字面量（"..."）内不替换（工具闭环真 bug：IF(A1="B2",..) 的文本 "B2" 曾被篡改成 "$B$2"）。
        /// </summary>
        public static string ToAbsolute(string formula)
        {
            return TransformOutsideStrings(formula, m => "$" + m.Groups[1].Value.ToUpperInvariant() + "$" + m.Groups[2].Value);
        }
        /// <summary>A1 引用绝对→相对（字符串字面量内不替换，同 ToAbsolute 口径）。</summary>
        public static string ToRelative(string formula)
        {
            return TransformOutsideStrings(formula, m => m.Groups[1].Value + m.Groups[2].Value);
        }
        /// <summary>
        /// 仅对「受保护区域」之外的片段做正则替换（工具分类审查第六批 v2 组合状态机）：
        /// 受保护区域 = 双引号字面量（"..."）+ 单引号工作表名（'...'，含 '' 转义）——
        /// v3.8.1.96 只按 " 切分保护了双引号，'Sales S2'!$B:$C 的工作表名内 "S2" 曾被篡改成 "$S$2"
        /// （rewrite_formulas to_absolute 损坏公式）。注意 ' 在双引号内是文本内容不是定界符
        /// （="it's" 合法），须先判双引号再判单引号。
        /// </summary>
        private static string TransformOutsideStrings(string formula, System.Text.RegularExpressions.MatchEvaluator replace)
        {
            if (string.IsNullOrEmpty(formula)) return formula;

            // 第一遍：标记受保护区域（双引号字面量优先于单引号工作表名；"" 与 '' 成对转义）
            var quoted = new bool[formula.Length];
            var inDouble = false;
            var inSingle = false;
            for (var i = 0; i < formula.Length; i++)
            {
                var c = formula[i];
                if (inDouble)
                {
                    quoted[i] = true;
                    if (c == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"') { quoted[i + 1] = true; i++; }
                        else inDouble = false;
                    }
                    continue;
                }
                if (inSingle)
                {
                    quoted[i] = true;
                    if (c == '\'')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '\'') { quoted[i + 1] = true; i++; }
                        else inSingle = false;
                    }
                    continue;
                }
                if (c == '"') { inDouble = true; quoted[i] = true; }
                else if (c == '\'') { inSingle = true; quoted[i] = true; }
            }

            // 第二遍：受保护区域原样保留，其余段做替换
            var sb = new System.Text.StringBuilder(formula.Length);
            var last = 0;
            for (var i = 0; i < formula.Length; i++)
            {
                if (quoted[i] && (i == 0 || !quoted[i - 1]))
                {
                    // 保护段开始：先替换前一段，再原样回填保护段
                    if (i > last) sb.Append(AbsRelRx.Replace(formula.Substring(last, i - last), replace));
                    var j = i;
                    while (j < formula.Length && quoted[j]) j++;
                    sb.Append(formula, i, j - i);
                    i = j - 1;
                    last = j;
                }
            }
            if (last < formula.Length) sb.Append(AbsRelRx.Replace(formula.Substring(last), replace));
            return sb.ToString();
        }
        /// <summary>A1 引用模式（ToAbsolute/ToRelative 共用）。</summary>
        private static readonly Regex AbsRelRx =
            new Regex(@"(?<![A-Za-z0-9_$!''])\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\d(])", RegexOptions.Compiled);

        /// <summary>把一条公式整体包上 IFERROR( x, fallback )。已包 IFERROR 开头的不重复包。</summary>
        public static string WrapIfError(string formula, string fallback = "0")
        {
            var body = formula.TrimStart().TrimStart('=');
            if (body.StartsWith("IFERROR(", StringComparison.OrdinalIgnoreCase)) return "=" + body;
            return "=IFERROR(" + body + "," + (fallback ?? "0") + ")";
        }

        /// <summary>
        /// 试算黑名单：CALL/REGISTER 系可执行外部代码，任何 Evaluate 通道（工具/UI 预览）统一拦截。
        /// 返回 null 表示安全；返回字符串为被禁函数名。
        /// </summary>
        public static string FindBlockedFunction(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return null;
            var upper = formula.ToUpperInvariant();
            foreach (var banned in new[] { "CALL(", "REGISTER(", "REGISTER.ID(", "EXEC(" })
            {
                if (upper.Contains(banned)) return banned.TrimEnd('(');
            }
            return null;
        }

        /// <summary>criteria 智能形态：数字直传；运算符开头直传；单元格引用直传；文本加引号。</summary>
        public static string NormalizeCriteria(string criteria)
        {
            if (string.IsNullOrWhiteSpace(criteria)) return "\"\"";
            var t = criteria.Trim();
            double num;
            if (double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out num)) return t;
            if (t.StartsWith(">=", StringComparison.Ordinal) || t.StartsWith("<=", StringComparison.Ordinal) ||
                t.StartsWith("<>", StringComparison.Ordinal) || t.StartsWith(">", StringComparison.Ordinal) ||
                t.StartsWith("<", StringComparison.Ordinal) || t.StartsWith("=", StringComparison.Ordinal))
            {
                return "\"" + t + "\"";
            }
            if (Regex.IsMatch(t, @"^\$?[A-Za-z]{1,3}\$?\d+$")) return t; // 单元格引用
            if (Regex.IsMatch(t, @"^(?:'[一-龥\w ]+'|[一-龥\w]+)!\$?[A-Za-z]{1,3}\$?\d+$")) return t; // 跨表单元格引用
            return "\"" + t.Replace("\"", "\"\"") + "\"";
        }

        /// <summary>构造参数。</summary>
        public sealed class BuildOptions
        {
            public string Template { get; set; }
            public string SourceAddress { get; set; }
            public int ValueColumn { get; set; } = 1;
            public int? CriteriaColumn { get; set; }
            public string Criteria { get; set; }
            public int? WeightColumn { get; set; }
            public string SheetName { get; set; }
            public int HeaderRows { get; set; } = 1;
            public bool WrapIfErrorForRates { get; set; } = true;
            public string IfErrorFallback { get; set; } = "0";
            // 双区域模板（growth_rate/diff/ratio）：上期/目标区域与值列
            public string SecondSourceAddress { get; set; }
            public int SecondValueColumn { get; set; } = 1;
            // rank 排序方向：0=降序（默认，最大为 1），1=升序
            public int RankOrder { get; set; } = 0;
        }

        /// <summary>意图解析结果。</summary>
        public sealed class ParsedIntent
        {
            public string Template { get; set; }
            public int ValueColumn { get; set; } = 1;
            public int? CriteriaColumn { get; set; }
            public string Criteria { get; set; }
            public int? WeightColumn { get; set; }
            public int HeaderRows { get; set; } = 1;
            public int RankOrder { get; set; } = 0;
            public string MatchedKeyword { get; set; }
        }

        /// <summary>
        /// 大白话需求 → 模板意图（A1 构造器优先链路第一步，纯逻辑确定性）。
        /// keywords 按优先级排列；headers 为数据选区表头（用于「按XX列」的列名定位，1 基）。
        /// 返回 null 表示未命中任何模板（走模型兜底）。
        /// </summary>
        public static ParsedIntent ParseIntent(string requirement, string[] headers, int sourceWidth)
        {
            if (string.IsNullOrWhiteSpace(requirement)) return null;
            var text = requirement.Trim();

            string FindColumn(string keyword)
            {
                if (headers == null || keyword == null) return null;
                for (var i = 0; i < headers.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(headers[i]) &&
                        headers[i].IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return (i + 1).ToString(CultureInfo.InvariantCulture);
                    }
                }
                return null;
            }

            // 关键词 → 模板（顺序即优先级：具体的在前，泛化的在后）
            // 形态：「按<条件列>统计/求和<值列>」「<值列>的排名」「累计」「占比」「同环比/增长率」…
            var intent = new ParsedIntent { HeaderRows = 1 };

            // 去重计数
            if ((text.Contains("多少个不同") || text.Contains("去重") || text.Contains("不重复")) &&
                (text.Contains("多少") || text.Contains("计数") || text.Contains("数量") || text.Contains("几个")))
            {
                intent.Template = "distinct_count";
                intent.CriteriaColumn = 1;
                intent.MatchedKeyword = "去重计数";
                return intent;
            }

            // 排名
            if (text.Contains("排名") || text.Contains("排第几") || text.Contains("第几名"))
            {
                intent.Template = "rank";
                intent.RankOrder = (text.Contains("升序") || text.Contains("从小到大") || text.Contains("最小为1")) ? 1 : 0;
                var col = FindColumn(ExtractRankedColumn(text));
                if (col != null) intent.ValueColumn = int.Parse(col, CultureInfo.InvariantCulture);
                return intent;
            }

            // 累计
            if (text.Contains("累计") || text.Contains("到本月为止") || text.Contains("累加"))
            {
                intent.Template = "cumulative_sum";
                return intent;
            }

            // 逐行占比
            if (text.Contains("占比") && (text.Contains("每行") || text.Contains("占合计") || text.Contains("占总额") || text.Contains("占总")))
            {
                intent.Template = "percent_of_total";
                return intent;
            }

            // 加权平均
            if (text.Contains("加权"))
            {
                intent.Template = "weighted_avg";
                var wcol = FindColumn(ExtractWeightColumn(text));
                if (wcol != null) intent.WeightColumn = int.Parse(wcol, CultureInfo.InvariantCulture);
                return intent;
            }

            // 中位数 / 标准差
            if (text.Contains("中位数")) { intent.Template = "median"; return intent; }
            if (text.Contains("标准差")) { intent.Template = "stdev"; return intent; }

            // 条件聚合：按XX（条件列）统计/求和/计数/平均YY（值列）
            var condAgg = Regex.Match(text, @"(?:按|根据|以)(?<cond>[\u4e00-\u9fa5A-Za-z0-9]{1,12}?)的?(?:来)?(?:统计|求和|合计|加总|计数|数量|平均|均值|客单价)");
            if (condAgg.Success)
            {
                var condName = condAgg.Groups["cond"].Value;
                var condCol = FindColumn(condName);
                if (condCol == null) return null; // 条件列在表头找不到 → 不硬猜，走模型
                intent.CriteriaColumn = int.Parse(condCol, CultureInfo.InvariantCulture);
                intent.Criteria = ExtractCriteriaValue(text, condName) ?? string.Empty;
                if (string.IsNullOrWhiteSpace(intent.Criteria)) return null; // 条件值缺失 → 走模型

                if (text.Contains("平均") || text.Contains("均值") || text.Contains("客单价"))
                {
                    intent.Template = "averageifs";
                }
                else if (text.Contains("计数") || text.Contains("数量") || text.Contains("多少"))
                {
                    intent.Template = "countifs";
                }
                else
                {
                    intent.Template = "sumifs";
                }
                // 值列：需求里「XX的销售额/金额/数量」或表头第二数值列启发
                var valCol = FindValueColumnForAgg(text, headers);
                if (valCol != null) intent.ValueColumn = int.Parse(valCol, CultureInfo.InvariantCulture);
                return intent;
            }

            // 裸聚合：求和/合计/平均/最大/最小/计数（无「按」字）
            if (text.Contains("求和") || text.Contains("合计") || text.Contains("加总") || text.Contains("总和"))
            { intent.Template = "sum"; return intent; }
            if (text.Contains("平均") || text.Contains("均值")) { intent.Template = "average"; return intent; }
            if (text.Contains("最大") || text.Contains("最高")) { intent.Template = "max"; return intent; }
            if (text.Contains("最小") || text.Contains("最低")) { intent.Template = "min"; return intent; }
            if (text.Contains("计数") || text.Contains("多少行") || text.Contains("多少条")) { intent.Template = "count"; return intent; }

            return null; // 未命中 → 模型兜底
        }

        private static string ExtractRankedColumn(string text)
        {
            // 「按销售额排名」「金额的排名」→ 提取列名关键词
            var m = Regex.Match(text, @"(?:按|对)([\u4e00-\u9fa5A-Za-z0-9]{1,10}?)(?:列)?(?:的)?(?:排名|排第几|第几名)");
            if (m.Success) return m.Groups[1].Value;
            var m2 = Regex.Match(text, @"([\u4e00-\u9fa5A-Za-z0-9]{1,10}?)(?:的)?(?:排名|排第几|第几名)");
            return m2.Success ? m2.Groups[1].Value : null;
        }

        private static string ExtractWeightColumn(string text)
        {
            var m = Regex.Match(text, @"(?:按|以|用)([\u4e00-\u9fa5A-Za-z0-9]{1,10}?)(?:加权|作为权重)");
            return m.Success ? m.Groups[1].Value : null;
        }

        private static string ExtractCriteriaValue(string text, string condName)
        {
            // 「按部门统计华东区的销售额」→ 条件值取条件词后面的引号内容或「区/部」结尾词
            var m = Regex.Match(text, "[\u4e00-\u9fa5A-Za-z0-9]*?[\u201c\u0022]([^\u201d\u0022]+)[\u201d\u0022]");
            if (m.Success) return m.Groups[1].Value;
            // 无引号：找条件列名之后到值列名/统计词之间的片段（如「部门统计华东区的」→ 华东区）
            var m2 = Regex.Match(text, @"(?:统计|求和|合计|计数|平均|均值|客单价)(?:每个|各)?([\u4e00-\u9fa5A-Za-z0-9]{1,12}?)(?:的|的销售额|的金额|的销量)");
            if (m2.Success && !string.IsNullOrWhiteSpace(m2.Groups[1].Value)) return m2.Groups[1].Value;
            return null;
        }

        private static string FindValueColumnForAgg(string text, string[] headers)
        {
            if (headers == null) return null;
            // 需求里出现的列名（销售额/金额/数量/利润…）优先
            for (var i = 0; i < headers.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(headers[i])) continue;
                if (text.Contains(headers[i])) return (i + 1).ToString(CultureInfo.InvariantCulture);
            }
            return null;
        }

        /// <summary>
        /// 按模板生成公式。返回公式文本（带 =）。模板不支持/入参缺失抛 ArgumentException（fail-closed）。
        /// </summary>
        public static string Build(BuildOptions o)
        {
            if (o == null || string.IsNullOrWhiteSpace(o.Template) || string.IsNullOrWhiteSpace(o.SourceAddress))
            {
                throw new ArgumentException("template 与 source_address 必填。");
            }
            if (!TryParseAddress(o.SourceAddress, out var srcSheet, out var c1, out var r1, out var c2, out var r2))
            {
                throw new ArgumentException("source_address 形态非法：" + o.SourceAddress);
            }
            var sheet = string.IsNullOrEmpty(o.SheetName) ? srcSheet : o.SheetName;
            var qualify = !string.IsNullOrEmpty(sheet);
            var width = c2 - c1 + 1;
            if (r2 - r1 + 1 <= o.HeaderRows)
            {
                throw new ArgumentException("源区域总行数（" + (r2 - r1 + 1) + "）不足表头行数 " + o.HeaderRows + "，数据区为空。");
            }
            var firstDataRow = r1 + Math.Max(0, o.HeaderRows);
            if (firstDataRow > r2) firstDataRow = r2;
            if (o.ValueColumn < 1 || o.ValueColumn > width)
            {
                throw new ArgumentException("value_column=" + o.ValueColumn + " 超出源区域列宽 " + width + "。");
            }

            string ValueRef()
            {
                var col = c1 + Math.Max(1, o.ValueColumn) - 1;
                return RangeRef(col, firstDataRow, col, r2, sheet, true, qualify);
            }
            string ColRange(int colInSource)
            {
                var col = c1 + Math.Max(1, colInSource) - 1;
                return RangeRef(col, firstDataRow, col, r2, sheet, true, qualify);
            }

            var t = o.Template.Trim().ToLowerInvariant();

            // 双区域模板：第二区域（上期/目标）解析与行数一致性校验（fail-closed）
            string secondValueRefRow(int row) // 行相对（逐行填充联动）
            {
                if (string.IsNullOrWhiteSpace(o.SecondSourceAddress))
                {
                    throw new ArgumentException(t + " 需要第二数据区域 second_source_address（上期/目标区域）。");
                }
                if (!TryParseAddress(o.SecondSourceAddress, out var s2, out var c1b, out var r1b, out var c2b, out var r2b))
                {
                    throw new ArgumentException("second_source_address 形态非法：" + o.SecondSourceAddress);
                }
                if (r2b - r1b + 1 <= o.HeaderRows || r2b - r1b != r2 - r1)
                {
                    throw new ArgumentException("第二区域行数与源区域不一致（" + (r2b - r1b + 1) + " vs " + (r2 - r1 + 1) + " 行，均含表头），逐行计算无法对齐。");
                }
                if (o.SecondValueColumn < 1 || c1b + o.SecondValueColumn - 1 > c2b)
                {
                    throw new ArgumentException("second_value_column=" + o.SecondValueColumn + " 超出第二区域列宽 " + (c2b - c1b + 1) + "。");
                }
                var colB = c1b + o.SecondValueColumn - 1;
                var firstDataB = r1b + Math.Max(0, o.HeaderRows);
                var qualifyB = !string.IsNullOrEmpty(s2);
                return CellRefMixed(colB, row, s2, qualifyB);
            }

            switch (t)
            {
                case "sum":
                    return "=SUM(" + ValueRef() + ")";
                case "count":
                    return "=COUNT(" + ValueRef() + ")";
                case "max":
                    return "=MAX(" + ValueRef() + ")";
                case "min":
                    return "=MIN(" + ValueRef() + ")";
                case "average":
                    return "=IFERROR(AVERAGE(" + ValueRef() + ")," + (o.IfErrorFallback ?? "0") + ")";
                case "sumifs":
                case "countifs":
                case "averageifs":
                {
                    if (!o.CriteriaColumn.HasValue || string.IsNullOrWhiteSpace(o.Criteria))
                    {
                        throw new ArgumentException(t + " 需要 criteria_column 与 criteria。");
                    }
                    if (o.CriteriaColumn.Value < 1 || c1 + o.CriteriaColumn.Value - 1 > c2)
                    {
                        throw new ArgumentException("criteria_column=" + o.CriteriaColumn.Value + " 超出源区域列宽 " + (c2 - c1 + 1) + "。");
                    }
                    var critRange = ColRange(o.CriteriaColumn.Value);
                    var crit = NormalizeCriteria(o.Criteria);
                    if (t == "sumifs") return "=SUMIFS(" + ValueRef() + "," + critRange + "," + crit + ")";
                    if (t == "countifs") return "=COUNTIFS(" + critRange + "," + crit + ")";
                    return "=IFERROR(AVERAGEIFS(" + ValueRef() + "," + critRange + "," + crit + ")," + (o.IfErrorFallback ?? "0") + ")";
                }
                case "weighted_avg":
                {
                    if (!o.WeightColumn.HasValue)
                    {
                        throw new ArgumentException("weighted_avg 需要 weight_column。");
                    }
                    if (o.WeightColumn.Value < 1 || o.WeightColumn.Value > c2 - c1 + 1)
                    {
                        throw new ArgumentException("weight_column=" + o.WeightColumn.Value + " 超出源区域列宽 " + (c2 - c1 + 1) + "。");
                    }
                    var v = ValueRef();
                    var w = ColRange(o.WeightColumn.Value);
                    return "=IFERROR(SUMPRODUCT(" + v + "," + w + ")/SUM(" + w + ")," + (o.IfErrorFallback ?? "0") + ")";
                }
                case "distinct_count":
                {
                    var labelCol = o.CriteriaColumn ?? o.ValueColumn;
                    var range = ColRange(labelCol);
                    return "=IFERROR(SUMPRODUCT(1/COUNTIF(" + range + "," + range + ")),0)";
                }
                case "rank":
                {
                    // RANK.EQ 逐行排名：值格混合引用（$X2），值域绝对，order 0=降序（默认）
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    var vcell = CellRefMixed(vcol, firstDataRow, sheet, qualify);
                    var vrange = RangeRef(vcol, firstDataRow, vcol, r2, sheet, true, qualify);
                    return "=IFERROR(RANK.EQ(" + vcell + "," + vrange + "," + o.RankOrder + ")," + (o.IfErrorFallback ?? "0") + ")";
                }
                case "cumulative_sum":
                {
                    // 累计求和：首端绝对 + 尾端行相对（填充联动）
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    return "=SUM(" + CellRef(vcol, firstDataRow, sheet, true, qualify) + ":" + CellRefMixed(vcol, firstDataRow, sheet, qualify) + ")";
                }
                case "percent_of_total":
                {
                    // 逐行占比：行值 / 全列总和，IFERROR 兜底
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    var vcell = CellRefMixed(vcol, firstDataRow, sheet, qualify);
                    var vrange = RangeRef(vcol, firstDataRow, vcol, r2, sheet, true, qualify);
                    return "=IFERROR(" + vcell + "/SUM(" + vrange + ")," + (o.IfErrorFallback ?? "0") + ")";
                }
                case "median":
                    return "=MEDIAN(" + ValueRef() + ")";
                case "stdev":
                    return "=IFERROR(STDEV.P(" + ValueRef() + ")," + (o.IfErrorFallback ?? "0") + ")";
                case "growth_rate":
                {
                    // 同环比增长率：(本期-上期)/上期，双列混合引用逐行，除零兜底
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    var cur = CellRefMixed(vcol, firstDataRow, sheet, qualify);
                    var prev = secondValueRefRow(firstDataRow);
                    return "=IFERROR((" + cur + "-" + prev + ")/" + prev + "," + (o.IfErrorFallback ?? "0") + ")";
                }
                case "diff":
                {
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    var cur = CellRefMixed(vcol, firstDataRow, sheet, qualify);
                    var prev = secondValueRefRow(firstDataRow);
                    return "=" + cur + "-" + prev;
                }
                case "ratio":
                {
                    // 达成率/占比（双区域相除）：实际/目标，除零兜底
                    var vcol = c1 + Math.Max(1, o.ValueColumn) - 1;
                    var cur = CellRefMixed(vcol, firstDataRow, sheet, qualify);
                    var prev = secondValueRefRow(firstDataRow);
                    return "=IFERROR(" + cur + "/" + prev + "," + (o.IfErrorFallback ?? "0") + ")";
                }
                default:
                    throw new ArgumentException("未知模板：" + o.Template + "。支持 sum/count/max/min/average/sumifs/countifs/averageifs/weighted_avg/distinct_count/rank/cumulative_sum/percent_of_total/median/stdev/growth_rate/diff/ratio。");
            }
        }
    }
}
