using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 四类操作结果自检的纯静态判据（不依赖 COM，可单元测试）：
    /// 单元格格式 / 条件格式 / 排序筛选 / 公式。
    /// 统一「预期值 vs 实际值」判定模式：AI 操作时传入预期参数，自检读取实际属性后逐项比对，
    /// 产出 verdict + issues + selfcheck 三要素（与 verify_chart 一致）。
    /// </summary>
    public static class ExcelVerifyJudgments
    {
        /// <summary>字符串是否仍含花括号（未解析的 {name} 占位符），用于跳过“断言值缺失”的假告警。</summary>
        public static bool ContainsBraces(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return value.IndexOf('{') >= 0 || value.IndexOf('}') >= 0;
        }

        /// <summary>给 verdict 追加后缀标记（verdict=="ok" 时直接取该标记名，否则追加 +name，不重复）。</summary>
        public static void AppendVerdictFlag(ref string verdict, string name)
        {
            if (string.Equals(verdict, "ok", StringComparison.Ordinal))
            {
                verdict = name;
            }
            else if (verdict.IndexOf("+" + name, StringComparison.Ordinal) < 0)
            {
                verdict += "+" + name;
            }
        }

        // ===================== 单元格格式 =====================

        /// <summary>数字格式比对：忽略大小写与首尾空白；未给预期返回 true（未断言）。</summary>
        public static bool NumberFormatMatches(string actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            return string.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>布尔属性比对（加粗/斜体/锁定等）：未给预期返回 true（未断言）。</summary>
        public static bool BoolMatches(bool actual, bool? expected)
        {
            return !expected.HasValue || actual == expected.Value;
        }

        /// <summary>颜色比对：实际 BGR int 与预期（BGR int 或 #RRGGBB）是否一致；未给预期返回 true。</summary>
        public static bool ColorMatches(int actualBgr, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            var exp = ParseExpectedColor(expected);
            return exp.HasValue && actualBgr == exp.Value;
        }

        /// <summary>对齐比对：把实际 xlHAlign/xlVAlign 常量与预期（left/center/right/justify 或 top/middle/bottom）归一化后比较。</summary>
        public static bool AlignmentMatches(int actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            var norm = NormalizeAlignment(expected);
            if (norm == null)
            {
                return true; // 无法识别的预期 → 不断言
            }

            return actual == norm.Value;
        }

        /// <summary>边框线型比对：实际 LineStyle 常量（xlContinuous=1/xlNone=-4142 等）与预期（continuous/none/dashed 等）归一化后比较。</summary>
        public static bool BorderStyleMatches(int actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            var norm = NormalizeBorderStyle(expected);
            if (norm == null)
            {
                return true;
            }

            return actual == norm.Value;
        }

        /// <summary>解析预期颜色：接受 Excel BGR int（含负数）或 #RRGGBB / RRGGBB 十六进制，返回 BGR int。</summary>
        public static int? ParseExpectedColor(string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return null;
            }

            var s = expected.Trim();
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intVal))
            {
                return intVal;
            }

            if (s[0] == '#')
            {
                s = s.Substring(1);
            }

            if (s.Length == 6 && int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            {
                // Excel 用 BGR 布局：R 在最低字节
                var bgr = (int)(((rgb >> 16) & 0xFF) | (((rgb >> 8) & 0xFF) << 8) | ((rgb & 0xFF) << 16));
                if (bgr < 0)
                {
                    unchecked { bgr = (int)((uint)bgr); }
                }

                return bgr;
            }

            return null;
        }

        /// <summary>归一化对齐预期值 → xlHAlign/xlVAlign 常量；无法识别返回 null。</summary>
        public static int? NormalizeAlignment(string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return null;
            }

            switch (expected.Trim().ToLowerInvariant())
            {
                case "left":
                case "left_align":
                case "左对齐":
                case "靠左":
                    return -4131; // xlHAlignLeft
                case "center":
                case "center_align":
                case "居中":
                case "水平居中":
                    return -4108; // xlHAlignCenter
                case "right":
                case "right_align":
                case "右对齐":
                case "靠右":
                    return -4152; // xlHAlignRight
                case "justify":
                case "两端对齐":
                    return -4130; // xlHAlignJustify
                case "top":
                case "top_align":
                case "顶端对齐":
                case "靠上":
                    return -4160; // xlVAlignTop
                case "middle":
                case "middle_align":
                case "垂直居中":
                case "居中_垂直":
                    return -4108; // xlVAlignCenter
                case "bottom":
                case "bottom_align":
                case "底端对齐":
                case "靠下":
                    return -4107; // xlVAlignBottom
                default:
                    return null;
            }
        }

        /// <summary>归一化边框线型预期值 → xlLineStyle 常量；无法识别返回 null。</summary>
        public static int? NormalizeBorderStyle(string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return null;
            }

            switch (expected.Trim().ToLowerInvariant())
            {
                case "continuous":
                case "solid":
                case "实线":
                    return 1; // xlContinuous
                case "none":
                case "无":
                case "无边框":
                    return -4142; // xlLineStyleNone
                case "dashed":
                case "虚线":
                    return -4115; // xlDash
                case "dotted":
                case "点线":
                    return -4118; // xlDot
                case "double":
                case "双线":
                    return -4119; // xlDouble
                case "dashdot":
                case "点划线":
                    return 4; // xlDashDot
                default:
                    return null;
            }
        }

        /// <summary>单元格格式综合判定：逐项比对数字/对齐/字体/边框/填充/保护，产出 verdict/issues/selfcheck。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeCellFormatVerdict(
            string actualNumberFormat,
            bool? actualBold,
            bool? actualItalic,
            string actualFontName,
            int? actualFontSize,
            int? actualFontColor,
            int? actualBackColor,
            int? actualHAlign,
            int? actualVAlign,
            int? actualBorderStyle,
            bool? actualLocked,
            string expectedNumberFormat,
            bool? expectedBold,
            bool? expectedItalic,
            string expectedFontName,
            int? expectedFontSize,
            string expectedFontColor,
            string expectedBackColor,
            string expectedHAlign,
            string expectedVAlign,
            string expectedBorderStyle,
            bool? expectedLocked)
        {
            var issues = new List<string>();
            var verdict = "ok";

            // 断言型 verify 的 fail-closed 口径（§4A）：完全没给任何 expected_* 时不得返回 ok，
            // 否则"设置了格式但忘了给期望"会被当成自检通过（假证据）。
            var hasAnyExpectation = !string.IsNullOrWhiteSpace(expectedNumberFormat)
                || expectedBold.HasValue || expectedItalic.HasValue
                || !string.IsNullOrWhiteSpace(expectedFontName) || expectedFontSize.HasValue
                || !string.IsNullOrWhiteSpace(expectedFontColor) || !string.IsNullOrWhiteSpace(expectedBackColor)
                || !string.IsNullOrWhiteSpace(expectedHAlign) || !string.IsNullOrWhiteSpace(expectedVAlign)
                || !string.IsNullOrWhiteSpace(expectedBorderStyle) || expectedLocked.HasValue;
            if (!hasAnyExpectation)
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供任何 expected_* 期望值，格式断言未执行；请补充期望后再自检。");
                return (verdict, issues, "期望值缺失：格式断言未执行（fail-closed）。");
            }

            // M16 全量审查（fail-closed 统一）：提供了期望值但无法识别（拼写错误等）时，
            // 原实现 ColorMatches/AlignmentMatches/BorderStyleMatches 返回 true（不断言）→ 假通过。
            // 改为显式报 invalid_expectation 并使该断言失效。
            if (!string.IsNullOrWhiteSpace(expectedHAlign) && NormalizeAlignment(expectedHAlign) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_h_align：「{expectedHAlign}」，对齐断言未执行。");
                expectedHAlign = null;
            }

            if (!string.IsNullOrWhiteSpace(expectedVAlign) && NormalizeAlignment(expectedVAlign) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_v_align：「{expectedVAlign}」，对齐断言未执行。");
                expectedVAlign = null;
            }

            if (!string.IsNullOrWhiteSpace(expectedBorderStyle) && NormalizeBorderStyle(expectedBorderStyle) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_border_style：「{expectedBorderStyle}」，边框断言未执行。");
                expectedBorderStyle = null;
            }

            if (!string.IsNullOrWhiteSpace(expectedFontColor) && ParseExpectedColor(expectedFontColor) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_font_color：「{expectedFontColor}」，颜色断言未执行。");
                expectedFontColor = null;
            }

            if (!string.IsNullOrWhiteSpace(expectedBackColor) && ParseExpectedColor(expectedBackColor) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_back_color：「{expectedBackColor}」，颜色断言未执行。");
                expectedBackColor = null;
            }

            if (!NumberFormatMatches(actualNumberFormat, expectedNumberFormat))
            {
                AppendVerdictFlag(ref verdict, "number");
                issues.Add($"数字格式与预期不符：预期「{expectedNumberFormat}」，实际「{actualNumberFormat}」。");
            }

            if (!BoolMatches(actualBold ?? false, expectedBold))
            {
                AppendVerdictFlag(ref verdict, "bold");
                issues.Add($"字体加粗与预期不符：预期 {(expectedBold == true ? "加粗" : "不加粗")}，实际 {(actualBold == true ? "加粗" : "不加粗")}。");
            }

            if (!BoolMatches(actualItalic ?? false, expectedItalic))
            {
                AppendVerdictFlag(ref verdict, "italic");
                issues.Add($"字体斜体与预期不符：预期 {(expectedItalic == true ? "斜体" : "非斜体")}，实际 {(actualItalic == true ? "斜体" : "非斜体")}。");
            }

            if (!string.IsNullOrWhiteSpace(expectedFontName)
                && !string.Equals(actualFontName, expectedFontName, StringComparison.OrdinalIgnoreCase))
            {
                AppendVerdictFlag(ref verdict, "font");
                issues.Add($"字体名称与预期不符：预期「{expectedFontName}」，实际「{actualFontName}」。");
            }

            if (expectedFontSize.HasValue && actualFontSize.HasValue && actualFontSize.Value != expectedFontSize.Value)
            {
                AppendVerdictFlag(ref verdict, "font");
                issues.Add($"字号与预期不符：预期 {expectedFontSize.Value}，实际 {actualFontSize.Value}。");
            }

            if (!ColorMatches(actualFontColor ?? 0, expectedFontColor))
            {
                AppendVerdictFlag(ref verdict, "font");
                issues.Add($"字体颜色与预期不符：预期「{expectedFontColor}」，实际 {FormatColor(actualFontColor)}。");
            }

            if (!ColorMatches(actualBackColor ?? 0, expectedBackColor))
            {
                AppendVerdictFlag(ref verdict, "fill");
                issues.Add($"填充颜色与预期不符：预期「{expectedBackColor}」，实际 {FormatColor(actualBackColor)}。");
            }

            if (!AlignmentMatches(actualHAlign ?? 0, expectedHAlign))
            {
                AppendVerdictFlag(ref verdict, "align");
                issues.Add($"水平对齐与预期不符：预期「{expectedHAlign}」，实际 {DescribeAlignment(actualHAlign)}。");
            }

            if (!AlignmentMatches(actualVAlign ?? 0, expectedVAlign))
            {
                AppendVerdictFlag(ref verdict, "align");
                issues.Add($"垂直对齐与预期不符：预期「{expectedVAlign}」，实际 {DescribeAlignment(actualVAlign)}。");
            }

            if (!BorderStyleMatches(actualBorderStyle ?? -4142, expectedBorderStyle))
            {
                AppendVerdictFlag(ref verdict, "border");
                issues.Add($"边框线型与预期不符：预期「{expectedBorderStyle}」，实际 {DescribeBorderStyle(actualBorderStyle)}。");
            }

            if (!BoolMatches(actualLocked ?? false, expectedLocked))
            {
                AppendVerdictFlag(ref verdict, "protect");
                issues.Add($"单元格锁定与预期不符：预期 {(expectedLocked == true ? "锁定" : "未锁定")}，实际 {(actualLocked == true ? "锁定" : "未锁定")}。");
            }

            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过")
              .Append("：单元格格式");
            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        private static string FormatColor(int? bgr)
        {
            if (!bgr.HasValue)
            {
                return "(未读取)";
            }

            var v = bgr.Value;
            var r = v & 0xFF;
            var g = (v >> 8) & 0xFF;
            var b = (v >> 16) & 0xFF;
            return "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static string DescribeAlignment(int? value)
        {
            if (!value.HasValue)
            {
                return "(未读取)";
            }

            return value switch
            {
                -4131 => "左对齐",
                -4108 => "居中",
                -4152 => "右对齐",
                -4130 => "两端对齐",
                -4160 => "顶端对齐",
                -4107 => "底端对齐",
                _ => "常量 " + value.Value
            };
        }

        private static string DescribeBorderStyle(int? value)
        {
            if (!value.HasValue)
            {
                return "(未读取)";
            }

            return value switch
            {
                1 => "实线",
                -4142 => "无边框",
                -4115 => "虚线",
                -4118 => "点线",
                -4119 => "双线",
                4 => "点划线",
                _ => "常量 " + value.Value
            };
        }

        // ===================== 条件格式 =====================

        /// <summary>归一化条件格式规则类型：data_bar/duplicates/greater/less/between 及中文别名、xlFormatConditionType 数字常量 → 规范名。</summary>
        public static string NormalizeRuleType(string rule)
        {
            if (string.IsNullOrWhiteSpace(rule))
            {
                return null;
            }

            switch (rule.Trim().ToLowerInvariant())
            {
                case "data_bar":
                case "databar":
                case "数据条":
                case "4": // xlDatabar
                    return "data_bar";
                case "duplicates":
                case "duplicate":
                case "重复值":
                case "8": // xlUniqueValues
                    return "duplicates";
                case "greater":
                case "greater_than":
                case "大于":
                    return "greater";
                case "less":
                case "less_than":
                case "小于":
                    return "less";
                case "between":
                case "介于":
                    return "between";
                case "1": // xlCellValue（具体运算符由 Operator 决定）
                    return "cell_value";
                default:
                    return null;
            }
        }

        /// <summary>条件格式综合判定：比对规则类型/公式/运算符/填充色/字体色，产出 verdict/issues/selfcheck。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeConditionalFormatVerdict(
            string actualRuleType,
            string actualFormula1,
            string actualFormula2,
            int? actualOperator,
            int? actualFillColor,
            int? actualFontColor,
            string expectedRule,
            string expectedFormula1,
            string expectedFormula2,
            int? expectedOperator,
            string expectedFillColor,
            string expectedFontColor)
        {
            var issues = new List<string>();
            var verdict = "ok";
            var expectedRuleNorm = NormalizeRuleType(expectedRule);

            if (string.IsNullOrWhiteSpace(expectedRuleNorm))
            {
                if (!string.IsNullOrWhiteSpace(expectedRule))
                {
                    // M16 全量审查（fail-closed 统一）：提供了 expected_rule 但无法识别（拼写错误等），
                    // 不得静默当"未指定"放行——显式报 invalid_expectation。
                    AppendVerdictFlag(ref verdict, "invalid_expectation");
                    issues.Add($"无法识别的 expected_rule：「{expectedRule}」，断言未执行。");
                    return (verdict, issues, "期望值非法：条件格式规则类型断言未执行。");
                }

                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供 expected_rule，条件格式规则类型断言未执行。");
                return (verdict, issues, "期望值缺失：条件格式规则类型断言未执行（fail-closed）。");
            }

            var actualRuleNorm = NormalizeRuleType(actualRuleType);
            if (!string.Equals(actualRuleNorm, expectedRuleNorm, StringComparison.OrdinalIgnoreCase))
            {
                AppendVerdictFlag(ref verdict, "rule");
                issues.Add($"条件格式规则类型与预期不符：预期「{expectedRuleNorm}」，实际「{actualRuleType ?? "(无规则)"}」。");
            }

            if (!string.IsNullOrWhiteSpace(expectedFormula1)
                && !string.Equals(NormalizeFormula(actualFormula1), NormalizeFormula(expectedFormula1), StringComparison.OrdinalIgnoreCase))
            {
                AppendVerdictFlag(ref verdict, "formula");
                issues.Add($"条件格式公式与预期不符：预期「{expectedFormula1}」，实际「{actualFormula1 ?? "(无)"}」。");
            }

            if (!string.IsNullOrWhiteSpace(expectedFormula2)
                && !string.Equals(NormalizeFormula(actualFormula2), NormalizeFormula(expectedFormula2), StringComparison.OrdinalIgnoreCase))
            {
                AppendVerdictFlag(ref verdict, "formula");
                issues.Add($"条件格式第二公式与预期不符：预期「{expectedFormula2}」，实际「{actualFormula2 ?? "(无)"}」。");
            }

            if (expectedOperator.HasValue && actualOperator.HasValue && actualOperator.Value != expectedOperator.Value)
            {
                AppendVerdictFlag(ref verdict, "operator");
                issues.Add($"条件格式运算符与预期不符：预期 {expectedOperator.Value}，实际 {actualOperator.Value}。");
            }

            if (!ColorMatches(actualFillColor ?? 0, expectedFillColor))
            {
                AppendVerdictFlag(ref verdict, "fill");
                issues.Add($"条件格式填充色与预期不符：预期「{expectedFillColor}」，实际 {FormatColor(actualFillColor)}。");
            }

            if (!ColorMatches(actualFontColor ?? 0, expectedFontColor))
            {
                AppendVerdictFlag(ref verdict, "font");
                issues.Add($"条件格式字体色与预期不符：预期「{expectedFontColor}」，实际 {FormatColor(actualFontColor)}。");
            }

            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过")
              .Append("：条件格式规则「").Append(actualRuleNorm ?? actualRuleType ?? "(无)").Append("」");
            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        // ===================== 排序与筛选 =====================

        /// <summary>归一化排序方向：asc/desc 及中文 → 规范名。</summary>
        public static string NormalizeSortDirection(string direction)
        {
            if (string.IsNullOrWhiteSpace(direction))
            {
                return null;
            }

            switch (direction.Trim().ToLowerInvariant())
            {
                case "asc":
                case "ascending":
                case "升序":
                    return "asc";
                case "desc":
                case "descending":
                case "降序":
                    return "desc";
                default:
                    return null;
            }
        }

        /// <summary>排序结果数据级判定：检查排序列身体行是否按预期方向单调；支持次排序列组成的复合字典序（与 sort_range 工具的主/次排序语义一致）。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeSortVerdict(
            object[,] matrix,
            int headerRows,
            int? expectedColumn,
            string expectedDirection,
            int? expectedColumn2,
            string expectedDirection2)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (!expectedColumn.HasValue || string.IsNullOrWhiteSpace(expectedDirection))
            {
                // 单侧缺期望只跳过排序断言；工具级"两侧都没有期望"由 VerifySortFilter 统一判 no_expected。
                return (verdict, issues, "未指定排序依据列与方向，跳过排序断言。");
            }

            // M16 全量审查（fail-closed 统一）：方向值无法识别（拼写错误等）不得静默当升序。
            var dir = NormalizeSortDirection(expectedDirection);
            if (dir == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_direction：「{expectedDirection}」，断言未执行。");
                return (verdict, issues, "期望值非法：排序方向断言未执行。");
            }

            if (expectedDirection2 != null && NormalizeSortDirection(expectedDirection2) == null)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add($"无法识别的 expected_direction_2：「{expectedDirection2}」，断言未执行。");
                return (verdict, issues, "期望值非法：次排序方向断言未执行。");
            }

            var sortCols = matrix.GetLength(1);
            var hasSecondary = expectedColumn2.HasValue;

            var primaryInvalid = expectedColumn.Value < 1 || expectedColumn.Value > sortCols;
            var secondaryInvalid = hasSecondary && (expectedColumn2.Value < 1 || expectedColumn2.Value > sortCols);

            if (primaryInvalid || secondaryInvalid)
            {
                AppendVerdictFlag(ref verdict, "sort");
                issues.Add(primaryInvalid
                    ? $"主排序列号非法：第 {expectedColumn.Value} 列超出区域列数（{sortCols}）。"
                    : $"次排序列号非法：第 {expectedColumn2.Value} 列超出区域列数（{sortCols}）。");
            }
            else if (!IsCompositeSorted(
                matrix, headerRows, expectedColumn.Value, dir,
                hasSecondary ? expectedColumn2.Value : (int?)null,
                hasSecondary ? (NormalizeSortDirection(expectedDirection2) ?? "asc") : null))
            {
                AppendVerdictFlag(ref verdict, "sort");
                var expectTxt = hasSecondary
                    ? $"应主要按第 {expectedColumn.Value} 列{(dir == "asc" ? "升序" : "降序")}，并按第 {expectedColumn2.Value} 列{(NormalizeSortDirection(expectedDirection2) == "desc" ? "降序" : "升序")}"
                    : $"应按第 {expectedColumn.Value} 列{(dir == "asc" ? "升序" : "降序")}";
                issues.Add($"数据并未按预期排序：{expectTxt}，但实际数据不满足该复合排序。");
            }

            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过")
              .Append("：按第 ").Append(expectedColumn.Value).Append(" 列").Append(dir == "asc" ? "升序" : "降序");
            if (hasSecondary)
            {
                sb.Append("，次按第 ").Append(expectedColumn2.Value).Append(" 列").Append(NormalizeSortDirection(expectedDirection2) == "desc" ? "降序" : "升序");
            }

            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        /// <summary>检查矩阵第 col 列从 headerRows(含) 起的身体行按 direction(asc/desc) 单调；给出 col2 时进一步对主列并列的行按 (col2,direction2) 次排序检查（复合字典序）。</summary>
        public static bool IsCompositeSorted(object[,] matrix, int headerRows, int col, string direction,
            int? col2, string direction2)
        {
            if (matrix == null)
            {
                return false;
            }

            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            if (col < 1 || col > cols || headerRows < 0 || headerRows >= rows)
            {
                return headerRows < 0 ? true : rows - headerRows <= 1;
            }

            if (col2.HasValue && (col2.Value < 1 || col2.Value > cols))
            {
                return false;
            }

            var asc = !string.Equals(direction, "desc", StringComparison.OrdinalIgnoreCase);
            var asc2 = col2.HasValue && !string.Equals(direction2, "desc", StringComparison.OrdinalIgnoreCase);
            for (var r = headerRows + 1; r < rows; r++)
            {
                var c = CompareCell(matrix[r, col - 1], matrix[r - 1, col - 1]);
                if (c != 0)
                {
                    if ((asc && c < 0) || (!asc && c > 0))
                    {
                        return false;
                    }

                    continue;
                }

                if (col2.HasValue)
                {
                    var c2 = CompareCell(matrix[r, col2.Value - 1], matrix[r - 1, col2.Value - 1]);
                    if ((asc2 && c2 < 0) || (!asc2 && c2 > 0))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>筛选结果数据级判定：检查区域每个数据行是否都满足 (列, 运算符, 条件) 谓词（与内存 filter_excel_data 工具语义一致）。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeFilterVerdict(
            object[,] matrix,
            int headerRows,
            int? expectedField,
            string expectedCriteria,
            string expectedOperator)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (!expectedField.HasValue || string.IsNullOrEmpty(expectedCriteria))
            {
                // 单侧缺期望只跳过筛选断言；工具级"两侧都没有期望"由 VerifySortFilter 统一判 no_expected。
                return (verdict, issues, "未指定筛选列与条件，跳过筛选断言。");
            }

            var failRows = 0;
            var total = 0;
            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            if (expectedField.Value < 1 || expectedField.Value > cols)
            {
                AppendVerdictFlag(ref verdict, "filter");
                issues.Add($"筛选列号非法：第 {expectedField.Value} 列超出区域列数（{cols}）。");
            }
            else
            {
                for (var r = headerRows; r < rows; r++)
                {
                    total++;
                    var cell = matrix[r, expectedField.Value - 1];
                    if (!FilterMatches(cell, expectedOperator, expectedCriteria))
                    {
                        failRows++;
                    }
                }

                if (failRows > 0)
                {
                    AppendVerdictFlag(ref verdict, "filter");
                    issues.Add($"筛选并未完全生效：区域内 {failRows}/{total} 个数据行不满足条件（第 {expectedField.Value} 列 {DescribeFilterOp(expectedOperator)}「{expectedCriteria}」）。");
                }
            }

            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过")
              .Append("：筛选第 ").Append(expectedField.Value).Append(" 列").Append(DescribeFilterOp(expectedOperator)).Append("「").Append(expectedCriteria).Append("」");
            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        /// <summary>检查矩阵第 col 列从 headerRows(含) 起的身体行是否按 direction(asc/desc) 单调排列。</summary>
        public static bool IsColumnSorted(object[,] matrix, int col, int headerRows, string direction)
        {
            if (matrix == null)
            {
                return false;
            }

            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            if (col < 1 || col > cols || headerRows < 0 || headerRows >= rows)
            {
                return headerRows < 0 ? true : rows - headerRows <= 1;
            }

            var asc = string.Equals(direction, "desc", StringComparison.OrdinalIgnoreCase) ? false : true;
            for (var r = headerRows + 1; r < rows; r++)
            {
                var c = CompareCell(matrix[r, col - 1], matrix[r - 1, col - 1]);
                if (asc && c < 0)
                {
                    return false;
                }

                if (!asc && c > 0)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>单元格是否满足筛选谓词：equals/not_equals/contains/empty/not_empty/gt/gte/lt/lte。</summary>
        public static bool FilterMatches(object cell, string op, string value)
        {
            var opNorm = string.IsNullOrWhiteSpace(op) ? "equals" : op.Trim().ToLowerInvariant();
            var text = cell == null ? string.Empty : Convert.ToString(cell);
            switch (opNorm)
            {
                case "equals":
                case "=":
                    return string.Equals(text ?? string.Empty, value, StringComparison.OrdinalIgnoreCase);
                case "not_equals":
                case "<>":
                    return !string.Equals(text ?? string.Empty, value, StringComparison.OrdinalIgnoreCase);
                case "contains":
                    return (text ?? string.Empty).IndexOf(value ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0;
                case "empty":
                    return IsBlank(cell);
                case "not_empty":
                    return !IsBlank(cell);
                case "gt":
                case "gte":
                case "lt":
                case "lte":
                    return CompareNumeric(cell, value, opNorm);
                default:
                    return string.Equals(text ?? string.Empty, value, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool CompareNumeric(object left, string right, string op)
        {
            if (!double.TryParse(right, NumberStyles.Any, CultureInfo.InvariantCulture, out var rightNumber))
            {
                return false;
            }

            double leftNumber;
            if (!double.TryParse(Convert.ToString(left), NumberStyles.Any, CultureInfo.InvariantCulture, out leftNumber))
            {
                return false;
            }

            return op switch
            {
                "gt" => leftNumber > rightNumber,
                "gte" => leftNumber >= rightNumber,
                "lt" => leftNumber < rightNumber,
                "lte" => leftNumber <= rightNumber,
                _ => false
            };
        }

        private static bool IsBlank(object cell)
        {
            return cell == null
                || (cell is string s && string.IsNullOrWhiteSpace(s))
                || (cell is DBNull);
        }

        /// <summary>单元格比较：null 最小，数字&lt;文本，数字按数值比较，文本按习惯序比较。</summary>
        private static int CompareCell(object a, object b)
        {
            var aNum = TryNumber(a, out var ad);
            var bNum = TryNumber(b, out var bd);
            if (aNum && bNum)
            {
                return ad.CompareTo(bd);
            }

            if (aNum != bNum)
            {
                return aNum ? -1 : 1; // 数字 < 文本
            }

            var as_ = a == null ? string.Empty : Convert.ToString(a);
            var bs_ = b == null ? string.Empty : Convert.ToString(b);
            return string.Compare(as_, bs_, StringComparison.CurrentCulture);
        }

        private static bool TryNumber(object v, out double d)
        {
            d = 0;
            if (v == null || v is DBNull)
            {
                return false;
            }

            if (v is IConvertible c && !(v is string))
            {
                try
                {
                    d = c.ToDouble(CultureInfo.InvariantCulture);
                    return !double.IsNaN(d);
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        private static string DescribeFilterOp(string op)
        {
            switch (string.IsNullOrWhiteSpace(op) ? "equals" : op.Trim().ToLowerInvariant())
            {
                case "equals": case "=": return "等于";
                case "not_equals": case "<>": return "不等于";
                case "contains": return "包含";
                case "empty": return "为空";
                case "not_empty": return "非空";
                case "gt": return "大于";
                case "gte": return "大于等于";
                case "lt": return "小于";
                case "lte": return "小于等于";
                default: return "等于";
            }
        }

        // ===================== 公式 =====================

        /// <summary>归一化公式：去除首尾空白与开头的 =，移除全部空白字符，忽略大小写，用于宽松比对。</summary>
        public static string NormalizeFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return null;
            }

            var s = formula.Trim();
            if (s.StartsWith("=", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            return s.Replace(" ", string.Empty).Replace("\t", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).ToUpperInvariant();
        }

        /// <summary>公式自检综合判定：语法错误值 + 预期公式比对 + 计算值校验（浮点容差），产出 verdict/issues/selfcheck。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeFormulaVerdict(
            bool hasFormula,
            string actualFormula,
            string actualValue,
            string expectedFormula,
            string expectedValue,
            double? tolerance)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (!hasFormula)
            {
                AppendVerdictFlag(ref verdict, "noformula");
                issues.Add("目标单元格不含公式（has_formula=false），说明公式并未真正写入。");
            }

            if (!string.IsNullOrWhiteSpace(actualFormula) && IsErrorValue(actualFormula))
            {
                AppendVerdictFlag(ref verdict, "error");
                issues.Add($"单元格返回错误值 {actualFormula}：公式可能存在引用失效（#REF!）、除零（#DIV/0!）、类型不匹配（#VALUE!）等问题。");
            }
            else if (!string.IsNullOrWhiteSpace(actualValue) && IsErrorValue(actualValue))
            {
                AppendVerdictFlag(ref verdict, "error");
                issues.Add($"单元格计算值返回错误值 {actualValue}：公式可能存在引用失效（#REF!）、除零（#DIV/0!）、类型不匹配（#VALUE!）等问题。");
            }

            if (!string.IsNullOrWhiteSpace(expectedFormula)
                && !string.Equals(NormalizeFormula(actualFormula), NormalizeFormula(expectedFormula), StringComparison.OrdinalIgnoreCase))
            {
                AppendVerdictFlag(ref verdict, "formula");
                issues.Add($"公式内容与预期不符：预期「{expectedFormula}」，实际「{actualFormula ?? "(无)"}」。");
            }

            if (!string.IsNullOrWhiteSpace(expectedValue))
            {
                var valueOk = ValuesMatch(actualValue, expectedValue, tolerance);
                if (!valueOk)
                {
                    AppendVerdictFlag(ref verdict, "value");
                    issues.Add($"公式计算值与预期不符：预期「{expectedValue}」，实际「{actualValue ?? "(无)"}」。");
                }
            }

            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过")
              .Append("：公式");
            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        /// <summary>计算值比对：数值按浮点容差比较（默认 1e-9），文本忽略大小写与首尾空白；无法解析为数值时按文本比较。</summary>
        public static bool ValuesMatch(string actual, string expected, double? tolerance)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            var a = actual.Trim();
            var e = expected.Trim();
            if (double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out var ad)
                && double.TryParse(e, NumberStyles.Any, CultureInfo.InvariantCulture, out var ed))
            {
                var tol = tolerance ?? 1e-9;
                return Math.Abs(ad - ed) <= tol;
            }

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判断字符串是否为 Excel 错误值：#REF! 等错误文本，或 .NET COM 互操作下的 “Error 2007”（xlErrDiv0）等其他形态。</summary>
        public static bool IsErrorValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var s = value.Trim();
            return s.StartsWith("#", StringComparison.Ordinal) && s.EndsWith("!", StringComparison.Ordinal)
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("!", StringComparison.Ordinal) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("?", StringComparison.Ordinal) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("NAME", StringComparison.OrdinalIgnoreCase) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("VALUE", StringComparison.OrdinalIgnoreCase) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("DIV/0", StringComparison.OrdinalIgnoreCase) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("N/A", StringComparison.OrdinalIgnoreCase) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("NUM", StringComparison.OrdinalIgnoreCase) > 0
                || s.StartsWith("#", StringComparison.Ordinal) && s.IndexOf("NULL", StringComparison.OrdinalIgnoreCase) > 0
                || (s.StartsWith("Error", StringComparison.OrdinalIgnoreCase)
                    && s.Length > 5
                    && int.TryParse(s.Substring(5), NumberStyles.Integer, CultureInfo.InvariantCulture, out _));
        }

        // ===================== 写入结果自检 =====================

    /// <summary>写入结果综合判定：期望值矩阵逐格比对（数值容差/文本忽略大小写与空白），期望非空单元格数比对。
    /// type_aware 开启时额外做类型感知回读：期望是数值而实际是文本型数字（如 "3"）即报 issue——
    /// 这类写入「看起来对」但无法参与求和/排序，是高发的静默错误。</summary>
    public static (string verdict, List<string> issues, string selfcheck) ComputeWriteResultVerdict(
        object[,] actual,
        object[,] expected,
        int? expectedCount,
        double? tolerance,
        string address,
        bool typeAware = false)
    {
        var issues = new List<string>();
        var verdict = "ok";
        var topLeft = FirstCellOf(address);

        if (expected == null && !expectedCount.HasValue)
        {
            AppendVerdictFlag(ref verdict, "no_expected");
            issues.Add("未提供 expected_values 或 expected_count，写入自检无判定依据");
        }

        if (expected != null)
        {
            var rows = Math.Max(actual == null ? 0 : actual.GetLength(0), expected.GetLength(0));
            var cols = Math.Max(actual == null ? 0 : actual.GetLength(1), expected.GetLength(1));
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var av = (actual != null && r < actual.GetLength(0) && c < actual.GetLength(1)) ? actual[r, c] : null;
                    var ev = (r < expected.GetLength(0) && c < expected.GetLength(1)) ? expected[r, c] : null;
                    if (!ValueMatches(av, ev, tolerance))
                    {
                        AppendVerdictFlag(ref verdict, "value");
                        issues.Add(OffsetAddress(topLeft, r, c) + "：实际值 " + TextOf(av) + "，期望值 " + TextOf(ev));
                    }
                    else if (typeAware && IsTextNumber(av, ev))
                    {
                        AppendVerdictFlag(ref verdict, "type");
                        issues.Add(OffsetAddress(topLeft, r, c) + "：实际为文本型数字 \"" + TextOf(av)
                            + "\"（期望是数值），无法参与求和/排序，请按数值类型重写");
                    }
                }
            }
        }

            if (expectedCount.HasValue)
            {
                var count = CountNonBlank(actual);
                if (count != expectedCount.Value)
                {
                    AppendVerdictFlag(ref verdict, "count");
                    issues.Add("区域非空单元格数 " + count + " 与期望 " + expectedCount.Value + " 不符");
                }
            }

            return BuildSelfcheck("写入结果", verdict, issues);
        }

        // ===================== 数据清洗自检 =====================

        /// <summary>数据清洗自检的单个检查项。</summary>
        public sealed class CleanCheck
        {
            /// <summary>no_blank_rows / no_blanks / no_duplicates / dates_valid / no_text_numbers / no_text_matches / all_blank。</summary>
            public string Type { get; set; }

            /// <summary>适用的列号（从 1 开始）；null 表示全部列。</summary>
            public int[] Columns { get; set; }

            /// <summary>no_text_matches 专用：判定区域文本是否仍匹配该 .NET 正则（删除类操作的确定性期望）。</summary>
            public string Pattern { get; set; }
        }

        /// <summary>数据清洗综合判定：对区域身体行（跳过表头）逐检查项扫描，任一发现问题即追加 issue 与 verdict 标记。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeDataCleanVerdict(
            object[,] values,
            IReadOnlyList<CleanCheck> checks,
            int headerRows,
            string address)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (checks == null || checks.Count == 0)
            {
                AppendVerdictFlag(ref verdict, "no_checks");
                issues.Add("未指定任何清洗检查项（checks），清洗自检无判定依据");
            }

            if (checks != null)
            {
                foreach (var check in checks)
                {
                    if (check == null || string.IsNullOrWhiteSpace(check.Type))
                    {
                        continue;
                    }

                    switch (NormalizeCleanType(check.Type))
                    {
                        case "no_blank_rows":
                            var blankRows = CountBlankRows(values, headerRows);
                            if (blankRows > 0)
                            {
                                AppendVerdictFlag(ref verdict, "blank_row");
                                issues.Add("仍存在整行空白 " + blankRows + " 行");
                            }

                            break;
                        case "no_blanks":
                            var blanks = CountBlanksInColumns(values, headerRows, check.Columns);
                            if (blanks > 0)
                            {
                                AppendVerdictFlag(ref verdict, "blank");
                                issues.Add("指定列仍存在空白单元格 " + blanks + " 个");
                            }

                            break;
                        case "no_duplicates":
                            var duplicates = CountDuplicateGroups(values, headerRows, check.Columns);
                            if (duplicates > 0)
                            {
                                AppendVerdictFlag(ref verdict, "duplicate");
                                issues.Add("仍存在重复记录 " + duplicates + " 组");
                            }

                            break;
                        case "dates_valid":
                            var invalidDates = CountInvalidDates(values, headerRows, check.Columns);
                            if (invalidDates > 0)
                            {
                                AppendVerdictFlag(ref verdict, "date");
                                issues.Add("存在无法解析为日期的文本 " + invalidDates + " 个");
                            }

                            break;
                        case "no_text_numbers":
                            var textNumbers = CountTextNumbers(values, headerRows, check.Columns);
                            if (textNumbers > 0)
                            {
                                AppendVerdictFlag(ref verdict, "text_number");
                                issues.Add("存在文本型数字 " + textNumbers + " 个");
                            }

                            break;
                        case "no_text_matches":
                            // 删除类操作（delete_text_matches）的确定性期望：全区域（含表头行）不应残留匹配文本。
                            if (string.IsNullOrWhiteSpace(check.Pattern))
                            {
                                AppendVerdictFlag(ref verdict, "no_pattern");
                                issues.Add("no_text_matches 检查缺少 pattern，无判定依据");
                                break;
                            }

                            int matches;
                            try
                            {
                                matches = CountRegexMatches(values, check.Pattern, out var badPattern);
                                if (badPattern != null)
                                {
                                    AppendVerdictFlag(ref verdict, "bad_pattern");
                                    issues.Add("no_text_matches 的 pattern 非法，无法判定： " + badPattern);
                                    break;
                                }
                            }
                            catch (ArgumentException ex)
                            {
                                AppendVerdictFlag(ref verdict, "bad_pattern");
                                issues.Add("no_text_matches 的 pattern 非法，无法判定： " + ex.Message);
                                break;
                            }

                            if (matches > 0)
                            {
                                AppendVerdictFlag(ref verdict, "text_match");
                                issues.Add("仍存在匹配 '" + check.Pattern + "' 的文本 " + matches + " 处，疑似删除未净");
                            }

                            break;
                        case "all_blank":
                            // 清空类操作（clear_excel_range）的确定性期望：整个区域（含表头行）应为空。
                            var nonBlank = CountNonBlankCells(values);
                            if (nonBlank > 0)
                            {
                                AppendVerdictFlag(ref verdict, "not_blank");
                                issues.Add("区域仍残留非空单元格 " + nonBlank + " 个，疑似清空未净");
                            }

                            break;
                    }
                }
            }

            return BuildSelfcheck("数据清洗", verdict, issues);
        }

        // ===================== 透视表自检 =====================

        /// <summary>透视表自检的单个抽查行：标签 + 期望汇总值。</summary>
        public sealed class PivotExpectedRow
        {
            public string Label { get; set; }

            public double? Value { get; set; }
        }

        /// <summary>透视表综合判定：输出区域首行为表头，数据行按第一列标签匹配，比对第二列汇总值。</summary>
    /// <summary>透视表综合判定：按 expected_rows 逐行抽查汇总值、expected_count 核对行数；
    /// 提供 sourceSum/targetSum 时做「明细合计 == 汇总合计」的跨区域守恒对账，拦截透视漏行/汇总截断。</summary>
    public static (string verdict, List<string> issues, string selfcheck) ComputePivotVerdict(
        object[,] values,
        IReadOnlyList<PivotExpectedRow> expectedRows,
        int? expectedCount,
        double? tolerance,
        string address,
        double? sourceSum = null,
        double? targetSum = null)
    {
        var issues = new List<string>();
        var verdict = "ok";
        var rows = values == null ? 0 : values.GetLength(0);
        var dataRows = rows > 0 ? rows - 1 : 0;

        if ((expectedRows == null || expectedRows.Count == 0) && !expectedCount.HasValue
            && !sourceSum.HasValue)
        {
            AppendVerdictFlag(ref verdict, "no_expected");
            issues.Add("未提供 expected_rows、expected_count 或 source_address 守恒对账，透视自检无判定依据");
        }

        if (expectedCount.HasValue && dataRows != expectedCount.Value)
        {
            AppendVerdictFlag(ref verdict, "count");
            issues.Add("数据行数 " + dataRows + " 与期望 " + expectedCount.Value + " 不符");
        }

        if (sourceSum.HasValue && targetSum.HasValue)
        {
            var tol = tolerance ?? 1e-6;
            if (Math.Abs(sourceSum.Value - targetSum.Value) > tol)
            {
                AppendVerdictFlag(ref verdict, "conservation");
                issues.Add("数量守恒不符：明细列合计 " + sourceSum.Value + " ≠ 汇总列合计 " + targetSum.Value
                    + "（差 " + Math.Abs(sourceSum.Value - targetSum.Value) + "），疑似透视漏行或汇总块被截断");
            }
        }

            if (expectedRows != null)
            {
                foreach (var exp in expectedRows)
                {
                    if (exp == null || string.IsNullOrWhiteSpace(exp.Label))
                    {
                        continue;
                    }

                    if (!TryFindPivotRow(values, exp.Label, out var actualValue))
                    {
                        AppendVerdictFlag(ref verdict, "missing");
                        issues.Add("未找到标签行：'" + exp.Label + "'");
                        continue;
                    }

                    if (exp.Value.HasValue
                        && !ValuesMatch(Convert.ToString(actualValue), Convert.ToString(exp.Value.Value), tolerance))
                    {
                        AppendVerdictFlag(ref verdict, "value");
                        issues.Add("标签 '" + exp.Label + "' 值 " + TextOf(actualValue) + " 与期望 " + exp.Value.Value + " 不符");
                    }
                }
            }

            return BuildSelfcheck("透视表", verdict, issues);
        }

        // ===================== 标记自检 =====================

        /// <summary>标记自检的一个已标记单元格：绝对地址 + 实际填充色（Excel BGR）。</summary>
        public sealed class MarkedCell
        {
            public string Address { get; set; }

            public uint Color { get; set; }
        }

        /// <summary>标记综合判定：被标记总数、实际标记色集合 vs 预期类型色集合、期望地址逐个命中。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeMarkerVerdict(
            IReadOnlyList<MarkedCell> markedCells,
            int? expectedCount,
            IReadOnlyDictionary<string, string> expectedFillColors,
            IReadOnlyList<string> expectedMarked,
            string address)
        {
            var issues = new List<string>();
            var verdict = "ok";
            var cells = markedCells ?? (IReadOnlyList<MarkedCell>)new MarkedCell[0];

            if (!expectedCount.HasValue
                && (expectedFillColors == null || expectedFillColors.Count == 0)
                && (expectedMarked == null || expectedMarked.Count == 0))
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供 expected_count / expected_fill_colors / expected_marked，标记自检无判定依据");
            }

            if (expectedCount.HasValue && cells.Count != expectedCount.Value)
            {
                AppendVerdictFlag(ref verdict, "count");
                issues.Add("被标记单元格数 " + cells.Count + " 与期望 " + expectedCount.Value + " 不符");
            }

            if (expectedFillColors != null && expectedFillColors.Count > 0)
            {
                var expectedSet = new HashSet<uint>();
                var known = true;
                foreach (var kv in expectedFillColors)
                {
                    if (!TryParseHexToBgr(kv.Value, out var bgr))
                    {
                        known = false;
                        break;
                    }

                    expectedSet.Add(bgr);
                }

                if (known)
                {
                    var actualSet = new HashSet<uint>();
                    foreach (var m in cells)
                    {
                        actualSet.Add(m.Color);
                    }

                    if (actualSet.Any(c => !expectedSet.Contains(c)))
                    {
                        AppendVerdictFlag(ref verdict, "color");
                        issues.Add("存在预期之外的标记颜色");
                    }

                    if (expectedSet.Any(c => !actualSet.Contains(c)))
                    {
                        AppendVerdictFlag(ref verdict, "color");
                        issues.Add("部分预期标记颜色未出现");
                    }
                }
            }

            if (expectedMarked != null)
            {
                var actualAddrs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var m in cells)
                {
                    actualAddrs.Add(NormalizeAddr(m.Address));
                }

                foreach (var addr in expectedMarked)
                {
                    if (!string.IsNullOrWhiteSpace(addr) && !actualAddrs.Contains(NormalizeAddr(addr)))
                    {
                        AppendVerdictFlag(ref verdict, "mark");
                        issues.Add("地址 " + addr + " 未被标记");
                    }
                }
            }

            return BuildSelfcheck("标记", verdict, issues);
        }

        // ===================== 迷你图自检 =====================

        /// <summary>迷你图组只读快照：类型（1=折线 2=柱状 3=胜负）、落点地址、数据源地址。</summary>
        public sealed class SparklineInfo
        {
            public int? Type { get; set; }

            public string Location { get; set; }

            public string SourceData { get; set; }
        }

        /// <summary>迷你图综合判定：按 dest_range 找覆盖落点的迷你图组，断言类型与数据源。
        /// 无 expected_type/expected_data_source 时 fail-closed（no_expected），不静默放行。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeSparklineVerdict(
            IReadOnlyList<SparklineInfo> groups,
            string destRange,
            string expectedType,
            string expectedSource)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (string.IsNullOrWhiteSpace(destRange))
            {
                AppendVerdictFlag(ref verdict, "no_target");
                issues.Add("未提供 dest_range，迷你图自检无目标可比对");
            }

            if (string.IsNullOrWhiteSpace(expectedType) && string.IsNullOrWhiteSpace(expectedSource))
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供 expected_type 或 expected_data_source，迷你图自检无判定依据");
            }

            if (verdict.IndexOf("no_expected", StringComparison.Ordinal) >= 0
                || verdict.IndexOf("no_target", StringComparison.Ordinal) >= 0)
            {
                return BuildSelfcheck("迷你图", verdict, issues);
            }

            var found = groups?.FirstOrDefault(g => AddressMatches(g.Location, destRange));

            if (found == null)
            {
                AppendVerdictFlag(ref verdict, "missing");
                issues.Add("未在落点 " + destRange + " 找到迷你图（可能未生成或被清除）");
                return BuildSelfcheck("迷你图", verdict, issues);
            }

            var typeMap = new Dictionary<string, int?>(StringComparer.Ordinal)
            {
                ["line"] = 1,
                ["column"] = 2,
                ["winloss"] = 3
            };
            if (!string.IsNullOrWhiteSpace(expectedType))
            {
                if (!typeMap.TryGetValue(expectedType.Trim().ToLowerInvariant(), out var expectedSparkType))
                {
                    // 拼写错误/别名不得静默跳过断言（M16 fail-closed 统一口径）。
                    AppendVerdictFlag(ref verdict, "invalid_expectation");
                    issues.Add("无法识别的 expected_type：「" + expectedType + "」，可用 line/column/winloss。");
                }
                else if (found.Type.HasValue && found.Type.Value != expectedSparkType.Value)
                {
                    AppendVerdictFlag(ref verdict, "type");
                    issues.Add("迷你图类型实际 " + found.Type.Value + " 与预期 " + expectedType + "(=" + expectedSparkType.Value + ") 不符");
                }
            }

            if (!string.IsNullOrWhiteSpace(expectedSource)
                && !string.IsNullOrWhiteSpace(found.SourceData)
                && found.SourceData.Replace("$", string.Empty).IndexOf(expectedSource.Replace("$", string.Empty), StringComparison.OrdinalIgnoreCase) < 0)
            {
                AppendVerdictFlag(ref verdict, "source");
                issues.Add("迷你图数据源实际 '" + found.SourceData + "' 与预期 '" + expectedSource + "' 不符");
            }

            return BuildSelfcheck("迷你图", verdict, issues);
        }

        /// <summary>落点地址匹配：$ 归一化后整段相等，或目标地址以边界（!/冒号/空白）完整出现在落点内——
        /// 防止 dest="B1" 误匹配 location="AB1" 这类子串假阳性。</summary>
        private static bool AddressMatches(string location, string target)
        {
            if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(target))
            {
                return false;
            }

            var loc = location.Replace("$", string.Empty);
            var dest = target.Replace("$", string.Empty);
            if (string.Equals(loc, dest, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var index = loc.IndexOf(dest, StringComparison.OrdinalIgnoreCase);
            while (index >= 0)
            {
                var before = index == 0 ? '!' : loc[index - 1];
                var endIndex = index + dest.Length;
                var after = endIndex >= loc.Length ? '!' : loc[endIndex];
                if (!char.IsLetterOrDigit(before) && !char.IsLetterOrDigit(after))
                {
                    return true;
                }

                index = loc.IndexOf(dest, index + 1, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        // ===================== 数据验证自检 =====================

        /// <summary>数据验证只读快照：COM Validation.Type / Operator / Formula1 / Formula2。</summary>
        public sealed class DataValidationInfo
        {
            /// <summary>Validation.Type 原值（xlValidateInputOnly=0 表示无验证规则）。</summary>
            public int? Type { get; set; }

            public int? Operator { get; set; }

            public string Formula1 { get; set; }

            public string Formula2 { get; set; }
        }

        /// <summary>数据验证综合判定：expected_type(list/interval)+可选 source/operator/value 断言。
        /// 无任何 expected 时 fail-closed（no_expected）。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeDataValidationVerdict(
            DataValidationInfo actual,
            string expectedType,
            string expectedSource,
            string expectedOperator,
            string expectedValue1,
            string expectedValue2)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (string.IsNullOrWhiteSpace(expectedType)
                && string.IsNullOrWhiteSpace(expectedSource)
                && string.IsNullOrWhiteSpace(expectedOperator))
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供 expected_type/expected_source/expected_operator，数据验证自检无判定依据");
                return BuildSelfcheck("数据验证", verdict, issues);
            }

            if (actual == null || !actual.Type.HasValue || actual.Type.Value == 0)
            {
                AppendVerdictFlag(ref verdict, "missing");
                issues.Add("目标区域不存在数据验证规则（可能未设置成功）");
                return BuildSelfcheck("数据验证", verdict, issues);
            }

            if (!string.IsNullOrWhiteSpace(expectedType))
            {
                var t = expectedType.Trim().ToLowerInvariant();
                var ok = t == "list" ? actual.Type.Value == 3
                    : t == "interval" ? actual.Type.Value == 1 || actual.Type.Value == 2
                    : (bool?)null;
                if (ok == null)
                {
                    AppendVerdictFlag(ref verdict, "bad_expected");
                    issues.Add("expected_type 仅支持 list/interval，当前为 " + expectedType);
                }
                else if (!ok.Value)
                {
                    AppendVerdictFlag(ref verdict, "type");
                    issues.Add("验证类型实际 " + actual.Type.Value + " 与预期 " + expectedType + " 不符");
                }
            }

            if (!string.IsNullOrWhiteSpace(expectedSource)
                && !string.IsNullOrWhiteSpace(actual.Formula1)
                && NormalizeFormula(actual.Formula1).IndexOf(NormalizeFormula(expectedSource), StringComparison.OrdinalIgnoreCase) < 0)
            {
                AppendVerdictFlag(ref verdict, "source");
                issues.Add("验证来源实际 '" + actual.Formula1 + "' 与预期 '" + expectedSource + "' 不符");
            }

            var operatorMap = new Dictionary<string, int?>(StringComparer.Ordinal)
            {
                ["between"] = 1,
                ["greater"] = 5,
                ["less"] = 6
            };
            if (!string.IsNullOrWhiteSpace(expectedOperator))
            {
                if (!operatorMap.TryGetValue(expectedOperator.Trim().ToLowerInvariant(), out var expectedOp))
                {
                    // 拼写错误/别名不得静默跳过断言（M16 fail-closed 统一口径）。
                    AppendVerdictFlag(ref verdict, "invalid_expectation");
                    issues.Add("无法识别的 expected_operator：「" + expectedOperator + "」，可用 between/greater/less。");
                }
                else if (actual.Operator.HasValue && actual.Operator.Value != expectedOp.Value)
                {
                    AppendVerdictFlag(ref verdict, "operator");
                    issues.Add("比较运算符实际 " + actual.Operator.Value + " 与预期 " + expectedOperator + "(=" + expectedOp.Value + ") 不符");
                }
            }

            foreach (var pair in new[] { ("value1", expectedValue1, actual.Formula1), ("value2", expectedValue2, actual.Formula2) })
            {
                if (string.IsNullOrWhiteSpace(pair.Item2))
                {
                    continue;
                }

                // 给了期望边界就必须能对上：实际缺失（Formula 未设置）也算不符，fail-closed 不静默放行。
                if (string.IsNullOrWhiteSpace(pair.Item3) || NormalizeFormula(pair.Item3) != NormalizeFormula(pair.Item2))
                {
                    AppendVerdictFlag(ref verdict, pair.Item1);
                    issues.Add("边界 " + pair.Item1 + " 实际 '" + (string.IsNullOrWhiteSpace(pair.Item3) ? "<缺失>" : pair.Item3)
                        + "' 与预期 '" + pair.Item2 + "' 不符");
                }
            }

            return BuildSelfcheck("数据验证", verdict, issues);
        }

        // ===================== 合并状态自检 =====================

        /// <summary>合并状态综合判定：Range.MergeCells（true=整区已合并 / false=无合并 / null=混合）。
        /// 无 expected_merged 时 fail-closed。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeMergeVerdict(
            bool? actualMergeCells,
            bool? expectedMerged,
            string address)
        {
            var issues = new List<string>();
            var verdict = "ok";

            if (!expectedMerged.HasValue)
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供 expected_merged，合并自检无判定依据");
                return BuildSelfcheck("合并状态", verdict, issues);
            }

            if (string.IsNullOrWhiteSpace(address))
            {
                AppendVerdictFlag(ref verdict, "no_target");
                issues.Add("未提供 address，合并自检无目标可比对");
            }

            if (expectedMerged.Value)
            {
                if (actualMergeCells == false)
                {
                    AppendVerdictFlag(ref verdict, "not_merged");
                    issues.Add("区域 " + (address ?? "目标") + " 未处于合并状态");
                }
                else if (actualMergeCells == null)
                {
                    // 混合（同类项合并的合法产物）：不算失败，给提示性 issue 由模型/用户复核。
                    issues.Add("区域为混合合并状态（部分单元格已合并），符合同类项合并的预期形态，请复核分组是否符合意图");
                }
            }
            else
            {
                if (actualMergeCells == true)
                {
                    AppendVerdictFlag(ref verdict, "still_merged");
                    issues.Add("区域 " + (address ?? "目标") + " 仍处于合并状态，疑似取消合并未生效");
                }
                else if (actualMergeCells == null)
                {
                    AppendVerdictFlag(ref verdict, "still_mixed");
                    issues.Add("区域仍存在合并单元格（混合状态），疑似拆分未净");
                }
            }

            return BuildSelfcheck("合并状态", verdict, issues);
        }

        // ===================== 清洗判据内部实现 =====================

        private static string NormalizeCleanType(string type)
        {
            var t = (type ?? string.Empty).Trim().ToLowerInvariant();
            switch (t)
            {
                case "no_blank_rows":
                case "no_blank_row":
                case "去空行":
                case "无整行空白":
                    return "no_blank_rows";
                case "no_blanks":
                case "no_blank":
                case "填空白":
                case "无空白":
                    return "no_blanks";
                case "no_duplicates":
                case "duplicates":
                case "去重":
                case "无重复":
                    return "no_duplicates";
                case "dates_valid":
                case "date":
                case "修日期":
                case "日期有效":
                    return "dates_valid";
                case "no_text_numbers":
                case "text_number":
                case "文本型数字":
                case "无文本数字":
                    return "no_text_numbers";
                case "no_text_matches":
                case "text_matches":
                case "无匹配文本":
                    return "no_text_matches";
                case "all_blank":
                case "all_empty":
                case "全空白":
                case "已清空":
                    return "all_blank";
                default:
                    return t;
            }
        }

        /// <summary>从表头行之后统计仍为整行空白（全部列空白）的行数。</summary>
        private static int CountBlankRows(object[,] values, int headerRows)
        {
            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            for (var r = Math.Max(0, headerRows); r < rows; r++)
            {
                var allBlank = true;
                for (var c = 0; c < cols; c++)
                {
                    if (!IsBlankValue(values[r, c]))
                    {
                        allBlank = false;
                        break;
                    }
                }

                if (allBlank)
                {
                    count++;
                }
            }

            return count;
        }

        private static int CountBlanksInColumns(object[,] values, int headerRows, int[] columns)
        {
            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            for (var r = Math.Max(0, headerRows); r < rows; r++)
            {
                foreach (var c in EffectiveColumns(columns, cols))
                {
                    if (c >= 0 && c < cols && IsBlankValue(values[r, c]))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        /// <summary>统计重复记录组数（按指定列或整行判重，从表头行之后）。</summary>
        private static int CountDuplicateGroups(object[,] values, int headerRows, int[] columns)
        {
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var duplicates = 0;
            for (var r = Math.Max(0, headerRows); r < rows; r++)
            {
                if (IsBlankRow(values, r, columns))
                {
                    continue;
                }

                var key = RowKey(values, r, columns);
                if (!seen.Add(key))
                {
                    duplicates++;
                }
            }

            return duplicates;
        }

        private static int CountInvalidDates(object[,] values, int headerRows, int[] columns)
        {
            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            // 显式指定日期列：列内所有非空文本都必须能解析为日期（严格），"not-a-date" 这类坏值必须命中；
            // 未指定列（全区域扫描）才用「疑似日期文本」启发式，避免姓名/备注等文本列被误报为坏日期。
            var strict = columns != null && columns.Length > 0;
            for (var r = Math.Max(0, headerRows); r < rows; r++)
            {
                foreach (var c in EffectiveColumns(columns, cols))
                {
                    if (c < 0 || c >= cols || values[r, c] == null || !(values[r, c] is string))
                    {
                        continue;
                    }

                    var text = values[r, c].ToString().Trim();
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    if (!strict && !IsLikelyDateText(text))
                    {
                        continue;
                    }

                    if (!DateTime.TryParse(text, out _))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountTextNumbers(object[,] values, int headerRows, int[] columns)
        {
            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            for (var r = Math.Max(0, headerRows); r < rows; r++)
            {
                foreach (var c in EffectiveColumns(columns, cols))
                {
                    if (c < 0 || c >= cols || values[r, c] == null)
                    {
                        continue;
                    }

                    var v = values[r, c];
                    if (v is string s && s.Trim().Length > 0
                        && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        /// <summary>文本若同时含数字与连字符/斜杠/点（日期形态），才尝试解析；纯文本不误报为坏日期。</summary>
        private static bool IsLikelyDateText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var s = text.Trim();
            var hasDigit = false;
            foreach (var ch in s)
            {
                if (char.IsDigit(ch))
                {
                    hasDigit = true;
                    break;
                }
            }

            if (!hasDigit)
            {
                return false;
            }

            return s.IndexOf('-') >= 0 || s.IndexOf('/') >= 0 || s.IndexOf('.') >= 0 || s.IndexOf('年') >= 0;
        }

        private static int[] EffectiveColumns(int[] columns, int cols)
        {
            if (columns == null || columns.Length == 0)
            {
                var all = new int[cols];
                for (var i = 0; i < cols; i++)
                {
                    all[i] = i;
                }

                return all;
            }

            return columns.Select(c => c - 1).ToArray();
        }

        private static bool IsBlankRow(object[,] values, int row, int[] columns)
        {
            var cols = values.GetLength(1);
            foreach (var c in EffectiveColumns(columns, cols))
            {
                if (c >= 0 && c < cols && !IsBlankValue(values[row, c]))
                {
                    return false;
                }
            }

            return true;
        }

        private static string RowKey(object[,] values, int row, int[] columns)
        {
            var cols = values.GetLength(1);
            var parts = new List<string>();
            foreach (var c in EffectiveColumns(columns, cols))
            {
                parts.Add(c >= 0 && c < cols ? NormalizeCellText(values[row, c]) : "");
            }

            return string.Join("\u0001", parts);
        }

        // ===================== 透视表判据内部实现 =====================

        /// <summary>在数据行（首行为表头）中按第一列标签查找，返回第二列汇总值。</summary>
        private static bool TryFindPivotRow(object[,] values, string label, out object actualValue)
        {
            actualValue = null;
            if (values == null || values.GetLength(0) <= 1)
            {
                return false;
            }

            for (var r = 1; r < values.GetLength(0); r++)
            {
                var rowLabel = values[r, 0] == null ? null : Convert.ToString(values[r, 0]);
                if (string.Equals(rowLabel?.Trim(), label.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    actualValue = values[r, 1];
                    return true;
                }
            }

            return false;
        }

        // ===================== 通用辅助 =====================

        private static (string verdict, List<string> issues, string selfcheck) BuildSelfcheck(
            string name, string verdict, List<string> issues)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("自检").Append(verdict == "ok" ? "通过" : "未通过").Append("：").Append(name);
            if (issues.Count > 0)
            {
                sb.Append("；问题：").Append(string.Join("；", issues));
            }

            return (verdict, issues, sb.ToString());
        }

        /// <summary>单元格原始值比对：空值期望需实际也为空；数值按容差；其余文本忽略大小写与首尾空白。</summary>
        private static bool ValueMatches(object actual, object expected, double? tolerance)
        {
            if (IsBlankValue(expected))
            {
                return IsBlankValue(actual);
            }

            if (IsBlankValue(actual))
            {
                return false;
            }

            var a = NormalizeCellText(actual);
            var e = NormalizeCellText(expected);
            if (double.TryParse(a, NumberStyles.Any, CultureInfo.InvariantCulture, out var ad)
                && double.TryParse(e, NumberStyles.Any, CultureInfo.InvariantCulture, out var ed))
            {
                var tol = tolerance ?? 1e-9;
                return Math.Abs(ad - ed) <= tol;
            }

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>类型感知回读：期望是真数值类型（CLR 数值），而实际是文本型数字（字符串形式可解析数值）。
        /// 文本型数字「看起来对」但无法参与求和/排序/图表绘制，是静默错误的高发形态。</summary>
        private static bool IsTextNumber(object actual, object expected)
        {
            if (!(actual is string text) || expected == null || expected is string)
            {
                return false;
            }

            var type = expected.GetType();
            var numericExpected = type == typeof(double) || type == typeof(float) || type == typeof(decimal)
                || type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte);
            if (!numericExpected)
            {
                return false;
            }

            var trimmed = text.Trim();
            return double.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out _)
                || double.TryParse(trimmed, NumberStyles.Any, CultureInfo.CurrentCulture, out _);
        }

        private static bool IsBlankValue(object value)
        {
            return value == null || value is string s && string.IsNullOrWhiteSpace(s);
        }

        // ===================== 智能表自检 =====================

        /// <summary>工作表上单个智能表的只读快照。</summary>
        public sealed class TableInfo
        {
            public string Name { get; set; }

            public string Address { get; set; }

            public int Rows { get; set; }

            public int Columns { get; set; }

            public string[] Headers { get; set; }
        }

        /// <summary>
        /// 智能表判定：按名称或地址定位目标表，核对存在性/表头/行列数。
        /// 无任何期望参数时仅断言「目标表存在」，避免无判定依据的假通过。
        /// </summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeTableVerdict(
            string expectedName,
            string address,
            string[] expectedHeaders,
            int? expectedRows,
            int? expectedColumns,
            IReadOnlyList<TableInfo> tables)
        {
            var verdict = "ok";
            var issues = new List<string>();

            if (tables == null || tables.Count == 0)
            {
                AppendVerdictFlag(ref verdict, "no_tables");
                issues.Add("工作表上没有任何智能表");
                return BuildSelfcheck("智能表", verdict, issues);
            }

            TableInfo target = null;
            if (!string.IsNullOrWhiteSpace(expectedName))
            {
                foreach (var table in tables)
                {
                    if (string.Equals(table.Name, expectedName, StringComparison.OrdinalIgnoreCase))
                    {
                        target = table;
                        break;
                    }
                }

                if (target == null)
                {
                    AppendVerdictFlag(ref verdict, "no_target");
                    issues.Add("未找到名为 '" + expectedName + "' 的智能表（现有：" +
                               string.Join("、", System.Linq.Enumerable.Select(tables, t => t.Name)) + "）");
                }
            }
            else
            {
                var query = ParseRange(address);
                if (query.HasValue)
                {
                    foreach (var table in tables)
                    {
                        var span = ParseRange(table.Address);
                        if (span.HasValue
                            && span.Value.r1 <= query.Value.r1 && query.Value.r2 <= span.Value.r2
                            && span.Value.c1 <= query.Value.c1 && query.Value.c2 <= span.Value.c2)
                        {
                            target = table;
                            break;
                        }
                    }
                }
                else
                {
                    // 既无名称也无地址：只有一张表时取它，多张时无判定依据按 fail-closed。
                    if (tables.Count == 1)
                    {
                        target = tables[0];
                    }
                    else
                    {
                        AppendVerdictFlag(ref verdict, "no_target");
                        issues.Add("未提供 name 或 address，且工作表上有 " + tables.Count + " 张智能表，无法唯一定位");
                    }
                }

                if (target == null && verdict == "ok")
                {
                    AppendVerdictFlag(ref verdict, "no_target");
                    issues.Add("未在地址 '" + address + "' 上找到智能表（现有：" +
                               string.Join("、", System.Linq.Enumerable.Select(tables, t => t.Address)) + "）");
                }
            }

            if (target != null)
            {
                if (expectedHeaders != null && expectedHeaders.Length > 0)
                {
                    var actual = target.Headers ?? Array.Empty<string>();
                    for (var index = 0; index < expectedHeaders.Length; index++)
                    {
                        var expected = NormalizeCellText(expectedHeaders[index]);
                        var actualText = index < actual.Length ? NormalizeCellText(actual[index]) : null;
                        if (!string.Equals(expected, actualText, StringComparison.OrdinalIgnoreCase))
                        {
                            AppendVerdictFlag(ref verdict, "headers");
                            issues.Add("表头第 " + (index + 1) + " 列实际 '" +
                                       (actualText ?? "<空>") + "' 与预期 '" + expected + "' 不符");
                        }
                    }
                }

                if (expectedRows.HasValue && target.Rows != expectedRows.Value)
                {
                    AppendVerdictFlag(ref verdict, "rows");
                    issues.Add("表行数实际 " + target.Rows + " 与预期 " + expectedRows.Value + " 不符");
                }

                if (expectedColumns.HasValue && target.Columns != expectedColumns.Value)
                {
                    AppendVerdictFlag(ref verdict, "columns");
                    issues.Add("表列数实际 " + target.Columns + " 与预期 " + expectedColumns.Value + " 不符");
                }
            }

            return BuildSelfcheck("智能表", verdict, issues);
        }

        // ===================== 保护状态自检 =====================

        /// <summary>工作表/工作簿保护状态的只读快照。</summary>
        public sealed class ProtectionInfo
        {
            /// <summary>工作簿结构保护（增删/重命名工作表被禁）。</summary>
            public bool ProtectStructure { get; set; }

            /// <summary>目标工作表保护（ProtectContents）。</summary>
            public bool SheetProtected { get; set; }

            /// <summary>目标单元格是否锁定（Locked；保护表上锁定单元格拒绝写入）。</summary>
            public bool? CellLocked { get; set; }

            /// <summary>探测的单元格地址（未提供为 null）。</summary>
            public string CellAddress { get; set; }

            /// <summary>保护表是否允许筛选/排序（未保护时无意义）。</summary>
            public bool? AllowFiltering { get; set; }

            public bool? AllowSorting { get; set; }
        }

        /// <summary>
        /// 保护状态判定：expected_protected / expected_writable 两个期望至少给一个才有判定语义；
        /// 都不给时仅回报状态、verdict=ok（探测型用法，写入前置检查的主要调用方式）。
        /// </summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeProtectionVerdict(
            ProtectionInfo info,
            bool? expectedProtected,
            bool? expectedWritable)
        {
            var verdict = "ok";
            var issues = new List<string>();

            if (expectedProtected.HasValue && info.SheetProtected != expectedProtected.Value)
            {
                AppendVerdictFlag(ref verdict, "protect_state");
                issues.Add(info.SheetProtected
                    ? "工作表实际处于保护状态，与预期「未保护」不符"
                    : "工作表实际未保护，与预期「已保护」不符");
            }

            if (expectedWritable.HasValue)
            {
                if (expectedWritable.Value)
                {
                    if (info.SheetProtected && info.CellLocked == true)
                    {
                        AppendVerdictFlag(ref verdict, "write_blocked");
                        issues.Add("工作表受保护且单元格 " + (info.CellAddress ?? "") + " 已锁定，写入将被拒绝（0x800A03EC 的典型成因）");
                    }
                }
                else if (!info.SheetProtected)
                {
                    AppendVerdictFlag(ref verdict, "not_protected");
                    issues.Add("工作表未保护，与预期「不可写入（需保护）」不符");
                }
            }

            return BuildSelfcheck("保护状态", verdict, issues);
        }

        /// <summary>规范化区域地址：去 $、去空格。</summary>
        private static string NormalizeAddress(string address)
        {
            return (address ?? string.Empty).Replace("$", string.Empty).Trim();
        }

        /// <summary>解析 A1 风格区域为 (r1,c1,r2,c2) 四元组；单格按零面积区域；解析失败返回 null。</summary>
        // ===================== 图片自检 =====================

        /// <summary>工作表图片（msoPicture/msoLinkedPicture）的只读快照。坐标与尺寸单位为磅。</summary>
        public sealed class ImageInfo
        {
            public string Name { get; set; }

            public double Left { get; set; }

            public double Top { get; set; }

            public double Width { get; set; }

            public double Height { get; set; }
        }

        /// <summary>锚点单元格的页面坐标边界（磅）：Left/Top 为单元格左上角，Right/Bottom 为右下角。</summary>
        public sealed class CellBounds
        {
            public double Left { get; set; }

            public double Top { get; set; }

            public double Right { get; set; }

            public double Bottom { get; set; }
        }

        /// <summary>
        /// 图片判定：按 expected_count 核对图片数量；按 anchor（单元格页面坐标）定位图片左上角
        /// 落在锚点单元格内的目标图，并可按 expected_width/expected_height 核对尺寸（磅，容差 2pt）。
        /// 无任何期望参数时仅断言「至少存在一张图片」，避免无判定依据的假通过。
        /// </summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeImageVerdict(
            int? expectedCount,
            string anchor,
            CellBounds anchorBounds,
            double? expectedWidth,
            double? expectedHeight,
            IReadOnlyList<ImageInfo> pictures)
        {
            const double tolerancePt = 2.0;
            var verdict = "ok";
            var issues = new List<string>();
            var snapshot = pictures ?? Array.Empty<ImageInfo>();

            if (expectedCount.HasValue && snapshot.Count != expectedCount.Value)
            {
                AppendVerdictFlag(ref verdict, "count_mismatch");
                issues.Add("图片数量实际 " + snapshot.Count + " 张与预期 " + expectedCount.Value + " 张不符");
            }

            ImageInfo target = null;
            if (!string.IsNullOrWhiteSpace(anchor))
            {
                if (anchorBounds == null)
                {
                    AppendVerdictFlag(ref verdict, "invalid_anchor");
                    issues.Add("锚点地址 '" + anchor + "' 无法解析为有效单元格坐标");
                }
                else
                {
                    foreach (var picture in snapshot)
                    {
                        var withinHorizontally = picture.Left >= anchorBounds.Left - tolerancePt
                                                 && picture.Left < anchorBounds.Right + tolerancePt;
                        var withinVertically = picture.Top >= anchorBounds.Top - tolerancePt
                                               && picture.Top < anchorBounds.Bottom + tolerancePt;
                        if (withinHorizontally && withinVertically)
                        {
                            target = picture;
                            break;
                        }
                    }

                    if (target == null)
                    {
                        AppendVerdictFlag(ref verdict, "no_target");
                        issues.Add("未找到左上角落在锚点 '" + anchor + "' 内的图片（现有图片左上角坐标：" +
                                   (snapshot.Count == 0
                                       ? "无"
                                       : string.Join("、", System.Linq.Enumerable.Select(
                                           snapshot,
                                           p => "(" + Math.Round(p.Left, 1) + "," + Math.Round(p.Top, 1) + ")"))) +
                                   "，锚点单元格左上角 (" + Math.Round(anchorBounds.Left, 1) + "," +
                                   Math.Round(anchorBounds.Top, 1) + ")）");
                    }
                }
            }
            else if (expectedWidth.HasValue || expectedHeight.HasValue)
            {
                // fail-closed：尺寸断言必须能定位目标图，否则比对无意义。
                AppendVerdictFlag(ref verdict, "size_without_anchor");
                issues.Add("expected_width/expected_height 必须与 anchor 搭配使用（需锚点定位目标图片后才能核对尺寸）");
            }

            if (target != null)
            {
                if (expectedWidth.HasValue && Math.Abs(target.Width - expectedWidth.Value) > tolerancePt)
                {
                    AppendVerdictFlag(ref verdict, "size");
                    issues.Add("图片宽度实际 " + Math.Round(target.Width, 1) + " 磅与预期 " +
                               expectedWidth.Value + " 磅不符");
                }

                if (expectedHeight.HasValue && Math.Abs(target.Height - expectedHeight.Value) > tolerancePt)
                {
                    AppendVerdictFlag(ref verdict, "size");
                    issues.Add("图片高度实际 " + Math.Round(target.Height, 1) + " 磅与预期 " +
                               expectedHeight.Value + " 磅不符");
                }
            }

            if (!expectedCount.HasValue && string.IsNullOrWhiteSpace(anchor) && snapshot.Count == 0)
            {
                // 探测型口径（§4A）：无期望值时只回报观察结果，不判失败（否则"探测一张空表"会被当成自检未通过）。
                issues.Add("工作表上没有任何图片（探测结果，未提供期望值）。");
            }

            return BuildSelfcheck("图片", verdict, issues);
        }

        // ===================== 冻结窗格自检 =====================

        /// <summary>活动窗口冻结窗格状态的只读快照。</summary>
        public sealed class FreezePaneInfo
        {
            /// <summary>当前窗口是否处于冻结状态（FreezePanes）。</summary>
            public bool Frozen { get; set; }

            /// <summary>冻结分割行号（=冻结的行数；未冻结为 0）。</summary>
            public int SplitRow { get; set; }

            /// <summary>冻结分割列号（=冻结的列数；未冻结为 0）。</summary>
            public int SplitColumn { get; set; }
        }

        /// <summary>
        /// 冻结窗格判定：expected_frozen 核对冻结状态；expected_cell（如 B2）核对冻结位置——
        /// 期望分割行列 = 锚点单元格行列号 - 1。expected_cell 与 expected_frozen=false 组合矛盾，按 fail-closed 处理。
        /// 无任何期望参数时仅回报状态、verdict=ok（探测型用法）。
        /// </summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeFreezeVerdict(
            bool? expectedFrozen,
            string expectedCell,
            FreezePaneInfo actual)
        {
            var verdict = "ok";
            var issues = new List<string>();
            var cell = expectedCell;
            var hasCell = !string.IsNullOrWhiteSpace(cell);
            (int r1, int c1, int r2, int c2)? parsedCell = null;

            if (hasCell)
            {
                parsedCell = ParseRange(cell);
                if (!parsedCell.HasValue || parsedCell.Value.r1 != parsedCell.Value.r2 || parsedCell.Value.c1 != parsedCell.Value.c2)
                {
                    AppendVerdictFlag(ref verdict, "invalid_cell");
                    issues.Add("冻结锚点 '" + cell + "' 无法解析为有效单元格地址");
                    cell = null;
                    hasCell = false;
                }
                else if (parsedCell.Value.r1 == 1 && parsedCell.Value.c1 == 1)
                {
                    AppendVerdictFlag(ref verdict, "invalid_cell");
                    issues.Add("A1 不能作为冻结锚点（其上方无行、左侧无列）");
                    cell = null;
                    hasCell = false;
                }
                else if (expectedFrozen.HasValue && !expectedFrozen.Value)
                {
                    AppendVerdictFlag(ref verdict, "invalid_expectation");
                    issues.Add("expected_cell 与 expected_frozen=false 组合矛盾：期望冻结位置意味着期望已冻结");
                    hasCell = false;
                }
            }

            if (expectedFrozen.HasValue && actual.Frozen != expectedFrozen.Value)
            {
                AppendVerdictFlag(ref verdict, "freeze_state");
                issues.Add(actual.Frozen
                    ? "窗口实际处于冻结状态，与预期「未冻结」不符"
                    : "窗口实际未冻结，与预期「已冻结」不符");
            }

            if (hasCell)
            {
                var expectedSplitRow = parsedCell.Value.r1 - 1;
                var expectedSplitColumn = parsedCell.Value.c1 - 1;
                if (!actual.Frozen)
                {
                    AppendVerdictFlag(ref verdict, "freeze_state");
                    issues.Add("期望冻结在 '" + cell + "'（冻结 " + expectedSplitRow + " 行 " +
                               expectedSplitColumn + " 列），但窗口实际未冻结");
                }
                else if (actual.SplitRow != expectedSplitRow || actual.SplitColumn != expectedSplitColumn)
                {
                    AppendVerdictFlag(ref verdict, "freeze_position");
                    issues.Add("冻结位置实际为 " + actual.SplitRow + " 行 " + actual.SplitColumn +
                               " 列，与预期锚点 '" + cell + "'（" + expectedSplitRow + " 行 " +
                               expectedSplitColumn + " 列）不符");
                }
            }

            return BuildSelfcheck("冻结窗格", verdict, issues);
        }

        // ===================== 超链接自检 =====================

        /// <summary>工作表超链接的只读快照（锚点单元格已归一化为无 $ 的 A1 地址）。</summary>
        public sealed class HyperlinkInfo
        {
            /// <summary>锚点单元格地址（如 B2）。</summary>
            public string Cell { get; set; }

            /// <summary>外部网址（内部链接为空串）。</summary>
            public string Url { get; set; }

            /// <summary>工作簿内部位置（如 Sheet1!A1；外部链接为空串）。</summary>
            public string SubAddress { get; set; }

            /// <summary>显示文本。</summary>
            public string Text { get; set; }
        }

        /// <summary>
        /// 超链接判定：expected_count 核对整表总数；address 定位目标单元格/区域内的超链接，
        /// 并可按 expected_url/expected_subaddress/expected_text 核对内容（fail-closed：内容断言必须与
        /// address 搭配；expected_present=false 与内容断言组合矛盾）。address 无任何期望时断言
        /// 「区域内至少存在一个超链接」，避免无判定依据的假通过。
        /// </summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeHyperlinkVerdict(
            int? expectedCount,
            string address,
            bool? expectedPresent,
            string expectedUrl,
            string expectedSubAddress,
            string expectedText,
            IReadOnlyList<HyperlinkInfo> links)
        {
            var verdict = "ok";
            var issues = new List<string>();
            var snapshot = links ?? Array.Empty<HyperlinkInfo>();
            var hasAddress = !string.IsNullOrWhiteSpace(address);
            var hasContentAssert = expectedUrl != null || expectedSubAddress != null || expectedText != null;

            if (expectedCount.HasValue && snapshot.Count != expectedCount.Value)
            {
                AppendVerdictFlag(ref verdict, "count_mismatch");
                issues.Add("超链接总数实际 " + snapshot.Count + " 个与预期 " + expectedCount.Value + " 个不符");
            }

            if (hasContentAssert && !hasAddress)
            {
                // fail-closed：内容断言没有目标锚点就无法定位比对对象。
                AppendVerdictFlag(ref verdict, "target_without_address");
                issues.Add("expected_url/expected_subaddress/expected_text 必须与 address 搭配（需定位目标单元格后才能核对内容）");
            }

            if (expectedPresent.HasValue && !hasAddress)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add("expected_present 必须与 address 搭配（整表数量断言请用 expected_count）");
            }

            if (expectedPresent == false && hasContentAssert)
            {
                AppendVerdictFlag(ref verdict, "invalid_expectation");
                issues.Add("expected_present=false（期望不存在）与 expected_url/expected_subaddress/expected_text 组合矛盾");
            }

            if (hasAddress)
            {
                var parsedTarget = ParseRange(address);
                if (!parsedTarget.HasValue)
                {
                    AppendVerdictFlag(ref verdict, "invalid_address");
                    issues.Add("目标地址 '" + address + "' 无法解析为有效单元格/区域");
                }
                else
                {
                    var matches = new List<HyperlinkInfo>();
                    foreach (var link in snapshot)
                    {
                        var linkCell = ParseRange(link.Cell);
                        if (linkCell.HasValue
                            // M16 全量审查：按链接锚点矩形（r1..r2, c1..c2）与目标区域求交，
                            // 与写侧 CollectHyperlinks 的多格链接口径对齐（原实现只看左上角，
                            // address 指向多格链接非首格时会误报 no_target）。
                            && linkCell.Value.r2 >= parsedTarget.Value.r1 && linkCell.Value.r1 <= parsedTarget.Value.r2
                            && linkCell.Value.c2 >= parsedTarget.Value.c1 && linkCell.Value.c1 <= parsedTarget.Value.c2)
                        {
                            matches.Add(link);
                        }
                    }

                    if (expectedPresent == false)
                    {
                        if (matches.Count > 0)
                        {
                            AppendVerdictFlag(ref verdict, "should_be_absent");
                            issues.Add("预期 '" + address + "' 内不存在超链接，实际仍有 " + matches.Count +
                                       " 个（锚点：" + DescribeCells(matches) + "）");
                        }
                    }
                    else if (matches.Count == 0)
                    {
                        AppendVerdictFlag(ref verdict, "no_target");
                        issues.Add("'" + address + "' 内未找到超链接（整表现有 " + snapshot.Count +
                                   " 个，锚点：" + DescribeCells(snapshot) + "）");
                    }
                    else
                    {
                        var target = matches[0];
                        if (expectedUrl != null && !UrlEquals(target.Url, expectedUrl))
                        {
                            AppendVerdictFlag(ref verdict, "url_mismatch");
                            issues.Add("超链接地址实际 '" + (target.Url ?? string.Empty) +
                                       "' 与预期 '" + expectedUrl + "' 不符");
                        }

                        if (expectedSubAddress != null && !SubAddressEquals(target.SubAddress, expectedSubAddress))
                        {
                            AppendVerdictFlag(ref verdict, "subaddress_mismatch");
                            issues.Add("超链接内部位置实际 '" + (target.SubAddress ?? string.Empty) +
                                       "' 与预期 '" + expectedSubAddress + "' 不符");
                        }

                        if (expectedText != null && !string.Equals(
                                (target.Text ?? string.Empty).Trim(),
                                expectedText.Trim(),
                                StringComparison.OrdinalIgnoreCase))
                        {
                            AppendVerdictFlag(ref verdict, "text_mismatch");
                            issues.Add("超链接显示文本实际 '" + (target.Text ?? string.Empty) +
                                       "' 与预期 '" + expectedText + "' 不符");
                        }
                    }
                }
            }

            return BuildSelfcheck("超链接", verdict, issues);
        }

        private static string DescribeCells(IReadOnlyList<HyperlinkInfo> links)
        {
            if (links.Count == 0)
            {
                return "无";
            }

            var parts = new List<string>();
            foreach (var link in links)
            {
                parts.Add(link.Cell);
            }

            return string.Join("、", parts);
        }

        /// <summary>外部网址宽松比对：忽略大小写与末尾斜杠（宿主可能对地址做大小写归一）。</summary>
        private static bool UrlEquals(string actual, string expected)
        {
            return string.Equals(
                (actual ?? string.Empty).Trim().TrimEnd('/'),
                (expected ?? string.Empty).Trim().TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>内部位置宽松比对：忽略大小写与 '$'、'\'' 定位符（Sheet1!A1 ≈ 'Sheet1'!$A$1）。</summary>
        private static bool SubAddressEquals(string actual, string expected)
        {
            return string.Equals(
                NormalizeSubAddress(actual),
                NormalizeSubAddress(expected),
                StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeSubAddress(string value)
        {
            return (value ?? string.Empty).Replace("'", string.Empty).Replace("$", string.Empty).Trim();
        }

        private static (int r1, int c1, int r2, int c2)? ParseRange(string address)
        {
            var text = NormalizeAddress(address);
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var parts = text.Split(':');
            if (parts.Length > 2)
            {
                return null;
            }

            if (!TryParseCell(parts[0], out var r1, out var c1))
            {
                return null;
            }

            var (r2, c2) = parts.Length == 2 && TryParseCell(parts[1], out var r2p, out var c2p)
                ? (r2p, c2p)
                : (r1, c1);
            return (Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
        }

        private static bool TryParseCell(string cell, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(cell))
            {
                return false;
            }

            var index = 0;
            while (index < cell.Length && char.IsLetter(cell[index]))
            {
                var upper = char.ToUpperInvariant(cell[index]);
                if (upper < 'A' || upper > 'Z')
                {
                    return false;
                }

                column = column * 26 + (upper - 'A' + 1);
                index++;
            }

            if (index == 0 || index == cell.Length || !int.TryParse(cell.Substring(index), out row) || row < 1)
            {
                return false;
            }

            return true;
        }

        /// <summary>统计全矩阵（含表头行）非空单元格数——all_blank 检查用。</summary>
        private static int CountNonBlankCells(object[,] values)
        {
            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    if (!IsBlankValue(values[r, c]))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        /// <summary>统计全矩阵（含表头行）匹配正则的非空文本数；pattern 非法时经 out 参数带回原异常文案。</summary>
        private static int CountRegexMatches(object[,] values, string pattern, out string badPattern)
        {
            badPattern = null;
            System.Text.RegularExpressions.Regex regex;
            try
            {
                regex = new System.Text.RegularExpressions.Regex(
                    pattern, System.Text.RegularExpressions.RegexOptions.None, TimeSpan.FromMilliseconds(500));
            }
            catch (ArgumentException ex)
            {
                badPattern = ex.Message;
                return 0;
            }

            var count = 0;
            var rows = values == null ? 0 : values.GetLength(0);
            var cols = values == null ? 0 : values.GetLength(1);
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var text = values[r, c] as string;
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    try
                    {
                        if (regex.IsMatch(text))
                        {
                            count++;
                        }
                    }
                    catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
                    {
                        // 超时按不匹配处理（fail-closed 由 pattern 复杂度提示补足）。
                    }
                }
            }

            return count;
        }

        private static string NormalizeCellText(object value)
        {
            return Convert.ToString(value).Trim();
        }

        private static string TextOf(object value)
        {
            return IsBlankValue(value) ? "(空)" : Convert.ToString(value);
        }

        private static int CountNonBlank(object[,] values)
        {
            var count = 0;
            if (values == null)
            {
                return 0;
            }

            for (var r = 0; r < values.GetLength(0); r++)
            {
                for (var c = 0; c < values.GetLength(1); c++)
                {
                    if (!IsBlankValue(values[r, c]))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        /// <summary>取区域地址左上角（去 $ 与工作表前缀），供偏移定位。</summary>
        private static string FirstCellOf(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return null;
            }

            var s = address.Trim();
            var bang = s.IndexOf('!');
            if (bang >= 0)
            {
                s = s.Substring(bang + 1);
            }

            var colon = s.IndexOf(':');
            // 去掉全部 '$'（"$N$1"→"N1"）：OffsetAddress 需要解析行号，残留的 '$' 会让所有
            // issue 地址错误地落在左上角。此前只 TrimStart 前导 '$'，带 $ 的地址（schema 示例形态）必踩。
            return (colon >= 0 ? s.Substring(0, colon) : s).Replace("$", string.Empty).Trim();
        }

        /// <summary>在区域左上角基础上按行/列偏移（0 基）计算绝对地址。</summary>
        private static string OffsetAddress(string topLeft, int dr, int dc)
        {
            if (string.IsNullOrWhiteSpace(topLeft))
            {
                return null;
            }

            var letters = "";
            var i = 0;
            while (i < topLeft.Length && char.IsLetter(topLeft[i]))
            {
                letters += topLeft[i];
                i++;
            }

            var digits = topLeft.Substring(i);
            if (letters.Length == 0
                || !int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row))
            {
                return topLeft;
            }

            return ColumnName(ColumnIndexOf(letters) + dc) + (row + dr);
        }

        private static int ColumnIndexOf(string letters)
        {
            var n = 0;
            foreach (var ch in letters)
            {
                n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            }

            return n;
        }

        private static string ColumnName(int col)
        {
            var name = "";
            while (col > 0)
            {
                var rem = (col - 1) % 26;
                name = (char)('A' + rem) + name;
                col = (col - 1) / 26;
            }

            return name;
        }

        /// <summary>把 "#RRGGBB" 解析为 Excel BGR 长整型。</summary>
        private static bool TryParseHexToBgr(string hex, out uint bgr)
        {
            bgr = 0;
            if (string.IsNullOrWhiteSpace(hex))
            {
                return false;
            }

            var h = hex.Trim().TrimStart('#');
            if (h.Length != 6
                || !int.TryParse(h.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
                || !int.TryParse(h.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
                || !int.TryParse(h.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
            {
                return false;
            }

            bgr = (uint)((b << 16) | (g << 8) | r);
            return true;
        }

        private static string NormalizeAddr(string address)
        {
            return (address ?? string.Empty).Replace("$", string.Empty).Trim().ToUpperInvariant();
        }

        // ===================== 结构自检 =====================

        /// <summary>结构自检综合判定：工作表存在性（新增/重命名/删除后核对）与区域维度（行列数）。
        /// 全部事实由 service 读回传入，本判定不接触 COM。</summary>
        public static (string verdict, List<string> issues, string selfcheck) ComputeStructureVerdict(
            IReadOnlyList<string> sheetNames,
            string expectedSheetExists,
            string expectedSheetAbsent,
            int? expectedRowCount,
            int? expectedColumnCount,
            int? actualRowCount,
            int? actualColumnCount,
            string address)
        {
            var issues = new List<string>();
            var verdict = "ok";
            var anyExpectation = false;

            if (!string.IsNullOrWhiteSpace(expectedSheetExists))
            {
                anyExpectation = true;
                var exists = sheetNames != null && sheetNames.Any(
                    name => string.Equals(name, expectedSheetExists.Trim(), StringComparison.OrdinalIgnoreCase));
                if (!exists)
                {
                    AppendVerdictFlag(ref verdict, "missing_sheet");
                    issues.Add("工作表 \"" + expectedSheetExists + "\" 不存在"
                        + (sheetNames != null && sheetNames.Count > 0 ? "（现有：" + string.Join("、", sheetNames) + "）" : "（工作簿没有任何工作表）"));
                }
            }

            if (!string.IsNullOrWhiteSpace(expectedSheetAbsent))
            {
                anyExpectation = true;
                var exists = sheetNames != null && sheetNames.Any(
                    name => string.Equals(name, expectedSheetAbsent.Trim(), StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    AppendVerdictFlag(ref verdict, "sheet_present");
                    issues.Add("工作表 \"" + expectedSheetAbsent + "\" 仍存在，删除/重命名未真正生效");
                }
            }

            if (expectedRowCount.HasValue || expectedColumnCount.HasValue)
            {
                anyExpectation = true;
                if (string.IsNullOrWhiteSpace(address))
                {
                    AppendVerdictFlag(ref verdict, "no_address");
                    issues.Add("核对了行列数期望但未提供 address，无法读取区域维度");
                }
                else
                {
                    if (expectedRowCount.HasValue && actualRowCount.HasValue && actualRowCount.Value != expectedRowCount.Value)
                    {
                        AppendVerdictFlag(ref verdict, "rows");
                        issues.Add("区域行数 " + actualRowCount.Value + " 与期望 " + expectedRowCount.Value + " 不符（" + address + "）");
                    }

                    if (expectedColumnCount.HasValue && actualColumnCount.HasValue && actualColumnCount.Value != expectedColumnCount.Value)
                    {
                        AppendVerdictFlag(ref verdict, "cols");
                        issues.Add("区域列数 " + actualColumnCount.Value + " 与期望 " + expectedColumnCount.Value + " 不符（" + address + "）");
                    }
                }
            }

            if (!anyExpectation)
            {
                AppendVerdictFlag(ref verdict, "no_expected");
                issues.Add("未提供任何结构期望（expected_sheet_exists / expected_sheet_absent / expected_row_count / expected_column_count），结构自检无判定依据");
            }

            return BuildSelfcheck("结构", verdict, issues);
        }
    }
}
