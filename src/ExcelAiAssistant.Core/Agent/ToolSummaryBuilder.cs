using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 把工具执行结果（ToolResult.Data）转成一行人性化摘要，供前端「思考过程」工具卡片展示。
    /// 避免把整段原始 JSON 直接回显到「范围 / 结果」行（verify_chart 等大对象被截断后既不可读也不可解析）。
    /// 规则：工具专属摘要优先，通用回退取「范围类字段 + 少量短标量」。
    /// </summary>
    public static class ToolSummaryBuilder
    {
        public static string Summarize(string tool, object data)
        {
            if (data == null)
            {
                return "完成";
            }

            if (data is string text)
            {
                return string.IsNullOrWhiteSpace(text) ? "完成" : text.Trim();
            }

            JObject o;
            try
            {
                o = JObject.FromObject(data);
            }
            catch
            {
                AppLog.Warn("ToolSummaryBuilder", "[LogCover] ToolSummaryBuilder.cs:31  silent-catch covered (no exception variable)");
                return data.ToString();
            }

            var specific = SummarizeSpecific(tool, o);
            if (specific != null)
            {
                return specific;
            }

            return SummarizeGeneric(o);
        }

        /// <summary>把工具结果序列化为完整 JSON（供前端「详情」展开查看原始数据）。</summary>
        public static string SerializeDetail(object data)
        {
            if (data == null)
            {
                return null;
            }

            if (data is string text)
            {
                return text;
            }

            try
            {
                return JObject.FromObject(data).ToString(Newtonsoft.Json.Formatting.None);
            }
            catch
            {
                AppLog.Warn("ToolSummaryBuilder", "[LogCover] ToolSummaryBuilder.cs:62  silent-catch covered (no exception variable)");
                return data.ToString();
            }
        }

        private static string SummarizeSpecific(string tool, JObject o)
        {
            if (string.IsNullOrEmpty(tool))
            {
                return null;
            }

            if (tool.IndexOf("verify_chart", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyChart(o);
            }

            if (tool.IndexOf("verify_cell_format", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyCellFormat(o);
            }

            if (tool.IndexOf("verify_conditional_format", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyConditionalFormat(o);
            }

            if (tool.IndexOf("verify_sort_filter", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifySortFilter(o);
            }

            if (tool.IndexOf("verify_formula", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyFormula(o);
            }

            if (tool.IndexOf("verify_write_result", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyWithHead(o, "写入自检");
            }

            if (tool.IndexOf("verify_data_clean", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyWithHead(o, "清洗自检");
            }

            if (tool.IndexOf("verify_pivot", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyWithHead(o, "透视自检");
            }

            if (tool.IndexOf("verify_marker", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyWithHead(o, "标记自检");
            }

            if (tool.IndexOf("verify_structure", StringComparison.Ordinal) >= 0)
            {
                return SummarizeVerifyWithHead(o, "结构自检");
            }

            if (tool.IndexOf("create_excel_chart", StringComparison.Ordinal) >= 0
                || tool.IndexOf("insert_recommended_chart", StringComparison.Ordinal) >= 0)
            {
                return SummarizeCreateChart(o);
            }

            if (tool.IndexOf("sparkline", StringComparison.Ordinal) >= 0)
            {
                return SummarizeSparkline(o);
            }

            if (tool.IndexOf("export", StringComparison.Ordinal) >= 0
                && (tool.IndexOf("chart", StringComparison.Ordinal) >= 0 || tool.IndexOf("image", StringComparison.Ordinal) >= 0))
            {
                return SummarizeExportImage(o);
            }

            if (tool.IndexOf("pivot", StringComparison.Ordinal) >= 0)
            {
                return SummarizePivot(o);
            }

            if (tool.IndexOf("recommend", StringComparison.Ordinal) >= 0)
            {
                return SummarizeRecommend(o);
            }

            return null;
        }

        private static string SummarizeVerifyChart(JObject o)
        {
            var parts = new List<string>();
            var verdict = Str(o, "verdict");
            if (string.Equals(verdict, "ok", StringComparison.Ordinal))
            {
                parts.Add("自检通过");
            }
            else if (!string.IsNullOrEmpty(verdict))
            {
                parts.Add("自检未通过(" + verdict + ")");
            }

            var typeName = Str(o, "chart_type_name");
            if (!string.IsNullOrEmpty(typeName))
            {
                parts.Add(typeName);
            }

            if (o["series_count"] != null)
            {
                parts.Add(o["series_count"].ToString() + " 系列");
            }

            if (o["point_total"] != null)
            {
                parts.Add(o["point_total"].ToString() + " 点");
            }

            if (Bool(o, "title_ok") == false)
            {
                parts.Add("标题不符");
            }

            if (Bool(o, "data_source_ok") == false)
            {
                parts.Add("数据源不符");
            }

            if (Bool(o, "data_rendered") == false)
            {
                parts.Add("未渲染");
            }

            return parts.Count > 0 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizeVerifyCellFormat(JObject o)
        {
            var parts = new List<string>();
            var verdict = Str(o, "verdict");
            parts.Add(string.Equals(verdict, "ok", StringComparison.Ordinal) ? "格式自检通过" : "格式自检未通过(" + verdict + ")");
            var nf = Str(o, "actual_number_format");
            if (!string.IsNullOrEmpty(nf))
            {
                parts.Add("数字格式 " + nf);
            }

            if (Bool(o, "actual_bold") == true)
            {
                parts.Add("加粗");
            }

            if (Bool(o, "actual_locked") == false)
            {
                parts.Add("未锁定");
            }

            return string.Join(" · ", parts);
        }

        private static string SummarizeVerifyConditionalFormat(JObject o)
        {
            var parts = new List<string>();
            var verdict = Str(o, "verdict");
            parts.Add(string.Equals(verdict, "ok", StringComparison.Ordinal) ? "条件格式自检通过" : "条件格式自检未通过(" + verdict + ")");
            var rule = Str(o, "actual_rule_type");
            if (!string.IsNullOrEmpty(rule))
            {
                parts.Add("规则 " + rule);
            }

            if (o["rule_count"] != null)
            {
                parts.Add(o["rule_count"].ToString() + " 条规则");
            }

            return string.Join(" · ", parts);
        }

        private static string SummarizeVerifySortFilter(JObject o)
        {
            var verdict = Str(o, "verdict");
            var head = string.Equals(verdict, "ok", StringComparison.Ordinal) ? "排序筛选自检通过" : "排序筛选自检未通过(" + verdict + ")";
            var selfcheck = Str(o, "selfcheck");
            return string.IsNullOrEmpty(selfcheck) ? head : head + "：" + selfcheck;
        }

        private static string SummarizeVerifyFormula(JObject o)
        {
            var parts = new List<string>();
            var verdict = Str(o, "verdict");
            parts.Add(string.Equals(verdict, "ok", StringComparison.Ordinal) ? "公式自检通过" : "公式自检未通过(" + verdict + ")");
            if (Bool(o, "has_formula") == false)
            {
                parts.Add("无公式");
            }

            var formula = Str(o, "actual_formula");
            if (!string.IsNullOrEmpty(formula))
            {
                parts.Add(formula);
            }

            var value = Str(o, "actual_value");
            if (!string.IsNullOrEmpty(value))
            {
                parts.Add("=" + value);
            }

            return string.Join(" · ", parts);
        }

        /// <summary>通用自检摘要：头部「X自检通过/未通过(verdict)」+ 追加 issues 计数（有则）。</summary>
        private static string SummarizeVerifyWithHead(JObject o, string prefix)
        {
            var verdict = Str(o, "verdict");
            var parts = new List<string>();
            parts.Add(string.Equals(verdict, "ok", StringComparison.Ordinal)
                ? prefix + "通过"
                : prefix + "未通过(" + (verdict ?? "?") + ")");
            if (o["issues"] is JArray issues && issues.Count > 0)
            {
                parts.Add(issues.Count + " 项问题");
            }

            return string.Join(" · ", parts);
        }

        private static string SummarizeCreateChart(JObject o)
        {
            var parts = new List<string> { "已创建图表" };
            var name = Str(o, "chart_name");
            if (!string.IsNullOrEmpty(name))
            {
                parts.Add(name);
            }

            var type = Str(o, "chart_type");
            if (!string.IsNullOrEmpty(type))
            {
                parts.Add(type);
            }

            var source = Str(o, "source_address");
            if (!string.IsNullOrEmpty(source))
            {
                parts.Add("数据源 " + source);
            }

            var title = Str(o, "title");
            if (!string.IsNullOrEmpty(title))
            {
                parts.Add("标题「" + title + "」");
            }

            return parts.Count > 1 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizeSparkline(JObject o)
        {
            var parts = new List<string> { "已插入" };
            var type = Str(o, "type");
            if (!string.IsNullOrEmpty(type))
            {
                parts.Add(type + " 迷你图");
            }

            var dest = Str(o, "dest_range");
            if (!string.IsNullOrEmpty(dest))
            {
                parts.Add("→ " + dest);
            }

            var data = Str(o, "data_range");
            if (!string.IsNullOrEmpty(data))
            {
                parts.Add("数据 " + data);
            }

            return parts.Count > 1 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizeExportImage(JObject o)
        {
            var parts = new List<string> { "已导出" };
            var name = Str(o, "chart_name");
            if (!string.IsNullOrEmpty(name))
            {
                parts.Add("图表 " + name);
            }

            var path = Str(o, "file_path");
            if (!string.IsNullOrEmpty(path))
            {
                parts.Add(path);
            }

            return parts.Count > 1 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizePivot(JObject o)
        {
            var parts = new List<string>();
            if (o["row_count"] != null)
            {
                parts.Add(o["row_count"].ToString() + " 行汇总");
            }

            var target = Str(o, "target_address");
            if (!string.IsNullOrEmpty(target))
            {
                parts.Add("→ " + target);
            }

            return parts.Count > 0 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizeRecommend(JObject o)
        {
            var parts = new List<string>();
            var type = Str(o, "recommended") ?? Str(o, "chart_type") ?? Str(o, "chart_type_name");
            if (!string.IsNullOrEmpty(type))
            {
                parts.Add("推荐 " + type);
            }

            var reason = Str(o, "reason");
            if (!string.IsNullOrEmpty(reason))
            {
                parts.Add(reason);
            }

            return parts.Count > 0 ? string.Join(" · ", parts) : null;
        }

        private static string SummarizeGeneric(JObject o)
        {
            foreach (var key in new[] { "address", "target_address", "dest_range", "source_address", "range", "sheet" })
            {
                var v = o[key];
                if (v != null && v.Type != JTokenType.Null)
                {
                    var s = v.ToString();
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        return s;
                    }
                }
            }

            var parts = new List<string>();
            foreach (var prop in o.Properties())
            {
                if (parts.Count >= 3)
                {
                    break;
                }

                var val = prop.Value;
                if (val == null || val.Type == JTokenType.Null || val.Type == JTokenType.Object || val.Type == JTokenType.Array)
                {
                    continue;
                }

                var str = val.ToString();
                if (string.IsNullOrWhiteSpace(str) || str.Length > 40)
                {
                    continue;
                }

                if (string.Equals(str, "True", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(str, "False", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                parts.Add(prop.Name + "=" + str);
            }

            return parts.Count > 0 ? string.Join(" · ", parts) : "完成";
        }

        private static string Str(JObject o, string key)
        {
            var v = o[key];
            if (v == null || v.Type == JTokenType.Null)
            {
                return null;
            }

            var s = v.ToString();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private static bool? Bool(JObject o, string key)
        {
            var v = o[key];
            if (v == null || v.Type != JTokenType.Boolean)
            {
                return null;
            }

            return (bool)v;
        }
    }
}
