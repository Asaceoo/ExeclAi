using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 公式构造与验证闭环四工具的 JSON Schema。
    /// </summary>
    public static class FormulaSchemas
    {
        public static Dictionary<string, object> BuildFormula()
        {
            return Object(
                Names("template", "source_address"),
                Property("template", "string", "模板：sum/count/max/min/average/sumifs/countifs/averageifs/weighted_avg/distinct_count。"),
                Property("source_address", "string", "源数据区域（如 $A$1:$F$21，含表头行）。"),
                IntProperty("value_column", "值列在源区域内的列序（1 基），默认 1。"),
                IntProperty("criteria_column", "条件列在源区域内的列序（1 基）；sumifs/countifs/averageifs/distinct_count 用。"),
                Property("criteria", "string", "条件值：数字/运算符形态(>=100)/单元格引用/文本 自动识别形态。"),
                IntProperty("weight_column", "权重列列序（1 基）；weighted_avg 必填。"),
                Property("sheet_name", "string", "公式引用的工作表名；缺省=源区域自带限定，未限定则同表裸地址。"),
                IntProperty("header_rows", "源区域表头行数，默认 1（数据区自动从表头之后开始）。"),
                Property("second_source_address", "string", "第二数据区域（上期/目标），growth_rate/diff/ratio 双区域模板必填，行数须与源区域一致。"),
                IntProperty("second_value_column", "第二区域的值列列序（1 基），默认 1。"),
                IntProperty("rank_order", "rank 模板排序方向：0=降序（默认，最大值为 1），1=升序。"));
        }

        public static Dictionary<string, object> EvalFormula()
        {
            return Object(
                Names("formula"),
                Property("formula", "string", "要试算的公式（带或不带 =）。不写入单元格，Application.Evaluate 直接求值。"));
        }

        public static Dictionary<string, object> DiagnoseFormula()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要诊断的区域（≤200 格）。"),
                Property("sheet", "string", "工作表名，缺省活动工作表。"));
        }

        public static Dictionary<string, object> RewriteFormulas()
        {
            return Object(
                Names("address", "operation"),
                Property("address", "string", "要批改的区域（≤500 格）。"),
                Property("operation", "string", "批改操作：to_absolute（引用转绝对）/ to_relative（转相对）/ wrap_iferror（整条公式包 IFERROR）。"),
                Property("sheet", "string", "工作表名，缺省活动工作表。"),
                Property("iferror_fallback", "string", "wrap_iferror 的兜底值，默认 0。"));
        }

        // ---- 与 ThirdPhaseSchemas 相同形态的助手 ----

        private static List<object> Names(params string[] names)
        {
            var list = new List<object>();
            foreach (var n in names) list.Add(n);
            return list;
        }

        private static Dictionary<string, object> Object(List<object> required, params Dictionary<string, object>[] properties)
        {
            var props = new Dictionary<string, object>();
            foreach (var p in properties)
            {
                foreach (var kv in p) props[kv.Key] = kv.Value;
            }
            return new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", props },
                { "required", required },
                { "additionalProperties", false }
            };
        }

        private static Dictionary<string, object> Property(string name, string type, string description)
        {
            return new Dictionary<string, object>
            {
                { name, new Dictionary<string, object> { { "type", type }, { "description", description } } }
            };
        }

        private static Dictionary<string, object> IntProperty(string name, string description)
        {
            return new Dictionary<string, object>
            {
                { name, new Dictionary<string, object> { { "type", "integer" }, { "description", description } } }
            };
        }

        private static Dictionary<string, object> BoxArray(string name, string description, params Dictionary<string, object>[] items)
        {
            var itemProps = new Dictionary<string, object>();
            foreach (var p in items)
            {
                foreach (var kv in p) itemProps[kv.Key] = kv.Value;
            }
            return new Dictionary<string, object>
            {
                { name, new Dictionary<string, object>
                {
                    { "type", "array" },
                    { "description", description },
                    { "items", new Dictionary<string, object> { { "type", "object" }, { "properties", itemProps } } }
                } }
            };
        }
    }
}
