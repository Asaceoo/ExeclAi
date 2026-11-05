using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>增强工具（rate_change_analysis / cell_comment / data_validation）参数 Schema。</summary>
    internal static class ExcelEnhanceSchemas
    {
        public static Dictionary<string, object> RateChangeAnalysis()
        {
            return Object(
                Names("address"),
                Property("address", "string", "源区域地址，应包含表头行（第 1 行为表头）。"),
                Property("label_field", "integer", "标签/周期列号（默认 1）。"),
                Property("value_field", "integer", "数值列号（默认 2），将逐行与上一行比较。"),
                Property("output_address", "string", "可选，将 标签/数值/绝对变化/百分比变化 写回此处。"),
                Property("sheet", "string", "目标工作表名，缺省为当前活动工作表。"));
        }

        public static Dictionary<string, object> CellComment()
        {
            return Object(
                Names("address"),
                Property("address", "string", "要操作批注的单元格或区域地址。"),
                Property("action", "string", "操作类型：get（读取，默认）/ set（写入）/ clear（清除）。"),
                Property("text", "string", "set 操作时的批注文本。"),
                Property("sheet", "string", "目标工作表名，缺省为当前活动工作表。"));
        }

        public static Dictionary<string, object> DataValidation()
        {
            return Object(
                Names("address"),
                Property("address", "string", "设置数据有效性的区域地址。"),
                Property("type", "string", "验证类型：list（下拉列表，默认）/ interval（数值范围）。"),
                Property("source", "string", "list 时必填：单元格范围（如 A1:A5）或用逗号分隔的列表值。"),
                Property("operator", "string", "interval 时可选：between（默认）/ greater / less。"),
                Property("value1", "string", "interval 的数值边界值 1（必填，如 10）。"),
                Property("value2", "string", "interval 的数值边界值 2（between 时必填，如 100）。"),
                Property("sheet", "string", "目标工作表名，缺省为当前活动工作表。"));
        }

        private static string[] Names(params string[] names)
        {
            return names;
        }

        private static Dictionary<string, object> Object(string[] required, params Dictionary<string, object>[] properties)
        {
            var map = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                map[Convert.ToString(property["name"])] = property;
            }

            var schema = new Dictionary<string, object> { ["type"] = "object", ["properties"] = map };
            if (required != null && required.Length > 0)
            {
                schema["required"] = required;
            }

            return schema;
        }

        private static Dictionary<string, object> Property(string name, string type, string description)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }
    }
}