using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>数据智能填充（P6-C'）工具参数 Schema。分析只读 + 写入受控，分离成两个工具。</summary>
    internal static class ExcelAiFillToolSchemas
    {
        public static Dictionary<string, object> AnalyzeBlanks()
        {
            return Object(
                Names("address"),
                Property("address", "string", "含表头+数据的源区域地址，如 A1:E40（内部空白单元格将被识别为待填）。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("has_header", "boolean", "首行是否为表头，默认 true。"),
                EnumProperty("reference_scope", new[] { "none", "same_sheet", "whole_workbook" }, "参考上下文范围，默认 same_sheet。none=不取参考样本；same_sheet=取本表首行样本；whole_workbook=再额外列出工作簿全部工作表名（不展开各表内容）。"),
                Property("max_blanks", "integer", "最多返回的空缺格数量上限，默认 5000；超过时仅返回前 N 个并置 has_more=true。"));
        }

        public static Dictionary<string, object> ApplyFills()
        {
            return Object(
                Names("address", "fills"),
                Property("address", "string", "源区域地址（与 ai_fill_blank_cells 分析时一致），如 A1:E40。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("fills", "object", "待填入的空缺格映射，键为单元格地址、值为建议值，如 {\"B3\":\"华东区\",\"C7\":85}。只允许填充本区域内的空白格，非空格会被拒绝。"));
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

        private static Dictionary<string, object> EnumProperty(string name, string[] values, string description = null)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }
    }
}