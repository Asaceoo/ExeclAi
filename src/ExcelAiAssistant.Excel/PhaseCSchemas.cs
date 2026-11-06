using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// P4-C 工具 JSON Schema：推荐图表一键插入 与 模板化文字报告。
    /// </summary>
    public static class PhaseCSchemas
    {
        public static Dictionary<string, object> InsertRecommendedChart()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要分析的数值/分类数据区域。"),
                Property("title", "string", "图表标题；缺省自动生成中文标题。"),
                Property("anchor", "string", "图表左上角锚点单元格（如 F2）；缺省自动放到数据区右侧。"),
                IntProperty("width", "图表宽度（像素），默认 420。"),
                IntProperty("height", "图表高度（像素），默认 260。"));
        }

        public static Dictionary<string, object> GenerateAnalysisReport()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要分析的数据区域。"),
                EnumProperty("output", new[] { "worksheet", "markdown", "html" }),
                Property("sheet_name", "string", "output=worksheet 时新工作表名，默认「分析报告」。"),
                IntProperty("anomaly_std", "异常判定的标准差倍数 k，默认 2.0。"),
                IntProperty("max_columns", "参与分析的列数上限，默认 20。"),
                Property("sections", "string", "逗号分隔的章节白名单（概览/趋势/异常/重点列摘要/相关性/建议），缺省全部章节。"),
                Property("correlation_threshold", "number", "相关性判定阈值 |r|，范围 0.05~0.95，默认 0.7。"),
                Property("output_path", "string", "output=markdown/html 时的输出文件路径；缺省写入用户文档目录。"));
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

        private static Dictionary<string, object> Property(string name, string type, string description = null)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }

        private static Dictionary<string, object> IntProperty(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["minimum"] = 0, ["description"] = description };
        }

        private static Dictionary<string, object> EnumProperty(string name, string[] values)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
        }
    }
}