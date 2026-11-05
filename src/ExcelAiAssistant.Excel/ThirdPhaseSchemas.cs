using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第三期工具 JSON Schema：导入外部文件/批量处理 与 数据分析/可视化。
    /// </summary>
    public static class ThirdPhaseSchemas
    {
        // ---- 导入外部文件与批量处理 ----

        public static Dictionary<string, object> PreviewExternalFile()
        {
            return Object(
                Names("file_path"),
                Property("file_path", "string", "要预览的外部文件绝对路径。"),
                IntProperty("sample_rows", "要采样的行数，默认 10。"),
                Property("delimiter", "string", "分隔符，auto=自动识别（默认）。"),
                Property("encoding", "string", "编码：auto(默认)/utf-8/gbk/ascii。"));
        }

        public static Dictionary<string, object> ImportDelimitedFile()
        {
            return Object(
                Names("file_path", "target_address"),
                Property("file_path", "string", "要导入的分隔文本文件绝对路径。"),
                Property("target_sheet", "string", "目标工作表，默认活动工作表。"),
                Property("target_address", "string", "写入区域的左上角地址。"),
                Property("delimiter", "string", "分隔符，auto=自动识别（默认）。"),
                Property("encoding", "string", "编码：auto(默认)/utf-8/gbk/ascii。"),
                Property("skip_empty_rows", "boolean", "跳过整行空白的行，默认 true。"));
        }

        public static Dictionary<string, object> ImportExcelFileSheet()
        {
            return Object(
                Names("file_path"),
                Property("file_path", "string", "要导入的 Excel 文件绝对路径。"),
                BoxArray("sheets", "要导入的工作表名；缺省导入全部。", Property("name", "string", "工作表名")),
                Property("copy_as_new_sheet", "boolean", "作为新工作表插入，默认 true；false 时按 target 复制取值。"),
                Property("target_sheet", "string", "copy_as_new_sheet=false 时的目标工作表。"),
                Property("target_address", "string", "copy_as_new_sheet=false 时写入区域的左上角地址。"));
        }

        public static Dictionary<string, object> MergeDelimitedFiles()
        {
            return Object(
                Names("file_paths", "output_address"),
                BoxArray("file_paths", "待合并的分隔文本文件绝对路径列表。", Property("path", "string", "文件路径")),
                Property("output_sheet", "string", "输出工作表，默认活动工作表。"),
                Property("output_address", "string", "输出区域的左上角地址。"),
                Property("delimiter", "string", "分隔符，auto=自动识别（默认）。"),
                Property("encoding", "string", "编码：auto(默认)/utf-8/gbk/ascii。"),
                Property("has_header", "boolean", "每个文件首行是否为表头，默认 true；后续文件表头被跳过。"));
        }

        // ---- 数据分析与可视化 ----

        /// <summary>FR-K1 业务口径规则文件读取：无必填参数，目标固定为活动工作簿同目录 _口径.md。</summary>
        public static Dictionary<string, object> GlossaryRead()
        {
            return Object(Names());
        }

        public static Dictionary<string, object> RangeStatistics()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要分析的数据区域。"));
        }

        public static Dictionary<string, object> Correlation()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要分析的数值数据区域。"),
                EnumProperty("method", new[] { "pearson", "rank" }),
                Property("threshold", "number", "强相关判定阈值 |r|，范围 0.05~0.95，默认 0.7；达标的配对会标记 strong=true。"));
        }

        public static Dictionary<string, object> TrendForecast()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "时间/序列标签列 + 数值列构成的连续区域（首行为表头）。"),
                IntProperty("time_column", "时间/序列标签列（区域内第几列），默认 1。"),
                IntProperty("value_column", "数值列（区域内第几列），默认 2。"),
                IntProperty("periods", "向前预测的期数，1~12，默认 3。"),
                EnumProperty("method", new[] { "linear", "moving_average" }),
                IntProperty("window", "moving_average 的窗口期数，2~12，默认 3。"));
        }

        public static Dictionary<string, object> HypothesisTest()
        {
            return Object(
                Names("address", "test"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "检验数据区域（首行为表头）。"),
                EnumProperty("test", new[] { "two_sample_t", "paired_t", "anova", "chi_square" }),
                Property("alpha", "number", "显著性水平 α，0.001~0.2，默认 0.05。"),
                IntProperty("value_column", "数值列（two_sample_t/anova/paired_t），默认 2。"),
                IntProperty("group_column", "分组标签列（two_sample_t/anova），默认 1。"),
                IntProperty("value_column2", "配对第二数值列（paired_t），默认 3。"),
                IntProperty("column_a", "类别列 A（chi_square），默认 1。"),
                IntProperty("column_b", "类别列 B（chi_square），默认 2。"));
        }

        public static Dictionary<string, object> ClusterAnalysis()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "特征数据区域（首行为表头）。"),
                Property("columns", "string", "特征列号逗号分隔（如 2,3,4）；缺省自动选数值占比≥60% 的列。"),
                IntProperty("k", "显式指定簇数 2~6；缺省按平均轮廓系数自动定 k。"),
                IntProperty("k_max", "自动定 k 的上限 2~6，默认 6。"),
                EnumProperty("standardize", new[] { "on", "off" }));
        }

        public static Dictionary<string, object> RegressionAnalysis()
        {
            return Object(
                Names("address", "target_column"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "回归数据区域（首行为表头，含因变量与特征列）。"),
                IntProperty("target_column", "因变量列号（从 1 开始）；缺省取区域最后一列。"),
                Property("columns", "string", "特征列号逗号分隔（如 2,3）；缺省自动选除因变量外数值占比≥60% 的列。"));
        }

        public static Dictionary<string, object> TextVariants()
        {
            return Object(
                Names("address", "column"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "文本数据区域（首行为表头）。"),
                IntProperty("column", "要聚类变体的文本列号（从 1 开始）。"),
                IntProperty("max_distance", "归一化后编辑距离阈值 1~4，默认 2。"));
        }

        public static Dictionary<string, object> AddInsightCard()
        {
            return Object(
                Names("title", "body"),
                Property("sheet", "string", "目标工作表（如看板表），默认活动工作表。"),
                Property("cell", "string", "卡片左上角锚点单元格（如 H20），默认 A1。"),
                Property("title", "string", "洞察卡标题（一句话结论，如「华东区贡献 42% 销售额」）。"),
                Property("body", "string", "洞察卡正文（支撑结论的量化依据，过长自动截断）。"),
                Property("color", "string", "卡片颜色：blue/green/orange/red/gray，默认 blue。"),
                IntProperty("width_pt", "卡片宽度（pt），默认 220。"),
                EnumProperty("replace_existing", new[] { "true", "false" }));
        }

        public static Dictionary<string, object> ValueDistribution()
        {
            return Object(
                Names("address", "column"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "数据区域。"),
                IntProperty("column", "要统计分布的列号（从 1 开始）。"),
                IntProperty("top", "返回出现次数最多的前 N 个值，默认 10。"));
        }

        public static Dictionary<string, object> ColumnProfile()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "含标题行的数据区域。"),
                IntProperty("column", "要画像的列号（从 1 开始），默认 1。"),
                IntProperty("top", "分布返回前 N 个高频值，默认 10。"));
        }

        public static Dictionary<string, object> RecommendChart()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要分析的数据区域。"));
        }

        public static Dictionary<string, object> Sparkline()
        {
            return Object(
                Names("data_range", "dest_range"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("data_range", "string", "作为迷你图数据源的行/列区域。"),
                Property("dest_range", "string", "放置迷你图的单元格区域。"),
                EnumProperty("type", new[] { "line", "column", "winloss" }),
                Property("color", "string", "迷你图颜色（AARRGGBB 十六进制，可选）。"));
        }

        public static Dictionary<string, object> ExportChartImage()
        {
            return Object(
                Names("file_path"),
                Property("sheet", "string", "图表所在工作表，默认活动工作表。"),
                Property("chart_name", "string", "要导出的图表名；缺省导出当前活动图表。"),
                Property("file_path", "string", "输出图片文件的完整路径。目标已存在时需显式 overwrite=true（默认 OVERWRITE_DENIED）。"),
                Property("overwrite", "boolean", "目标图片已存在时是否覆盖（默认 false）。"),
                IntProperty("width", "导出宽度像素。"),
                IntProperty("height", "导出高度像素。"));
        }

        public static Dictionary<string, object> VerifyChart()
        {
            return Object(
                new string[0],
                Property("sheet", "string", "待校验图表所在工作表，默认活动工作表。"),
                Property("chart_name", "string", "待校验的图表名；缺省校验当前活动工作表最后一张图表。"),
                Property("expected_chart_type", "string", "预期的图表类型（Column/Bar/Line/Pie/Scatter/Area 或中文如“柱形图/饼图/散点图”），用于交叉核对生成类型是否符合需求。"),
                Property("expected_title", "string", "预期的图表标题文本，用于核对是否写对了标题（缺省不做标题断言）。"),
                Property("expected_data_source", "string", "预期的数据源区域（如 A1:B8），只要任一系列数据源与之相交即判定数据源命中（缺省不做数据源断言）。"),
                IntProperty("expected_series", "预期系列数量（用于校验数据源系列是否齐全）；不确定可省略。"));
        }

        private static string[] Names(params string[] names)
        {
            return names;
        }

        private static Dictionary<string, object> Object(string[] required, Dictionary<string, object> box, params Dictionary<string, object>[] properties)
        {
            var map = new Dictionary<string, object>();
            if (box != null)
            {
                map[Convert.ToString(box["name"])] = box;
            }

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

        private static Dictionary<string, object> Object(string[] required, params Dictionary<string, object>[] properties)
        {
            return Object(required, null, properties);
        }

        private static Dictionary<string, object> BoxArray(string name, string description, Dictionary<string, object> items)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "array", ["description"] = description, ["items"] = items };
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
        /// <summary>图表读回诊断（v3.7.1.37）：无期望全面体检+根因归类+修复建议+记录。</summary>
        public static Dictionary<string, object> ChartDiagnose()
        {
            return Object(Names("chart_name"),
                Property("sheet", "string", "图表所在工作表名，默认活动工作表。"),
                Property("chart_name", "string", "图表名（list 级工具可查）；缺省取该表最后一张图。"),
                Property("note", "string", "诊断报告：verdict/root_cause/findings（依据+severity）/suggestions（建议动作）/chart_image（AI 读图）。诊断已记录。"));
        }

        /// <summary>图表读回修复（v3.7.1.37）：按诊断修复（重绑/重建/补标题）+前后对比。</summary>
        public static Dictionary<string, object> ChartRepair()
        {
            return Object(Names("chart_name"),
                Property("sheet", "string", "图表所在工作表名，默认活动工作表。"),
                Property("chart_name", "string", "图表名；缺省取该表最后一张图。"),
                Property("source_address", "string", "重绑源数据区域（含表头）；原生新图重建时必填。"),
                EnumProperty("type", new[] { "treemap", "histogram", "waterfall", "sunburst", "boxwhisker", "pareto" }),
                Property("title", "string", "补全标题（提供且缺失时设置）。"));
        }

    }
}