using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 四类操作结果自检工具 JSON Schema：单元格格式 / 条件格式 / 排序筛选 / 公式。
    /// 全部为只读工具，AI 操作时传入预期参数（expected_*），自检读取实际属性后逐项比对。
    /// </summary>
    public static class VerifySchemas
    {
        public static Dictionary<string, object> VerifyCellFormat()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的单元格区域地址（如 A1 或 A1:C3）。"),
                Property("expected_number_format", "string", "预期的数字格式代码（如 0.00、#,##0、yyyy-mm-dd、@），用于核对数字格式是否设置正确。"),
                Property("expected_bold", "boolean", "预期是否加粗。"),
                Property("expected_italic", "boolean", "预期是否斜体。"),
                Property("expected_font_name", "string", "预期字体名称（如 微软雅黑、Calibri）。"),
                IntProperty("expected_font_size", "预期字号。"),
                Property("expected_font_color", "string", "预期字体颜色（#RRGGBB 或 Excel 颜色 int）。"),
                Property("expected_back_color", "string", "预期填充颜色（#RRGGBB 或 Excel 颜色 int）。"),
                Property("expected_h_align", "string", "预期水平对齐：left/center/right/justify（或中文 左对齐/居中/右对齐/两端对齐）。"),
                Property("expected_v_align", "string", "预期垂直对齐：top/middle/bottom（或中文 顶端对齐/垂直居中/底端对齐）。"),
                Property("expected_border_style", "string", "预期边框线型：continuous/none/dashed/dotted/double（或中文 实线/无/虚线/点线/双线）。"),
                Property("expected_locked", "boolean", "预期是否锁定（保护相关）。"));
        }

        public static Dictionary<string, object> VerifyConditionalFormat()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的条件格式区域地址。"),
                IntProperty("rule_index", "要校验的规则序号（从 1 开始，默认 1）。"),
                Property("expected_rule", "string", "预期的规则类型：data_bar/duplicates/greater/less/between（或中文 数据条/重复值/大于/小于/介于）。"),
                Property("expected_formula1", "string", "预期的规则公式/阈值（如 =$A1>100 或 100）。"),
                Property("expected_formula2", "string", "预期的第二公式/阈值（between 规则用）。"),
                IntProperty("expected_operator", "预期的运算符常量（xlBetween=1/xlGreater=5/xlLess=6 等，可选）。"),
                Property("expected_fill_color", "string", "预期填充颜色（#RRGGBB 或 Excel 颜色 int）。"),
                Property("expected_font_color", "string", "预期字体颜色（#RRGGBB 或 Excel 颜色 int）。"));
        }

        public static Dictionary<string, object> VerifySortFilter()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的排序/筛选区域地址。"),
                IntProperty("header_rows", "表头行数（默认 1，从此行以下为数据身体）。"),
                IntProperty("expected_sort_column", "预期的主排序依据列号（从 1 开始）。"),
                Property("expected_sort_direction", "预期的主排序方向：asc/desc（或中文 升序/降序）。"),
                IntProperty("expected_filter_field", "预期的筛选列号（从 1 开始，可选）。"),
                Property("expected_filter_criteria1", "预期的筛选条件值（如 北京 或 100）。"),
                Property("expected_filter_operator", "预期的筛选谓词：equals/not_equals/contains/empty/not_empty/gt/gte/lt/lte（默认 equals）。"));
        }

        public static Dictionary<string, object> VerifyFormula()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的公式单元格地址。"),
                Property("expected_formula", "string", "预期的公式内容（如 =B2*C2，忽略大小写与空白）。"),
                Property("expected_value", "string", "预期的计算值（数值按容差比较，文本忽略大小写）。"),
                Property("tolerance", "number", "数值比较容差，默认 1e-9。"));
        }

        public static Dictionary<string, object> VerifyWriteResult()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的写入区域地址。"),
                Property("expected_values", "array", "期望写入的二维值矩阵（外层=行，内层=列，与 address 区域同形），逐格比对；空值表示该格期望为空。"),
                IntProperty("expected_count", "期望区域内非空单元格数（可选）。"),
                Property("tolerance", "number", "数值比较容差，默认 1e-9。"),
                Property("type_aware", "boolean", "类型感知回读：期望是数值而实际是文本型数字（如 \"3\"）时判不通过，默认 false。"));
        }

        public static Dictionary<string, object> VerifyDataClean()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要校验的清洗区域地址。"),
                IntProperty("header_rows", "表头行数（默认 1，从此行以下为数据身体）。"),
                Property("checks", "array", "要执行的检查项数组，每项 { type, columns, pattern }；type ∈ no_blank_rows(无整行空白)/no_blanks(指定列无空白)/no_duplicates(无重复)/dates_valid(日期列有效)/no_text_numbers(无文本型数字)/no_text_matches(无匹配 pattern 的文本，需 pattern)/all_blank(全区域为空)；columns 为从 1 开始的列号数组（缺省=全部列）。"));
        }

        public static Dictionary<string, object> VerifySparkline()
        {
            return Object(
                Names("dest_range"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("dest_range", "string", "放置迷你图的单元格区域。"),
                EnumProperty("expected_type", new[] { "line", "column", "winloss" }),
                Property("expected_data_source", "string", "期望的数据源区域地址（可选）。"));
        }

        public static Dictionary<string, object> VerifyDataValidation()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要核对数据验证的区域地址。"),
                EnumProperty("expected_type", new[] { "list", "interval" }),
                Property("expected_source", "string", "list 类型期望的来源（单元格范围或逗号分隔列表值，可选）。"),
                Property("expected_operator", "string", "interval 期望的比较符：between/greater/less（可选）。"),
                Property("expected_value1", "string", "期望的边界值 1（可选）。"),
                Property("expected_value2", "string", "期望的边界值 2（可选）。"));
        }

        public static Dictionary<string, object> VerifyMerge()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要核对合并状态的区域地址。"),
                Property("expected_merged", "string", "期望状态：true=已合并 / false=无合并（同类项合并的混合状态按通过+提示处理）。"));
        }

        public static Dictionary<string, object> VerifyTable()
        {
            return Object(
                Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("name", "string", "要核对的智能表名称（与 address 二选一；都缺省且仅有一张表时取该表）。"),
                Property("address", "string", "智能表所在区域地址（与 name 二选一）。"),
                Property("expected_headers", "array", "预期表头文本数组（如 [\"姓名\",\"金额\"]），逐列比对。"),
                IntProperty("expected_row_count", "预期的表总行数（含表头）。"),
                IntProperty("expected_column_count", "预期的表总列数。"));
        }

        public static Dictionary<string, object> VerifyProtection()
        {
            return Object(
                Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要探测锁定状态的单元格地址（可选；用于判断写入是否会被拒绝）。"),
                Property("expected_protected", "boolean", "预期工作表是否处于保护状态。"),
                Property("expected_writable", "boolean", "预期目标单元格可写入（true 时若受保护且锁定则判不符）。"));
        }

        public static Dictionary<string, object> VerifyImage()
        {
            return Object(
                Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("expected_count", "integer", "预期的图片数量（仅统计图片类 Shape，不含图表/文本框）。"),
                Property("anchor", "string", "预期图片左上角所在锚点单元格地址（如 F2）。"),
                Property("expected_width", "number", "预期图片宽度（磅，需与 anchor 搭配；容差 2 磅）。"),
                Property("expected_height", "number", "预期图片高度（磅，需与 anchor 搭配；容差 2 磅）。"));
        }

        public static Dictionary<string, object> VerifyFreezePanes()
        {
            return Object(
                Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表（读取前会切换到该工作表，冻结状态按表存储）。"),
                Property("expected_frozen", "boolean", "预期窗口是否处于冻结状态。"),
                Property("expected_cell", "string", "预期冻结锚点单元格（如 B2 → 冻结 1 行 1 列）。"));
        }

        public static Dictionary<string, object> VerifyPivot()
        {
            return Object(
                Names("target_address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("target_address", "string", "透视表汇总输出区域（首行为表头 Row Labels/Total）。"),
                Property("expected_rows", "array", "抽查行数组，每项 { label, value }；按第一列标签匹配，比对第二列汇总值。"),
                IntProperty("expected_count", "期望的数据行数（不含表头，可选）。"),
                Property("tolerance", "number", "数值比较容差，默认 1e-9。"),
                Property("source_address", "string", "明细数据区域（可选，配合 source_value_column 做数量守恒对账：明细列合计==汇总列合计）。"),
                IntProperty("source_value_column", "明细数值列号（从 1 开始，可选；与 source_address 搭配使用）。"),
                IntProperty("source_label_column", "明细标签列号（从 1 开始，默认 1；标签为空的行不计入明细合计，与透视聚合口径一致）。"));
        }

        public static Dictionary<string, object> VerifyStructure()
        {
            return Object(
                new string[0],
                Property("sheet", "string", "核对区域维度时的目标工作表，默认活动工作表。"),
                Property("expected_sheet_exists", "string", "期望存在的工作表名（新增/重命名后核对，可选）。"),
                Property("expected_sheet_absent", "string", "期望不存在的工作表名（删除/重命名后核对，可选）。"),
                Property("address", "string", "要核对行列数的区域地址（可选，配合 expected_row_count/expected_column_count）。"),
                IntProperty("expected_row_count", "期望的区域行数（可选，需提供 address）。"),
                IntProperty("expected_column_count", "期望的区域列数（可选，需提供 address）。"));
        }

        public static Dictionary<string, object> VerifyMarker()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要扫描标记的区域地址（区域过大时会被拒绝，上限 20000 格）。"),
                IntProperty("expected_count", "期望被标记（非默认填充色）的单元格总数（可选）。"),
                Property("expected_fill_colors", "object", "各标记类型期望的填充色映射（type→\"#RRGGBB\"），与实际标记色集合比对（可选）。"),
                Property("expected_marked", "array", "期望被标记的地址列表（可选）。"));
        }

        public static Dictionary<string, object> VerifyHyperlink()
        {
            return Object(
                new string[0],
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "目标单元格/区域地址（如 B2 或 A1:A5），用于定位要核对的超链接。"),
                Property("expected_present", "boolean", "期望目标区域是否应存在超链接（默认存在；false 用于断言删除生效）。"),
                Property("expected_url", "string", "预期的外部网址（需与 address 搭配；忽略大小写与末尾斜杠）。"),
                Property("expected_subaddress", "string", "预期的工作簿内部位置（需与 address 搭配，如 Sheet1!A1；忽略 $ 与引号）。"),
                Property("expected_text", "string", "预期的显示文本（需与 address 搭配；忽略大小写与首尾空白）。"),
                IntProperty("expected_count", "预期的整表超链接总数（可选，独立于 address 断言）。"));
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
