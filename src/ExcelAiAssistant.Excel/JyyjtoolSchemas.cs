using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// JYYJtool 借鉴工具（文本/数字/表格/工作表/复制粘贴）的 JSON Schema 定义。
    /// </summary>
    public static class JyyjtoolSchemas
    {
        // ---------- ③ 文本语义 ----------

        public static Dictionary<string, object> ExtractTextCharacters()
        {
            return Object("address", Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要处理的单列区域。"),
                EnumProp("kind", new[] { "chinese", "english", "number", "alnum", "special" }),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> InsertTextIntoCell()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要处理的单个单元格。"),
                Prop("insertion", "要插入的子串。"),
                EnumProp("when", new[] { "char", "separator" }),
                EnumProp("where", new[] { "before", "after" }),
                IntProp("index", "字符下标（0 起）或分隔符第 N 次出现（1 起，缺省 1）。"),
                Prop("separator", "when=separator 时的分隔符，如逗号、顿号。"),
                Prop("target_address", "输出单元格；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> DeleteTextMatches()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要处理的区域。"),
                Prop("pattern", "要删除的正则表达式（.NET 语法）。"),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> ConvertTextToPinyin()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要转换的文本区域。"),
                EnumProp("style", new[] { "full", "initial" }),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> ConvertChineseVariant()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要转换的文本区域。"),
                EnumProp("direction", new[] { "simplified_to_traditional", "traditional_to_simplified" }),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> AmountToUppercase()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "存放数值金额的区域（如合计行）。"),
                Prop("target_address", "输出区域；缺省写到原区域右侧相邻列。"));
        }

        // ---------- ③ 数字 ----------

        public static Dictionary<string, object> BankersRoundValues()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要四舍五入的数值区域。IncludeCells(=bankers rounding)。"),
                IntProp("decimals", "保留小数位数，默认 2。"),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> PickNumbersToSum()
        {
            return Object("candidates_address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("candidates_address", "候选数字所在区域。"),
                NumberProp("target", "要找的目标和。"),
                NumberProp("tolerance", "允许误差，默认 0.01。"),
                Prop("output_address", "写入选中结果的位置；缺省只返回不写入。"));
        }

        public static Dictionary<string, object> FillRandomNumbers()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要填充随机数的目标区域。"),
                NumberProp("min", "最小值，默认 0。"),
                NumberProp("max", "最大值，默认 100。"),
                IntProp("decimals", "小数位数，默认 0。"));
        }

        public static Dictionary<string, object> FillNumberSequence()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "序列起始单元格。"),
                NumberProp("start", "起始值；缺省用当前单元格值。"),
                NumberProp("step", "步长，默认 1。"),
                IntProp("count", "序列长度，默认 10。"),
                EnumProp("direction", new[] { "down", "right" })); 
        }

        public static Dictionary<string, object> GeneratePermutationList()
        {
            return Object("output_address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                IntProp("n", "元素个数（1..最大值，超出将截断）。"),
                Prop("output_address", "排列结果输出区域的起始单元格。"),
                BoolProp("combination", "true=组合（取 2 组合数），false=排列，默认 false。"));
        }

        // ---------- ③ 行列表格 ----------

        public static Dictionary<string, object> InsertBlankRowsEveryN()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "受影响的数据区域。"),
                IntProp("interval", "每隔 N 行插入，默认 1。"),
                IntProp("count", "每次插入的行数，默认 1。"));
        }

        public static Dictionary<string, object> DeleteBlankRowsRange()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要清理的数据区域。"),
                BoolProp("only_fully_blank", "true=仅整行为空的才删，false=首列为空的删，默认 false。"));
        }

        public static Dictionary<string, object> MakePayslipRows()
        {
            return Object("header_address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("header_address", "表头区域（工资条重复的表头行）。"),
                Prop("data_address", "数据区域（每行数据上方插入表头）。"),
                Prop("output_address", "结果起始单元格；缺省原区域原位展开。"));
        }

        public static Dictionary<string, object> FillBlankCellsDown()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要向下填充空单元格的区域。"),
                BoolProp("copy_value", "true=空单元格沿用上方非空值，false=沿用上方公式，默认 true。"));
        }

        public static Dictionary<string, object> MergeDuplicateCells()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要合并同类单元格的单列区域。"),
                BoolProp("first_only", "true=每个分组仅保留左上单元格并合并，默认 true。"));
        }

        public static Dictionary<string, object> FuzzyMatchColumns()
        {
            return Object("first_address", "second_address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("first_address", "第一列数据（基准）。"),
                Prop("second_address", "第二列数据（被匹配）。"),
                Prop("output_address", "匹配结果输出区域（成对输出，含是否达标列）。"),
                NumberProp("threshold", "相似度阈值 0..1，默认 0.6。"));
        }

        public static Dictionary<string, object> WrapFormulasWithIfError()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要包裹 IFERROR 的公式区域。"),
                Prop("fallback", "错误时返回的值，默认 \"\"。"),
                Prop("target_address", "输出区域；缺省原位覆盖。"));
        }

        public static Dictionary<string, object> Convert2dTo1d()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "宽表数据区域。"),
                IntProperty("key_columns", "前 N 列作为标识列，默认 1。"),
                Prop("output_address", "长表输出起始单元格；缺省写到原区域下方。"));
        }

        public static Dictionary<string, object> Convert1dTo2d()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "一维数据区域。"),
                IntProperty("columns_per_row", "每行元素个数，默认 3。"),
                Prop("output_address", "二维输出起始单元格；缺省原位。"));
        }

        public static Dictionary<string, object> CreatePivotTable()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "明细数据区域（首行为表头）。"),
                IntPosProp("row_field", "行标签列号（1 起）。"),
                IntPosProp("value_field", "数值统计列号（1 起）。"),
                EnumProp("aggregation", new[] { "sum", "count", "average" }),
                Prop("output_address", "透视结果起始单元格；缺省写到原区域下方。"));
        }

        public static Dictionary<string, object> CrossReferenceColumns()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "源数据区域（含键列与待回填列）。"),
                IntPosProp("key_column", "源区域中作为匹配键的列号（1 起）。"),
                IntPosProp("value_column", "源区域中回填匹配值的列号（1 起，可为空白列）。"),
                Prop("lookup_address", "查找区域（同一工作表内，含键列与值列）。"),
                IntPosProp("lookup_key_column", "查找区域中键列号（1 起）。"),
                IntPosProp("lookup_value_column", "查找区域中值列号（1 起）。"),
                BoolProp("last_match", "true=重复键取最后一行，false=取第一行，默认 false。"));
        }

        // ---------- ③ 工作表 ----------

        public static Dictionary<string, object> SplitSheetByField()
        {
            return Object("address",
                Prop("sheet", "源工作表，默认活动工作表。"),
                Prop("address", "数据区域（首行为表头）。"),
                IntPosProp("field_column", "用于拆分的字段列号（1 起）。"),
                Prop("sheet_prefix", "新工作表明称前缀，默认取字段值（自动去重/截断）。"),
                BoolProp("keep_header", "新表是否保留表头，默认 true。"));
        }

        public static Dictionary<string, object> CreateSheetByValue()
        {
            return Object("name_cell",
                Prop("sheet", "存有工作表名的单元格所在工作表，默认活动工作表。"),
                Prop("name_cell", "其值将作为新工作表明称（如 A1）。"),
                Prop("base_sheet", "要复制到新表的工作表；缺省仅新建空白表。"),
                EnumProp("position", new[] { "end", "before", "after" }),
                Prop("reference", "position=before/after 时的参考工作表名。"));
        }

        public static Dictionary<string, object> CompareTwoSheets()
        {
            return Object("sheet_b",
                Prop("sheet_a", "基准工作表，默认活动工作表。"),
                Prop("address_a", "基准数据区域。"),
                Prop("sheet_b", "对照工作表。"),
                Prop("address_b", "对照数据区域。"),
                Prop("output_address", "对比差异输出起始单元格；缺省只返回不写入。"));
        }

        public static Dictionary<string, object> ProtectSheetWatermark()
        {
            return Object("action",
                EnumProp("action", new[] { "protect", "unprotect", "watermark" }),
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("password", "保护密码；缺省无密码。"),
                Prop("watermark", "action=watermark 时写入的水印文本。"),
                Prop("watermark_cell", "水印文本写入的单元格，如 A1。"));
        }

        // ---------- ③ 复制粘贴增强 ----------

        public static Dictionary<string, object> CopyRangeAsImage()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要复制为图片的区域。"),
                EnumProp("appearance", new[] { "screen", "as_shown", "printer" }),
                Prop("anchor", "粘贴图片的锚点单元格，默认 A1。"));
        }

        public static Dictionary<string, object> CopyVisibleRangeOnly()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "在筛选状态下的可见数据区域。"),
                Prop("target_address", "复制可见行到目标区域起始单元格。"));
        }

        public static Dictionary<string, object> SmartPasteDelimiter()
        {
            return Object("address",
                Prop("sheet", "目标工作表，默认活动工作表。"),
                Prop("address", "要拆分的数据区域。"),
                EnumProp("delimiter", new[] { "auto", "comma", "semicolon", "tab", "pipe", "space", "custom" }),
                Prop("custom", "delimiter=custom 时使用的自定义分隔符，可多字符（为空则报 INVALID_DELIMITER）。"),
                Prop("output_address", "拆分结果输出起始单元格。"));
        }

        // ---- 各字段类型工具 ----
        private static Dictionary<string, object> Object(params object[] props)
        {
            var map = new Dictionary<string, object>();
            string required = null;
            foreach (var prop in props)
            {
                if (prop is string s)
                {
                    required = s;
                    continue;
                }

                var dict = (Dictionary<string, object>)prop;
                map[Convert.ToString(dict["name"])] = dict;
            }

            var schema = new Dictionary<string, object> { ["type"] = "object", ["properties"] = map };
            var names = (required ?? string.Empty)
                .Split(',')
                .Where(n => !string.IsNullOrEmpty(n))
                .ToArray();
            if (names.Length > 0)
            {
                schema["required"] = names;
            }

            return schema;
        }

        private static Dictionary<string, object> Prop(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["description"] = description };
        }

        private static Dictionary<string, object> EnumProp(string name, string[] values)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
        }

        private static Dictionary<string, object> IntProp(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["minimum"] = 0, ["description"] = description };
        }

        private static Dictionary<string, object> IntPosProp(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["minimum"] = 1, ["description"] = description };
        }

        private static Dictionary<string, object> IntProperty(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["description"] = description };
        }

        private static Dictionary<string, object> NumberProp(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "number", ["description"] = description };
        }

        private static Dictionary<string, object> BoolProp(string name, string description)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "boolean", ["description"] = description };
        }
    }
}