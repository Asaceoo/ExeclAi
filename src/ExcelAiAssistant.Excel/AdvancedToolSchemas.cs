using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    public static class AdvancedToolSchemas
    {
        public static Dictionary<string, object> SheetList()
        {
            return Object(null);
        }

        public static Dictionary<string, object> PreviewSheet()
        {
            return Object(
                null,
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                IntProperty("rows", "预读行数，默认 8（前后各取若干行）。"));
        }

        public static Dictionary<string, object> ReadRange()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要读取的区域地址。"),
                IntProperty("max_cells", "超过该单元格数只返回摘要，默认 2000。"));
        }

        public static Dictionary<string, object> ReadSelection()
        {
            return Object(
                null,
                IntProperty("max_cells", "超过该单元格数只返回摘要，默认 2000。"));
        }

        public static Dictionary<string, object> SelectRange()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要选中并激活的区域地址。"),
                Property("activate", "boolean", "是否激活对应工作表，默认 true。"));
        }

        public static Dictionary<string, object> FindCells()
        {
            return Object(
                Names("what"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "搜索区域，默认整表 UsedRange。"),
                Property("what", "string", "搜索的关键字。"),
                EnumProperty("mode", new[] { "value", "formula" }),
                Property("match_case", "boolean", "是否区分大小写，默认 false。"));
        }

        public static Dictionary<string, object> Filter()
        {
            return Object(
                Names("action"),
                EnumProperty("action", new[] { "set", "clear" }),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "数据区域（set 时必填，首行为标题）。"),
                IntProperty("column", "从 1 开始的列号（set 时必填）。"),
                EnumProperty("operator", new[] { "equals", "not_equals", "contains", "gt", "gte", "lt", "lte", "empty", "not_empty" }),
                Property("value", "string", "筛选阈值。"));
        }

        public static Dictionary<string, object> NamedRange()
        {
            return Object(
                Names("action", "name"),
                EnumProperty("action", new[] { "query", "create", "update", "delete" }),
                Property("name", "string", "命名区域名称。"),
                Property("sheet", "string", "引用所在工作表（创建/更新时使用）。"),
                Property("reference", "string", "引用区域地址，如 A1:B5。"));
        }

        public static Dictionary<string, object> BatchWrite()
        {
            return Object(
                Names("cells"),
                BoxArray("cells", "要写入的单元格列表。",
                    Object(null,
                        Property("address", "string", "单元格地址或区域。"),
                        Property("value", "string", "要写入的值。"),
                        EnumProperty("value_mode", new[] { "value", "formula" }))));
        }

        public static Dictionary<string, object> AppendScratchpad()
        {
            return Object(
                null,
                Property("scratch_sheet", "string", "AI 计算草稿工作表名，默认 _AI_CALC。"),
                Property("input", "string", "输入数据。"),
                Property("formula", "string", "计算公式。"),
                Property("expected_output", "string", "期望输出。"),
                Property("note", "string", "备注。"));
        }

        public static Dictionary<string, object> ClearScratchpad()
        {
            return Object(
                null,
                Property("scratch_sheet", "string", "AI 计算草稿工作表名，默认 _AI_CALC。"));
        }

        public static Dictionary<string, object> Dimension()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "定位行列，如 A1 表示 A 列第 1 行、B:B 表示整列、3:3 表示整行。"),
                IntProperty("quantity", "插入/删除的行列数量，默认 1。"),
                EnumProperty("dimension", new[] { "row", "column" }));
        }

        public static Dictionary<string, object> CopyPaste()
        {
            return Object(
                Names("source_address", "target_address"),
                Property("source_sheet", "string", "源工作表，默认活动工作表。"),
                Property("source_address", "string", "源区域地址。"),
                Property("target_sheet", "string", "目标工作表，默认活动工作表。"),
                Property("target_address", "string", "目标区域左上角地址。"),
                Property("paste_values", "boolean", "仅粘贴值，默认 false。"));
        }

        public static Dictionary<string, object> Toc()
        {
            return Object(
                null,
                Property("sheet", "string", "目录页名称，默认“目录”。"),
                Property("include_hidden", "boolean", "是否包含隐藏工作表，默认 false。"));
        }

        // 内联注册在 ExcelToolHost 的基础读写/格式工具参数 schema
        //（此前用 4 参构造注册 → Parameters 为 null → 模型看不到任何参数，工具“能用但总不对”。）
        public static Dictionary<string, object> WriteRange()
        {
            return Object(
                Names("address", "values"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "目标区域地址，例如 $A$1:$C$10。"),
                TwoDimArray("values", "要写入的二维数据（外层=行，内层=列）。"),
                EnumProperty("value_mode", new[] { "auto", "formula", "value" }));
        }

        public static Dictionary<string, object> WriteCell()
        {
            return Object(
                Names("address", "value"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "目标单元格地址，例如 B3。"),
                CellValue("value", "要写入的值或公式文本（数字/布尔原样写，文本与公式按串写）。"),
                EnumProperty("value_mode", new[] { "auto", "formula", "value" }));
        }

        public static Dictionary<string, object> ClearRange()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要清空内容的区域地址，例如 $A$1:$C$10。"));
        }

        public static Dictionary<string, object> AppendRows()
        {
            return Object(
                Names("start_cell", "values"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("start_cell", "string", "锚点起始单元格，例如 A1；将从此列的末行下方继续追加。"),
                TwoDimArray("values", "要追加的二维数据（外层=行，内层=列）。"));
        }

        public static Dictionary<string, object> PreviewRange()
        {
            return Object(
                null,
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要预览的区域地址；省略则预览当前选区前 20 行。"));
        }

        public static Dictionary<string, object> FormatRange()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要设置格式的区域地址，例如 $A$1:$C$10。"),
                Property("number_format", "string", "Excel 数字格式串，例如 “0.00”、“yyyy-mm-dd”。"),
                Property("bold", "boolean", "是否加粗（true/false）。"),
                Property("italic", "boolean", "是否斜体（true/false）。"),
                Property("font_color", "string", "字体颜色：Excel 颜色序号、#RRGGBB 或颜色中文名。"),
                Property("back_color", "string", "填充背景色：Excel 颜色序号、#RRGGBB 或颜色中文名。"),
                Property("auto_fit_columns", "boolean", "是否自动调整列宽（true/false）。"));
        }

        public static Dictionary<string, object> UndoOperation()
        {
            return Object(
                Names("operation_id"),
                Property("operation_id", "string", "要回滚的操作记录 ID（来自写入结果或历史列表）。"));
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

            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = map
            };

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
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "array",
                ["description"] = description,
                ["items"] = items
            };
        }

        /// <summary>Y1：单元格值参数 = string/number/boolean/null 联合类型（Property 的放宽版）。</summary>
        private static Dictionary<string, object> CellValue(string name, string description)
        {
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = new[] { "string", "number", "boolean", "null" },
                ["description"] = description
            };
        }

        private static Dictionary<string, object> TwoDimArray(string name, string description)
        {
            // Y1 修复：单元格值放宽为 string/number/boolean/null——原 type:string 会引导严格校验客户端
            // 把数值转成 "1" 字符串写入（文本型数字坑），工具端 NormalizeValues/WriteRangeCore 本就收 any。
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "array",
                ["description"] = description,
                ["items"] = new Dictionary<string, object>
                {
                    ["type"] = "array",
                    ["items"] = new Dictionary<string, object> { ["type"] = new[] { "string", "number", "boolean", "null" } }
                }
            };
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
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "integer",
                ["minimum"] = 1,
                ["description"] = description
            };
        }

        private static Dictionary<string, object> EnumProperty(string name, string[] values)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
        }
    }
}