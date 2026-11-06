using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    public static class SecondPhaseSchemas
    {
        public static Dictionary<string, object> ExportMarkdown()
        {
            return Object(
                null,
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要导出的区域；缺省用当前选区。"),
                Property("include_headers", "boolean", "是否把首行作为表头，默认 true。"));
        }

        public static Dictionary<string, object> ImportMarkdown()
        {
            return Object(
                Names("markdown", "address"),
                Property("markdown", "string", "Markdown 表格文本（以 | 分隔）。"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "写入区域的左上角地址。"),
                EnumProperty("value_mode", new[] { "auto", "value", "formula" }));
        }

        public static Dictionary<string, object> MergeSplit()
        {
            return Object(
                Names("action", "address"),
                EnumProperty("action", new[] { "merge", "unmerge", "merge_similar", "fill_merged" }),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "处理区域地址。"),
                EnumProperty("direction", new[] { "column", "row" }));
        }

        public static Dictionary<string, object> ResetValueType()
        {
            return Object(
                Names("address", "target_type"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要重置类型的区域。"),
                EnumProperty("target_type", new[] { "text", "number", "date" }));
        }

        public static Dictionary<string, object> TransformFormulas()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "包含公式的区域。"),
                BoxArray("actions", "按顺序应用的公式转换动作。",
                    Object(null,
                        EnumProperty("type", new[] { "iferror", "ifna", "round", "custom" }),
                        Property("fallback", "string", "iferror/ifna 的回退值。"),
                        Property("digits", "integer", "round 的小数位数。"),
                        Property("function", "string", "custom 模式的函数名。"),
                        Property("args", "string", "custom 模式的附加参数。"))));
        }

        public static Dictionary<string, object> ConvertRefs()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "包含公式的区域。"),
                EnumProperty("mode", new[] { "to_absolute", "to_relative" }));
        }

        public static Dictionary<string, object> TraceRefs()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "含公式的目标单元格。"));
        }

        public static Dictionary<string, object> DrillRange()
        {
            return Object(
                Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "要下钻展开的单元格或汇总区。"));
        }

        public static Dictionary<string, object> CopyVisibleRows()
        {
            return Object(
                Names("source_address"),
                Property("source_sheet", "string", "源工作表，默认活动工作表。"),
                Property("source_address", "string", "筛选后的数据区域。"),
                Property("target_sheet", "string", "目标工作表。"),
                Property("target_address", "string", "目标区域左上角地址。"));
        }

        public static Dictionary<string, object> PageSetup()
        {
            return Object(
                null,
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                EnumProperty("orientation", new[] { "portrait", "landscape" }),
                EnumProperty("paper_size", new[] { "a4", "a3", "a5", "letter", "legal" }),
                IntProperty("fit_to_pages_wide", "缩放为 N 页宽。"),
                IntProperty("fit_to_pages_tall", "缩放为 N 页高。"),
                IntProperty("scale", "缩放比例 10-400。"),
                BoxArray("margins", "页边距（cm）。",
                    Object(null,
                        Property("top", "number", "上边距 cm。"),
                        Property("bottom", "number", "下边距 cm。"),
                        Property("left", "number", "左边距 cm。"),
                        Property("right", "number", "右边距 cm。"))),
                Property("print_area", "string", "打印区域地址。"),
                Property("center_horizontally", "boolean", "水平居中。"),
                Property("center_vertically", "boolean", "垂直居中。"));
        }

        public static Dictionary<string, object> ExportPdf()
        {
            return Object(
                Names("file_path"),
                Property("file_path", "string", "输出 PDF 的完整路径。目标已存在时需显式 overwrite=true（默认 OVERWRITE_DENIED）。"),
                Property("overwrite", "boolean", "目标 PDF 已存在时是否覆盖（默认 false）。"),
                BoxArray("sheets", "要导出的工作表名；缺省导出整本。", Property("name", "string", "工作表名")));
        }

        public static Dictionary<string, object> MergeWorkbooks()
        {
            return Object(
                Names("source_workbooks"),
                BoxArray("source_workbooks", "要合入的工作簿文件路径列表。", Property("path", "string", "文件路径")));
        }

        public static Dictionary<string, object> SplitWorkbook()
        {
            return Object(
                Names("output_folder"),
                Property("output_folder", "string", "拆分结果输出目录。"),
                Property("overwrite", "boolean", "目标 xlsx 已存在时是否覆盖（默认 false：跳过该表并计入 skipped_sheets，不视为失败）。"),
                BoxArray("sheets", "要拆分的工作表名；缺省全部。", Property("name", "string", "工作表名")));
        }

        public static Dictionary<string, object> SaveWorkbook()
        {
            return Object(
                Names(),
                Property("file_path", "string", "另存路径（.xlsx/.xlsm/.xls）。缺省保存回当前文件；当前工作簿从未存盘时必须提供，否则报 NEED_FILE_PATH。"),
                Property("overwrite", "boolean", "目标文件已存在且不是当前工作簿文件时，必须显式 true 才允许覆盖（默认 false，报 OVERWRITE_DENIED）。"));
        }

        public static Dictionary<string, object> ListExternalLinks()
        {
            return Object(null);
        }

        public static Dictionary<string, object> ReplaceExternalLinks()
        {
            return Object(
                Names("old_link", "new_link"),
                Property("old_link", "string", "原外部链接。"),
                Property("new_link", "string", "替换后的链接。"));
        }

        public static Dictionary<string, object> VbaValidate()
        {
            return Object(
                Names("snippet"),
                Property("snippet", "string", "VBA 代码片段。"));
        }

        public static Dictionary<string, object> VbaRun()
        {
            return Object(
                Names("snippet"),
                Property("snippet", "string", "VBA 代码片段。"),
                Property("macro_name", "string", "要执行的宏名（缺省自动生成入口）。"),
                Property("args", "object", "JSON 数据（对象/数组/标量，支持嵌套），注入 VBA 模块级变量 gXaiArgs；代码内用 XaiArg(\"key\") 读取（缺键返回 Null）。"),
                Property("state_clear", "boolean", "为 true 时先清空 VBA 持久状态（XaiStateGet/XaiStateSet 读写），默认 false。"),
                IntProperty("result_limit_kb", "结构化结果（XaiReturn）上限 KB，默认 1024。"));
        }

        public static Dictionary<string, object> RunPython()
        {
            return Object(
                null,
                Property("code", "string", "要执行的 Python 代码。"),
                Property("script_path", "string", "已有脚本文件路径（与 code 二选一）。"),
                Property("input_data", "object", "注入 Python 命名空间的 input_data 变量（JSON 对象/数组/标量），代码内直接使用。"),
                Property("persist", "boolean", "为 true 时保留执行命名空间跨调用（变量/自定义函数可被下次 persist 调用复用），默认 false 每次隔离。"),
                Property("state_clear", "boolean", "为 true 时先清空持久状态（xai_state 与 persist 命名空间），默认 false。"),
                IntProperty("timeout_ms", "超时毫秒。"),
                IntProperty("output_limit_kb", "输出截断上限 KB。"),
                IntProperty("result_limit_kb", "结构化结果（xai_result）上限 KB，默认 1024。"));
        }

        public static Dictionary<string, object> RunPythonExcel()
        {
            return Object(
                Names("code"),
                Property("code", "string", "要执行的 Python 代码；已自动前置 import xlwings as xw，可直接读写当前活动工作簿（如 xw.books.active.sheets[0].range(\"A1\").value = 42）。"),
                Property("input_data", "object", "注入 Python 命名空间的 input_data 变量（JSON 对象/数组/标量），代码内直接使用。"),
                Property("persist", "boolean", "为 true 时保留执行命名空间跨调用（变量/自定义函数可被下次 persist 调用复用），默认 false 每次隔离。"),
                Property("state_clear", "boolean", "为 true 时先清空持久状态（xai_state 与 persist 命名空间），默认 false。"),
                IntProperty("timeout_ms", "超时毫秒。"),
                IntProperty("output_limit_kb", "输出截断上限 KB。"),
                IntProperty("result_limit_kb", "结构化结果（xai_result）上限 KB，默认 1024。"),
                BoxArray("write_ranges", "AI 声明要写入的区域（执行前快照，失败自动回滚）。支持 'Sheet!A1:B5' 字符串或 {sheet, address} 对象。",
                    Object(null,
                        Property("sheet", "string", "工作表名，缺省活动表。"),
                        Property("address", "string", "区域地址。"))));
        }

        public static Dictionary<string, object> RunCli()
        {
            return Object(
                Names("command"),
                Property("command", "string", "可执行命令名（必须在白名单内）。"),
                BoxArray("args", "命令行参数。", Property("value", "string", "参数")),
                IntProperty("timeout_ms", "超时毫秒。"),
                IntProperty("output_limit_kb", "输出截断上限 KB。"),                Property("cwd", "string", "工作目录。"));
        }

        public static Dictionary<string, object> RunExe()
        {
            return Object(
                Names("path"),
                Property("path", "string", "可执行文件绝对路径（必须在白名单内）。"),
                BoxArray("args", "命令行参数。", Property("value", "string", "参数")),
                IntProperty("timeout_ms", "超时毫秒。"),
                IntProperty("output_limit_kb", "输出截断上限 KB。"));
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

        public static Dictionary<string, object> ListPythonPackages()
        {
            return Object(
                null,
                Property("filter", "string", "按库名过滤的关键字（可选，不区分大小写）。"));
        }

        public static Dictionary<string, object> InstallPythonPackages()
        {
            return Object(
                Names("packages"),
                BoxArray("packages", "要安装的 pip 包名列表（须来自 python_setup_guide 返回的常用库目录，或经用户明确指定）。", Property("pkg", "string", "pip 包名，如 pandas。")),
                IntProperty("timeout_ms", "单个库安装超时毫秒，默认 280000。"));
        }

        public static Dictionary<string, object> PythonSetupGuide()
        {
            return Object(null);
        }
    }

    /// <summary>自定义工具库（沉淀工具）相关参数 schema。</summary>
    internal static class CustomLibrarySchemas
    {
        public static Dictionary<string, object> SaveVbaTool()
        {
            return new Dictionary<string, object>
            {
                ["type"] = "object",
                ["required"] = new[] { "name", "description", "code" },
                ["properties"] = new Dictionary<string, object>
                {
                    ["name"] = Object_("string", "唯一工具名：字母开头，字母/数字/下划线，≤40 字符。"),
                    ["description"] = Object_("string", "用途说明；AI 据此判断何时调用。"),
                    ["code"] = Object_("string", "VBA 代码。")
                }
            };
        }

        private static Dictionary<string, object> Object_(string type, string description)
        {
            return new Dictionary<string, object> { ["type"] = type, ["description"] = description };
        }
    }
}