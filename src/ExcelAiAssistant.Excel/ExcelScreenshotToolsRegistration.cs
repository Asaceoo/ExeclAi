using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>截图工具 schema 与注册：take_screenshot（写入级，渲染工作表/区域/窗口成图片）。</summary>
    public static class ExcelScreenshotToolsRegistration
    {
        public static void RegisterScreenshotTools(ToolRegistry registry, dynamic application)
        {
            var service = new ExcelScreenshotService((object)application);
            // M33：只读渲染——不碰剪贴板、不插图表对象，只读模式下也能"看图"
            var renderer = new ExcelRangeRenderer((object)application);
            registry.Register(new ToolDescriptor(
                "render_range_image",
                "把工作表区域渲染成 PNG 图片（只读渲染：不碰剪贴板、不插入图表对象、不修改工作簿），"
                + "输出 file_path 与内联 base64（MCP 以 image 内容块投递）。用于在只读模式下做视觉确认——"
                + "表头层级、空列、数字格式、明显错位。"
                + "\n仅还原：显示文本、加粗、底色、字色、水平对齐；列宽按文本自动测量。"
                + "\n不还原：合并单元格（合并区按普通网格绘制）、边框样式、条件格式、图表/图片/形状、行高与换行；"
                + "超过 120 字的文本会被截断——图上文字不等于单元格真实值，不能作为「格式无误」的证据。"
                + "\n返回：text 载荷含 file_path/width/height/rows/columns/rendered_rows/rendered_columns/truncated；"
                + "图片以 MCP image 内容块投递（多模态模型可直接看到），>3MB 时不内联并给 inline_omitted。"
                + "\n上限 300 行 / 60 列 / 4000 格（联合约束，如 60 列时最多 66 行），超出截断并置 truncated=true。"
                + "\n需要像素级外观请用 take_screenshot（写入级，需 --allow-write）；若工具列表里没有它，说明当前是只读模式，请勿调用。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => renderer.Render(args)),
                RenderRangeImage()));

            registry.Register(new ToolDescriptor(
                "take_screenshot",
                "把工作表/指定区域/Excel 窗口渲染成图片文件（png 或 jpeg），用于模型\"看结果\"做视觉自检。"
                + "target=sheet 截当前或指定工作表的已用区域（空表会返回 EMPTY_RANGE 错误）；"
                + "target=range 截 address 指定区域；target=active_window 截 Excel 主窗口（无头模式下自动降级为工作表截图并给出 warning）。"
                + "返回 file_path/size/宽高；return_base64=true 时附带内联图片数据（仅多模态模型设 true，MCP 会以 image 内容块投递；"
                + "纯文本模型保持 false 并用 file_path 获取截图）。不改动单元格，但注意：sheet/range 模式经 CopyPicture 实现，"
                + "会临时占用并覆盖系统剪贴板内容（用户正在使用剪贴板时须先提示）、临时插入并在结束后删除图表对象（进入撤销栈，可能使工作簿变脏）；"
                + "不改变任何单元格的值或格式。"
                + "截图落盘默认写入 %TEMP%\\JYYJ\\screenshots，超过 24 小时的历史截图会在下次截图时自动清理；"
                + "如需长期保存请显式传 file_path。临时图表对象在截图结束后自动删除，不会改动任何单元格的值或格式。"
                + "本工具只出图、不出判定；需要结构化 passed/failed 请先调 run_self_check（纯只读），通过后按需再截图——两次调用，顺序自检在前。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.TakeScreenshot(args)),
                TakeScreenshot()));
        }

        public static Dictionary<string, object> RenderRangeImage()
        {
            return Object(
                null,
                Property("sheet", "string", "目标工作表名；缺省为活动工作表。"),
                Property("address", "string", "要渲染的区域地址（如 A1:F20）；缺省渲染该表已用区域。"),
                Property("title", "string", "可选：图片顶部标题，缺省为「表名!区域」。"),
                Property("return_base64", "boolean", "默认 true：内联图片数据（>3MB 时不内联，只给 file_path）。纯文本客户端可设 false 只取路径。"));
        }

        public static Dictionary<string, object> TakeScreenshot()
        {
            return Object(
                null,
                Property("target", "string", "截图对象：sheet（默认，工作表已用区域）/ range（address 指定区域）/ active_window（Excel 主窗口）。", new[] { "sheet", "range", "active_window" }),
                Property("sheet", "string", "目标工作表名；缺省为活动工作表。"),
                Property("address", "string", "target=range 时必填的区域地址，如 A1:F20。"),
                Property("file_path", "string", "输出图片路径；缺省写入系统临时目录下的 JYYJ/screenshots（自动补扩展名）。"),
                Property("format", "string", "图片格式：png（默认）/ jpeg。", new[] { "png", "jpeg" }),
                Property("return_base64", "boolean", "是否在结果中内联 base64 图片数据（仅多模态模型设 true；>3MB 时不返回，改用 file_path）。"));
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

        private static Dictionary<string, object> Property(string name, string type, string description, string[] allowed = null)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            if (allowed != null)
            {
                result["enum"] = allowed;
            }

            return result;
        }
    }
}