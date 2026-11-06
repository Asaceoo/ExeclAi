using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>截图工具 schema 与注册：take_screenshot（只读，渲染工作表/区域/窗口成图片）。</summary>
    public static class ExcelScreenshotToolsRegistration
    {
        public static void RegisterScreenshotTools(ToolRegistry registry, dynamic application)
        {
            var service = new ExcelScreenshotService((object)application);
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
                + "路由：只要图不要结论 → 直接调本工具（一次调用）；既要结构化 passed/failed 又要图 → 调 run_self_check 传 screenshot=true（一次调用同时返回）。不要两步重复截图。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.TakeScreenshot(args)),
                TakeScreenshot()));
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