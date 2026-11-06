using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>AI 自主闭环工具注册：run_self_check（显式自检）与 apply_correction（有限修正阶梯）。</summary>
    public static class ExcelSelfCheckToolsRegistration
    {
        public static void RegisterSelfCheckTools(ToolRegistry registry, dynamic application)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            var service = new ExcelSelfCheckService(registry);
            var shots = application == null ? null : new ExcelScreenshotService((object)application);

            registry.Register(new ToolDescriptor(
                "run_self_check",
                "对一次主操作做确定性回读自检：传入主工具名与其执行时的实参（实参必须嵌套在 args 对象内，勿平铺在顶层），"
                + "服务端按 AutoVerifyPolicy 推导 verify_* 计划并执行，汇总 passed/failed/skipped 与每条 verdict。"
                + "用于 AI 在写入后\"自验\"结果是否符合预期；只有全部计划通过才 overall_ok=true。"
                + "\n判定语义：error_code=NO_VERIFY_PLAN / INVALID_ARGS 表示「这组工具+实参推导不出自检计划」，"
                + "不代表写入失败、更不代表数据没写进去——不要重写或回滚，改用显式 verify_* 工具。"
                + "\nscreenshot 缺省 false：不传时只做 JSON 回读，不碰剪贴板、不插图表、不落盘——这是默认用法，每次写入后都可放心调用。"
                + "\n仅当 screenshot=true 时才走截图通道：临时占用系统剪贴板、临时插入并删除图表对象（进入撤销栈、可能令工作簿变脏），"
                + "并在 %TEMP%\\JYYJ\\screenshots 落一个 PNG（24 小时后自动清理）——但不会改动任何单元格的值或格式。"
                + "\n截图失败（data.screenshot_error / screenshot_message）不影响 overall_ok 与 passed/failed 判定——不要因此重写或回滚数据；"
                + "最多单独重试一次 take_screenshot。"
                + "\n路由：只要图不要结论 → take_screenshot；结论和图都要 → 本工具 screenshot=true（一次调用同时返回）。不要两步重复截图。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => MaybeAttachScreenshot(service.RunSelfCheck(args), args, shots)),
                RunSelfCheck()));

            registry.Register(new ToolDescriptor(
                "apply_correction",
                "有限修正阶梯（AI 自我修改）：按顺序执行 attempts 里每组 {tool,args,verify_tool,verify_args}，"
                + "每组执行后立即用 verify_* 判定，首个「已验证通过」的轮次即停，最多 5 轮。"
                + "\n**判定只看 verified / verified_attempt**：verified=true 才算通过；ok=true 仅表示子工具执行成功，"
                + "verified=false 一律视为「未验证」，不得向用户说已完成 / 已修正 / 确认无误。"
                + "\n安全边界：不执行 Dangerous 级子工具；verify 必须已注册且只读（填错或非只读会让该轮 verify_skipped 直接判失败）；写工具在只读模式下不可用。"
                + "\nverify_args 结构以该 verify 工具 inputSchema 为准（verify_write_result 需 sheet/address/expected_values，"
                + "expected_values 为与 address 同形的二维数组，如 [[\"姓名\",95]]）。",
                RiskLevel.NormalWrite,
                args => ExcelToolHost.Execute(() => service.ApplyCorrection(args)),
                ApplyCorrection()));
        }


        /// <summary>AI 自主闭环：自检结果可选附带一张截图（screenshot=true），让多模态模型
        /// 直接"看见"结果而不是只读 JSON。截图失败不影响自检结论，只在结果里标注。</summary>
        private static ExcelAiAssistant.Core.Models.ToolResult MaybeAttachScreenshot(
            ExcelAiAssistant.Core.Models.ToolResult result,
            System.Collections.Generic.Dictionary<string, object> args,
            ExcelScreenshotService shots)
        {
            if (result == null || shots == null || !Flag(args, "screenshot")) { return result; }

            var shotArgs = new System.Collections.Generic.Dictionary<string, object>();
            // 审查 R28：必须显式要 base64，否则截图只回 file_path，image 内容块永远发不出去（实测空转）。
            shotArgs["return_base64"] = true;
            if (args != null)
            {
                if (args.TryGetValue("screenshot_target", out var target) && target != null) { shotArgs["target"] = target; }
                if (args.TryGetValue("screenshot_sheet", out var sheet) && sheet != null) { shotArgs["sheet"] = sheet; }
                if (args.TryGetValue("screenshot_address", out var addr) && addr != null)
                {
                    shotArgs["address"] = addr;
                    // 只给 address 不给 target 时按区域截，否则会被静默降级成整表
                    if (!shotArgs.ContainsKey("target")) { shotArgs["target"] = "range"; }
                }
            }

            try
            {
                var shot = shots.TakeScreenshot(shotArgs);
                var data = result.Data as System.Collections.Generic.Dictionary<string, object>;
                if (data == null)
                {
                    // diagnosis 而非 previous：模型会把 previous 误读成"上一次的数据/已回滚"
                    data = new System.Collections.Generic.Dictionary<string, object> { ["diagnosis"] = result.Data };
                    result.Data = data;
                }
                data["screenshot"] = shot.Data;
                // MCP 的 image 通道只认 Data 顶层 base64/mime_type —— 必须平铺一份，不能只嵌套
                if (shot.Data is System.Collections.Generic.Dictionary<string, object> sd)
                {
                    if (sd.TryGetValue("base64", out var b64) && b64 != null) { data["base64"] = b64; }
                    if (sd.TryGetValue("format", out var fmt) && fmt != null) { data["mime_type"] = "image/" + fmt; }
                    if (sd.TryGetValue("file_path", out var fp) && fp != null) { data["screenshot_path"] = fp; }
                }
                if (!shot.Ok)
                {
                    data["screenshot_error"] = shot.ErrorCode;
                    data["screenshot_message"] = shot.Warning;
                }
            }
            catch (Exception ex)
            {
                // 截图失败绝不能影响自检结论；但信息必须有地方落——Data 是字符串时也要能记录
                var d = result.Data as System.Collections.Generic.Dictionary<string, object>;
                if (d == null)
                {
                    d = new System.Collections.Generic.Dictionary<string, object> { ["diagnosis"] = result.Data };
                    result.Data = d;
                }
                d["screenshot_error"] = "SCREENSHOT_EXCEPTION";
                d["screenshot_message"] = ex.Message;
            }
            return result;
        }

        private static bool Flag(System.Collections.Generic.Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out var v) || v == null) { return false; }
            if (v is bool b) { return b; }
            var text = Convert.ToString(v).Trim();
            if (text == "1") { return true; }
            if (text == "0") { return false; }
            return bool.TryParse(text, out var parsed) && parsed;
        }

        public static Dictionary<string, object> RunSelfCheck()
        {
            return Object(
                new[] { "tool" },
                Property("tool", "string", "刚执行过的主工具名（不能是 verify_* 自检工具本身）。"),
                ObjectProperty("args", "主工具执行时使用的实参对象（用于推导确定性期望）。"),
                ArrayProperty("affected_ranges", "string", "主工具返回的受影响区域（如 \"Sheet1!A1:B2\"），用于追加行等需要运行时地址的自检富化。"),
                ObjectProperty("result_data", "主工具返回的 data 对象（提供输出地址/行数/落盘路径等运行时信息）。"),
                IntProperty("max_plans", "最多执行的自检计划条数，默认 10，上限 10。"),
                Property("screenshot", "boolean", "可选：true 时自检完成后附带一张截图（base64 以 image 内容块回传），让多模态模型直接看图确认。默认 false。"),
                Property("screenshot_target", "string", "可选截图范围：sheet（整表）/range（区域）/active_window（窗口）。默认 sheet。"),
                Property("screenshot_sheet", "string", "可选：截图目标工作表名，默认活动表。"),
                Property("screenshot_address", "string", "可选：当 screenshot_target=range 时的区域地址，如 A1:F20。"));
        }

        public static Dictionary<string, object> ApplyCorrection()
        {
            var attemptItems = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = new Dictionary<string, object>
                {
                    ["tool"] = new Dictionary<string, object> { ["type"] = "string", ["description"] = "本轮要执行的工具名。" },
                    ["args"] = new Dictionary<string, object> { ["type"] = "object", ["description"] = "本轮工具实参。" },
                    ["verify_tool"] = new Dictionary<string, object> { ["type"] = "string", ["description"] = "本轮执行后用于判定的只读 verify 工具名。强烈建议必填：省略即 verified=false（只按主工具返回码通过），不得当作已验证落盘。" },
                    ["verify_args"] = new Dictionary<string, object>
                    {
                        ["type"] = "object",
                        // M29 审查（U3）：指明 verify_args 的内层结构以该 verify 工具的 inputSchema 为准，
                        // 并给出最常见形态示例——AI 只看本 schema 也能构造正确参数。
                        ["description"] = "verify 工具的实参；内层结构以该 verify 工具在 tools/list 中的 inputSchema 为准。"
                            + "例：verify_write_result 需 sheet/address/expected_values，其中 expected_values 为与 address 同形的"
                            + "二维数组（外层行、内层列），如 [[\"姓名\",95]]。"
                    }
                },
                ["required"] = new[] { "tool" }
            };

            return Object(
                new[] { "attempts" },
                new Dictionary<string, object>
                {
                    ["name"] = "attempts",
                    ["type"] = "array",
                    ["description"] = "修正阶梯：按顺序尝试的对象数组，元素 {tool, args, verify_tool, verify_args}。",
                    ["items"] = attemptItems
                },
                IntProperty("max_attempts", "最多执行几轮，默认 3，上限 5。"),
                Property("stop_on_first_pass", "boolean", "首个通过即停止，默认 true。"));
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

        private static Dictionary<string, object> ObjectProperty(string name, string description)
        {
            return Property(name, "object", description);
        }

        private static Dictionary<string, object> ArrayProperty(string name, string itemType, string description)
        {
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "array",
                ["description"] = description,
                ["items"] = new Dictionary<string, object> { ["type"] = itemType }
            };
        }

        private static Dictionary<string, object> IntProperty(string name, string description)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = "integer", ["minimum"] = 0 };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }
    }
}