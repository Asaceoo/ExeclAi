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

            registry.Register(new ToolDescriptor(
                "run_self_check",
                "对一次主操作做确定性回读自检：传入主工具名与其执行时的实参（实参必须嵌套在 args 对象内，勿平铺在顶层），"
                + "服务端按 AutoVerifyPolicy 推导 verify_* 计划并执行，汇总 passed/failed/skipped 与每条 verdict。"
                + "用于 AI 在写入后\"自验\"结果是否符合预期；只有全部计划通过才 overall_ok=true。"
                + "\n判定语义：error_code=NO_VERIFY_PLAN / INVALID_ARGS 表示「这组工具+实参推导不出自检计划」，"
                + "不代表写入失败、更不代表数据没写进去——不要重写或回滚，改用显式 verify_* 工具。"
                + "\n本工具是纯只读：不做任何截图、不碰剪贴板、不插图表、不落盘——每次写入后都可放心调用。"
                + "\n需要图片时另调 take_screenshot（写入级，需 --allow-write）。",
                RiskLevel.ReadOnly,
                args => ExcelToolHost.Execute(() => service.RunSelfCheck(args)),
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


        public static Dictionary<string, object> RunSelfCheck()
        {
            return Object(
                new[] { "tool" },
                Property("tool", "string", "刚执行过的主工具名（不能是 verify_* 自检工具本身）。"),
                ObjectProperty("args", "主工具执行时使用的实参对象（用于推导确定性期望）。"),
                ArrayProperty("affected_ranges", "string", "主工具返回的受影响区域（如 \"Sheet1!A1:B2\"），用于追加行等需要运行时地址的自检富化。"),
                ObjectProperty("result_data", "主工具返回的 data 对象（提供输出地址/行数/落盘路径等运行时信息）。"),
                IntProperty("max_plans", "最多执行的自检计划条数，默认 10，上限 10。"));
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