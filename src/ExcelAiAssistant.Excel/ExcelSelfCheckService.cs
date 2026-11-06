using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// AI 自主闭环的两个服务端原语：
    /// - run_self_check：按主工具 + 实参推导确定性自检计划（复用 AutoVerifyPolicy 中央口径），显式执行 verify_* 并汇总判定。
    /// - apply_correction：按"修正阶梯"顺序尝试多组 工具+参数，每组执行后自带 verify 判定，首个通过即停——把
    ///   "自检失败 → 换一种写法再试 → 直到符合预期"这一循环从模型侧猜谜收敛为服务端可审计的有限重试。
    ///
    /// 安全边界：
    /// - 只执行 registry 中已注册的工具（写模式关闭时写工具本就不在 registry，天然拒绝）；
    /// - 自检计划只允许 ReadOnly 的 verify_*，非只读一律跳过（防组合工具演变成写后门）；
    /// - apply_correction 拒绝 Dangerous 子工具（必须由外层单独走确认门，不得借用组合工具绕过）。
    /// </summary>
    public sealed class ExcelSelfCheckService
    {
        /// <summary>单次 run_self_check 最多执行的计划条数（防一次批量写生成几十次回读拖垮回合）。</summary>
        public const int MaxSelfCheckPlans = 10;

        /// <summary>apply_correction 的单轮尝试上限（硬顶，防模型给超长阶梯）。</summary>
        public const int MaxCorrectionAttempts = 5;

        /// <summary>结果 Data 上报的文本上限（防大区域回读把响应撑爆）。</summary>
        public const int MaxReportedTextChars = 4000;

        private readonly ToolRegistry registry;

        public ExcelSelfCheckService(ToolRegistry registry)
        {
            this.registry = registry ?? throw new ArgumentNullException(nameof(registry));
        }

        // ===================== run_self_check =====================

        public ToolResult RunSelfCheck(Dictionary<string, object> args)
        {
            var tool = Required(args, "tool");
            if (AutoVerifyPolicy.IsVerifyTool(tool))
            {
                return Invalid("INVALID_ARGS", "tool 必须是主操作工具，不能是 verify_* 自检工具本身。");
            }

            var primaryArgs = NormalizeDictionary(Get(args, "args")) ?? new Dictionary<string, object>();
            var affectedRanges = NormalizeStringList(Get(args, "affected_ranges"));
            var resultData = Normalize(Get(args, "result_data"));
            var maxPlans = Math.Max(1, Math.Min(MaxSelfCheckPlans, PInt(args, "max_plans", MaxSelfCheckPlans)));

            var plans = new List<AutoVerifyPlan>();
            AutoVerifyPolicy.TryPlanAll(tool, primaryArgs, plans);
            if (plans.Count == 0)
            {
                // M29 审查（U6）：先排查"参数平铺在顶层"的误诊——真实病因是结构错误而非无可断言期望
                if (primaryArgs.Count == 0 && ContainsForeignTopLevelKeys(args))
                {
                    // M32：screenshot* 是已迁移参数，不能按"平铺实参"指导模型改 args（越改越错）
                    var deprecatedKey = new[] { "screenshot", "screenshot_target", "screenshot_sheet", "screenshot_address" }
                        .FirstOrDefault(args.ContainsKey);
                    return new ToolResult
                    {
                        Ok = false,
                        ErrorCode = deprecatedKey != null ? "PARAM_MOVED" : "INVALID_ARGS",
                        Data = deprecatedKey != null
                            ? ("参数 " + deprecatedKey + " 已于 M32 从 run_self_check 移除：自检现在纯只读、不产出图片。"
                                + "需要图片请单独调用 take_screenshot（写入级工具，需 --allow-write）。")
                            : ("检测到主工具实参被平铺在顶层（如 sheet/address 等直接放在 run_self_check 参数里）。"
                                + "正确形态是嵌套：{\"tool\":\"<主工具>\",\"args\":{<主工具实参>}}。请把顶层实参移入 args 对象后重试。")
                    };
                }

                // M29 审查（U2）：豁免命中时透传字典里的人话理由，把"可能属豁免"变成确定诊断
                if (AutoVerifyPolicy.DocumentedWriteToolExemptions.TryGetValue(tool, out var exemptReason))
                {
                    return new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "NO_VERIFY_PLAN",
                        Data = "工具 " + tool + " 属自动自检豁免清单（设计如此）：" + exemptReason
                            + " 请按业务意图显式调用对应 verify_* 工具核对。"
                    };
                }

                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "NO_VERIFY_PLAN",
                    Data = "工具 " + tool + " 没有可从实参推导的确定性自检计划（本参数组合不具备可断言期望）。"
                        + "请按业务意图显式调用对应 verify_* 工具核对。"
                };
            }

            // 主工具结果富化（输出地址/行数/落盘路径只有结果知道）；富化失败即丢弃该计划（宁缺勿误报）。
            var enriched = new List<AutoVerifyPlan>();
            foreach (var plan in plans)
            {
                if (AutoVerifyPolicy.EnrichFromResult(plan, tool, resultData, affectedRanges))
                {
                    enriched.Add(plan);
                }
            }

            if (enriched.Count == 0)
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = "NO_VERIFY_PLAN",
                    Data = "自检计划生成后全部被富化阶段丢弃（缺少输出地址/落盘路径等运行时信息）。"
                        + "请把主工具返回的 result_data 或 affected_ranges 一并传入，或显式调用 verify_* 核对。"
                };
            }

            if (enriched.Count > maxPlans)
            {
                enriched = enriched.Take(maxPlans).ToList();
            }

            var checks = new List<Dictionary<string, object>>();
            var passed = 0;
            var failed = 0;
            var skipped = 0;
            foreach (var plan in enriched)
            {
                var entry = new Dictionary<string, object>
                {
                    ["verify_tool"] = plan.VerifyTool,
                    ["args"] = plan.Args,
                    ["reason"] = plan.Reason
                };

                if (!registry.TryGet(plan.VerifyTool, out var verifyTool))
                {
                    entry["skipped"] = "verify 工具未注册（可能被宿主过滤）";
                    skipped++;
                    checks.Add(entry);
                    continue;
                }

                if (verifyTool.RiskLevel != RiskLevel.ReadOnly)
                {
                    entry["skipped"] = "verify 工具非只读，组合自检拒绝执行";
                    skipped++;
                    checks.Add(entry);
                    continue;
                }

                ToolResult verifyResult;
                try
                {
                    verifyResult = verifyTool.Execute(plan.Args ?? new Dictionary<string, object>());
                }
                catch (Exception exception)
                {
                    verifyResult = new ToolResult { Ok = false, ErrorCode = "VERIFY_EXCEPTION", Data = exception.Message };
                }

                var ok = AutoVerifyPolicy.VerifyPassed(verifyResult);
                // M29 审查（U7）：verdict 归一化透传——取不到显式 verdict 时按判定结论回填
                // （ok→"ok" / 失败→error_code 或 "failed"），AI 读到的恒为非空可决策枚举。
                var verdict = AutoVerifyPolicy.ExtractVerdict(verifyResult);
                entry["ok"] = ok;
                entry["verdict"] = !string.IsNullOrEmpty(verdict)
                    ? verdict
                    : (ok ? "ok" : (verifyResult.ErrorCode ?? "failed"));
                entry["error_code"] = verifyResult.ErrorCode;
                entry["data"] = BoundText(verifyResult.Data);
                if (ok)
                {
                    passed++;
                }
                else
                {
                    failed++;
                }

                checks.Add(entry);
            }

            var overall = passed > 0 && failed == 0 && skipped == 0;
            return new ToolResult
            {
                Ok = overall,
                ErrorCode = overall ? null : (failed > 0 ? "SELF_CHECK_FAILED" : "SELF_CHECK_SKIPPED"),
                Data = new Dictionary<string, object>
                {
                    ["tool"] = tool,
                    ["total"] = checks.Count,
                    ["passed"] = passed,
                    ["failed"] = failed,
                    ["skipped"] = skipped,
                    ["overall_ok"] = overall,
                    ["checks"] = checks
                },
                Warning = overall ? null : "自检未全部通过：请按 checks 中的 verdict / error_code 修正后重试。"
            };
        }

        // ===================== apply_correction =====================

        public ToolResult ApplyCorrection(Dictionary<string, object> args)
        {
            var attempts = NormalizeList(Get(args, "attempts"));
            if (attempts == null || attempts.Count == 0)
            {
                return Invalid("INVALID_ARGS", "attempts 必须是非空数组，元素形如 {tool, args, verify_tool, verify_args}。");
            }

            var maxAttempts = Math.Max(1, Math.Min(MaxCorrectionAttempts, PInt(args, "max_attempts", 3)));
            var stopOnFirstPass = PBool(args, "stop_on_first_pass", true);

            var log = new List<Dictionary<string, object>>();
            var passedAttempt = 0;
            var attemptIndex = 0;
            var executedTools = 0; // M29 审查（B2）：真正执行过主工具的次数（skipped/UNKNOWN/DENIED 不计）
            foreach (var raw in attempts)
            {
                attemptIndex++;
                if (attemptIndex > maxAttempts)
                {
                    log.Add(new Dictionary<string, object>
                    {
                        ["attempt"] = attemptIndex,
                        ["skipped"] = "超过 max_attempts=" + maxAttempts
                    });
                    continue;
                }

                var attempt = NormalizeDictionary(raw);
                if (attempt == null)
                {
                    log.Add(new Dictionary<string, object> { ["attempt"] = attemptIndex, ["error_code"] = "INVALID_ARGS", ["error"] = "attempt 不是对象" });
                    continue;
                }

                var tool = GetString(attempt, "tool");
                if (string.IsNullOrWhiteSpace(tool) || !registry.TryGet(tool, out var descriptor))
                {
                    log.Add(new Dictionary<string, object>
                    {
                        ["attempt"] = attemptIndex,
                        ["tool"] = tool,
                        ["error_code"] = "UNKNOWN_TOOL",
                        ["error"] = "工具未注册（只读模式下写工具不可用）"
                    });
                    continue;
                }

                if (descriptor.RiskLevel == RiskLevel.Dangerous)
                {
                    log.Add(new Dictionary<string, object>
                    {
                        ["attempt"] = attemptIndex,
                        ["tool"] = tool,
                        ["error_code"] = "CONFIRMATION_REQUIRED",
                        ["error"] = "apply_correction 不执行 Dangerous 工具；请单独调用该工具并走确认门。"
                    });
                    continue;
                }

                var toolArgs = NormalizeDictionary(Get(attempt, "args")) ?? new Dictionary<string, object>();
                executedTools++; // M29 审查（B2）：到达 Execute 的才算真正执行
                ToolResult mainResult;
                try
                {
                    mainResult = descriptor.Execute(toolArgs);
                }
                catch (Exception exception)
                {
                    mainResult = new ToolResult { Ok = false, ErrorCode = "TOOL_EXCEPTION", Data = exception.Message };
                }

                var entry = new Dictionary<string, object>
                {
                    ["attempt"] = attemptIndex,
                    ["tool"] = tool,
                    ["main_ok"] = mainResult.Ok,
                    ["main_error_code"] = mainResult.ErrorCode,
                    ["main_warning"] = mainResult.Warning,
                    ["main_data"] = BoundText(mainResult.Data)
                };

                var passed = mainResult.Ok;
                // R31（三方审查）：无 verify_tool 的通过只是"执行成功"，不是"验证通过"。
                // 必须在返回里留下显式降级痕迹，否则模型拿着 ok=true 对用户宣称"已确认无误"（虚假自验）。
                var verified = false;
                var verifyToolName = GetString(attempt, "verify_tool");
                if (!string.IsNullOrWhiteSpace(verifyToolName))
                {
                    entry["verify_tool"] = verifyToolName;
                    if (!registry.TryGet(verifyToolName, out var verifyTool))
                    {
                        entry["verify_skipped"] = "verify 工具未注册";
                        passed = false;
                    }
                    else if (verifyTool.RiskLevel != RiskLevel.ReadOnly)
                    {
                        entry["verify_skipped"] = "verify 工具非只读，拒绝执行";
                        passed = false;
                    }
                    else
                    {
                        var verifyArgs = NormalizeDictionary(Get(attempt, "verify_args")) ?? new Dictionary<string, object>();
                        ToolResult verifyResult;
                        try
                        {
                            verifyResult = verifyTool.Execute(verifyArgs);
                        }
                        catch (Exception exception)
                        {
                            verifyResult = new ToolResult { Ok = false, ErrorCode = "VERIFY_EXCEPTION", Data = exception.Message };
                        }

                        var verifyPassed = AutoVerifyPolicy.VerifyPassed(verifyResult);
                        // R32（审查者实测）：verify 通过不能掩盖主工具失败——否则"COM 报错但目标状态恰好符合"
                        // 会被判成"已执行且已验证"，模型据此宣称成功。
                        passed = mainResult.Ok && verifyPassed;
                        verified = passed;
                        if (mainResult.Ok != verifyPassed)
                        {
                            entry["main_verify_mismatch"] = true;
                        }
                        entry["verify_ok"] = passed;
                        entry["verdict"] = AutoVerifyPolicy.ExtractVerdict(verifyResult);
                        entry["verify_error_code"] = verifyResult.ErrorCode;
                        entry["verify_data"] = BoundText(verifyResult.Data);
                    }
                }

                entry["passed"] = passed;
                entry["verified"] = verified; // false=仅按主工具返回码通过，未经回读验证
                log.Add(entry);

                if (passed)
                {
                    if (passedAttempt == 0)
                    {
                        passedAttempt = attemptIndex;
                    }

                    // R32：只有"已验证通过"才提前停——否则后面带 verify 的尝试根本不会执行，
                    // verified_attempt 恒为 0，等于自己把验证机会跳过了。
                    if (stopOnFirstPass && verified)
                    {
                        break;
                    }
                }
            }

            var overallOk = passedAttempt > 0;
            // R31：首个"带 verify 且通过"的尝试编号；0=没有任何一轮真正验证过（通过只凭返回码）
            var verifiedAttempt = log.Where(e => e.TryGetValue("verified", out var v) && v is true
                    && e.TryGetValue("passed", out var p) && p is true)
                .Select(e => Convert.ToInt32(e["attempt"]))
                .FirstOrDefault();
            var anyPassedUnverified = overallOk && verifiedAttempt == 0;
            // R32：主工具执行成功但验证未完成（verify 拼错/非只读/回读未过）——副作用可能已落盘
            var mainOkButUnverified = log.Count(e =>
                e.TryGetValue("main_ok", out var mo) && mo is true
                && e.TryGetValue("passed", out var ps) && ps is false);

            return new ToolResult
            {
                Ok = overallOk,
                ErrorCode = overallOk ? null : "CORRECTION_FAILED",
                Data = new Dictionary<string, object>
                {
                    ["ok"] = overallOk,
                    ["passed_attempt"] = passedAttempt,
                    // verified_attempt=0 表示通过的尝试没有一轮带 verify——不得视为"已验证落盘"
                    ["verified_attempt"] = verifiedAttempt,
                    ["verified"] = verifiedAttempt > 0,
                    ["executed_attempts"] = executedTools, // M29 审查（B2）：真实执行数（不再混入 skipped/拒绝条目）
                    ["logged_attempts"] = log.Count, // 全部条目数（含 skipped/UNKNOWN/DENIED，审计完整）
                    ["attempts"] = log
                },
                Warning = overallOk
                    ? (anyPassedUnverified
                        ? "本次通过的那一轮未携带可执行的 verify_tool（未填 / 未注册 / 非只读 / 回读未通过），仅凭主工具返回码，未经回读验证；不得向用户表述为「已确认无误」。详见 attempts 的 verify_skipped / verify_error_code。"
                        : null)
                    : (mainOkButUnverified > 0
                        ? "有 " + mainOkButUnverified + " 轮主工具已执行成功但验证未完成（verify_tool 未注册/非只读或回读未过），副作用可能已落盘——请先看 attempts 再决定是否重试，不要直接重复写入。"
                        : "所有修正尝试均未通过校验：请检查各 attempt 的 main_error_code / verdict 后调整参数或换工具。")
            };
        }

        // ===================== 归一化与参数辅助 =====================

        /// <summary>M29 审查（U6）：args 顶层是否混入了 run_self_check 自身参数之外的键（平铺误诊检测）。</summary>
        private static bool ContainsForeignTopLevelKeys(Dictionary<string, object> args)
        {
            if (args == null)
            {
                return false;
            }

            // M32：screenshot* 已整体移到 take_screenshot，这里要给人话指引，
            // 否则会被当成"主工具实参平铺"而让模型一直改 args。
            foreach (var deprecated in new[] { "screenshot", "screenshot_target", "screenshot_sheet", "screenshot_address" })
            {
                if (args.ContainsKey(deprecated))
                {
                    return true; // 由上层统一报 INVALID_ARGS（见下方专用文案）
                }
            }

            var known = new HashSet<string>(StringComparer.Ordinal)
            {
                "tool", "args", "affected_ranges", "result_data", "max_plans",
            };
            foreach (var key in args.Keys)
            {
                if (!known.Contains(key))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>把 JObject / IDictionary 归一化为 Dictionary&lt;string, object&gt;（递归归一化值）。</summary>
        private static Dictionary<string, object> NormalizeDictionary(object node)
        {
            if (node == null)
            {
                return null;
            }

            if (node is JObject jobject)
            {
                var dict = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (var property in jobject.Properties())
                {
                    dict[property.Name] = Normalize(property.Value);
                }

                return dict;
            }

            if (node is IDictionary<string, object> map)
            {
                var dict = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (var pair in map)
                {
                    dict[pair.Key] = Normalize(pair.Value);
                }

                return dict;
            }

            // 任意对象：走 JSON 往返，保持与 MCP/CLI 入参同构。
            try
            {
                var token = JToken.FromObject(node);
                return token is JObject obj ? NormalizeDictionary(obj) : null;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static List<object> NormalizeList(object node)
        {
            if (node == null)
            {
                return null;
            }

            if (node is JArray jarray)
            {
                return jarray.Select(Normalize).ToList();
            }

            if (node is IEnumerable<object> sequence && !(node is string))
            {
                return sequence.Select(Normalize).ToList();
            }

            return null;
        }

        private static List<string> NormalizeStringList(object node)
        {
            var list = NormalizeList(node);
            if (list == null)
            {
                return null;
            }

            var result = new List<string>();
            foreach (var item in list)
            {
                var text = item == null ? null : Convert.ToString(item);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    result.Add(text);
                }
            }

            return result;
        }

        private static object Normalize(object node)
        {
            switch (node)
            {
                case null:
                    return null;
                case JObject _:
                    return NormalizeDictionary(node);
                case JArray _:
                    return NormalizeList(node);
                case JValue jvalue:
                    return jvalue.Value;
                case string _:
                    return node;
                case IDictionary<string, object> _:
                    return NormalizeDictionary(node);
                case IEnumerable<object> sequence:
                    return sequence.Select(Normalize).ToList();
                default:
                    return node;
            }
        }

        private static object Get(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var value))
            {
                return value;
            }

            return null;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            var value = Get(args, key);
            return value == null ? null : Convert.ToString(value);
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var text = GetString(args, key);
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new ArgumentException("缺少必填参数：" + key);
            }

            return text;
        }

        private static int PInt(Dictionary<string, object> args, string key, int fallback)
        {
            var value = Get(args, key);
            if (value != null && int.TryParse(Convert.ToString(value), out var parsed))
            {
                return parsed;
            }

            return fallback;
        }

        private static bool PBool(Dictionary<string, object> args, string key, bool fallback)
        {
            var value = Get(args, key);
            if (value != null && bool.TryParse(Convert.ToString(value), out var parsed))
            {
                return parsed;
            }

            return fallback;
        }

        /// <summary>把结果 Data 压到可读长度（JSON 序列化后截断），防大区域回读撑爆响应。</summary>
        private static object BoundText(object data)
        {
            if (data == null)
            {
                return null;
            }

            string text;
            try
            {
                text = data as string ?? JsonConvert.SerializeObject(data);
            }
            catch (JsonException)
            {
                text = Convert.ToString(data);
            }

            if (text != null && text.Length > MaxReportedTextChars)
            {
                return text.Substring(0, MaxReportedTextChars) + "…（已截断，共 " + text.Length + " 字符）";
            }

            return data;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}