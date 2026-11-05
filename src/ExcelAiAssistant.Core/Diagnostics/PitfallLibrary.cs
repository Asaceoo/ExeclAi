using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>一条机器可读的踩坑规则：签名匹配 + 规避建议，源自 docs/踩坑记录与规避手册。</summary>
    public sealed class PitfallRule
    {
        /// <summary>手册章节号 + 短标题，如「3.6 VBA 探针首轮瞬态宏不可用」。</summary>
        public string Id { get; set; }

        /// <summary>症状描述（对应手册「现象」）。</summary>
        public string Symptom { get; set; }

        /// <summary>签名关键字（大小写不敏感，日志任意位置包含即命中）。</summary>
        public IReadOnlyList<string> Signatures { get; set; }

        /// <summary>规避/修复建议（对应手册「规避」「修复」）。</summary>
        public string Advice { get; set; }

        public string DocSection { get; set; }
    }

    /// <summary>踩坑库命中结果。</summary>
    public sealed class PitfallHit
    {
        public PitfallRule Rule { get; set; }

        /// <summary>命中的具体签名。</summary>
        public string MatchedSignature { get; set; }
    }

    /// <summary>
    /// 踩坑失败模式库：把踩坑手册里的已知坑提炼为「日志签名 → 规避建议」规则表，
    /// 供一键自检 AI 研判与失败诊断逐条比对，把泛泛的日志总结升级为精确的已知失败模式匹配。
    /// 新坑沉淀：在手册补充章节后，同步在此追加规则即可被自动识别。
    /// </summary>
    public static class PitfallLibrary
    {
        private static readonly IReadOnlyList<PitfallRule> Rules = new List<PitfallRule>
        {
            new PitfallRule
            {
                Id = "1.1 RCW 被他人 FinalRelease 销毁",
                Symptom = "后续 COM 动态访问抛「COM 对象与其基础 RCW 分开」。",
                Signatures = new[] { "COM 对象与其基础 RCW 分开", "RCW 分开" },
                Advice = "RCW 按对象标识缓存，FinalReleaseComObject 会连带销毁他人持有的 RCW；改为每次操作前重新解析对象、用完即放，绝不跨调用共享/缓存 RCW。",
                DocSection = "踩坑手册 §1.1"
            },
            new PitfallRule
            {
                Id = "1.2 跨程序集 COM 释放不可达",
                Symptom = "编译失败：ExcelContextService.Release 为 internal。",
                Signatures = new[] { "inaccessible due to its protection level", "无法访问，因为它具有一定的保护级别" },
                Advice = "跨程序集释放 COM 用本类私有 SafeRelease（Marshal.IsComObject + FinalReleaseComObject + try/catch），不要引用其他程序集的 internal 释放方法。",
                DocSection = "踩坑手册 §1.2"
            },
            new PitfallRule
            {
                Id = "2.1/2.2 dynamic 分派丢运行时成员",
                Symptom = "RuntimeBinderException：未包含某成员的定义（元组元素名/匿名类成员名）。",
                Signatures = new[] { "RuntimeBinderException", "未包含", "does not contain a definition" },
                Advice = "dynamic 形参会传染下游动态分派；在源头把动态结果显式还原静态类型，跨程序集读匿名结果统一走 JsonConvert + JArray/JObject，不要用 dynamic 导航元素名。",
                DocSection = "踩坑手册 §2.1/§2.2"
            },
            new PitfallRule
            {
                Id = "3.1 WebView2 跨线程访问",
                Symptom = "CoreWebView2 can only be accessed from the UI thread；Ribbon 按钮无反应。",
                Signatures = new[] { "can only be accessed from the UI thread", "InjectRibbonPrompt failed" },
                Advice = "Ribbon onAction 本就在 UI 线程，依赖 WebView2/UI 控件的操作绝不能 Task.Run 丢线程池；同步 GetAwaiter().GetResult()。",
                DocSection = "踩坑手册 §3.1"
            },
            new PitfallRule
            {
                Id = "3.2 VBE 反复增删模块导致慢/窗口句柄失效",
                Symptom = "VBA 执行慢或抛 0x80070578 无效窗口句柄。",
                Signatures = new[] { "0x80070578", "无效的窗口句柄" },
                Advice = "VBA 执行复用常驻标准模块 JYYJHelper（存在则清空重写，不再 Remove），避免 VBE 工程反复重编译；模块/过程名不能以下划线开头。",
                DocSection = "踩坑手册 §3.2"
            },
            new PitfallRule
            {
                Id = "3.3 工作表名双引号 VBA 注入",
                Symptom = "工作表名含双引号时注入的 VBA 字符串被拆断。",
                Signatures = new[] { "Compile error", "VBA 编译错误" },
                Advice = "注入 VBA 前对工作表名按 VBA 规则转义（\" → \"\"），或改用索引/代号访问工作表。",
                DocSection = "踩坑手册 §3.3"
            },
            new PitfallRule
            {
                Id = "3.4 宏被禁用返回晦涩错误",
                Symptom = "宏禁用时 VBA_RUN_ERROR 反复重试无果。",
                Signatures = new[] { "VBA_MACROS_DISABLED", "所有宏都被禁用", "宏不可用" },
                Advice = "先区分「宏被禁用」（引导用户启用宏或改走内置工具）与「首轮瞬态未就绪」（自愈重试一次，见 §3.6）；真禁用时不要反复重试。",
                DocSection = "踩坑手册 §3.4/§3.6"
            },
            new PitfallRule
            {
                Id = "3.5 out-of-proc VBA 弹窗假性超时",
                Symptom = "独立 Excel 进程里 VBA 报错弹模态框，Application.Run 永久阻塞。",
                Signatures = new[] { "harness 超时", "假性超时" },
                Advice = "out-of-proc 自动化无法抑制 VBA 弹窗；错误拦截验证改为断言纯逻辑护栏（正则），不要真跑会报错的 VBA。",
                DocSection = "踩坑手册 §3.5"
            },
            new PitfallRule
            {
                Id = "3.6 VBA 探针首轮瞬态宏不可用",
                Symptom = "首轮 Application.Run 抛 0x800A9C68，第二轮才成功。",
                Signatures = new[] { "0x800A9C68", "cannot run the macro", "may not be available" },
                Advice = "该瞬态≠宏禁用：模块刚改写后 VBE 未就绪。按 ProbeLikelyUnavailable 判定后重追加探针自愈重试一次即可。",
                DocSection = "踩坑手册 §3.6"
            },
            new PitfallRule
            {
                Id = "3.7 失败日志缺真实错误正文",
                Symptom = "日志只有错误码（如 VBA_RUN_ERROR），看不出真实原因。",
                Signatures = new[] { "VBA_RUN_ERROR" },
                Advice = "诊断日志必须并入真实错误消息（CompactLogText 折叠截断），只记 ErrorCode 无法定位问题。",
                DocSection = "踩坑手册 §3.7"
            },
            new PitfallRule
            {
                Id = "3.8 2D 区域 AutoFit 必抛 1004",
                Symptom = "多行多列矩形 .AutoFit 运行时抛 1004。",
                Signatures = new[] { "AutoFit" },
                Advice = "2D .Range(...).AutoFit 是非法写法（静态校验已升为阻断）；改为 Rows.AutoFit / Columns.AutoFit 或单行/单列。",
                DocSection = "踩坑手册 §3.8"
            },
            new PitfallRule
            {
                Id = "3.9 VBA 真机验证触发 VBE 模态",
                Symptom = "SaveAs/重开工作簿触发 VBE 模态或 0x800706BE。",
                Signatures = new[] { "0x800706BE" },
                Advice = "真机验证 VBA 用「新建未存盘内存工作簿 + AutomationSecurity=1」，避免中途 SaveAs/重开。",
                DocSection = "踩坑手册 §3.9"
            },
            new PitfallRule
            {
                Id = "4.1 前端内联 JS 语法错误废掉侧边栏",
                Symptom = "注入脚本一个语法错误导致整个 WebView2 侧边栏白屏/无响应。",
                Signatures = new[] { "SyntaxError", "脚本错误" },
                Advice = "注入脚本先做语法/括号配平校验，注入后跑 WebView2 往返 ping 自检；不要在 C# 中裸拼接复杂 JS。",
                DocSection = "踩坑手册 §4.1"
            },
            new PitfallRule
            {
                Id = "4.4 ExecuteScriptAsync 不等待 Promise",
                Symptom = "把 ExecuteScriptAsync 返回值当执行结果，异步 Promise 结果丢失。",
                Signatures = new[] { "ExecuteScriptAsync" },
                Advice = "ExecuteScriptAsync 不等待返回的 Promise；异步结果需走事件回调或轮询确认，不要据返回值断言成功。",
                DocSection = "踩坑手册 §4.4"
            },
            new PitfallRule
            {
                Id = "确认超时被误报为用户取消",
                Symptom = "用户在等待确认时被系统超时中止，却被报成「用户取消」。",
                Signatures = new[] { "CONFIRM_TIMEOUT", "USER_DENIED" },
                Advice = "CONFIRM_TIMEOUT（超时系统中止）与 USER_DENIED（用户主动取消）语义不同；提示语须引导用户重试并在确认框点「允许执行」。",
                DocSection = "确认三态语义"
            },
            new PitfallRule
            {
                Id = "自检链路失效",
                Symptom = "回合内有可自检写操作但 0 条自检计划执行——写操作全部未经核对。",
                Signatures = new[] { "SELF_CHECK_LINK_SUSPECT", "自检链路疑似失效" },
                Advice = "检查 verify_* 工具是否注册且为只读风险级、AutoVerify 开关是否被关闭、结果富化是否失败；修复后重跑一次写操作，确认 auto_verify 观察消息恢复。",
                DocSection = "自动自检元监控"
            },
            new PitfallRule
            {
                Id = "写入/自检回读不符",
                Symptom = "工具报成功但回读值与预期不符（值时/类型/公式计算错位）。",
                Signatures = new[] { "回读不符", "回读不一致", "自检未通过" },
                Advice = "回读不符说明写入未真正生效或类型被 Excel 重解析（文本型数字/日期）；先用 verify_write_result 逐格核对 expected vs actual，再按类型重写。",
                DocSection = "写入自检闭环"
            }
        };

        /// <summary>在日志/文本中匹配已知踩坑签名，返回全部命中（同一规则可能因多个签名重复命中，已按规则去重）。</summary>
        public static IReadOnlyList<PitfallHit> Match(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return Array.Empty<PitfallHit>();
            }

            var hits = new List<PitfallHit>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var rule in Rules)
            {
                foreach (var signature in rule.Signatures)
                {
                    if (text.IndexOf(signature, StringComparison.OrdinalIgnoreCase) >= 0 &&
                        seen.Add(rule.Id))
                    {
                        hits.Add(new PitfallHit { Rule = rule, MatchedSignature = signature });
                        break;
                    }
                }
            }

            return hits;
        }

        public static IReadOnlyList<PitfallRule> AllRules => Rules;

        /// <summary>把命中结果格式化为可并入研判报告/提示词的行。</summary>
        public static string FormatHits(IEnumerable<PitfallHit> hits)
        {
            var lines = hits
                .Select(h => "· [" + h.Rule.Id + "] " + h.Rule.Symptom + " → 建议：" + h.Rule.Advice + "（" + h.Rule.DocSection + "）")
                .ToList();
            return lines.Count == 0 ? string.Empty : string.Join(Environment.NewLine, lines);
        }
    }
}
