using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Net;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Skills;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// SidebarBridge 分域 partial（阶段 A 机械拆分，见 docs/AI交接文档.md）。
    /// 公共 API 即 WebView2 前端契约（hostObjects.excelAiBridge.*），签名不得改动。
    /// </summary>
    public sealed partial class SidebarBridge : IDisposable
    {
        public void Dispose()
        {
            InjectPrompt = null;
            NavigateView = null;
            ExcelHostFactory = null;
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
        }

        private static byte[] ProtectApiKey(string apiKey)
        {
            return ProtectedData.Protect(
                Encoding.UTF8.GetBytes(apiKey),
                Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                DataProtectionScope.CurrentUser);
        }

        private static string UnprotectApiKey(byte[] cipherBytes)
        {
            var plain = ProtectedData.Unprotect(
                cipherBytes,
                Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plain);
        }

        private async Task<ConfirmationOutcome> ConfirmTool(ToolEnvelope envelope, RiskLevel riskLevel)
        {
            var code = ExtractCode(envelope.Tool, envelope.Args);
            if (code != null)
            {
                return await ConfirmCodePreview(envelope, riskLevel, code);
            }

            if (riskLevel == RiskLevel.ReadOnly)
            {
                return ConfirmationOutcome.Confirmed;
            }

            // 除高危(Dangerous)外一律直接执行，不再弹确认框；仅高危操作交由用户确认。
            if (riskLevel != RiskLevel.Dangerous)
            {
                return ConfirmationOutcome.Confirmed;
            }

            var target = DescribeTarget(envelope.Args);
            return await Task.Run(() => ShowNativeConfirm(
                envelope,
                riskLevel.ToString(),
                target,
                envelope.Reason ?? "执行模型请求的操作",
                code: null,
                language: null));
        }

        /// <summary>
        /// 以原生置顶模态框请求用户确认：替换原先「侧边栏内 overlay + 120s 静默超时」的方式。
        /// 原生框会置顶、闪烁任务栏并播放提示音，保证用户一定能看到；不自动超时（除非用户关闭 Excel）。
        /// </summary>
        private ConfirmationOutcome ShowNativeConfirm(
            ToolEnvelope envelope,
            string riskText,
            string target,
            string reason,
            string code,
            string language)
        {
            try
            {
                return NativeConfirmDialog.ShowDialog(
                    "确认执行",
                    envelope.Tool,
                    target,
                    reason,
                    riskText,
                    code,
                    language);
            }
            catch (Exception ex)
            {
                DiagnosticsLog.Write($"Native confirm fallback to overlay: {ex.Message}");
                // 原生框打开失败（极端环境）时回退到前端 overlay 等待，避免高危操作静默放行。
                return FallbackOverlayConfirm(envelope, reason, riskText).GetAwaiter().GetResult();
            }
        }

        /// <summary>原生确认框不可用时的兜底：走前端 overlay，仍有 120s 超时（不误报为用户取消，仅提示超时）。</summary>
        private async Task<ConfirmationOutcome> FallbackOverlayConfirm(ToolEnvelope envelope, string reason, string riskText)
        {
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (pendingConfirmations)
            {
                pendingConfirmations[envelope.CallId] = completion;
            }

            Publish("confirmation-request", new
            {
                call_id = envelope.CallId,
                tool = envelope.Tool,
                target = DescribeTarget(envelope.Args),
                reason,
                risk_level = riskText
            });

            var completed = await Task.WhenAny(completion.Task, Task.Delay(TimeSpan.FromSeconds(120)));
            lock (pendingConfirmations)
            {
                pendingConfirmations.Remove(envelope.CallId);
            }

            if (completed != completion.Task)
            {
                return ConfirmationOutcome.TimedOut;
            }

            return completion.Task.Result ? ConfirmationOutcome.Confirmed : ConfirmationOutcome.Declined;
        }

        /// <summary>
        /// 脚本/命令类工具：把代码放进确认面板，由用户审阅后选择执行/取消（复制的代码由前端处理）。
        /// </summary>
        private Task<ConfirmationOutcome> ConfirmCodePreview(ToolEnvelope envelope, RiskLevel riskLevel, CodeSnippet code)
        {
            if (riskLevel == RiskLevel.ReadOnly)
            {
                return Task.FromResult(ConfirmationOutcome.Confirmed);
            }

            return Task.FromResult(ShowNativeConfirm(
                envelope,
                "Code",
                DescribeTarget(envelope.Args),
                envelope.Reason ?? "执行模型请求的脚本/命令",
                code.Code,
                code.Language));
        }

        private static CodeSnippet ExtractCode(string tool, IReadOnlyDictionary<string, object> args)
        {
            var key = tool switch
            {
                "run_vba_snippet" => "snippet",
                "run_python_code" => "code",
                "run_python_excel" => "code",
                "run_cli_command" => "command",
                "run_exe_command" => "command",
                _ => null
            };
            if (key == null || args == null || !args.TryGetValue(key, out var raw) || raw == null)
            {
                return null;
            }

            var text = Convert.ToString(raw);
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var language = tool switch
            {
                "run_vba_snippet" => "VBA",
                "run_python_code" => "Python",
                "run_python_excel" => "Python",
                _ => "Shell"
            };
            return new CodeSnippet { Language = language, Code = text };
        }

        private sealed class CodeSnippet
        {
            public string Language { get; set; }

            public string Code { get; set; }
        }

        /// <summary>从代码里提取第一个 Sub/Function 名作为可读工具名，否则用安全前缀。</summary>
        private static string AutoToolName(string code)
        {
            if (!string.IsNullOrWhiteSpace(code))
            {
                const string pattern = @"(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property)\s+([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*)";
                var match = System.Text.RegularExpressions.Regex.Match(code, pattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Multiline);
                if (match.Success)
                {
                    var proc = match.Groups[1].Value;
                    return "proc_" + (proc.Length > 30 ? proc.Substring(0, 30) : proc);
                }
            }

            return "ct_vba_" + Guid.NewGuid().ToString("N").Substring(0, 6);
        }

        /// <summary>沉淀时缺省描述模板：从代码提取过程名与首段注释语义，生成“做什么、何时可用”的一句话说明。</summary>
        private static string BuildToolDescription(string code, string name)
        {
            var proc = "";
            if (!string.IsNullOrWhiteSpace(code))
            {
                const string pattern = @"(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property)\s+([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]*)";
                var match = System.Text.RegularExpressions.Regex.Match(code, pattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Multiline);
                if (match.Success)
                {
                    proc = match.Groups[1].Value;
                }
            }

            var hint = ExtractFirstComment(code);
            if (!string.IsNullOrWhiteSpace(hint))
            {
                return hint;
            }

            return string.IsNullOrEmpty(proc)
                ? $"{name}：自动沉淀的 VBA，可直接复用。"
                : $"{proc}：自动沉淀的 VBA，可直接复用。";
        }

        /// <summary>沉淀时缺省名：优先取代码首个 ' 注释（通常是开发者写的功能说明），凝练成一句话；
        /// 无注释时回落为过程/工具名。保证模板库展示的是“能看懂做什么”的简短描述，而非模板套话。</summary>
        private static string ExtractFirstComment(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return null;
            }

            foreach (var raw in code.Split('\n'))
            {
                var line = raw.Trim();
                // 跳过插入选区标记（'#XAI_SCOPE:…）等以 # 开头的内部注释，避免把标记当描述
                if (line.StartsWith("'") && line.Length > 1 && !line.StartsWith("'#"))
                {
                    var text = line.TrimStart('\'').Trim();
                    return text.Length > 24 ? text.Substring(0, 24) : text;
                }
            }

            return null;
        }

        private static string DescribeTarget(IReadOnlyDictionary<string, object> args)
        {
            foreach (var key in new[] { "address", "source_address", "target_address", "sheet", "name", "path", "file_path" })
            {
                if (args.TryGetValue(key, out var value) && value != null)
                {
                    return key == "sheet" ? $"工作表 {value}" : Convert.ToString(value);
                }
            }

            return "当前选区";
        }

        private string ApplySkillDirective(ToolRegistry registry, string prompt, out string skillBody)
        {
            return ApplySkillDirective(registry, prompt, out skillBody, out _);
        }

        private string ApplySkillDirective(ToolRegistry registry, string prompt, out string skillBody, out string skillSource)
        {
            skillBody = null;
            skillSource = null;
            var match = Regex.Match(
                prompt ?? string.Empty,
                @"/(skill|skill:)([\w-]+)",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return "none";
            }

            var skillName = match.Groups[2].Value;
            var skills = LoadSkills();
            var skill = skills.FirstOrDefault(item =>
                string.Equals(item.Name, skillName, StringComparison.OrdinalIgnoreCase) && item.Valid);
            if (skill == null)
            {
                throw new InvalidOperationException($"未找到技能：{skillName}");
            }

            skillBody = skill.Body ?? skill.Description;
            skillSource = string.IsNullOrWhiteSpace(skill.Source) ? (skill.IsBuiltIn ? "builtin" : "local") : skill.Source;
            if (skill.Tools.Count > 0)
            {
                registry.RestrictTo(skill.Tools);
            }

            return skillName;
        }

        internal static string SummarizeToolResult(ToolResult result)
        {
            if (result?.Data == null)
            {
                return "完成";
            }

            var text = result.Data is string value ? value : JsonConvert.SerializeObject(result.Data);
            return text.Length <= 120 ? text : text.Substring(0, 120) + "…";
        }

        private string BuildContextAnchor()
        {
            var context = GetCurrentContextObject();
            if (context == null)
            {
                return null;
            }

            var parts = new List<string>();
            if (!string.IsNullOrEmpty(context.Workbook?.Name))
            {
                parts.Add("工作簿：" + context.Workbook.Name);
            }

            var sheet = context.Worksheet?.Name ?? context.ActiveWorksheet?.Name;
            if (!string.IsNullOrEmpty(sheet))
            {
                parts.Add("当前工作表：" + sheet);
            }

            if (context.Selection != null && !string.IsNullOrEmpty(context.Selection.Address))
            {
                parts.Add($"当前选区:{context.Selection.Address}（{context.Selection.RowCount}×{context.Selection.ColumnCount}）");
            }

            if (parts.Count == 0)
            {
                return null;
            }

            var text = "当前 Excel 锚点 — " + string.Join("；", parts) +
                       "。后续用户未指明位置的操作默认在此范围内，除非用户另行指定。";

            return text + BuildExternalAnchorSuffix();
        }

        /// <summary>
        /// P5-C2 上下文锚点联动：把已连接的只读外部工作簿句柄追加到系统提示，
        /// 让模型清楚可经 read_external_range 引用的数据来源。
        /// </summary>
        private string BuildExternalAnchorSuffix()
        {
            var lines = new List<string>();
            try
            {
                foreach (var info in externalBooks.List())
                {
                    var sheets = info.Sheets.Count > 0 ? "，工作表：" + string.Join("、", info.Sheets.Take(5)) : "";
                    lines.Add($"{info.Token}（{info.Name}）{sheets}");
                }
            }
            catch { }

            if (lines.Count == 0)
            {
                return "";
            }

            var suffix = "\n已连接只读外部工作簿（可经 read_external_range 读取引用，句柄仅读、不会改动原文件）：";
            suffix += string.Join("；", lines) + "。";
            if (!string.IsNullOrEmpty(contextAnchorHandle) &&
                lines.Any(l => l.StartsWith(contextAnchorHandle + "（", StringComparison.Ordinal)))
            {
                suffix += "当前上下文锚点句柄：" + contextAnchorHandle + "。";
            }

            return suffix;
        }

        private string CreateSystemPrompt(ToolRegistry registry, ModelProfile profile, string skillBody)
        {
            var tools = registry.Tools.Select(tool => new
            {
                name = tool.Name,
                description = tool.Description,
                risk_level = tool.RiskLevel.ToString(),
                parameters = tool.Parameters
            });
            var builder = new StringBuilder();
            var anchor = BuildContextAnchor();
            if (!string.IsNullOrEmpty(anchor))
            {
                builder.AppendLine(anchor);
            }

            var memoryText = memoryStore.BuildInjectionText(ResolveSessionKey());
            if (!string.IsNullOrEmpty(memoryText))
            {
                builder.AppendLine();
                builder.AppendLine("长期记忆（来自此前会话，供参考）：");
                builder.AppendLine(memoryText);
            }

            builder.AppendLine("你是嵌入 Excel 的操作助手，名叫 XAI。");
            builder.AppendLine("【计划模式（P2-3）】当任务预计需要 3 步以上工具调用、或含不可逆写入（删行/覆盖区域/批量回写/结构变更）时：先用一段编号计划（每步=要调用的工具+预期影响范围）回复用户，等待用户明确认可（如\"可以/继续/执行\"）后再开始调用工具；简单任务（≤2 步、可回滚）无需计划直接执行。计划要短——每步一行，不复述用户原话。用户中途改主意时重新出计划。");
            builder.AppendLine("只通过提供的工具操作当前 Excel 工作簿；先理解结构和选区，再决定是否读取更多数据。");
            builder.AppendLine("大区域不得整表读取；写公式使用 A1 引用，绝对/相对引用要正确。");
            builder.AppendLine("任何修改必须说明影响范围；无法确定目标位置时先询问用户。");
            builder.AppendLine("不要编造数据、绕过权限或伪造结果。");
            builder.AppendLine("当用户需求过于模糊、前提或目标不明确，或你无法确定要操作哪个区域/字段时，先向用户提出简短、具体的确认问题，必要时给出 2-3 个候选，等用户答复后再执行；不要臆测意图、不要自行假设数据、更不要声称执行了你并未实际执行的步骤。");
            builder.AppendLine("生成或插入任何图表（create_excel_chart / insert_recommended_chart）后，必须调用 verify_chart 自检，核对：①图表是否真正生成；②系列与数据点是否有效（避免空图、数据源区域选错、数值被当文本导致点数为 0）；③图表类型是否符合需求；④系列数量是否齐全。若 verify_chart 返回的 verdict 为 no_chart / empty / partial 或包含 mismatch，说明图表未达标，必须先修正数据源或重新生成，绝不能向用户声称图表已正确生成。最终回复须附图表自检结论：图表类型、系列数、总数据点、实际数据源区域（data_ranges）、以及是否通过（回显 verify_chart 的 selfcheck）。");
            builder.AppendLine("【图表可读性五件套硬性契约】每个图表生成成功后，最终回复必须做到五点，缺一不可：①【结论先行】回复第一句必须是一句话数据结论（如\"华东销售额最高 320 万，但环比下滑 12%，是唯一负增长地区\"），先给答案再展开；②【图表分析报告】图表类型及选择理由、系列构成（每系列名称与来源列）、数据点数、标题、实际数据源区域、锚点位置，并附一句读图指南（横轴/每根柱或每个扇区代表什么、和谁比、不适合看什么）；③【数据分析报告】对源数据的量化摘要（行列规模、按图表维度的关键统计——分类数/各系列合计或均值/极值）且注明统计口径（单位、含不含税退货等），有目标或同期列时必须给出\"好/差\"的判断基准；④【计算依据】报告里每个关键数字注明来源区域与计算方式（如\"320 万 = $C$10:$C$27 求和，共 18 行\"），用户可回表核对；⑤【图表自检报告】verify_chart（或 auto_check）的 verdict 与依据回显（类型/系列/点数/渲染核验/守恒对账），任一项未通过必须先修复重检再交付。调用 create_excel_chart 时 title 必须带单位与口径（如\"销售额（万元）· 2025 年 · 不含退货\"）；有目标/同期列时优先用 create_combo_chart（柱+折线）呈现参照线。全部数字基于工具返回的真实数据，禁止编造。");
            builder.AppendLine("【分析结论计算依据硬性契约（FR-K3）】analyze_range_statistics / analyze_column_correlation / analyze_value_distribution / analyze_column_profile 的返回值均含 basis 字段（数据源区域、明细行区间、统计口径、公式示例）：回复中引用分析工具结论的每个关键数字必须随句注明 basis 中的来源区域与统计口径（如\"均值 320 = AVERAGE('Sheet1'!$C$3:$C$21)，共 19 行\"），让用户可回表核对；禁止只报数字不报依据，禁止编造 basis 中不存在的口径；相关系数结论必须说明成对完整观测口径与样本数 n。");
            builder.AppendLine("【业务口径规则文件硬性契约（FR-K1）】生成看板/分析结论/统计摘要前，先调用 read_business_glossary 读取工作簿同目录 _口径.md：①命中口径条目时，涉及关键数字的回复必须逐条引用条目原文（术语与定义原样转述，让用户知道用了哪条口径）；②口径与用户当次指令冲突时，用户指令优先，但必须向用户明示差异，不得静默择一；③未找到口径文件时可建议用户创建（工具返回模板），但不得擅自创建或改动文件；④口径文件仅本地读取，禁止把整文粘贴到回复或上传任何外部服务。");
            builder.AppendLine("服务端会在图表生成（create_excel_chart / insert_recommended_chart）、格式设置（format_excel_range / format_excel_conditional）、排序（sort_range）、值写入（write_excel_cell / write_excel_range）、批量回写（write_excel_cells_batch，≤5 格逐格自检）、追加行（append_excel_rows）、清洗（clean_excel_data / dedup_range）、透视汇总（含明细/汇总数量守恒对账）、标记、工作表新增/重命名/删除等主工具成功后自动调用配套 verify_* 工具回读核对，并把带 auto_verify 标记的观察消息附加自检结论。看到 auto_verify 结果的 verdict 不为 ok 时，必须先修正（改数据源/重新写入/调整参数）再继续，绝不能向用户声称操作已成功。");
            builder.AppendLine("自动自检覆盖不到的场景需要你主动补一次自检并把 selfcheck 回显给用户，自检参数从操作时使用的实参推导（把目标值作为 expected_* 传入）：①写入公式（value_mode=formula、fill_formula 或含公式的区间写入）后调用 verify_formula，核对公式内容与计算值；②筛选（filter_excel_data）后调用 verify_sort_filter，地址用输出结果区域，核对输出行都满足筛选条件；③透视汇总需要核对具体行标签与汇总值时，把期望行作为 expected_rows 传入 verify_pivot；④标记需要核对期望被标记地址时，把 expected_marked / expected_count 传入 verify_marker；⑤批量回写超过 5 格时逐格核对，或需要更细断言时补调。最终回复须附各操作的自检结论（回显对应工具的 selfcheck）。");

            switch (CurrentWriteMode())
            {
                case "formula":
                    builder.AppendLine("写入计算或汇总结果时，默认优先使用公式（如 =B2*C2 引用源区域），保证可联动、可追溯；除非用户明确要求或内容是非公式文本。");
                    break;
                case "value":
                    builder.AppendLine("写入数据时直接写入计算后的数值/文本即可，不要写成公式；除非用户明确要求写入公式。");
                    break;
            }

            if (skillBody != null)
            {
                builder.AppendLine();
                builder.AppendLine("已加载技能并必须遵守其步骤：");
                builder.AppendLine(skillBody);
            }
            else
            {
                // AI Skill P1（FR-S4）：技能清单渐进披露——只注入名称+一句话描述+来源，
                // 技能全文仅在用户确认加载（/技能名）后经 skillBody 进入上下文。
                var skillListText = SkillListInjector.BuildInjectionText(LoadSkills());
                if (!string.IsNullOrEmpty(skillListText))
                {
                    builder.AppendLine();
                    builder.AppendLine("已安装技能清单（当用户需求与某技能描述匹配时，建议用户输入 /技能名 加载；未加载的技能只是参考信息，不是当前指令）：");
                    builder.AppendLine(skillListText);
                }
            }

            if (profile.ProtocolMode == AgentProtocolMode.Fallback)
            {
                builder.AppendLine();
                builder.AppendLine("必须只输出一个 JSON 对象，不要输出 Markdown。");
                builder.AppendLine("回答用户时输出：{\"action\":\"final\",\"content\":\"...\"}。");
                builder.AppendLine("需要操作时输出：{\"action\":\"tool\",\"tool\":\"名称\",\"args\":{},\"reason\":\"...\",\"call_id\":\"唯一ID\"}。");
                builder.AppendLine("每次只调用一个工具；写入前优先用只读工具确认结构。");
            }
            else
            {
                builder.AppendLine();
                builder.AppendLine("使用提供的函数调用（tools）来执行 Excel 操作；读操作可直接调用，写/结构操作等用户确认。");
                builder.AppendLine("无需调用工具时再直接给出最终文字回答。");
                if (registry.TryGet("run_vba_snippet", out _))
                {
                    builder.AppendLine("当现有工具都完成不了用户需求时，可以编写 VBA 代码片段，用 run_vba_snippet 工具执行（参数 snippet 为完整代码，执行前会展示给用户确认）以达到用户想要的效果——先理解任务，再写逻辑自包含的代码，直接操作 ActiveWorkbook/ActiveSheet/Selection。");
                    builder.AppendLine("写 VBA 必须追求高性能，否则大区域会明显卡顿：大批量读写一律用数组或 Range 整体赋值，先一次性读入内存（如 data = 区域.Value）再循环处理最后整体写回（区域.Value = data），严禁在 For Each 循环里逐单元格读写；过程开头设 Application.ScreenUpdating = False、Application.Calculation = xlCalculationManual，结束前恢复为 True 和 xlCalculationAutomatic，用 Try/Final 保证恢复；不要在循环内反复访问 Range/Worksheet/ActiveSheet 等 COM 属性，把会重复使用的对象先用 Set 存到局部变量。");
                    builder.AppendLine("居中、对齐、字体、颜色、边框等仅改外观样式的操作，VBA 只允许修改格式属性（HorizontalAlignment/VerticalAlignment、Font、Interior、Borders、NumberFormat、RowHeight/ColumnWidth、IndentLevel 等）；严禁任何清空或覆盖单元格数据的行为：不调用 Clear、ClearContents、ClearFormats、Delete、Rows(...).Delete、EntireRow.Delete，不用等价代码（如把区域整体赋值空值/重排数据），不要为“重建”“规范化”而清空原数据——保持选区内单元格的数值与内容完全不变，只改外观。");
                    builder.AppendLine("在你成功用 run_vba_snippet 执行了自写代码后，必须主动调用 save_vba_tool 把该代码沉淀为可复用的 VBA 模板（给出一句话凝练的 description：概括这段代码做什么、让人一眼看懂，如“表头加粗+底色”“修复手机号与身份证格式”，≤20 字；name 用简短英文过程名，如 fix_phone_format；严禁写成“由本次成功执行的代码生成，重复或批量场景可直接调用”这类模板套话），以便今后在「VBA」面板模板库中复用；不要因“逻辑不够通用”“只此一次”“影响计算”等自行判断而跳过沉淀——save_vba_tool 会先征求用户确认，是否保存由用户决定，你只需在任何一次成功后都主动发起沉淀即可。");
                }
                else
                {
                    // M4 宿主感知过滤联动（WPS 兼容版）：run_vba_snippet 未注册即 WPS 无 VBA 环境，
                    // 必须让 AI 知情并转告用户原因，严禁提示 AI 调用不存在的工具。
                    // 口径对齐 ExcelToolHost.CreateRegistry：save_vba_tool 无条件注册（仅写库文件），
                    // 真正被过滤的是 run_vba_snippet / validate_vba_snippet 与 ct_* 沉淀模板工具。
                    builder.AppendLine("注意：当前宿主为 WPS 表格且未安装 VBA 环境，run_vba_snippet / validate_vba_snippet 及 VBA 模板工具（ct_* 沉淀工具）均不可用（不在工具清单中）；save_vba_tool 虽在清单中（仅把 VBA 代码写入模板库文件），但 WPS 无 VBA 环境下模板无法执行，不要建议用户在 WPS 中运行 VBA 模板。用户提出需要 VBA、宏或沉淀模板的需求时，如实说明原因并建议：①在 WPS 中安装 VBA 环境（专业版/安装 VBA 支持库）后重试，或②改用 Microsoft Excel 使用本助手；绝对不要尝试调用清单中不存在的工具、不要伪装已执行、不要用其它工具变通绕过 VBA 能力边界。");
                }
            }

            builder.AppendLine("可用工具：");
            builder.AppendLine(JsonConvert.SerializeObject(tools, Formatting.None));
            return builder.ToString();
        }

        private string CreateUserPrompt(string prompt)
        {
            var context = GetCurrentContextObject();
            var payload = new Dictionary<string, object>
            {
                ["request"] = prompt,
                ["excel_context"] = context
            };
            if (references.Count > 0)
            {
                payload["manual_references"] = references.Select(ToReferenceObject).ToArray();
            }

            return JsonConvert.SerializeObject(payload);
        }

        private static readonly string ActiveProfileFile =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAiAssistant", "active-profile.txt");

        private void PersistActiveProfileId()
        {
            try
            {
                var directory = Path.GetDirectoryName(ActiveProfileFile);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(ActiveProfileFile, activeProfileId ?? string.Empty);
            }
            catch (Exception xaiEx0) {
                AppLog.Warn("infra", "SidebarBridge.Infra.cs:PersistActiveProfileId 静默兜底触发: " + xaiEx0.Message);
                // 持久化失败不阻断保存；下次切换会重新写入。
            }
        }

        private string ReadPersistedActiveProfileId()
        {
            try
            {
                return File.Exists(ActiveProfileFile) ? File.ReadAllText(ActiveProfileFile).Trim() : string.Empty;
            }
            catch (Exception xaiEx1) {
                AppLog.Warn("infra", "SidebarBridge.Infra.cs:ReadPersistedActiveProfileId 静默兜底触发: " + xaiEx1.Message);
                return string.Empty;
            }
        }

        private ModelProfile ResolveActiveProfile(IReadOnlyList<ModelProfile> profiles)
        {
            if (profiles.Count == 0)
            {
                return null;
            }

            if (string.IsNullOrEmpty(activeProfileId))
            {
                activeProfileId = ReadPersistedActiveProfileId();
            }

            var active = profiles.FirstOrDefault(item =>
                string.Equals(item.Id, activeProfileId, StringComparison.OrdinalIgnoreCase));
            if (active == null)
            {
                activeProfileId = profiles[0].Id;
                PersistActiveProfileId();
                active = profiles[0];
            }

            return active;
        }

        private ModelProfile RequireProfile()
        {
            var profile = ResolveActiveProfile(configurationStore.Load());
            if (profile == null)
            {
                throw new InvalidOperationException("请先在设置中添加并保存模型方案。");
            }

            return profile;
        }

        private object ToProfileObject(ModelProfile profile)
        {
            return new
            {
                id = profile.Id,
                name = profile.Name,
                profile.BaseUrl,
                profile.ModelName,
                protocol = profile.ProtocolMode.ToString(),
                authScheme = profile.AuthScheme.ToString(),
                temperature = profile.Temperature,
                api_key_set = !string.IsNullOrEmpty(profile.ApiKey)
            };
        }

        private object ToConfigObject(ModelProfile profile)
        {
            return new
            {
                profile.BaseUrl,
                profile.ModelName,
                profile.Temperature,
                protocol = profile.ProtocolMode.ToString(),
                ApiKeySet = !string.IsNullOrEmpty(profile.ApiKey)
            };
        }

        private static object ToReferenceObject(ReferenceRegion reference)
        {
            return new
            {
                reference.Id,
                reference.Address,
                reference.Sheet,
                reference.Range,
                reference.RowCount,
                reference.ColumnCount,
                estimated_tokens = reference.RowCount * reference.ColumnCount
            };
        }

        private string ResolveSessionKey()
        {
            return CreateWorkbookHash();
        }

        private int RestoreSession()
        {
            var key = ResolveSessionKey();
            activeSessionId = ResolveActiveSession(key);
            conversation.Clear();
            activeSummary = sessionStore.LoadSummary(key, activeSessionId);
            var restored = sessionStore.Load(key, activeSessionId);
            conversation.AddRange(restored);
            return restored.Count / 2;
        }

        private string ResolveActiveSession(string key)
        {
            var sessions = ListSessionInfo(key);
            if (sessions.Count == 0)
            {
                return sessionStore.Create(key);
            }

            if (string.IsNullOrWhiteSpace(activeSessionId) ||
                !sessions.Any(item => string.Equals(item.Id, activeSessionId, StringComparison.OrdinalIgnoreCase)))
            {
                return sessions[0].Id;
            }

            return activeSessionId;
        }

        private void SwitchToSession(string targetSessionId)
        {
            var key = ResolveSessionKey();
            activeSessionId = targetSessionId;
            conversation.Clear();
            activeSummary = sessionStore.LoadSummary(key, targetSessionId);
            conversation.AddRange(sessionStore.Load(key, targetSessionId));
        }

        private void PersistSession()
        {
            PersistSession(out _);
        }

        private void PersistSession(out int turnCount)
        {
            if (conversation.Count == 0)
            {
                turnCount = 0;
                return;
            }

            if (string.IsNullOrWhiteSpace(activeSessionId))
            {
                activeSessionId = sessionStore.Create(ResolveSessionKey());
            }

            var key = ResolveSessionKey();
            sessionStore.Save(key, activeSessionId, conversation);
            turnCount = conversation.Count / 2;
        }

        /// <summary>把会话消息裁剪到最近 24 条；调用方在主线程或工具回调中均可安全使用。</summary>
        private void TrimConversation()
        {
            if (conversation.Count > 24)
            {
                conversation.RemoveRange(0, conversation.Count - 24);
            }
        }

        /// <summary>把一笔工具调用的结果数据压缩成用于日志的短摘要（JSON，超长截断）。</summary>
        private static string RenderedToolSummary(object toolEvent)
        {
            try
            {
                var record = (dynamic)toolEvent;
                var data = record.Result?.Data;
                if (data == null)
                {
                    return string.Empty;
                }

                var text = Newtonsoft.Json.JsonConvert.SerializeObject(data);
                return text.Length <= 220 ? text : text.Substring(0, 220) + "…";
            }
            catch (Exception xaiEx2) {
                AppLog.Warn("infra", "SidebarBridge.Infra.cs:RenderedToolSummary 静默兜底触发: " + xaiEx2.Message);
                return string.Empty;
            }
        }

        /// <summary>把一段 VBA 代码压成用于日志的简短摘要（首个子过程名 + 行数），便于在设置页核对。</summary>
        private static string SummarizeSnippet(string code)
        {
            try
            {
                var lines = (code ?? string.Empty)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Trim())
                    .ToArray();
                var sub = lines.FirstOrDefault(l => l.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase))
                          ?? lines.FirstOrDefault(l => l.StartsWith("Function ", StringComparison.OrdinalIgnoreCase));
                var subName = string.IsNullOrWhiteSpace(sub)
                    ? ""
                    : sub.Substring(sub.IndexOf(' ') + 1).TrimEnd('(', ' ');
                return subName.Length > 32
                    ? subName.Substring(0, 32) + "…"
                    : (subName ?? "") + "【" + lines.Length + " 行】";
            }
            catch (Exception xaiEx3) {
                AppLog.Warn("infra", "SidebarBridge.Infra.cs:SummarizeSnippet 静默兜底触发: " + xaiEx3.Message);
                return "VBA 片段";
            }
        }

        private static string TruncateCode(string code, int limit = 6000)
        {
            if (string.IsNullOrWhiteSpace(code) || code.Length <= limit)
            {
                return code ?? string.Empty;
            }

            return code.Substring(0, limit) + "\n…[已截断，共 " + code.Length + " 字符]";
        }

        /// <summary>把多行错误正文压成单行供日志使用：折叠换行 + 截断，避免把整段异常堆栈塞进一行日志。</summary>
        private static string CompactLogText(string text, int limit)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var oneLine = string.Join(" ", text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
            return oneLine.Length <= limit ? oneLine : oneLine.Substring(0, limit) + "…";
        }

        /// <summary>
        /// 首次对话（标题仍为默认）时用模型生成可读中文标题；模型不可用或失败时静默回退到截断标题。
        /// 后台执行，不阻塞回复。
        /// </summary>
        private void TryAutoTitleSession(string firstPrompt)
        {
            if (string.IsNullOrWhiteSpace(activeSessionId))
            {
                return;
            }

            var key = ResolveSessionKey();
            var currentTitle = sessionStore.Title(key, activeSessionId);
            if (!string.IsNullOrWhiteSpace(currentTitle) && currentTitle != "新会话")
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(firstPrompt))
            {
                return;
            }

            try
            {
                var profile = RequireProfile();
                Task.Run(async () =>
                {
                    try
                    {
                        using var client = new OpenAiCompatibleClient(profile);
                        using var timeoutSource = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                        var input = new List<ChatMessage>
                        {
                            ChatMessage.System(
                                "你是 Excel JYYJ助手的会话标题生成器。根据用户第一条指令，生成一个不超过 12 个字的" +
                                "中文会话标题，概括其意图（如「销售数据汇总」「库存周转分析」）。只输出标题本身，不要引号、不要解释。"),
                            ChatMessage.User(firstPrompt)
                        };
                        var result = await client.CompleteAsync(input, timeoutSource.Token).ConfigureAwait(false);
                        var title = (result?.Content ?? string.Empty).Trim().Trim('"', '“', '”', '「', '」');
                        if (title.Length > 24)
                        {
                            title = title.Substring(0, 24);
                        }

                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            sessionStore.Rename(key, activeSessionId, title);
                        }
                    }
                    catch (Exception xaiEx4) {
                        AppLog.Warn("infra", "SidebarBridge.Infra.cs:TryAutoTitleSession 静默兜底触发: " + xaiEx4.Message);
                        // 模型不可用/超时：保留截断标题，静默回退
                    }
                });
            }
            catch (Exception xaiEx5) {
                AppLog.Warn("infra", "SidebarBridge.Infra.cs:TryAutoTitleSession 静默兜底触发: " + xaiEx5.Message);
                // 无可用模型配置：静默回退
            }
        }

        private object GetActiveSessionObject()
        {
            return new
            {
                id = activeSessionId,
                restored = conversation.Count > 0,
                turn_count = conversation.Count / 2,
                has_summary = !string.IsNullOrWhiteSpace(activeSummary),
                summary = activeSummary ?? string.Empty,
                last_updated = string.IsNullOrWhiteSpace(activeSessionId)
                    ? (DateTime?)null
                    : sessionStore.LastUpdated(ResolveSessionKey(), activeSessionId),
                messages = conversation.Select(message => (object)new
                {
                    role = message.Role,
                    content = message.Content ?? string.Empty,
                    tool_events = ProjectToolEvents(message.ToolEvents),
                    reasoning = message.Reasoning ?? string.Empty,
                    steps = message.Steps,
                    turn_label = message.TurnLabel ?? string.Empty,
                    auto_sinked_vba = message.AutoSinkedVba
                }).ToArray()
            };
        }

        private IReadOnlyList<SessionStore.SessionInfo> ListSessionInfo()
        {
            return ListSessionInfo(ResolveSessionKey());
        }

        private IReadOnlyList<SessionStore.SessionInfo> ListSessionInfo(string key)
        {
            return sessionStore.List(key);
        }

        private object[] ListSessionObjects()
        {
            return ListSessionObjects(false);
        }

        private object[] ListSessionObjects(bool includeArchived)
        {
            return ListSessionInfo()
                .Where(info => includeArchived || !info.Archived)
                .Select(info => (object)new
                {
                    info.Id,
                    info.Title,
                    info.TurnCount,
                    UpdatedAt = info.UpdatedAt,
                    info.Archived,
                    Active = string.Equals(activeSessionId, info.Id, StringComparison.OrdinalIgnoreCase)
                })
                .ToArray();
        }

        private IReadOnlyList<Skill> LoadSkills()
        {
            if (!skillCatalogLoaded)
            {
                skillCatalog = new SkillCatalog();
                var userSkills = Path.Combine(RootFolder, "skills");
                // 内置技能随加载项 DLL 分发（安装目录\skills）。不能用 AppDomain.CurrentDomain.BaseDirectory：
                // COM 加载项运行在 Excel 宿主进程的默认 AppDomain，该值指向 Excel.exe 目录而非加载项目录，
                // 会导致内置技能永远加载不到（设置页「本地技能」空空）。
                var builtInSkills = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty,
                    "skills");
                skillCatalog.AddSearchDirectory(userSkills, false);
                skillCatalog.AddSearchDirectory(builtInSkills, true);
                DiagnosticsLog.Write($"LoadSkills reload-start user={userSkills} builtIn={builtInSkills}");
                skillCatalog.Reload();
                skillCatalogLoaded = true;
                DiagnosticsLog.Write(
                    $"LoadSkills reload-done count={skillCatalog.All.Count}");
            }

            return skillCatalog.All;
        }

        private static (string, string) SplitSheetAndRange(string address)
        {
            var trimmed = address.Trim();
            var marker = trimmed.IndexOf('!');
            if (marker > 0)
            {
                var sheet = trimmed.Substring(0, marker).Trim('\'', '"');
                return (sheet, trimmed.Substring(marker + 1));
            }

            return (null, trimmed);
        }

        private static (int, int) EstimateDimensions(string range)
        {
            var match = Regex.Match(range ?? string.Empty, @"^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$");
            if (!match.Success)
            {
                return (1, 1);
            }

            var startColumn = ColumnIndex(match.Groups[1].Value);
            var startRow = int.TryParse(match.Groups[2].Value, out var first) ? first : 1;
            var endColumn = match.Groups[3].Success ? ColumnIndex(match.Groups[3].Value) : startColumn;
            var endRow = match.Groups[4].Success && int.TryParse(match.Groups[4].Value, out var last) ? last : startRow;
            return (Math.Max(1, endRow - startRow + 1), Math.Max(1, endColumn - startColumn + 1));
        }

        private static int ColumnIndex(string column)
        {
            var value = 0;
            foreach (var character in column.ToUpperInvariant())
            {
                value = value * 26 + (character - 'A' + 1);
            }

            return value;
        }

        private ExcelToolHost CreateHost(string turnLabel = null)
        {
            dynamic workbook = excelApplication.ActiveWorkbook;
            try
            {
                var workbookPath = Convert.ToString(workbook?.FullName) ?? string.Empty;
                var host = new ExcelToolHost(
                    excelApplication,
                    operationRepository,
                    ExcelToolHost.CreateWorkbookHash(workbookPath),
                    sessionId,
                    turnLabel);
                host.WriteModePreference = CurrentWriteMode();
                host.RunWorkflowTool = workflowManager.RunWorkflowTool;
                host.ExternalBooks = externalBooks;
                return host;
            }
            finally
            {
                if (workbook != null && Marshal.IsComObject(workbook))
                {
                    Marshal.FinalReleaseComObject(workbook);
                }
            }
        }

        private ToolRegistry CreateRegistry(string turnLabel = null)
        {
            // M4 宿主感知过滤：过滤触发时（仅 WPS 无 VBA 环境）在诊断日志留痕；Excel 宿主信息为空、零开销。
            var host = CreateHost(turnLabel);
            var registry = host.CreateRegistry();
            var filterInfo = host.HostFilterInfo;
            if (!string.IsNullOrEmpty(filterInfo))
            {
                DiagnosticsLog.Write("Tool filter applied: " + filterInfo);
            }

            return registry;
        }

        private string CreateWorkbookHash()
        {
            dynamic workbook = null;
            try
            {
                workbook = excelApplication.ActiveWorkbook;
                return ExcelToolHost.CreateWorkbookHash(Convert.ToString(workbook?.FullName) ?? string.Empty);
            }
            finally
            {
                if (workbook != null && Marshal.IsComObject(workbook))
                {
                    Marshal.FinalReleaseComObject(workbook);
                }
            }
        }

        private ExcelCurrentContext GetCurrentContextObject()
        {
            var service = new ExcelContextService(excelApplication);
            return service.GetCurrentContext();
        }

        /// <summary>Excel 启动初期可能没有打开的工作簿，GetCurrentContext 会抛异常。
        /// 初始化时若因此整体失败，技能/工具/会话等静态数据也一并加载不到（设置页「本地技能」空空）。
        /// 这里把上下文降级为 null，前端 renderContext 已兼容，其余数据照常返回。</summary>
        private ExcelCurrentContext SafeGetCurrentContextObject()
        {
            try
            {
                return GetCurrentContextObject();
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write("Initialize context unavailable: " + exception.Message);
                return null;
            }
        }

        private async Task<string> ExecuteAsync(Func<Task<string>> action)
        {
            try
            {
                Publish("busy", true);
                return await action();
            }
            catch (Exception exception) when (!(exception is OutOfMemoryException))
            {
                return Error(exception);
            }
            finally
            {
                Publish("busy", false);
            }
        }

        private static string Error(Exception exception)
        {
            return JsonConvert.SerializeObject(new { ok = false, error = exception.Message });
        }

        // AgentProgress 是强类型，Newtonsoft 默认序列化为 PascalCase；而前端 renderLiveProgress
        // 约定读 snake_case（kind/reasoning/tool_index）。不转换会导致 p.kind 恒为 undefined，
        // 流式思考永远渲染不出来（只显示空卡）。这里显式铺成前端约定的字段名。
        internal static object ProgressPayload(AgentProgress p)
        {
            // key 以来源命名空间 + 序号组成前端配对键：外围 agent 工具用 agent索引，
            // run_workflow 内层步骤用 workflow 命名空间，避免两侧序号互相覆盖导致步骤卡片消失；
            // 步骤自检用独立命名空间 + "v" 后缀，避免与同序号主步骤卡片互相覆盖。
            var ns = p == null ? null : p.Namespace;
            var idx = p == null ? -1 : p.ToolIndex;
            var key = p == null
                ? "f"
                : (string.Equals(ns, "workflow-verify", StringComparison.Ordinal)
                    ? "wf" + idx + "v"
                    : (string.Equals(ns, "workflow", StringComparison.Ordinal)
                        ? "wf" + idx
                        : (idx >= 0 ? "i" + idx : "f" + (p.Tool ?? ""))));
            return new
            {
                kind = p == null ? -1 : (int)p.Kind,
                tool = p == null ? null : p.Tool,
                ok = p != null && p.Ok,
                summary = p == null ? null : p.Summary,
                detail = p == null ? null : p.Detail,
                reasoning = p == null ? null : p.Reasoning,
                tool_index = p == null ? -1 : p.ToolIndex,
                tool_count = p == null ? 0 : p.ToolCount,
                ns = ns,
                key = key
            };
        }

        // ToolEventView 是强类型，序列化为 PascalCase（Tool/Ok/Summary...）；前端约定读
        // snake_case（tool/ok/summary...）。不投影会导致 autoReplyBody/readableStep 拿到 undefined，
        // 工具型回复的正文摘要空白 →「AI 回复空白」。统一投影为前端字段名。
        private static object[] ProjectToolEvents(IEnumerable<ToolEventView> evs)
        {
            if (evs == null) return new object[0];
            return evs
                .Where(e => e != null)
                .Select(e => (object)new
                {
                    tool = e.Tool,
                    ok = e.Ok,
                    confirmed = e.Confirmed,
                    error_code = e.ErrorCode,
                    error_message = e.ErrorMessage,
                    summary = e.Summary,
                    detail = e.Detail
                })
                .ToArray();
        }

        private void Publish(string type, object payload)
        {
            var core = webView?.CoreWebView2;
            if (core == null)
            {
                return;
            }

            var json = JsonConvert.SerializeObject(new { type, payload });

            // .NET 下 CoreWebView2.PostWebMessageAsJson 必须在 UI 线程调用；推理流式增量
            // 来自 SSE 读流后台线程，若直接发会被静默丢弃 → 前端只能等整轮结束才一次性回放。
            // BeginInvoke 异步投递，不阻塞、不占 Excel 消息泵，事件先后顺序在同一队列内保持。
            if (webView.IsHandleCreated && webView.InvokeRequired)
            {
                try
                {
                    webView.BeginInvoke((Action)(() => { try { core.PostWebMessageAsJson(json); } catch { } }));
                }
                catch
                {
                }

                return;
            }

            try { core.PostWebMessageAsJson(json); }
            catch
            {
            }
        }

        private object[] GetToolCatalogObjects()
        {
            var registry = CreateRegistry();
            return registry.Tools
                .OrderBy(tool => (int)tool.RiskLevel)
                .ThenBy(tool => tool.Name, StringComparer.OrdinalIgnoreCase)
                .Select(tool => (object)new
                {
                    name = tool.Name,
                    description = tool.Description,
                    risk_level = tool.RiskLevel.ToString(),
                    risk_rank = (int)tool.RiskLevel
                })
                .ToArray();
        }

        private string CurrentTurnLabel()
        {
            return "第 " + (conversation.Count / 2 + 1) + " 轮";
        }

        private string CurrentWriteMode()
        {
            if (cachedWriteMode == null)
            {
                cachedWriteMode = new WritePreferenceStore().Load().ValueMode
                    .Trim()
                    .ToLowerInvariant();
            }

            return cachedWriteMode;
        }

        private int CurrentMaxToolCalls()
        {
            if (cachedMaxToolCalls < 1 || cachedMaxToolCalls > 10000)
            {
                var stored = new AgentSettingsStore().Load().MaxToolCallsPerTurn;
                // v3.7.1.39：上限由 1000 放宽到 10000；旧存档停在上限 1000 视为未自定义，迁移到新上限。
                cachedMaxToolCalls = stored == 1000 ? 10000 : stored;
            }

            return cachedMaxToolCalls;
        }

        private object[] GetHistoryRecords()
        {
            return operationRepository
                .List(new OperationQuery
                {
                    WorkbookHash = CreateWorkbookHash(),
                    Limit = 50
                })
                .Select(record => new
                {
                    id = record.Id,
                    tool = record.Tool,
                    turn_label = record.TurnLabel ?? string.Empty,
                    created_at_utc = record.CreatedAtUtc,
                    affected = DeserializeAffected(record.AffectedJson),
                    can_undo = record.CanUndo,
                    undo_kind = record.UndoKind
                })
                .Cast<object>()
                .ToArray();
        }

        private static object DeserializeAffected(string affectedJson)
        {
            try
            {
                return JsonConvert.DeserializeObject(affectedJson ?? "{}");
            }
            catch (Exception exception) when (
                exception is JsonException ||
                exception is ArgumentException)
            {
                return new { address = "未知目标" };
            }
        }

        private sealed class ReferenceRegion
        {
            public string Id { get; set; }

            public string Address { get; set; }

            public string Sheet { get; set; }

            public string Range { get; set; }

            public int RowCount { get; set; }

            public int ColumnCount { get; set; }
        }

        public Task<string> SetContextAnchor(string handle)
        {
            return ExecuteAsync(() =>
            {
                if (string.IsNullOrEmpty(handle))
                {
                    contextAnchorHandle = null;
                    return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, anchored = false }));
                }

                if (!externalBooks.TryGet(handle, out var workbook) || workbook == null)
                {
                    return Task.FromResult(JsonConvert.SerializeObject(new
                    {
                        ok = false,
                        error = "句柄不存在或已关闭：" + handle
                    }));
                }

                contextAnchorHandle = handle;
                return Task.FromResult(JsonConvert.SerializeObject(new { ok = true, anchored = true, handle = handle }));
            });
        }

        /// <summary>关闭指定外部只读句柄；若正作为上下文锚点则一并清除。</summary>
        public Task<string> CloseExternalBook(string handle)
        {
            return ExecuteAsync(() =>
            {
                var closed = externalBooks.Close(handle ?? string.Empty);
                if (closed && string.Equals(contextAnchorHandle, handle, StringComparison.OrdinalIgnoreCase))
                {
                    contextAnchorHandle = null;
                }

                return Task.FromResult(JsonConvert.SerializeObject(new { ok = closed, handle = handle, closed = closed }));
            });
        }

    }
}
