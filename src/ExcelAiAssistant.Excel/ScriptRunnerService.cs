using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Security;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 受控执行策略：默认关闭；白名单 + 黑名单 + 超时 + 输出截断。任何脚本/进程类工具都先经过此策略门禁。
    /// </summary>
    public sealed class ScriptOptions
    {
        public bool Enabled { get; set; }

        public string[] ExeAllowlist { get; set; } = new string[0];

        public string[] ExeDenylist { get; set; } = new string[0];

        public string[] CliAllowlist { get; set; } = new string[0];

        public string[] CliDenylist { get; set; } = new string[0];

        public string[] ScriptAllowlist { get; set; } = new string[0];

        public string[] ScriptDenylist { get; set; } = new string[0];

        public int TimeoutSeconds { get; set; } = 30;

        public int OutputLimitKB { get; set; } = 64;

        public string PythonPath { get; set; } = "python";

        public void Apply(ControlledExecutionConfig config)
        {
            if (config == null)
            {
                return;
            }

            Enabled = config.Enabled;
            ExeAllowlist = config.ExeAllowlist ?? new string[0];
            ExeDenylist = config.ExeDenylist ?? new string[0];
            CliAllowlist = config.CliAllowlist ?? new string[0];
            CliDenylist = config.CliDenylist ?? new string[0];
            ScriptAllowlist = config.ScriptAllowlist ?? new string[0];
            ScriptDenylist = config.ScriptDenylist ?? new string[0];
            TimeoutSeconds = config.TimeoutSeconds > 0 ? config.TimeoutSeconds : 30;
            OutputLimitKB = config.OutputLimitKb > 0 ? config.OutputLimitKb : 64;
            PythonPath = string.IsNullOrWhiteSpace(config.PythonPath) ? "python" : config.PythonPath;
        }

        public ControlledExecutionConfig ToConfig()
        {
            return new ControlledExecutionConfig
            {
                Enabled = Enabled,
                ExeAllowlist = ExeAllowlist,
                ExeDenylist = ExeDenylist,
                CliAllowlist = CliAllowlist,
                CliDenylist = CliDenylist,
                ScriptAllowlist = ScriptAllowlist,
                ScriptDenylist = ScriptDenylist,
                TimeoutSeconds = TimeoutSeconds,
                OutputLimitKb = OutputLimitKB,
                PythonPath = PythonPath
            };
        }
    }

    public sealed class ScriptRunnerService
    {
        private readonly dynamic application;
        private readonly ScriptOptions options;

        /// <summary>run_vba_snippet 复用的常驻草稿模块名。保留一个即可复用，避免每轮 Add/Remove 触发 VBE 工程重编译。VBA 模块名不能以下划线开头，故不带下划线。</summary>
        private const string ScratchModuleName = "JYYJHelper";

        /// <summary>VBA 持久状态独立模块名（P3-D2）：创建后不再改写，故模块级变量 gXaiState 的值跨调用保留
        /// （草稿模块每轮整清重写，若把状态变量放其中，改写会丢值，故独立成模块）。</summary>
        private const string StateModuleName = "JYYJXaiState";

        /// <summary>P3-D2 每轮独立运行模块名前缀：创建（新建 JYYJRun_&lt;guid&gt;）与清理（CleanupRunModules）必须共用，
        /// 防止两处字面量漂移导致「新建前缀改、清理前缀没改」→ 运行模块泄漏进工作簿。状态模块名不得以此开头（见 IsRunModuleName 测试）。</summary>
        internal const string RunModulePrefix = "JYYJRun_";

        /// <summary>判定模块名是否属于「运行/遗留模块」应被清理：旧版本常驻草稿 JYYJHelper 或 P3-D2 每轮运行模块 JYYJRun_*。
        /// 命中即整模块 components.Remove（不触发整工程重编译，状态保留）。</summary>
        internal static bool IsRunModuleName(string name)
        {
            return string.Equals(name, ScratchModuleName, StringComparison.Ordinal)
                || (name != null && name.StartsWith(RunModulePrefix, StringComparison.Ordinal));
        }

        /// <summary>Python/VBA 输出截断标记，须与 PythonWorkerHost 内嵌 worker 源码中的截断后缀保持一致
        /// （ASCII 三点，非单字符省略号 U+2026），run_python_code 返回的 truncated 字段依赖此串判定。</summary>
        internal const string OutputTruncatedMarker = "...（已截断）";

        public ScriptRunnerService(dynamic application, ScriptOptions options)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.options = options ?? new ScriptOptions();
        }

        public ToolResult ValidateVbaSnippet(Dictionary<string, object> args)
        {
            var snippet = Required(args, "snippet");
            if (options.Enabled == false)
            {
                return Disabled("validate_vba_snippet");
            }

            dynamic vbProject = null;
            try
            {
                try
                {
                    vbProject = application.VBE?.ActiveVBProject;
                }
                catch
                {
                    vbProject = null;
                }

                if (vbProject == null)
                {
                    // 无法访问 VBA 工程对象模型时降级为静态校验。
                    return StaticVbaValidate(snippet);
                }

                dynamic components = null;
                dynamic module = null;
                dynamic codeModule = null;
                try
                {
                    components = vbProject.VBComponents;
                    module = components.Add(1); // vbext_ct_StdModule
                    codeModule = module.CodeModule;
                    var line = codeModule.CountOfLines + 1;
                    codeModule.InsertLines(line, snippet);
                    var ok = codeModule.CountOfLines >= line;
                    // 诚实标注：这里只验证「片段可写入模块」，并不触发 VBE 编译（自动编译会弹对话框、且
                    // 模块未运行时 VBE 不校验语法），故不能误导模型把 valid 当成「已通过编译」。
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            validation_mode = "write_only",
                            writes = ok,
                            lines = codeModule.CountOfLines,
                            note = "仅验证片段可写入模块；未执行语法编译，请以 run_vba_snippet 的静态护栏与探针运行结果为准。"
                        }
                    };
                }
                catch (Exception ex)
                {
                    AppLog.Warn("ScriptRunnerService", $"[LogCover] ScriptRunnerService.cs:160  {ex.GetType().Name}: {ex.Message}");
                    return new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "VBA_SYNTAX_ERROR",
                        Data = "VBA 代码无法写入/校验：" + ex.Message
                    };
                }
                finally
                {
                    RemoveComponent(vbProject, module);
                    // M16 全量审查：components/module/codeModule RCW 释放（原实现从不释放，每次校验泄漏 3 个）。
                    Release((object)codeModule);
                    Release((object)module);
                    Release((object)components);
                }
            }
            finally
            {
                Release((object)vbProject);
            }
        }

        public ToolResult RunVbaSnippet(Dictionary<string, object> args)
        {
            var snippet = Required(args, "snippet");
            var macroName = Optional(args, "macro_name");
            if (options.Enabled == false)
            {
                return Disabled("run_vba_snippet");
            }

            // P1-D 入参通道：args（JSON 对象）→ 注入模块级 gXaiArgs，VBA 内用 XaiArg("key") 读取。
            var argsRaw = args.TryGetValue("args", out var a) ? a : null;
            var argsJson = NormalizeArgsJson(argsRaw);
            if (argsJson.Length > VbaArgsLimitChars)
            {
                return Invalid("INPUT_TOO_LARGE", "args 数据量过大（超过 " + (VbaArgsLimitChars / 1024) + " KB 上限），请缩小参数范围或改用 run_python_code 的 input_data 传大数据。");
            }

            // P1-A/E：结构化回传体积上限（result_limit_kb，默认 1 MB），换算为 VBA 侧序列化字符上限。
            var resultLimitChars = Math.Min(Math.Max(GetInt(args, "result_limit_kb", 1024), 1), 16 * 1024) * 1024; // 上限 16MB（审查第五批护栏）
            var helperBlock = BuildVbaHelperBlock(argsJson, resultLimitChars);

            // 执行前硬闸门：确定性致命问题（未闭合引号、End 不配对、无过程入口、对多行多列矩形
            // Range 调 .AutoFit、Range.Alignment、格式化时 Clear/Delete 数据）直接拦截，避免把
            // 必失败的代码注入 Excel 白跑一次再撞运行时错误（如 AutoFit 1004）。仅拦截 error 级。
            var gateMsg = PreRunVbaGate(snippet);
            if (gateMsg != null)
            {
                return Invalid("VBA_INVALID", gateMsg);
            }

            // P3-D2 持久状态与「每轮新建运行模块」架构：
            // 旧实现复用常驻草稿模块 JYYJHelper，每轮 DeleteLines 大删「含可执行语句的过程区」会触发 VBE
            // 整工程重编译 → 重置所有模块级变量（含状态模块 JYYJXaiState 的 gXaiState），持久状态跨调用丢失。
            // VerifyVbaStateSemantics 决定性实验（步骤25/26/28）：
            //  - 步骤25：删除含可执行语句的过程块 → 整工程重编译 → 他模块状态被重置（复现生产根因）；
            //  - 步骤26：components.Remove 整模块移除 → 不触发整工程重编译，他模块状态保留；
            //  - 步骤28：每轮「新建运行模块→设状态→移除→再新建」→ 状态跨调用保留。
            // 故改为：每轮新建独立运行模块 JYYJRun_<guid>，写入声明区+辅助块+探针+片段后运行，
            // finally 整模块 components.Remove；状态模块创建后永不再改写。
            dynamic vbProject = null;
            dynamic components = null;
            dynamic module = null;
            dynamic codeModule = null;
            var restoredScreen = false;
            object prevScreen = null;
            var restoredAutoSec = false;
            object prevAutoSec = null;
            try
            {
                // 优先锚定数据工具实际操作的 ActiveWorkbook 工程；VBE 焦点漂移（编辑器/他簿窗口激活）时，
                // ActiveVBProject 可能与 ActiveWorkbook 不一致，把宏注入到错误/被禁用的工程后 Run 会报「宏不可用或被禁用」。
                vbProject = ResolveTargetVBProject();
                if (vbProject == null)
                {
                    return Invalid("VBA_UNAVAILABLE", "无法访问 VBA 工程对象模型，请确认已启用“信任对 VBA 工程对象模型的访问”。");
                }

                // 注入+运行全程临时置低自动化安全级别，结束时还原：规避 AutomationSecurity=ForceDisable 导致的「所有宏被禁用」。
                prevAutoSec = GetAutomationSecurity();
                if (prevAutoSec != null && Convert.ToInt32(prevAutoSec) != 1)
                {
                    try { application.AutomationSecurity = 1; restoredAutoSec = true; } catch { }
                }

                components = vbProject.VBComponents;
                // P3-D2 持久状态：确保独立状态模块就绪（创建后不再改写，模块级变量跨调用保留）。
                // P3-D4 显式清除：state_clear=true 时先清空状态再执行。
                var stateCreated = EnsureStateModule(components);
                if (GetBool(args, "state_clear", false))
                {
                    ClearVbaState();
                }

                // 清理历史遗留运行模块（旧版本 JYYJHelper / 崩溃残留 JYYJRun_*）：整模块移除不重置状态（步骤26）。
                CleanupRunModules(components);

                // 每轮新建独立运行模块（唯一名），写入顺序：声明区 → 辅助块(过程) → 探针 → 片段（片段最后，
                // 保证缺省入口探测命中片段自带过程而非辅助块过程）。名字必须用 RunModulePrefix，CleanupRunModules 靠它回收。
                module = components.Add(1); // vbext_ct_StdModule
                try
                {
                    module.Name = RunModulePrefix + Guid.NewGuid().ToString("N").Substring(0, 8);
                }
                catch
                {
                }

                codeModule = module.CodeModule;
                codeModule.AddFromString(VbaHelperDeclarations);
                codeModule.AddFromString(helperBlock);
                // 归一化后的片段文本单独保留：入口解析必须以「片段本身」为源，而非整个模块——
                // 真机实测 VBE 会把片段过程（如 Sub T）重排到声明区之后、辅助块过程之前，
                // 模块内最后一个过程变成 XaiProbe9F2E，按模块整体探测会把探针当入口 → 探针自调无参 → 449「参数不可选」。
                var normalizedSnippet = NormalizeVbaModule(snippet, macroName);
                codeModule.AddFromString(normalizedSnippet);

                // 片段未自带过程声明 → 上面已包成 RunSnippet，入口必为 RunSnippet；
                // 否则（自带过程）用 macroName 选定或从片段文本探测最后一个过程（与 VBE 重排无关，稳定）。
                var wrappedEntry = !SnippetDefinesProcedure(snippet.Trim());
                var proc = ResolveVbaProcedure(codeModule, macroName, wrappedEntry, normalizedSnippet);
                if (proc == null)
                {
                    return Invalid("VBA_NO_PROCEDURE", "未检测到可执行的 Sub/Function 入口。");
                }

                // 探针随本轮运行模块一起注入（带 On Error）：既捕获用户代码运行时错误的真实消息（回喂侧边栏，
                // 替代晦涩的 HRESULT），又因 On Error 接管而抑制 Excel 弹出的 VBA 模态错误对话框。
                var moduleName = Convert.ToString(codeModule.Parent.Name);
                var probeFQ = moduleName + ".XaiProbe9F2E";
                var advice = BuildStaticAdvice(snippet);

                object probeRaw = null;
                try
                {
                    prevScreen = GetScreenUpdating();
                    if (Convert.ToBoolean(prevScreen) == true)
                    {
                        application.ScreenUpdating = false;
                        restoredScreen = true;
                    }

                    probeRaw = RunVbaProbeWithRetry(codeModule, probeFQ, proc, helperBlock);
                }
                finally
                {
                    if (restoredScreen)
                    {
                        try
                        {
                            application.ScreenUpdating = prevScreen;
                        }
                        catch
                        {
                        }
                    }
                }

                var parsed = ParseVbaProbeResult(probeRaw);
                if (parsed.IsOk)
                {
                    object resultData = null;
                    if (!string.IsNullOrEmpty(parsed.ResultJson))
                    {
                        // 截断时回传截断后的 JSON 字符串预览（非合法 JSON）；完整时解析为结构化 JToken。
                        if (parsed.ResultTruncated)
                        {
                            resultData = parsed.ResultJson;
                        }
                        else
                        {
                            try
                            {
                                resultData = Newtonsoft.Json.Linq.JToken.Parse(parsed.ResultJson);
                            }
                            catch
                            {
                                resultData = parsed.ResultJson;
                            }
                        }
                    }

                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            procedure = proc,
                            ran = true,
                            module = moduleName,
                            result = resultData,
                            result_truncated = parsed.ResultTruncated,
                            state_reset = stateCreated
                        }
                    };
                }

                // 探针捕获到 VBA 运行时错误：把真实消息回喂（已抑制 Excel 模态弹窗）。
                return Invalid("VBA_RUN_ERROR",
                    "VBA 运行时错误 " + parsed.ErrorNumber + "：" + parsed.ErrorDescription
                    + "（已中断，无弹窗）。请修正代码后重试，或改用内置单元格/格式/流程工具。" + advice);
            }
            catch (VbaModuleUnavailableException)
            {
                var reason = DetectVbaCompileIssue(snippet);
                return Invalid("VBA_SYNTAX_ERROR",
                    "VBA 代码无法编译：整段代码（含执行探针）在 VBE 中未能被解析为可运行过程，通常是代码存在编译错误（变量未定义、语句不完整、End Sub 缺失等）。"
                    + (string.IsNullOrEmpty(reason) ? "" : " 已定位到最可能的编译问题：" + reason)
                    + " 请先自检代码语法并修正后再执行；若仅为取值/格式化，请优先改用内置单元格/格式/流程工具完成操作，避免编写易错的底层宏。");
            }
            catch (Exception ex)
            {
                // 把静态校验建议并入错误信息回喂模型：失败时模型可据此重写成熟代码替代，而非反复用错误代码重试。
                var advice = "";
                try
                {
                    var staticCheck = StaticVbaValidate(snippet);
                    if (staticCheck.Data != null)
                    {
                        var jo = Newtonsoft.Json.Linq.JObject.FromObject(staticCheck.Data);
                        var issues = jo["issues"] as Newtonsoft.Json.Linq.JArray;
                        if (issues != null && issues.Count > 0)
                        {
                            var sb = new System.Text.StringBuilder();
                            foreach (var it in issues)
                            {
                                sb.Append(it["message"]).Append(' ').Append(it["hint"]).Append("; ");
                            }

                            advice = " 请按以下静态建议重写代码：" + sb.ToString().Trim(' ', ';') + "（或改用现有工具/换一种写法）。";
                        }
                    }
                }
                catch
                {
                }

                var code = ClassifyVbaRunError(ex);
                var detail = code == "VBA_MACROS_DISABLED"
                    ? "VBA 执行失败：当前工作簿的宏被禁用或宏不可用。请在 Excel“文件→选项→信任中心→宏设置”中勾选“启用所有宏”并启用“信任对 VBA 工程对象模型的访问”；或改用内置单元格/格式/流程工具完成此操作（推荐，无需宏环境）。"
                    : "VBA 执行失败：" + ex.Message + "（已中断，未造成后续影响）。";
                return Invalid(code, detail + advice);
            }
            finally
            {
                // P3-D2 关键：无论成败都整模块移除本轮运行模块——components.Remove 不触发整工程重编译
                // （VerifyVbaStateSemantics 步骤26/28），状态模块 JYYJXaiState 的模块级变量跨调用保留。
                try
                {
                    RemoveComponent(vbProject, module);
                }
                catch
                {
                }

                if (restoredAutoSec)
                {
                    try { application.AutomationSecurity = prevAutoSec; } catch { }
                }

                Release((object)codeModule);
                Release((object)module);
                Release((object)components);
                Release((object)vbProject);
            }
        }

        /// <summary>清理历史遗留运行模块：旧版本常驻草稿 JYYJHelper，以及崩溃残留的 JYYJRun_* 模块。
        /// components.Remove 整模块移除不触发整工程重编译（VerifyVbaStateSemantics 步骤26），安全。</summary>
        private void CleanupRunModules(dynamic components)
        {
            try
            {
                var count = Convert.ToInt32(components.Count);
                for (var i = count; i >= 1; i--)
                {
                    dynamic c = null;
                    try
                    {
                        c = components.Item(i);
                        if (IsRunModuleName(Convert.ToString(c.Name)))
                        {
                            components.Remove(c);
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        Release((object)c);
                    }
                }
            }
            catch
            {
            }
        }

        /// <summary>解析运行宏所应锚定的 VB 工程：优先数据工具实际操作的 ActiveWorkbook，其次其第一个工作簿，
        /// 最后回退 VBE.ActiveVBProject。避免 VBE 焦点漂移时把宏注入到与数据所在不同的工程（导致 Run 报宏不可用）。</summary>
        private dynamic ResolveTargetVBProject()
        {
            dynamic vbProject = null;
            try
            {
                var wb = application.ActiveWorkbook;
                if (wb != null)
                {
                    try { vbProject = wb.VBProject; } catch { vbProject = null; }
                }
                else
                {
                    try { if (application.Workbooks != null && application.Workbooks.Count > 0) vbProject = application.Workbooks.Item(1).VBProject; } catch { vbProject = null; }
                }
            }
            catch
            {
                AppLog.Warn("ScriptRunnerService", "[LogCover] ScriptRunnerService.cs:479  silent-catch covered (no exception variable)");
                vbProject = null;
            }

            if (vbProject == null)
            {
                try { vbProject = application.VBE?.ActiveVBProject; } catch { vbProject = null; }
            }

            return vbProject;
        }

        private object GetAutomationSecurity()
        {
            try
            {
                return application.AutomationSecurity;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>区分「宏被禁用/不可用」的环境性失败与普通代码运行失败，返回对应错误码，便于模型与用户采取不同纠正。</summary>
        public static string ClassifyVbaRunError(Exception ex)
        {
            var msg = ex?.Message ?? string.Empty;
            if (msg.IndexOf("cannot run the macro", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("may not be available", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("all macros may be disabled", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("宏被禁用", StringComparison.Ordinal) >= 0
                || msg.IndexOf("在此工作簿中不可用", StringComparison.Ordinal) >= 0)
            {
                return "VBA_MACROS_DISABLED";
            }

            return "VBA_RUN_ERROR";
        }

        private object GetScreenUpdating()
        {
            try
            {
                return application.ScreenUpdating;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>按名查找草稿模块；不存在才新建并命名（vbext_ct_StdModule），返回后由调用方用完即放。</summary>
        private dynamic FindOrAddScratchModule(dynamic components)
        {
            dynamic module = null;
            try
            {
                module = components.Item(ScratchModuleName);
            }
            catch
            {
                module = null;
            }

            if (module == null)
            {
                module = components.Add(1); // vbext_ct_StdModule
                try
                {
                    module.Name = ScratchModuleName;
                }
                catch
                {
                    // 命名失败不影响本次执行；后续按默认名亦可继续（复用则退化为每次新建）。
                }
            }

            return module;
        }

        /// <summary>确保持久状态模块就绪：已存在且含预期标记 → 原样保留（模块级 gXaiState 值跨调用保留）；
        /// 不存在或代码过期 → 移除重建。返回是否（重）建了模块（供 state_reset 提示）。</summary>
        private bool EnsureStateModule(dynamic components)
        {
            dynamic module = null;
            try
            {
                module = components.Item(StateModuleName);
            }
            catch
            {
                module = null;
            }

            if (module != null)
            {
                var ready = false;
                try
                {
                    dynamic cm = module.CodeModule;
                    var count = Convert.ToInt32(cm.CountOfLines);
                    var code = count > 0 ? Convert.ToString(cm.Lines(1, count)) : "";
                    Release((object)cm);
                    ready = code.IndexOf("XaiStateRaw", StringComparison.Ordinal) >= 0;
                }
                catch
                {
                    ready = false;
                }

                if (ready)
                {
                    Release((object)module);
                    return false;
                }

                // 存在但代码过期（addin 升级场景）：移除重建，状态丢弃可接受（升级即清场）。
                try
                {
                    components.Remove(module);
                }
                catch
                {
                }
                Release((object)module);
                module = null;
            }

            dynamic created = components.Add(1); // vbext_ct_StdModule
            try
            {
                created.Name = StateModuleName;
            }
            catch
            {
            }
            dynamic createdCm = created.CodeModule;
            try
            {
                createdCm.AddFromString(StateModuleSource);
            }
            finally
            {
                Release((object)createdCm);
                Release((object)created);
            }

            return true;
        }

        /// <summary>清空 VBA 持久状态（JYYJXaiState.gXaiState）。state_clear=true 时执行；尽力而为，失败不阻断。</summary>
        private void ClearVbaState()
        {
            try
            {
                application.Run(StateModuleName + ".XaiStateClearRaw");
            }
            catch
            {
            }
        }

        /// <summary>P3-D2 关键：写片段前只清「过程区」，保留「声明区」（恒为 5 行静态辅助声明）。
        /// 破坏性整清重写（DeleteLines(1,count)）触发 VBE 工程级重编译，会把独立状态模块 JYYJXaiState
        /// 的模块级 gXaiState 一并重置（VerifyVbaStateSemantics 步骤8b 实证）；仅替换过程区则状态跨调用保留
        /// （步骤9b 实证）。声明区非静态（旧版本 const 或片段残留模块级声明，罕见）时整清重建，本轮状态重置可接受。
        /// 返回是否发生了整清重建（供 state_reset 提示）。</summary>
        private bool ReplaceModuleCode(dynamic codeModule, string snippet, string macroName)
        {
            var fullReset = false;
            try
            {
                var total = Convert.ToInt32(codeModule.CountOfLines);
                if (total > 0)
                {
                    var decl = Convert.ToInt32(codeModule.CountOfDeclarationLines);
                    var current = decl > 0 ? Convert.ToString(codeModule.Lines(1, decl)) : "";
                    if (IsStaticDeclarationSection(current))
                    {
                        // 正常路径：仅删过程区，声明区原样保留 → 不触发工程级重编译，持久状态存活。
                        if (total > decl)
                        {
                            codeModule.DeleteLines(decl + 1, total - decl);
                        }
                    }
                    else
                    {
                        fullReset = true;
                        codeModule.DeleteLines(1, total);
                    }
                }
            }
            catch
            {
                // 读取/清空失败则直接追加
            }

            codeModule.AddFromString(NormalizeVbaModule(snippet, macroName));
            return fullReset;
        }

        /// <summary>判断模块声明区是否为恒定静态声明（VbaHelperDeclarations）：逐行 Trim 去空后精确比对，
        /// 区分旧版本 const 版（Private Const XAI_RESULT_LIMIT As Long = N）与当前变量版，避免把旧声明当静态保留。</summary>
        internal static bool IsStaticDeclarationSection(string text)
        {
            return NormalizeDeclLines(text) == NormalizeDeclLines(VbaHelperDeclarations);
        }

        private static string NormalizeDeclLines(string text)
        {
            if (text == null)
            {
                return string.Empty;
            }

            return string.Join("\n", text.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0));
        }

        /// <summary>VBA 探针源码：包一层 On Error，把用户过程的运行时错误捕获为可读字符串返回；
        /// 成功时若片段调用过 XaiReturn，则按 OK|&lt;json&gt;（完整）/ OKT|&lt;json&gt;（截断）协议回传结构化结果。</summary>
        private const string VbaProbeSource =
            "\nPublic Function XaiProbe9F2E(p As String) As String\n" +
            "    On Error GoTo fail\n" +
            "    XaiInitArgs\n" +
            "    Application.Run p\n" +
            "    If Len(gXaiResult) > 0 Then\n" +
            "        If gXaiResultTruncated Then\n" +
            "            XaiProbe9F2E = \"OKT|\" & gXaiResult\n" +
            "        Else\n" +
            "            XaiProbe9F2E = \"OK|\" & gXaiResult\n" +
            "        End If\n" +
            "    Else\n" +
            "        XaiProbe9F2E = \"OK\"\n" +
            "    End If\n" +
            "    Exit Function\n" +
            "fail:\n" +
            "    XaiProbe9F2E = \"ERR|\" & CStr(Err.Number) & \"|\" & Err.Description\n" +
            "End Function";

        /// <summary>VBA 辅助块·稳定声明区（P3-D2）：只在草稿模块缺失/过期时补一次，此后每轮替换过程区不触碰。
        /// 声明区必须保持恒定（XAI_RESULT_LIMIT 为变量、由每轮 XaiInitArgs 赋值），
        /// 是 ReplaceModuleCode「保留声明区→避免整工程重编译重置持久状态」的锚点。恒 5 行：注释 + 3 个 Private + 上限变量。</summary>
        internal const string VbaHelperDeclarations = """""
' ===== XAI 结构化回传 / 入参辅助 =====
Private gXaiResult As String
Private gXaiResultTruncated As Boolean
Private gXaiArgs As String
Private XAI_RESULT_LIMIT As Long
""""";

        /// <summary>VBA 辅助块过程区（序列化器 + 解析器 + XaiReturn/XaiArg + 状态编解码）：随探针一起按轮重建。
        /// 声明区已拆分到 VbaHelperDeclarations，此处不再含模块级声明，避免每轮重复添加触发「重复声明」编译错误。</summary>
        private const string VbaHelperSource = """""
Public Sub XaiReturn(v As Variant)
    Dim s As String
    s = XaiJson(v, 0)
    If Len(s) > XAI_RESULT_LIMIT Then
        gXaiResult = Left$(s, XAI_RESULT_LIMIT) & ChrW(&H2026)
        gXaiResultTruncated = True
    Else
        gXaiResult = s
        gXaiResultTruncated = False
    End If
End Sub

Public Function XaiArg(key As String) As Variant
    Dim obj As Variant
    obj = XaiParse(gXaiArgs)
    If IsArray(obj) Then
        Dim i As Long
        ' 上界 -1 + Step 2：仅按 [键,值,键,值,…] 成对遍历，奇数长/纯数组输入不会访问越界元素
        For i = LBound(obj, 1) To UBound(obj, 1) - 1 Step 2
            If CStr(obj(i)) = key Then
                XaiArg = obj(i + 1)
                Exit Function
            End If
        Next i
    End If
    XaiArg = Null
End Function

Public Function XaiStateGet(key As String) As Variant
    Dim obj As Variant
    obj = XaiParse(XaiStateRaw)
    If IsArray(obj) Then
        Dim i As Long
        ' 键值对数组 [键,值,键,值,…]：成对遍历找键，越界防护同 XaiArg
        For i = LBound(obj, 1) To UBound(obj, 1) - 1 Step 2
            If CStr(obj(i)) = key Then
                XaiStateGet = obj(i + 1)
                Exit Function
            End If
        Next i
    End If
    XaiStateGet = Null
End Function

Public Sub XaiStateSet(key As String, v As Variant)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim obj As Variant
    obj = XaiParse(XaiStateRaw)
    If IsArray(obj) Then
        Dim i As Long
        For i = LBound(obj, 1) To UBound(obj, 1) - 1 Step 2
            d(CStr(obj(i))) = obj(i + 1)
        Next i
    End If
    d(key) = v
    XaiStateRawSet XaiJson(d, 0)
End Sub

Public Sub XaiStateClear()
    XaiStateClearRaw
End Sub

Private Function XaiJson(v As Variant, depth As Long) As String
    If depth > 32 Then
        XaiJson = """<max_depth>"""
        Exit Function
    End If
    If (VarType(v) And vbArray) <> 0 Then
        XaiJson = XaiArr(v, depth)
        Exit Function
    End If
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            XaiJson = "null"
        Case vbBoolean
            If v Then XaiJson = "true" Else XaiJson = "false"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            XaiJson = XaiNum(v)
        Case vbString
            XaiJson = """" & XaiEsc(CStr(v)) & """"
        Case vbDate
            XaiJson = """" & Format(v, "yyyy-mm-ddThh:nn:ss") & """"
        Case vbObject
            XaiJson = XaiObj(v, depth)
        Case Else
            XaiJson = "null"
    End Select
End Function

Private Function XaiNum(v As Variant) As String
    ' Str() 的小数点随系统区域设置（如 de-DE 用逗号），统一替换为 '.' 保证输出合法 JSON。
    XaiNum = Replace(Trim(Str(v)), ",", ".")
End Function

Private Function XaiEsc(s As String) As String
    Dim sb As String
    Dim i As Long
    Dim c As String
    sb = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case """": sb = sb & "\"""
            Case "\": sb = sb & "\\"
            Case vbCr: sb = sb & "\r"
            Case vbLf: sb = sb & "\n"
            Case vbTab: sb = sb & "\t"
            Case Else:
                If AscW(c) < 32 Then
                    sb = sb & "\u" & Right$("000" & Hex$(AscW(c)), 4)
                Else
                    sb = sb & c
                End If
        End Select
    Next i
    XaiEsc = sb
End Function

Private Function XaiArr(v As Variant, depth As Long) As String
    Dim nd As Long
    nd = XaiDims(v)
    Dim sb As String
    Dim i As Long
    Dim j As Long
    sb = "["
    If nd = 1 Then
        For i = LBound(v, 1) To UBound(v, 1)
            If Len(sb) > 1 Then sb = sb & ","
            sb = sb & XaiJson(v(i), depth + 1)
        Next i
    ElseIf nd >= 2 Then
        For i = LBound(v, 1) To UBound(v, 1)
            If Len(sb) > 1 Then sb = sb & ","
            sb = sb & "["
            For j = LBound(v, 2) To UBound(v, 2)
                If j > LBound(v, 2) Then sb = sb & ","
                sb = sb & XaiJson(v(i, j), depth + 1)
            Next j
            sb = sb & "]"
        Next i
    End If
    XaiArr = sb & "]"
End Function

Private Function XaiDims(v As Variant) As Long
    Dim n As Long
    n = 0
    On Error Resume Next
    Err.Clear
    Do
        n = n + 1
        Dim x As Long
        x = UBound(v, n)
    Loop While Err.Number = 0
    On Error GoTo 0
    XaiDims = n - 1
End Function

Private Function XaiObj(v As Variant, depth As Long) As String
    Dim tn As String
    tn = TypeName(v)
    If tn = "Dictionary" Then
        XaiObj = XaiDict(v, depth)
    ElseIf tn = "Collection" Then
        XaiObj = XaiColl(v, depth)
    Else
        XaiObj = "null"
    End If
End Function

Private Function XaiDict(v As Variant, depth As Long) As String
    Dim sb As String
    Dim k As Variant
    sb = "{"
    For Each k In v.Keys
        If Len(sb) > 1 Then sb = sb & ","
        sb = sb & """" & XaiEsc(CStr(k)) & """:" & XaiJson(v(k), depth + 1)
    Next k
    XaiDict = sb & "}"
End Function

Private Function XaiColl(v As Variant, depth As Long) As String
    Dim sb As String
    Dim it As Variant
    sb = "["
    For Each it In v
        If Len(sb) > 1 Then sb = sb & ","
        sb = sb & XaiJson(it, depth + 1)
    Next it
    XaiColl = sb & "]"
End Function

Private Function XaiParse(s As String) As Variant
    Dim p As Long
    p = 1
    XaiParse = XaiParseVal(s, p)
End Function

Private Function XaiParseVal(s As String, p As Long) As Variant
    Call XaiSkipWs(s, p)
    Dim c As String
    c = Mid$(s, p, 1)
    If c = "{" Then
        XaiParseVal = XaiParseObj(s, p)
    ElseIf c = "[" Then
        XaiParseVal = XaiParseArr(s, p)
    ElseIf c = """" Then
        XaiParseVal = XaiParseStr(s, p)
    ElseIf c = "t" Then
        p = p + 4
        XaiParseVal = True
    ElseIf c = "f" Then
        p = p + 5
        XaiParseVal = False
    ElseIf c = "n" Then
        p = p + 4
        XaiParseVal = Null
    Else
        XaiParseVal = XaiParseNum(s, p)
    End If
End Function

Private Function XaiParseObj(s As String, p As Long) As Variant
    Dim arr() As Variant
    Dim n As Long
    n = 0
    p = p + 1
    Call XaiSkipWs(s, p)
    If Mid$(s, p, 1) = "}" Then
        p = p + 1
        XaiParseObj = Array()
        Exit Function
    End If
    Do
        Call XaiSkipWs(s, p)
        Dim k As String
        k = XaiParseStr(s, p)
        Call XaiSkipWs(s, p)
        If Mid$(s, p, 1) = ":" Then p = p + 1
        Call XaiSkipWs(s, p)
        Dim v As Variant
        v = XaiParseVal(s, p)
        n = n + 1
        ReDim Preserve arr(1 To n * 2)
        arr(n * 2 - 1) = k
        arr(n * 2) = v
        Call XaiSkipWs(s, p)
        Dim sep As String
        sep = Mid$(s, p, 1)
        If sep = "," Then
            p = p + 1
        ElseIf sep = "}" Then
            p = p + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop
    XaiParseObj = arr
End Function

Private Function XaiParseArr(s As String, p As Long) As Variant
    Dim arr() As Variant
    Dim n As Long
    n = 0
    p = p + 1
    Call XaiSkipWs(s, p)
    If Mid$(s, p, 1) = "]" Then
        p = p + 1
        XaiParseArr = Array()
        Exit Function
    End If
    Do
        Call XaiSkipWs(s, p)
        Dim v As Variant
        v = XaiParseVal(s, p)
        n = n + 1
        ReDim Preserve arr(1 To n)
        arr(n) = v
        Call XaiSkipWs(s, p)
        Dim sep As String
        sep = Mid$(s, p, 1)
        If sep = "," Then
            p = p + 1
        ElseIf sep = "]" Then
            p = p + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop
    XaiParseArr = arr
End Function

Private Function XaiParseStr(s As String, p As Long) As String
    Dim sb As String
    sb = ""
    p = p + 1
    Do While p <= Len(s)
        Dim c As String
        c = Mid$(s, p, 1)
        If c = """" Then
            p = p + 1
            Exit Do
        ElseIf c = "\" Then
            p = p + 1
            Dim e As String
            e = Mid$(s, p, 1)
            Select Case e
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case "t": sb = sb & vbTab
                Case "b": sb = sb & Chr(8)
                Case "f": sb = sb & Chr(12)
                Case "/": sb = sb & "/"
                Case "\": sb = sb & "\\"
                Case """": sb = sb & """"
                Case "u":
                    sb = sb & ChrW(Val("&H" & Mid$(s, p + 1, 4)))
                    p = p + 4
                Case Else: sb = sb & e
            End Select
            p = p + 1
        Else
            sb = sb & c
            p = p + 1
        End If
    Loop
    XaiParseStr = sb
End Function

Private Function XaiParseNum(s As String, p As Long) As Variant
    Dim start As Long
    start = p
    Do While p <= Len(s)
        Dim c As String
        c = Mid$(s, p, 1)
        If c = "-" Or c = "+" Or c = "." Or c = "e" Or c = "E" Or (c >= "0" And c <= "9") Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
    Dim txt As String
    txt = Mid$(s, start, p - start)
    If InStr(txt, ".") > 0 Or InStr(txt, "e") > 0 Or InStr(txt, "E") > 0 Then
        XaiParseNum = Val(txt)
    Else
        ' CLng 超出 Long 范围（±21.47 亿）会溢出（VBA 错误 6），大整数回退 Double 而非让整段探针失败。
        Dim d As Double
        d = Val(txt)
        If d > 2147483647# Or d < -2147483648# Then
            XaiParseNum = d
        Else
            XaiParseNum = CLng(d)
        End If
    End If
End Function

Private Sub XaiSkipWs(s As String, p As Long)
    Do While p <= Len(s)
        Dim c As String
        c = Mid$(s, p, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
End Sub
""""";

        /// <summary>VBA 持久状态模块源码（P3-D2）：独立成模块、创建后不再改写，故模块级 gXaiState 值跨调用保留。
        /// 只暴露原始读写/清空三个最小函数；编解码逻辑（XaiStateGet/Set/Clear）留在每轮重建的草稿模块辅助块里。
        /// internal 供测试断言状态区契约（模块级变量 + 三个最小访问器）。</summary>
        internal const string StateModuleSource = """""
' ===== XAI 持久状态模块 v1（创建后不再改写，模块级变量跨调用保留）=====
Option Explicit
Private gXaiState As String

Public Function XaiStateRaw() As String
    XaiStateRaw = gXaiState
End Function

Public Sub XaiStateRawSet(v As String)
    gXaiState = v
End Sub

Public Sub XaiStateClearRaw()
    gXaiState = ""
End Sub
""""";

        /// <summary>组装本轮 VBA 辅助块：过程区源码 + 携带本次 args 与结果上限的 XaiInitArgs + 探针。
        /// XAI_RESULT_LIMIT 由 XaiInitArgs 赋值为模块级变量（声明区保持恒定，见 VbaHelperDeclarations）。</summary>
        internal static string BuildVbaHelperBlock(string argsJson, int resultLimitChars)
        {
            var init = "\nPublic Sub XaiInitArgs()\n"
                       + "    XAI_RESULT_LIMIT = " + resultLimitChars.ToString() + "\n"
                       + "    gXaiArgs = " + VbaStringLiteral(argsJson) + "\n"
                       + "    gXaiResult = \"\"\n"
                       + "    gXaiResultTruncated = False\n"
                       + "End Sub";
            return VbaHelperSource + init + "\n" + VbaProbeSource;
        }

        /// <summary>把任意文本安全嵌入 VBA 字符串字面量：引号转义为 ""，超长按 100 字符分块用 &amp; _ 续行拼接。
        /// 分块以原始字符为单位、逐块独立转义，避免边界切断 "" 转义对导致字面量未闭合（VBA 单行上限 1024 字符）。</summary>
        internal static string VbaStringLiteral(string text)
        {
            text = text ?? string.Empty;
            if (text.Length <= 100)
            {
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }

            var sb = new StringBuilder();
            sb.Append('"');
            var pos = 0;
            var first = true;
            while (pos < text.Length)
            {
                var end = Math.Min(pos + 100, text.Length);
                var chunk = text.Substring(pos, end - pos).Replace("\"", "\"\"");
                if (!first)
                {
                    sb.Append("\" & _\n    \"");
                }

                sb.Append(chunk);
                first = false;
                pos = end;
            }

            sb.Append('"');
            return sb.ToString();
        }

        private static void AppendVbaProbe(dynamic codeModule, string helperBlock)
        {
            try
            {
                // P3-D2：静态声明区缺失/过期（首次创建、整清重建、旧版本 const 残留）时先补写，
                // 保证声明区恒为 VbaHelperDeclarations，是 ReplaceModuleCode 保活状态的锚点。
                EnsureStaticDeclarations(codeModule);

                // 幂等：探针已在模块里就不再重复 AddFromString，否则重试路径会追加第二个 XaiProbe9F2E → 重复声明，整模块编译失败。
                var hasProbe = false;
                try
                {
                    var count = Convert.ToInt32(codeModule.CountOfLines);
                    if (count > 0)
                    {
                        var all = Convert.ToString(codeModule.Lines(1, count));
                        hasProbe = all.IndexOf("XaiProbe9F2E", StringComparison.Ordinal) >= 0;
                    }
                }
                catch
                {
                    hasProbe = false;
                }

                if (!hasProbe)
                {
                    codeModule.AddFromString(helperBlock);
                }
            }
            catch
            {
            }
        }

        /// <summary>确保持久辅助声明区就绪：缺失或非静态（旧版本 const / 残留）时清旧声明重写为 VbaHelperDeclarations。
        /// 删除声明行属局部编辑，不会重置其他模块的状态（VerifyVbaStateSemantics 步骤6 实证：仅改动本模块声明不殃及他模块）。</summary>
        private static void EnsureStaticDeclarations(dynamic codeModule)
        {
            try
            {
                var decl = Convert.ToInt32(codeModule.CountOfDeclarationLines);
                var current = decl > 0 ? Convert.ToString(codeModule.Lines(1, decl)) : "";
                if (IsStaticDeclarationSection(current))
                {
                    return;
                }

                if (decl > 0)
                {
                    codeModule.DeleteLines(1, decl);
                }

                codeModule.InsertLines(1, VbaHelperDeclarations);
            }
            catch
            {
                // 声明区维护尽力而为：失败则沿用既有内容
            }
        }

        /// <summary>首次 Application.Run(探针) 可能因模块刚写入/刚改写而探针尚未被 VBE 编译器就绪，
        /// 抛「宏不可用/not available」（如 HRESULT 0x800A9C68）。此时重新追加探针并重试一次自愈，
        /// 避免把这种瞬态失败当普通代码运行错误回喂，否则第一轮总报晦涩 HRESULT、第二轮才成功。
        /// 若第一、二次都报「宏/过程不可用」（探针的 On Error 已接管用户代码的运行时错误），则说明草稿模块
        /// 整体未能编译（用户代码存在编译错误使整个模块含探针都无法被 VBE 解析），判定为必失败并抛专用异常，
        /// 由调用方归为 VBA_SYNTAX_ERROR，而不是把晦涩的“异常来自 HRESULT:0x800A9C68”当运行时错误回喂。</summary>
        private object RunVbaProbeWithRetry(dynamic codeModule, string probeFQ, string proc, string helperBlock)
        {
            try
            {
                var r = application.Run(probeFQ, proc);
                return r;
            }
            catch (Exception ex)
            {
                if (!ProbeLikelyUnavailable(ex))
                {
                    throw;
                }

                AppendVbaProbe(codeModule, helperBlock);
                try
                {
                    var r2 = application.Run(probeFQ, proc);
                    return r2;
                }
                catch (Exception ex2)
                {
                    if (ProbeLikelyUnavailable(ex2))
                    {
                        // 连续两次都命中「过程不可用」：不是瞬态未就绪，而是整模块编译失败了。
                        throw VbaModuleUnavailableException.From(ex, probeFQ);
                    }

                    throw;
                }
            }
        }

        /// <summary>判断是否「宏尚未就绪/不可用」类瞬态错误，命中则值得重试一次探针（自愈第一轮 HRESULT 0x800A9C68）。</summary>
        public static bool ProbeLikelyUnavailable(Exception ex)
        {
            if (ex == null)
            {
                return false;
            }

            // 0x800A9C68：自动化调用显式宏/过程尚未就绪（模块刚建/刚改未编译）；常见于第一轮调用。
            if (unchecked((ushort)(0x0000FFFF & (uint)ex.HResult)) == 0x9C68)
            {
                return true;
            }

            var m = ex.Message ?? string.Empty;
            return m.IndexOf("cannot run the macro", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("may not be available", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("not available in this workbook", StringComparison.OrdinalIgnoreCase) >= 0
                // 中文版 Excel 的瞬态「宏尚未就绪」消息（模块刚写入未编译时第一轮 Run 必现）：
                // “无法运行“X”宏。可能是因为该宏在此工作簿中不可用，或者所有的宏都被禁用。”
                || m.IndexOf("无法运行", StringComparison.Ordinal) >= 0
                || m.IndexOf("在此工作簿中不可用", StringComparison.Ordinal) >= 0
                || m.IndexOf("宏都被禁用", StringComparison.Ordinal) >= 0;
        }

        /// <summary>解析探针返回：'OK' → 成功且无结构化结果；'OK|&lt;json&gt;' → 成功且带结构化结果；
        /// 'OKT|&lt;json&gt;' → 成功但结果超上限被截断；'ERR|num|desc' → 运行时错误（真实消息）。
        /// 空值/非预期 → 视为未知运行错误，isOk=false 且 errDesc 为原始串。</summary>
        public static (bool IsOk, int ErrorNumber, string ErrorDescription, string ResultJson, bool ResultTruncated) ParseVbaProbeResult(object raw)
        {
            var s = raw == null ? string.Empty : Convert.ToString(raw);
            s = s == null ? string.Empty : s.Trim();
            if (string.IsNullOrEmpty(s) || s == "OK")
            {
                return (true, 0, string.Empty, null, false);
            }

            if (s.StartsWith("OKT|", StringComparison.Ordinal))
            {
                return (true, 0, string.Empty, s.Substring(4), true);
            }

            if (s.StartsWith("OK|", StringComparison.Ordinal))
            {
                return (true, 0, string.Empty, s.Substring(3), false);
            }

            if (s.StartsWith("ERR|", StringComparison.Ordinal))
            {
                var parts = s.Split(new[] { '|' }, 3, StringSplitOptions.None);
                int num = 0;
                if (parts.Length > 1)
                {
                    int.TryParse(parts[1], out num);
                }

                return (false, num, parts.Length > 2 ? parts[2] : string.Empty, null, false);
            }

            return (false, 0, s, null, false);
        }

        /// <summary>对失败片段附加静态校验建议（重写指引），供回喂模型，避免其带着错误代码反复重试。</summary>
        private static string BuildStaticAdvice(string snippet)
        {
            try
            {
                var staticCheck = StaticVbaValidate(snippet);
                if (staticCheck.Data != null)
                {
                    var jo = Newtonsoft.Json.Linq.JObject.FromObject(staticCheck.Data);
                    var issues = jo["issues"] as Newtonsoft.Json.Linq.JArray;
                    if (issues != null && issues.Count > 0)
                    {
                        var sb = new System.Text.StringBuilder();
                        foreach (var it in issues)
                        {
                            sb.Append(it["message"]).Append(' ').Append(it["hint"]).Append("; ");
                        }

                        return " 请按以下静态建议重写代码：" + sb.ToString().Trim(' ', ';') + "（或改用现有工具/换一种写法）。";
                    }
                }
            }
            catch
            {
            }

            return "";
        }

        public ToolResult RunPythonCode(Dictionary<string, object> args)
        {
            var code = Optional(args, "code");
            var scriptFile = Optional(args, "script_path");
            if (options.Enabled == false)
            {
                return Disabled("run_python_code");
            }

            try
            {
                var execFile = options.PythonPath;
                var generated = string.IsNullOrWhiteSpace(scriptFile);
                var targetFile = string.IsNullOrWhiteSpace(scriptFile)
                    ? WriteTempScript(code, "py")
                    : scriptFile;
                if (!generated)
                {
                    // 工具分类审查总复盘修复（GetFullPath 家族收口第 2 处）：script_path 非法路径
                    // 字符此前抛 ArgumentException 裸冒泡——补 try-catch 人话报错（对齐 CLI/EXE 口径）。
                    string fullScript;
                    try
                    {
                        fullScript = System.IO.Path.GetFullPath(targetFile);
                    }
                    catch (Exception pathEx)
                    {
                        return Invalid("INVALID_ARGS", "script_path 路径非法：" + pathEx.Message);
                    }
                    var scriptDecision = ControlledExecutionPolicy.Evaluate(fullScript, options.ScriptAllowlist, options.ScriptDenylist, fullPath: true);
                    if (scriptDecision == PolicyDecision.Deny)
                    {
                        return Invalid("SCRIPT_DENIED", $"脚本 {fullScript} 在黑名单中，已被策略拒绝。");
                    }

                    if (scriptDecision == PolicyDecision.NotAllowlisted)
                    {
                        return Invalid("NOT_ALLOWLISTED", $"脚本 {fullScript} 不在脚本白名单内。");
                    }
                }

                if (generated)
                {
                    // 越界加速：AI 直接给代码时走常驻 Python worker，复用单个长活解释器，避免
                    // 每轮冷启动 python 进程 + 写临时脚本 + 进程启停。仅使用内存调试信息与纯 Python
                    // 计算；不改变原有白名单/黑名单语义（/code 分支本就只执行模型生成的临时代码）。
                    // P1-C/B：input_data 注入命名空间；P1-B/E：xai_result 结构化回传 + result_limit_kb 护栏。
                    var (inputData, inputError) = NormalizeInputData(args);
                    if (inputError != null)
                    {
                        return Invalid("INPUT_TOO_LARGE", inputError);
                    }

                    var resultLimitBytes = Math.Min(Math.Max(GetInt(args, "result_limit_kb", 1024), 1), 16 * 1024) * 1024; // 上限 16MB（审查第五批护栏）
                    // P3-D1：persist 保留命名空间跨调用；state_clear 显式清空持久状态。
                    var execOptions = new PythonWorkerHost.ExecuteOptions
                    {
                        Persist = GetBool(args, "persist", false),
                        StateClear = GetBool(args, "state_clear", false)
                    };
                    var workerOk = PythonWorkerHost.TryExecute(
                        code,
                        execFile,
                        ClampTimeout(GetInt(args, "timeout_ms", options.TimeoutSeconds * 1000)),
                        Math.Min(Math.Max(GetInt(args, "output_limit_kb", options.OutputLimitKB), 1), 16 * 1024) * 1024,
                        inputData,
                        resultLimitBytes,
                        execOptions,
                        out var ok,
                        out var output,
                        out var result,
                        out var resultTruncated,
                        out var stateReset,
                        out var exit);
                    if (!workerOk || !ok)
                    {
                        // 缺模块自愈（工具闭环拿真实报错 → 自动修复 → 重试闭环）：
                        // ModuleNotFoundError 命中安全白名单 → pip 装包（镜像兜底）→ 重建 worker（丢弃旧导入缓存）→ 重试一次。
                        // 注意：Python 代码抛异常时 TryExecute 返回 workerOk=false（reply.ok=false），自愈判断必须
                        // 覆盖 !workerOk 分支，否则自愈永远够不着（真机实证 T1 首跑漏自愈）。超时/worker 异常的
                        // output 不含 ModuleNotFoundError 签名，不会误入自愈。
                        if (TryResolveMissingModule(output, out var missingModule, out var pipPackage))
                        {
                            var installOk = InstallPythonPackage(execFile, pipPackage, out var installMsg);
                            if (installOk)
                            {
                                PythonWorkerHost.ResetWorker();
                                var retryOk = PythonWorkerHost.TryExecute(
                                    code,
                                    execFile,
                                    ClampTimeout(GetInt(args, "timeout_ms", options.TimeoutSeconds * 1000)),
                                    Math.Min(Math.Max(GetInt(args, "output_limit_kb", options.OutputLimitKB), 1), 16 * 1024) * 1024,
                                    inputData,
                                    resultLimitBytes,
                                    execOptions,
                                    out var ok2,
                                    out var output2,
                                    out var result2,
                                    out var resultTruncated2,
                                    out var stateReset2,
                                    out var exit2);
                                if (retryOk && ok2)
                                {
                                    AppLog.Info("python", $"缺模块自愈成功：{missingModule} → pip install {pipPackage} 后重试通过");
                                    return new ToolResult
                                    {
                                        Ok = true,
                                        Data = new
                                        {
                                            exit_success = exit2 == 0,
                                            exit_code = exit2,
                                            output = output2,
                                            output_chars = output2 == null ? 0 : output2.Length,
                                            truncated = output2 != null && output2.IndexOf(OutputTruncatedMarker, StringComparison.Ordinal) >= 0,
                                            result = result2,
                                            result_truncated = resultTruncated2,
                                            state_reset = stateReset2,
                                            deps_auto_installed = pipPackage
                                        }
                                    };
                                }

                                if (retryOk)
                                {
                                    output = output2;
                                }
                            }

                            output = output + $"\n[自愈尝试] 检测到缺模块 {missingModule}，自动 pip install {pipPackage} " + (installOk ? "成功但重试仍失败" : "失败：" + installMsg);
                            return new ToolResult { Ok = false, ErrorCode = "PYTHON_RUN_ERROR", Data = output };
                        }

                        return new ToolResult { Ok = false, ErrorCode = "PYTHON_RUN_ERROR", Data = output };
                    }

                    return new ToolResult
                    {
                        Ok = ok,
                        ErrorCode = ok ? null : "PYTHON_RUN_ERROR",
                        Data = new
                        {
                            exit_success = ok && exit == 0,
                            exit_code = exit,
                            output = output,
                            output_chars = output == null ? 0 : output.Length,
                            truncated = output != null && output.IndexOf(OutputTruncatedMarker, StringComparison.Ordinal) >= 0,
                            result = result,
                            result_truncated = resultTruncated,
                            state_reset = stateReset
                        }
                    };
                }

                try
                {
                    var result = RunProcess(execFile, new[] { "-u", targetFile }, args, out var truncated);
                    if (!result.Item1)
                    {
                        return new ToolResult { Ok = false, ErrorCode = "PYTHON_RUN_ERROR", Data = result.Item2 };
                    }

                    return new ToolResult { Ok = true, Data = ParseProcessOutput(result.Item2, truncated) };
                }
                finally
                {
                    if (generated)
                    {
                        TryDelete(targetFile);
                    }
                }
            }
            catch (Exception ex)
            {
                return Invalid("PYTHON_ERROR", "Python 运行失败：" + ex.Message);
            }
        }

        /// <summary>缺模块 → pip 包名安全白名单：只自动安装常见、无歧义的数据处理/图像/办公库，防止任意模块名被拼进 pip 参数。
        /// internal 供测试断言解析与白名单行为。</summary>
        internal static readonly Dictionary<string, string> MissingModulePipMap = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["PIL"] = "pillow",
            ["cv2"] = "opencv-python",
            ["sklearn"] = "scikit-learn",
            ["yaml"] = "pyyaml",
            ["dateutil"] = "python-dateutil",
            ["docx"] = "python-docx",
            ["pptx"] = "python-pptx",
            ["bs4"] = "beautifulsoup4",
            ["lxml"] = "lxml",
            ["openpyxl"] = "openpyxl",
            ["xlsxwriter"] = "xlsxwriter",
            ["seaborn"] = "seaborn",
            ["scipy"] = "scipy",
            ["matplotlib"] = "matplotlib",
            ["numpy"] = "numpy",
            ["pandas"] = "pandas",
            ["statsmodels"] = "statsmodels",
            ["jieba"] = "jieba",
            ["wordcloud"] = "wordcloud",
            ["snownlp"] = "snownlp",
            ["plotly"] = "plotly",
            ["xlwings"] = "xlwings",
            ["openpyxl_image_loader"] = "openpyxl-image-loader",
            ["requests"] = "requests",
            ["chardet"] = "chardet",
            ["pywintypes"] = "pywin32",
            ["win32com"] = "pywin32",
            ["pythoncom"] = "pywin32"
        };

        private static readonly System.Text.RegularExpressions.Regex RegexMissingModule =
            new System.Text.RegularExpressions.Regex(
                @"ModuleNotFoundError: No module named '(?<m>[A-Za-z_]\w*)'",
                System.Text.RegularExpressions.RegexOptions.Compiled);

        /// <summary>从 Python 报错里解析缺模块名；命中白名单返回 true 及对应 pip 包名。</summary>
        internal static bool TryResolveMissingModule(string errorOutput, out string moduleName, out string pipPackage)
        {
            moduleName = null;
            pipPackage = null;
            if (string.IsNullOrWhiteSpace(errorOutput))
            {
                return false;
            }

            var m = RegexMissingModule.Match(errorOutput);
            if (!m.Success)
            {
                return false;
            }

            moduleName = m.Groups["m"].Value;
            return MissingModulePipMap.TryGetValue(moduleName, out pipPackage);
        }

        /// <summary>pip 安装单个白名单包：worker subprocess 执行（环境健康实证），清华→阿里→默认源兜底，失败退 --user 重试。</summary>
        private bool InstallPythonPackage(string pythonPath, string pipPackage, out string message)
        {
            message = null;
            var mirrors = new[] { "https://pypi.tuna.tsinghua.edu.cn/simple", "https://mirrors.aliyun.com/pypi/simple/", null };
            foreach (var mirror in mirrors)
            {
                if (PythonWorkerHost.TryPipInstall(pythonPath, pipPackage, mirror, false, 240000, out var outText))
                {
                    message = "已通过 pip 安装 " + pipPackage + "（源：" + (mirror ?? "默认 PyPI") + "）。";
                    return true;
                }

                var combined = outText ?? "";
                if ((combined.IndexOf("permission", StringComparison.OrdinalIgnoreCase) >= 0
                     || combined.IndexOf("access is denied", StringComparison.OrdinalIgnoreCase) >= 0)
                    && !combined.Contains("--user"))
                {
                    if (PythonWorkerHost.TryPipInstall(pythonPath, pipPackage, mirror, true, 240000, out var outUser))
                    {
                        message = "已通过 pip --user 安装 " + pipPackage + "（源：" + (mirror ?? "默认 PyPI") + "）。";
                        return true;
                    }
                }

                message = combined;
            }

            return false;
        }

        public ToolResult RunCliCommand(Dictionary<string, object> args)
        {
            var command = Required(args, "command");
            var commandArgs = ArgsList(args, "args");
            if (options.Enabled == false)
            {
                return Disabled("run_cli_command");
            }

            // M16 全量审查（白名单绕过修复）：原实现只按 basename 评估白名单、却按原样执行
            // command——传 "C:\temp\python.exe" 可借 basename "python" 命中白名单执行任意位置
            // 的同名二进制。修复：带路径成分的 command 按完整路径评估（必须与白名单条目精确
            // 一致）；裸命令名（走 PATH 解析）才按 basename 评估。
            string evaluated;
            if (command.IndexOf('\\') >= 0 || command.IndexOf('/') >= 0)
            {
                string fullPath;
                try
                {
                    fullPath = System.IO.Path.GetFullPath(command);
                }
                catch (Exception ex)
                {
                    return Invalid("INVALID_ARGS", "command 路径非法：" + ex.Message);
                }

                evaluated = fullPath.ToLowerInvariant();
            }
            else
            {
                evaluated = (System.IO.Path.GetFileNameWithoutExtension(command) ?? command).ToLowerInvariant();
            }

            var cliDecision = ControlledExecutionPolicy.Evaluate(evaluated, options.CliAllowlist, options.CliDenylist);
            if (cliDecision == PolicyDecision.Deny)
            {
                return Invalid(
                    "CLI_DENIED",
                    $"CLI 命令 {evaluated} 在黑名单中，已被策略拒绝。");
            }

            if (cliDecision == PolicyDecision.NotAllowlisted)
            {
                return Invalid(
                    "NOT_ALLOWLISTED",
                    $"CLI 命令 {evaluated} 不在白名单内（可用的：" + string.Join(", ", options.CliAllowlist) + "）。");
            }

            var result = RunProcess(command, commandArgs, args, out var truncatedCli);
            if (!result.Item1)
            {
                return new ToolResult { Ok = false, ErrorCode = "CLI_FAILED", Data = result.Item2 };
            }

            return new ToolResult { Ok = true, Data = ParseProcessOutput(result.Item2, truncatedCli) };
        }

        public ToolResult RunExeCommand(Dictionary<string, object> args)
        {
            var path = Required(args, "path");
            var commandArgs = ArgsList(args, "args");
            if (options.Enabled == false)
            {
                return Disabled("run_exe_command");
            }

            // 工具分类审查第五批修复（不对称漏网）：RunCliCommand 的 GetFullPath 有 try-catch 人话
            // 报错（M16 审查补的），RunExeCommand 漏网——path 含非法路径字符（如 "|dir"）时抛
            // ArgumentException/NotSupportedException 裸异常冒泡。对齐 CLI 口径。
            string full;
            try
            {
                full = System.IO.Path.GetFullPath(path);
            }
            catch (Exception ex)
            {
                return Invalid("INVALID_ARGS", "path 路径非法：" + ex.Message);
            }
            var exeDecision = ControlledExecutionPolicy.Evaluate(full, options.ExeAllowlist, options.ExeDenylist, fullPath: true);
            if (exeDecision == PolicyDecision.Deny)
            {
                return Invalid(
                    "EXE_DENIED",
                    $"EXE {full} 在黑名单中，已被策略拒绝。");
            }

            if (exeDecision == PolicyDecision.NotAllowlisted)
            {
                return Invalid(
                    "NOT_ALLOWLISTED",
                    $"EXE {full} 不在白名单内。请在配置中加入该可执行文件绝对路径。");
            }

            var result = RunProcess(full, commandArgs, args, out var truncatedExe);
            if (!result.Item1)
            {
                return new ToolResult { Ok = false, ErrorCode = "EXE_FAILED", Data = result.Item2 };
            }

            return new ToolResult { Ok = true, Data = ParseProcessOutput(result.Item2, truncatedExe) };
        }

        private Tuple<bool, string> RunProcess(string fileName, IList<string> arguments, Dictionary<string, object> args, out bool truncated)
        {
            var timeoutMs = ClampTimeout(GetInt(args, "timeout_ms", options.TimeoutSeconds * 1000));
            var outputLimitBytes = GetInt(args, "output_limit_kb", options.OutputLimitKB) * 1024;
            var startInfo = new ProcessStartInfo
            {
                FileName = fileName,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            try
            {
                var isPython = fileName.EndsWith(".py", StringComparison.OrdinalIgnoreCase)
                               || System.IO.Path.GetFileName(fileName).IndexOf("python", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isPython)
                {
                    // Python 子进程默认按控制台代码页(GBK)输出，强制其以 UTF-8 写出，并让 C# 端按 UTF-8 解码，
                    // 保证中文输出不乱码。非 Python 的 CLI/EXE 保持系统 ANSI 解码，避免误伤 GBK 输出。
                    startInfo.StandardOutputEncoding = Encoding.UTF8;
                    startInfo.StandardErrorEncoding = Encoding.UTF8;
                    PythonWorkerHost.PreparePythonChildEnv(startInfo);
                }
            }
            catch
            {
            }
            var workingDir = Convert.ToString(args.TryGetValue("cwd", out var cwd) && cwd != null ? cwd : System.IO.Path.GetTempPath());
            if (!string.IsNullOrWhiteSpace(workingDir))
            {
                startInfo.WorkingDirectory = workingDir;
            }

            if (arguments.Count > 0)
            {
                startInfo.Arguments = string.Join(" ", arguments.Select(QuoteArg));
            }

            using var process = new Process { StartInfo = startInfo };

            truncated = false;
            var output = new StringBuilder();
            var error = new StringBuilder();
            var limit = Math.Max(outputLimitBytes, 1024);
            var counter = new object();
            var truncatedFlag = false;

            process.OutputDataReceived += (_, e) => { if (!string.IsNullOrEmpty(e.Data) && AppendLimited(output, e.Data, limit, counter)) truncatedFlag = true; };
            process.ErrorDataReceived += (_, e) => { if (!string.IsNullOrEmpty(e.Data) && AppendLimited(error, e.Data, limit, counter)) truncatedFlag = true; };

            if (!process.Start())
            {
                truncated = truncatedFlag;
                return Tuple.Create(false, "进程无法启动。");
            }

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            if (!process.WaitForExit(Math.Max(timeoutMs, 1000)))
            {
                try
                {
                    process.Kill();
                }
                catch
                {
                }

                truncated = truncatedFlag;
                return Tuple.Create(false, "执行超时（limit=" + timeoutMs + "ms），进程已终止。");
            }

            process.WaitForExit();
            var exit = process.ExitCode;
            var combined = exit == 0 ? output.ToString() : error.ToString();
            if (string.IsNullOrWhiteSpace(combined) && exit != 0)
            {
                combined = error.ToString() + output.ToString();
            }

            combined = combined.Trim();
            if (combined.Length > limit)
            {
                combined = combined.Substring(0, limit) + "\n" + OutputTruncatedMarker;
                truncatedFlag = true;
            }

            truncated = truncatedFlag;
            if (exit != 0)
            {
                return Tuple.Create(false, $"退出码 {exit}\n{combined}");
            }

            return Tuple.Create(true, combined);
        }

        private static string QuoteArg(string argument)
        {
            if (string.IsNullOrEmpty(argument))
            {
                return "\"\"";
            }

            if (argument.IndexOfAny(new[] { ' ', '"', '\t' }) < 0)
            {
                return argument;
            }

            // Windows CRT argv 转义规则：内嵌引号用双引号转义（与 PythonWorkerHost.Quote 一致）。
            // 原 \" 写法在"反斜杠后接引号"的参数上会触发引号语义翻转（M16 全量审查）。
            return "\"" + argument.Replace("\"", "\"\"") + "\"";
        }

        private static bool AppendLimited(StringBuilder builder, string text, int limit, object counter)
        {
            lock (counter)
            {
                if (builder.Length >= limit)
                {
                    return true; // 本次输出行被丢弃（截断）
                }

                builder.AppendLine(text);
                return false;
            }
        }

        private static object ParseProcessOutput(string output, bool truncated)
        {
            return new
            {
                exit_success = true,
                output = output,
                output_chars = output.Length,
                truncated = truncated
            };
        }

        private string NormalizeVbaModule(string snippet, string macroName)
        {
            var trimmed = snippet.Trim();
            // 统一不启用 Option Explicit：AI/用户生成的片段常省略 Dim 直接使用隐式变量，Option Explicit
            // 会把任何未声明变量判为硬编译错误（Variable not defined），导致整模块（含执行探针）编译失败 →
            // 误报 VBA_SYNTAX_ERROR。改用 VBA 默认宽松类型（未声明变量视为 Variant），显著降低合法片段假阳性。
            // ReplaceModuleCode 每轮只替换过程区（声明区恒为静态辅助声明），片段自带 Option Explicit/模块级声明时
            // 声明区会偏离静态形态，触发整清重建路径，因此历史 Option Explicit 残留不会累积成重复声明。
            if (SnippetDefinesProcedure(trimmed))
            {
                return trimmed;
            }

            // 片段未自带 Sub/Function/Property 声明时，无论是否给了 macroName 一律包裹进 RunSnippet。
            // 旧实现仅在 macroName 为空时才包裹：一旦模型「给了 macroName 却没给过程声明」，
            // 会把裸顶层语句直接 AddFromString → 模块出现模块级可执行语句 → VBE 编译失败 →
            // 误报 VBA_SYNTAX_ERROR。#ParameterizedEntry# 语义：macroName 只用于在已有过程里挑入口，
            // 绝不能让一个「本该被包裹的裸体」跳过包裹。
            return "Public Sub RunSnippet()\n" + WithLineNumbers(trimmed) + "\nEnd Sub";
        }

        /// <summary>片段是否自带过程声明（Sub/Function/Property）。决定了「直接使用」还是「包一层 RunSnippet」。</summary>
        private static bool SnippetDefinesProcedure(string trimmed)
        {
            return !string.IsNullOrWhiteSpace(trimmed) && RegexVbaIsProcedure.IsMatch(trimmed);
        }

        private static string WithLineNumbers(string code)
        {
            return string.Join("\n", code.Split('\n').Select(line => line.Length > 512 ? line.Substring(0, 512) : line));
        }

        private static string ResolveVbaProcedure(dynamic codeModule, string macroName, bool wrappedEntry, string snippetText)
        {
            var name = wrappedEntry
                ? "RunSnippet" // 裸体已被包成 RunSnippet：忽略 macroName，入口固定为 RunSnippet
                : !string.IsNullOrWhiteSpace(macroName)
                    ? macroName.Trim()
                    : FindEntryProcedure(snippetText); // 从用户片段文本解析最后一个过程，不受 VBE 模块级重排影响
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var moduleName = Convert.ToString(codeModule.Parent.Name);
            try
            {
                var count = Convert.ToInt32(codeModule.CountOfLines);
                if (count > 0)
                {
                    return moduleName + "." + name;
                }
            }
            catch
            {
            }

            return null;
        }

        /// <summary>从 VBA 代码文本里提取最后一个声明的过程名（Sub/Function/Property），用于缺省入口解析。
        /// 支持中文/自命名宏（\w 匹配 CJK）；无过程声明返回 null。</summary>
        public static string FindEntryProcedure(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return null;
            }

            Match match = null;
            foreach (Match m in RegexVbaProcedureDecl.Matches(code))
            {
                match = m;
            }

            return match == null ? null : match.Groups["name"].Value;
        }

        /// <summary>静态诊断「最可能导致 VBE 编译失败的写法」，返回首条可行动建议；找不到则返回 null。
        /// 供 run_vba_snippet 在 VBA_SYNTAX_ERROR 时回喂具体原因，让模型一次性改对，而非带着错误代码盲试。
        public static string DetectVbaCompileIssue(string snippet)
        {
            if (string.IsNullOrWhiteSpace(snippet))
            {
                return null;
            }

            // 1) 未闭合字符串字面量：引号不成对 → “预期: 语句结束”类编译错误。
            //    （注：VBA 是 Unicode 语言，中文/非 ASCII 过程名与变量名在中文版 Excel 完全合法，故不在此
            //      误报“中文标识符无法编译”——那只会错怪合法代码、把模型带偏去改名而非修真正的编译错误。）
            var unbalanced = FindUnbalancedQuote(snippet);
            if (unbalanced >= 0)
            {
                return "第 " + unbalanced + " 行存在未闭合的双引号字符串：VBA 字符串必须成对闭合，请补全结尾的 \" 。";
            }

            // 3) 过程声明与 End 不配对（缺少 End Sub/Function/Property）。
            var declared = CountMatches(RegexVbaProcedureDecl, snippet);
            var ended = CountMatches(RegexVbaIsEnd, snippet);
            if (declared > ended)
            {
                return "过程声明比过程结尾多（" + declared + " 个 Sub/Function/Property 但仅 " + ended + " 个 End）：存在未闭合的过程，请为每个过程补全对应的 End Sub/End Function/End Property。";
            }

            return null;
        }

        private static int FindUnbalancedQuote(string code)
        {
            var lines = code.Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                var inStr = false;
                foreach (var ch in lines[i])
                {
                    if (ch == '"')
                    {
                        inStr = !inStr;
                    }
                }

                if (inStr)
                {
                    return i + 1;
                }
            }

            return -1;
        }

        private static int CountMatches(System.Text.RegularExpressions.Regex rx, string text)
        {
            return rx.Matches(text ?? string.Empty).Count;
        }

        /// <summary>执行前硬闸门：返回非 null 表示片段被静态校验判为必失败（存在 error 级问题），已拦截未执行。
        /// 公开供单元测试直接断言阻断效果。</summary>
        public static string PreRunVbaGate(string snippet)
        {
            var gate = StaticVbaValidate(snippet);
            var gateData = gate.Ok ? gate.Data : null;
            if (gateData == null)
            {
                return null;
            }

            var gj = gateData as Newtonsoft.Json.Linq.JObject
                     ?? Newtonsoft.Json.Linq.JObject.FromObject(gateData);
            var validTok = gj["valid"];
            var block = validTok != null
                        && validTok.Type == Newtonsoft.Json.Linq.JTokenType.Boolean
                        && !Convert.ToBoolean(((Newtonsoft.Json.Linq.JValue)validTok).Value);
            if (!block)
            {
                return null;
            }

            var errMsgs = (gj["issues"] as Newtonsoft.Json.Linq.JArray)
                ?.Where(i => string.Equals(Convert.ToString(i["level"]), "error", StringComparison.OrdinalIgnoreCase))
                .Select(i => Convert.ToString(i["message"]) + " " + Convert.ToString(i["hint"]))
                .ToArray();
            return "VBA 代码被静态校验拦截（未执行），存在必失败写法："
                   + string.Join("；", errMsgs ?? new string[0])
                   + "。请按建议重写后再执行，或改用内置单元格/格式/流程工具。";
        }

        private static ToolResult StaticVbaValidate(string snippet)
        {
            var issues = new List<object>();
            var lines = (snippet ?? string.Empty).Split('\n');

            // 1) 未闭合的字符串字面量（引号成对检测）。
            if (lines.Sum(l => l.Count(ch => ch == '"')) % 2 != 0)
            {
                issues.Add(new
                {
                    level = "error",
                    line = FirstLineWithOddQuote(lines),
                    message = "检测到未闭合的字符串：字符串字面量需成对使用双引号，请补齐补足的 \"。",
                    hint = "例如 Dim s As String: s = \"完成\" 中的引号必须成对。"
                });
            }

            // 2) End Sub/Function 与过程声明不配对。
            var openings = 0;
            foreach (var item in lines)
            {
                if (RegexVbaIsProcedure.IsMatch(item))
                {
                    openings++;
                }
                else if (RegexVbaIsEnd.IsMatch(item))
                {
                    openings--;
                }
            }

            if (openings > 0)
            {
                issues.Add(new
                {
                    level = "error",
                    line = 0,
                    message = $"有 {openings} 个过程缺少对应的 End Sub/End Function，代码不会被视为完整过程，调不到入口。",
                    hint = "在每个 Sub/Function 的结尾补上 End Sub / End Function。"
                });
            }
            else if (openings < 0)
            {
                issues.Add(new
                {
                    level = "error",
                    line = 0,
                    message = "存在多余的 End Sub/End Function，请核对每个过程是否成对闭合。"
                });
            }

            // 3) 完全没有过程声明。
            var hasProc = RegexVbaIsProcedure.IsMatch(snippet ?? string.Empty);
            if (!hasProc)
            {
                issues.Add(new
                {
                    level = "warning",
                    line = 0,
                    message = "未检测到 Sub/Function 过程声明，执行器将自动包裹为 RunSnippet 执行。",
                    hint = "建议显式写成 Sub 名称() ... End Sub 以便指定入口与复用；不写也会被自动包裹执行。"
                });
            }

            // 4) 性能隐患：循环内逐格访问单元格，应改为数组批量处理。
            var slowPattern = new[]
            {
                new { desc = "Cells(...)", pro = "在循环里用 Cells()/Range().Value 单格读写", rx = @"\bFor\b[\s\S]{0,800}?Cells\(" },
                new { desc = ".Select/.Activate", pro = "使用了 .Select/.Activate，多数情况下可省并拖慢执行", rx = "\\.Select\\b|\\.Activate\\b" }
            }.ToList();
            foreach (var slow in slowPattern)
            {
                if (new System.Text.RegularExpressions.Regex(slow.rx, System.Text.RegularExpressions.RegexOptions.IgnoreCase).IsMatch(snippet ?? string.Empty))
                {
                    issues.Add(new
                    {
                        level = "suggestion",
                        line = 0,
                        message = slow.pro + "：批量读写请改用数组或 Range.Value2 整体赋值，避免逐单元格交互。",
                        hint = "先 arr = rng.Value2，处理数组，再把结果 rng.Value2 = arr 一次性写回。"
                    });
                }
            }

            // valid 仅由 error 级问题决定：无过程声明只是 warning（NormalizeVbaModule 会包裹为 RunSnippet），
            // 不应阻断——否则裸代码（run_vba_snippet 最常见用法）被 PreRunVbaGate 误拦，与包裹逻辑自相矛盾。
            var valid = issues.All(i => i.GetType().GetProperty("level").GetValue(i).ToString() != "error");

            // 4.5) 对跨行跨列的矩形 Range 直接调用 .AutoFit 必抛运行时错误 1004（.AutoFit 仅适用于单行/单列、
            // 整行/整列或 Rows/Columns 集合）。把这种确定失败的写法判为 error（阻断），提示改写；
            // 安全写法（.Range(...).Rows/Columns(...).AutoFit、.EntireRow/.EntireColumn、单格/单行/单列）仅作 suggestion 提醒。
            AddVbaAutoFitIssues(snippet, issues);

            // 5) 常见非法属性/危险写法：Range 没有 Alignment（会抛运行时错误 438），居中/格式化不得清空数据。
            var badPatterns = new[]
            {
                new { desc = "Range.Alignment", rx = @"\.\s*Alignment\s*=", hint = "Range 对象没有 Alignment 属性，必抛运行时错误 438：水平对齐用 .HorizontalAlignment，垂直对齐用 .VerticalAlignment。" },
                new { desc = "格式化时清空数据", rx = @"\.\s*Clear\b|\.\s*ClearContents\b|\.\s*ClearFormats\b|Rows\s*\([^)]*\)\s*\.\s*Delete\b|\.\s*EntireRow\s*\.\s*Delete\b", hint = "仅居中/对齐/字体/边框等外观操作时，严禁 Clear/ClearContents/Delete，会清空单元格数据。" }
            }.ToList();
            foreach (var bad in badPatterns)
            {
                if (new System.Text.RegularExpressions.Regex(bad.rx, System.Text.RegularExpressions.RegexOptions.IgnoreCase).IsMatch(snippet ?? string.Empty))
                {
                    issues.Add(new
                    {
                        level = "error",
                        line = 0,
                        message = "检测到错误写法/危险操作：" + bad.desc + "。",
                        hint = bad.hint
                    });
                }
            }

            valid = issues.All(i => i.GetType().GetProperty("level").GetValue(i).ToString() != "error");
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    validation_mode = "static",
                    valid = valid,
                    has_procedure = hasProc,
                    lines = lines.Length,
                    issues,
                    note = "未能访问 VBA 工程，执行静态检查。请优先打开下列可操作的修正建议。"
                }
            };
        }

        private static int FirstLineWithOddQuote(string[] lines)
        {
            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].Count(ch => ch == '"') % 2 != 0)
                {
                    return i + 1;
                }
            }

            return 0;
        }

        private void RemoveComponent(dynamic vbProject, dynamic module)
        {
            if (module == null)
            {
                return;
            }

            try
            {
                vbProject.VBComponents.Remove(module);
            }
            catch
            {
            }
        }

        private string WriteTempScript(string code, string extension)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                throw new ArgumentException("缺少必填参数：code。");
            }

            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "xai_" + Guid.NewGuid().ToString("N") + "." + extension);
            System.IO.File.WriteAllText(path, code, Encoding.UTF8);
            return path;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static List<string> ArgsList(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var raw) && raw is IEnumerable<object> list
                ? list.Select(item => Convert.ToString(item)).ToList()
                : new List<string>();
        }

        private static ToolResult Disabled(string tool)
        {
            return Invalid("DISABLED", $"{tool} 属于高危受控执行工具，默认关闭。请在设置中启用受控执行并配置白名单后再使用。");
        }

        private static readonly System.Text.RegularExpressions.Regex RegexVbaIsProcedure =
            new System.Text.RegularExpressions.Regex(
                @"^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function|Property)\s+\w+",
                System.Text.RegularExpressions.RegexOptions.Multiline);

        private static readonly System.Text.RegularExpressions.Regex RegexVbaIsEnd =
            new System.Text.RegularExpressions.Regex(
                @"^\s*End\s+(Sub|Function|Property)\s*$",
                System.Text.RegularExpressions.RegexOptions.Multiline);

        /// <summary>捕获所有 `?Range(...)...AutoFit` 调用：接收者前缀「可选」，链条记录 `(:)` 与 `)` 之间到
        /// `.AutoFit` 之前的中缀成员（如 `.Rows("1:3")`、`.Offset(1)`），用于精确区分“直接对矩形调 AutoFit(必错)”
        /// 与“安全改写(Rows/Columns/EntireRow/EntireColumn)”。匹配其对两个源码片段进行矩形判定。</summary>
        public static readonly System.Text.RegularExpressions.Regex RegexRangeAutoFitM =
            new System.Text.RegularExpressions.Regex(
                @"(?:[\w$]+\.|\b)Range\s*\(\s*(?<addr>[^()]*?)\s*\)\s*(?<chain>(?:\.[A-Za-z_]\w*(?:\s*\([^)]*\))?\s*)*)\s*\.AutoFit\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        /// <summary>公开兼容旧名（供测试/前端引用）：只要片段解析出“直接对跨行列矩形 Range 调 .AutoFit”即命中。
        /// 底层委托精确判定，避免旧正则误伤安全改写。</summary>
        public static readonly System.Text.RegularExpressions.Regex RegexVbaAutoFit2D =
            new System.Text.RegularExpressions.Regex(
                @"\.Range\s*\(.*?\)\s*(?:\.[^.]+)*\s*\.AutoFit\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        /// <summary>每个单元格地址片段（去掉 $ 前缀）按「X列Y行」拆分：完整表列引用（如 G、K）行号为全表，
        /// 不参与“跨行判定”。</summary>
        private static readonly System.Text.RegularExpressions.Regex RegexCellAddr =
            new System.Text.RegularExpressions.Regex(
                @"^\$?(?<col>[A-Z]+)\$?(?<row>\d+)$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        /// <summary>在 `Issues` 中补充 .AutoFit 相关的校验结果：对跨行跨列的矩形 Range 直接调 .AutoFit 判为
        /// error(阻断)，安全改写与单格/单行/单列仅作 suggestion。核心是“矩形块 = 多列 且 多行”（如 G12:K18）。</summary>
        private static void AddVbaAutoFitIssues(string snippet, List<object> issues)
        {
            if (string.IsNullOrWhiteSpace(snippet))
            {
                return;
            }

            var sawAutoFit = false;
            foreach (System.Text.RegularExpressions.Match m in RegexRangeAutoFitM.Matches(snippet))
            {
                sawAutoFit = true;
                var chain = m.Groups["chain"].Value;
                if (RegexSafeAutoFitChain.IsMatch(chain))
                {
                    continue; // .Rows(...)/.Columns(...)/.EntireRow/.EntireColumn → 安全改写，不视为必错
                }

                var addr = m.Groups["addr"].Value.Trim().Trim('"', '\'', ' ');
                if (IsRectCellRange(addr))
                {
                    issues.Add(new
                    {
                        level = "error",
                        line = 0,
                        message = "检测到对跨行跨列的矩形 Range 直接调用 .AutoFit：运行时会抛错误 1004 中断整个宏。请改为按行/列分别调用或设 RowHeight/ColumnWidth=0。",
                        hint = "非法：ws.Range('G12:K18').AutoFit。改：ws.Rows('12:18').RowHeight = 0: ws.Columns('G:K').ColumnWidth = 0（或 .Rows.AutoFit / .Columns.AutoFit）。"
                    });
                    return;
                }
            }

            if (sawAutoFit)
            {
                issues.Add(new
                {
                    level = "suggestion",
                    line = 0,
                    message = "检测到 .AutoFit：请确认接收对象是单行/单列 Range、整行/整列或 Rows/Columns 集合；对跨行跨列矩形 Range 直接调用会抛运行时错误 1004。",
                    hint = "安全写法：ws.Rows('12:18').AutoFit: ws.Columns('G:K').AutoFit，或设 RowHeight/ColumnWidth=0。"
                });
            }
        }

        /// <summary>匹配 `.AutoFit` 之前的中缀安全改写链：`.Rows(...)`/`.Columns(...)`/`.EntireRow`/`.EntireColumn`。</summary>
        private static readonly System.Text.RegularExpressions.Regex RegexSafeAutoFitChain =
            new System.Text.RegularExpressions.Regex(
                @"\.(?:Rows|Columns|EntireRow|EntireColumn)\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        /// <summary>是否“跨列 且 跨行”的单元格矩形（如 G12:K18、A1:C3）：两者都跨才算矩形块→AutoFit 必错；
        /// 单列多行(A1:A5)、单行多列(A1:E1)、整行(2:2)、完整表列(A:C)均不是矩形块（AutoFit 合法）。</summary>
        private static bool IsRectCellRange(string address)
        {
            if (string.IsNullOrWhiteSpace(address) || address.IndexOf(':') < 0)
            {
                return false; // 单格或非单元格地址（命名区域等）无法确定，不阻断
            }

            var parts = address.Split(':');
            if (parts.Length != 2)
            {
                return false;
            }

            var a = RegexCellAddr.Match(parts[0]);
            var b = RegexCellAddr.Match(parts[1]);
            if (!a.Success || !b.Success)
            {
                return false; // 含行段/列段（如 2:12、A:C）→ 整行/整列，不判为矩形块
            }

            var rowA = Convert.ToInt32(a.Groups["row"].Value);
            var rowB = Convert.ToInt32(b.Groups["row"].Value);
            var colA = a.Groups["col"].Value.ToUpperInvariant();
            var colB = b.Groups["col"].Value.ToUpperInvariant();
            var multiRow = rowA != rowB;
            var multiCol = colA != colB;
            return multiRow && multiCol;
        }

        private static readonly System.Text.RegularExpressions.Regex RegexVbaProcedureDecl =
            new System.Text.RegularExpressions.Regex(
                @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property)\s+(?<name>\w+)",
                System.Text.RegularExpressions.RegexOptions.Multiline);

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            if (value is Newtonsoft.Json.Linq.JValue jv && jv.Value != null)
            {
                return Convert.ToBoolean(jv.Value);
            }

            return Convert.ToBoolean(value);
        }

        /// <summary>VBA args 注入模块源码的字符上限：避免超大 args 撑爆 AddFromString / 模块体积。</summary>
        private const int VbaArgsLimitChars = 64 * 1024;

        /// <summary>把 run_vba_snippet 的 args 归一化为单行 JSON 字符串（缺省返回 "{}"）。
        /// 支持模型传 JSON 字符串或对象；非法 JSON 字符串降级为字符串值。</summary>
        internal static string NormalizeArgsJson(object raw)
        {
            if (raw == null)
            {
                return "{}";
            }

            Newtonsoft.Json.Linq.JToken token;
            if (raw is Newtonsoft.Json.Linq.JToken jt)
            {
                token = jt;
            }
            else if (raw is string s)
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    return "{}";
                }

                try
                {
                    token = Newtonsoft.Json.Linq.JToken.Parse(s);
                }
                catch
                {
                    token = new Newtonsoft.Json.Linq.JValue(s);
                }
            }
            else
            {
                try
                {
                    token = Newtonsoft.Json.Linq.JToken.FromObject(raw);
                }
                catch
                {
                    token = new Newtonsoft.Json.Linq.JValue(Convert.ToString(raw));
                }
            }

            return token.ToString(Newtonsoft.Json.Formatting.None);
        }

        /// <summary>归一化 run_python_code 的 input_data 并施加 B3 护栏：超过 10 万 cell 或 8 MB 序列化体积时
        /// 返回错误信息（由调用方回 Invalid("INPUT_TOO_LARGE")）；无 input_data 时返回 (null, null)。</summary>
        internal static (Newtonsoft.Json.Linq.JToken Token, string Error) NormalizeInputData(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("input_data", out var raw) || raw == null)
            {
                return (null, null);
            }

            Newtonsoft.Json.Linq.JToken token;
            if (raw is string s)
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    return (null, null);
                }

                try
                {
                    token = Newtonsoft.Json.Linq.JToken.Parse(s);
                }
                catch
                {
                    token = new Newtonsoft.Json.Linq.JValue(s);
                }
            }
            else
            {
                try
                {
                    token = Newtonsoft.Json.Linq.JToken.FromObject(raw);
                }
                catch
                {
                    return (null, "input_data 无法序列化为 JSON。");
                }
            }

            if (CountCells(token) > 100000)
            {
                return (null, "input_data 数据量过大（超过 10 万 cell 上限），请缩小输入范围或分批处理。");
            }

            var json = token.ToString(Newtonsoft.Json.Formatting.None);
            if (json.Length > 8 * 1024 * 1024)
            {
                return (null, "input_data 序列化后体积过大（超过 8 MB 上限），请缩小输入。");
            }

            return (token, null);
        }

        /// <summary>递归统计 JSON 中的标量 cell 数量（数组元素/对象属性值各计 1），用于 input_data 上限护栏。</summary>
        internal static int CountCells(Newtonsoft.Json.Linq.JToken token)
        {
            switch (token.Type)
            {
                case Newtonsoft.Json.Linq.JTokenType.Array:
                    var sum = 0;
                    foreach (var item in (Newtonsoft.Json.Linq.JArray)token)
                    {
                        sum += CountCells(item);
                    }

                    return sum;
                case Newtonsoft.Json.Linq.JTokenType.Object:
                    var objSum = 0;
                    foreach (var prop in (Newtonsoft.Json.Linq.JObject)token)
                    {
                        objSum += CountCells(prop.Value);
                    }

                    return objSum;
                default:
                    return 1;
            }
        }

        /// <summary>把代码执行工具的超时限定在 [1s, 300s]，防止模型传入超大 timeout_ms 把常驻 worker/进程钉死数小时。</summary>
        private static int ClampTimeout(int ms)
        {
            return Math.Max(1000, Math.Min(ms, 300000));
        }

        private static void Release(object value)
        {
            if (value == null)
            {
                return;
            }

            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(value))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
                }
            }
            catch
            {
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        /// <summary>草稿模块整体未能编译（探针过程连续两次不可用）时抛出的专用异常：调用方据其把
        /// “必然失败的代码”归为 VBA_SYNTAX_ERROR，而非把晦涩的 HRESULT 0x800A9C68 当运行时错误回喂。</summary>
        private sealed class VbaModuleUnavailableException : Exception
        {
            public VbaModuleUnavailableException(string message) : base(message) { }

            public static VbaModuleUnavailableException From(Exception inner, string probeFq)
            {
                return new VbaModuleUnavailableException(
                    "草稿模块未能编译，探针过程不可用：" + probeFq + "（" + inner.Message + "）");
            }
        }
    }
}