using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// run_python_excel：在受控 Python 中直接读写当前 Excel 活动工作簿（xlwings 主 API，pywin32 兜底）。
    /// 复用 PythonWorkerHost 常驻 worker；每次请求把「import xlwings as xw」前置连接代码与用户代码拼接后执行，
    /// 连接上下文逐请求重建，Excel 重启/worker 崩溃后自然重连重建。
    ///
    /// 安全契约（2.0.0.0 需求手册 3.3 / 四）：
    /// - Dangerous + 确认门：工具注册为 Dangerous，每次调用都过原生确认框。
    /// - 版本/位数校验：python 探测要求 ≥3.9 且位数与 Excel（进程位数）一致，不符给出安装指引，不撞 COM 失败。
    /// - 依赖自动安装：xlwings/pywin32 缺失时 pip 安装（清华/阿里镜像兜底 → 系统目录失败退 --user → 装后 import 自检）。
    /// - 写回纪律：AI 声明 write_ranges 时执行前快照该区域，执行失败自动回滚并回报。
    /// </summary>
    public sealed class PythonExcelService
    {
        private readonly dynamic application;
        private readonly ScriptOptions options;

        public PythonExcelService(dynamic application, ScriptOptions options)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.options = options ?? new ScriptOptions();
        }

        public ToolResult Run(Dictionary<string, object> args)
        {
            var code = OptionalString(args, "code");
            if (options.Enabled == false)
            {
                return Invalid("DISABLED", "run_python_excel 已被受控执行总开关禁用。");
            }

            if (string.IsNullOrWhiteSpace(code))
            {
                return Invalid("MISSING_CODE", "缺少要执行的 Python 代码（code）。");
            }

            var execFile = string.IsNullOrWhiteSpace(options.PythonPath) ? "python" : options.PythonPath;
            var snapshots = new List<RangeSnapshot>();
            try
            {
                // B3 入参护栏：input_data 超上限显式拒绝。
                var (inputData, inputError) = ScriptRunnerService.NormalizeInputData(args);
                if (inputError != null)
                {
                    return Invalid("INPUT_TOO_LARGE", inputError);
                }

                // C4 写回纪律：声明 write_ranges → 执行前快照；执行失败自动回滚。
                var declared = ParseWriteRanges(args);
                if (!declared.Ok)
                {
                    return Invalid("WRITE_RANGE_INVALID", declared.Error);
                }

                foreach (var wr in declared.Ranges)
                {
                    snapshots.Add(Snapshot(wr));
                }

                // 版本/位数/依赖探测（一次 python -c 拿全部信息）。
                var probe = Probe(execFile);
                if (!probe.Found)
                {
                    return Invalid("PYTHON_NOT_FOUND", "未找到 Python（" + execFile + "），无法执行 run_python_excel。" + PyInstallGuide
                        + (string.IsNullOrEmpty(probe.Raw) ? "" : "\n探测详情：" + probe.Raw));
                }

                if (!probe.ParseOk)
                {
                    return Invalid("PYTHON_PROBE_FAILED", "Python 探测失败，无法判断版本/位数/依赖：\n" + (probe.Raw ?? ""));
                }

                var excelBits = IntPtr.Size * 8;
                if (!VersionAtLeast39(probe.Version))
                {
                    return Invalid("PYTHON_VERSION_MISMATCH",
                        "当前 Python " + probe.Version + "（" + probe.Bits + " 位）低于 3.9。请安装 64 位 Python 3.10 或 3.11 并勾选 Add to PATH，再重试。" + PyInstallGuide);
                }

                if (probe.Bits != excelBits)
                {
                    return Invalid("PYTHON_BITNESS_MISMATCH",
                        "Python 为 " + probe.Bits + " 位，与当前 Excel（" + excelBits + " 位）不匹配，COM 直连无法工作。请安装与 Excel 同位的 64 位 Python 3.10/3.11 并勾选 Add to PATH。" + PyInstallGuide);
                }

                // C5 依赖自动安装：缺失即装，幂等（已装跳过）。工具已过确认门，描述明示安装行为。
                if (!probe.Xlwings || !probe.Win32com)
                {
                    var missing = new List<string>();
                    if (!probe.Xlwings) missing.Add("xlwings");
                    if (!probe.Win32com) missing.Add("pywin32");
                    var install = InstallDeps(execFile);
                    if (!install.Item1)
                    {
                        return Invalid("PYTHON_DEPS_INSTALL_FAILED", "缺少依赖 " + string.Join("、", missing) + "，自动安装失败：" + install.Item2);
                    }

                    // 依赖刚装完：丢弃常驻 worker，避免旧 worker 在 pywin32 刚就绪的瞬间 import xlwings
                    // 缓存「引擎未注册」失败态，导致后续一直报误导性的缺 pywin32。
                    PythonWorkerHost.ResetWorker();
                }

                // C7 worker 复用：常驻 worker 执行，前置连接代码逐请求重建连接上下文。
                var fullCode = PrependConnectPreamble(code);
                var timeoutMs = Clamp(GetInt(args, "timeout_ms", options.TimeoutSeconds * 1000));
                var outputLimitBytes = Math.Max(GetInt(args, "output_limit_kb", options.OutputLimitKB), 1) * 1024;
                var resultLimitBytes = Math.Max(GetInt(args, "result_limit_kb", 1024), 1) * 1024;
                // P3-D1：persist 保留命名空间跨调用；state_clear 显式清空持久状态。
                var execOptions = new PythonWorkerHost.ExecuteOptions
                {
                    Persist = GetBool(args, "persist", false),
                    StateClear = GetBool(args, "state_clear", false)
                };

                var workerOk = PythonWorkerHost.TryExecute(
                    fullCode,
                    execFile,
                    timeoutMs,
                    outputLimitBytes,
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
                    // 缺模块自愈（与 run_python_code 同口径）：ModuleNotFoundError 命中安全白名单 →
                    // pip 装包（镜像兜底）→ 重建 worker（丢弃旧导入缓存）→ 重试一次；重试失败才回滚声明区域。
                    // Python 代码抛异常时 TryExecute 返回 workerOk=false，自愈判断必须覆盖该分支（真机实证）。
                    if (ScriptRunnerService.TryResolveMissingModule(output, out var missingModule, out var pipPackage))
                    {
                        var installOk = InstallPythonModule(execFile, pipPackage, out var installMsg);
                        if (installOk)
                        {
                            PythonWorkerHost.ResetWorker();
                            var retryOk = PythonWorkerHost.TryExecute(
                                fullCode,
                                execFile,
                                timeoutMs,
                                outputLimitBytes,
                                inputData,
                                resultLimitBytes,
                                execOptions,
                                out var ok2,
                                out var output2,
                                out var result2,
                                out var rTrunc2,
                                out var sReset2,
                                out var exit2);
                            if (retryOk && ok2)
                            {
                                AppLog.Info("python", $"run_python_excel 缺模块自愈成功：{missingModule} → pip install {pipPackage} 后重试通过");
                                return new ToolResult
                                {
                                    Ok = true,
                                    Data = new
                                    {
                                        exit_success = exit2 == 0,
                                        exit_code = exit2,
                                        output = output2,
                                        output_chars = output2 == null ? 0 : output2.Length,
                                        truncated = output2 != null && output2.IndexOf(ScriptRunnerService.OutputTruncatedMarker, StringComparison.Ordinal) >= 0,
                                        result = result2,
                                        result_truncated = rTrunc2,
                                        state_reset = sReset2,
                                        write_ranges = declared.Ranges.ConvertAll(r => new { sheet = r.Sheet, address = r.Address }),
                                        write_ranges_snapshotted = snapshots.Count,
                                        deps_auto_installed = pipPackage
                                    }
                                };
                            }

                            if (retryOk)
                            {
                                output = output2;
                            }
                        }
                        else
                        {
                            output = output + "\n[自愈尝试] 检测到缺模块 " + missingModule + "，自动 pip install " + pipPackage + " 失败：" + installMsg;
                        }
                    }

                    Rollback(snapshots);
                    return Invalid("PYTHON_RUN_ERROR",
                        (workerOk ? "Python 执行失败：" : "") + output
                        + (snapshots.Count > 0 ? "\n已自动回滚声明区域 write_ranges。" : "")
                        + ClassifyConnectionError(output));
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        exit_success = ok && exit == 0,
                        exit_code = exit,
                        output = output,
                        output_chars = output == null ? 0 : output.Length,
                        truncated = output != null && output.IndexOf(ScriptRunnerService.OutputTruncatedMarker, StringComparison.Ordinal) >= 0,
                        result = result,
                        result_truncated = resultTruncated,
                        state_reset = stateReset,
                        write_ranges = declared.Ranges.ConvertAll(r => new { sheet = r.Sheet, address = r.Address }),
                        write_ranges_snapshotted = snapshots.Count
                    }
                };
            }
            catch (Exception ex)
            {
                Rollback(snapshots);
                return Invalid("PYTHON_EXCEL_ERROR", "run_python_excel 失败：" + ex.Message
                    + (snapshots.Count > 0 ? "\n已自动回滚声明区域 write_ranges。" : ""));
            }
        }

        private sealed class RangeSpec
        {
            public string Sheet;
            public string Address;
        }

        private sealed class RangeSnapshot
        {
            public string Sheet;
            public string Address;
            public object Value;
        }

        private const string PyInstallGuide =
            " 安装指引：到 python.org 下载 64 位 Python 3.10 或 3.11 安装包，安装时务必勾选 Add Python to PATH；"
            + "如已安装请在设置里核对 Python 路径与位数。";

        /// <summary>解析 write_ranges：元素可为字符串（Sheet!A1:B5 或 A1:B5，缺省活动表）或 {sheet,address} 对象。</summary>
        private static (bool Ok, string Error, List<RangeSpec> Ranges) ParseWriteRanges(Dictionary<string, object> args)
        {
            var list = new List<RangeSpec>();
            if (!args.TryGetValue("write_ranges", out var raw) || raw == null)
            {
                return (true, null, list);
            }

            var arr = raw as System.Collections.IEnumerable ?? new object[] { raw };
            foreach (var rawItem in arr)
            {
                if (rawItem == null)
                {
                    continue;
                }

                // JSON 反序列化产物归一化：字符串字面量到 C# 侧是 JValue 而非 System.String，
                // 不先归一化会误入 JObject 分支并把合法地址报成「格式无效」（真机实证 Sheet1!A52:C60）。
                var item = (object)rawItem;
                if (item is Newtonsoft.Json.Linq.JValue jv && jv.Type == Newtonsoft.Json.Linq.JTokenType.String)
                {
                    item = Convert.ToString(jv.Value);
                }

                if (item is string s)
                {
                    var spec = SplitRangeAddress(s.Trim());
                    if (spec == null)
                    {
                        return (false, "write_ranges 含无效地址：" + s, null);
                    }

                    list.Add(spec);
                    continue;
                }

                try
                {
                    var jo = item is JObject j ? j : JObject.FromObject(item);
                    var sheet = Convert.ToString(jo["sheet"]) ?? string.Empty;
                    var address = Convert.ToString(jo["address"]);
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        return (false, "write_ranges 项缺少 address。", null);
                    }

                    list.Add(new RangeSpec { Sheet = string.IsNullOrWhiteSpace(sheet) ? null : sheet.Trim(), Address = address.Trim() });
                }
                catch (Exception)
                {
                    AppLog.Warn("PythonExcelService", "[LogCover] PythonExcelService.cs:298  silent-catch covered (no exception variable)");
                    return (false, "write_ranges 项格式无效：" + Convert.ToString(item)
                        + "。合法格式：\"Sheet1!A52:C60\"（带表名）、\"A52:C60\"（活动表）或 {\"sheet\":\"Sheet1\",\"address\":\"A52:C60\"}。", null);
                }
            }

            return (true, null, list);
        }

        /// <summary>把 "Sheet1!A1:B5" / "A1:B5" 拆为 {Sheet, Address}；缺省表名用 null（活动表）。</summary>
        private static RangeSpec SplitRangeAddress(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var bang = text.IndexOf('!');
            if (bang >= 0)
            {
                var sheet = text.Substring(0, bang).Trim().Trim('\'', '"');
                var addr = text.Substring(bang + 1).Trim();
                if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(addr))
                {
                    return null;
                }

                return new RangeSpec { Sheet = sheet, Address = addr };
            }

            return new RangeSpec { Sheet = null, Address = text.Trim() };
        }

        /// <summary>解析后返回内部可复用的 RangeSpec 列表（供测试断言地址拆分）。</summary>
        internal static List<object> SplitRangeAddressList(IEnumerable<string> texts)
        {
            var list = new List<object>();
            foreach (var t in texts)
            {
                var s = SplitRangeAddress(t);
                list.Add(new { sheet = s?.Sheet, address = s?.Address });
            }

            return list;
        }

        /// <summary>测试钩子：复用真实 ParseWriteRanges 逻辑，返回可断言的匿名结果（避免重复实现导致语义漂移）。</summary>
        internal static (bool Ok, string Error, List<object> Ranges) ParseWriteRangesForTest(Dictionary<string, object> args)
        {
            var parsed = ParseWriteRanges(args);
            if (!parsed.Ok)
            {
                return (false, parsed.Error, null);
            }

            return (true, null, parsed.Ranges.ConvertAll(r => (object)new { sheet = r.Sheet, address = r.Address }));
        }

        /// <summary>前置连接代码：每次请求重建 xlwings 连接上下文（import xlwings as xw）。</summary>
        internal static string PrependConnectPreamble(string code)
        {
            return "import xlwings as xw\n" + (code ?? string.Empty);
        }

        private RangeSnapshot Snapshot(RangeSpec spec)
        {
            dynamic sheet = ResolveSheet(spec.Sheet);
            dynamic range = null;
            try
            {
                range = sheet.Range[spec.Address];
                var value = range.Value2;
                return new RangeSnapshot { Sheet = spec.Sheet, Address = spec.Address, Value = value };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        private void Rollback(List<RangeSnapshot> snapshots)
        {
            if (snapshots == null || snapshots.Count == 0)
            {
                return;
            }

            // 日志全覆盖批次：回滚是关键安全事件，开始/失败都必须留痕（失败静默曾是排查盲区）。
            AppLog.Warn("python", $"write_ranges 回滚开始：{snapshots.Count} 个区域");

            foreach (var snap in snapshots)
            {
                try
                {
                    dynamic sheet = ResolveSheet(snap.Sheet);
                    dynamic range = null;
                    try
                    {
                        range = sheet.Range[snap.Address];
                        range.Value2 = snap.Value;
                    }
                    finally
                    {
                        Release((object)range);
                        Release((object)sheet);
                    }
                }
                catch (Exception rbEx)
                {
                    // 回滚尽力而为，失败不掩盖原始错误；但失败本身必须留痕供事后定位。
                    AppLog.Error("python", $"回滚区域 {snap.Sheet}!{snap.Address} 失败：" + rbEx.Message);
                }
            }
        }

        private dynamic ResolveSheet(string sheetName)
        {
            var wb = application.ActiveWorkbook;
            try
            {
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    return application.ActiveSheet;
                }

                return wb.Worksheets[sheetName];
            }
            finally
            {
                if (wb != null)
                {
                    Release((object)wb);
                }
            }
        }

        /// <summary>一次 python -c 探测：版本/位数/xlwings/win32com 是否可导入。Found=false 表示 python 不可用。</summary>
        internal ProbeResult Probe(string pythonPath)
        {
            var script = "import sys,struct,importlib.util,json;"
                         + "print(json.dumps({'v':'%d.%d'%(sys.version_info.major,sys.version_info.minor),"
                         + "'b':struct.calcsize('P')*8,"
                         + "'x':importlib.util.find_spec('xlwings') is not None,"
                         + "'w':importlib.util.find_spec('win32com') is not None}))";
            var (found, exit, output, err) = RunCapture(pythonPath, new[] { "-c", script }, 15000);
            var pr = new ProbeResult { Found = found };
            if (!found)
            {
                pr.Raw = output + err;
                AppLog.Warn("python", $"探测未找到 python（{pythonPath}）：" + (pr.Raw ?? "").Trim());
                return pr;
            }

            if (exit != 0)
            {
                pr.Raw = output + err;
                AppLog.Warn("python", $"探测 exit={exit}（{pythonPath}）：" + (pr.Raw ?? "").Trim());
                return pr;
            }

            try
            {
                var jo = JObject.Parse(output.Trim());
                pr.Version = Convert.ToString(jo["v"]);
                pr.Bits = Convert.ToInt32(jo["b"]);
                pr.Xlwings = Convert.ToBoolean(jo["x"]);
                pr.Win32com = Convert.ToBoolean(jo["w"]);
                pr.ParseOk = true;
            }
            catch (Exception ex)
            {
                AppLog.Warn("PythonExcelService", $"[LogCover] PythonExcelService.cs:469  {ex.GetType().Name}: {ex.Message}");
                pr.Raw = output + "\n" + err + "\n" + ex.Message;
            }

            return pr;
        }

        internal sealed class ProbeResult
        {
            public bool Found;
            public bool ParseOk;
            public string Version;
            public int Bits;
            public bool Xlwings;
            public bool Win32com;
            public string Raw;
        }

        internal static bool VersionAtLeast39(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
            {
                return false;
            }

            var parts = version.Split('.');
            if (parts.Length < 2 || !int.TryParse(parts[0], out var major) || !int.TryParse(parts[1], out var minor))
            {
                return false;
            }

            return major > 3 || (major == 3 && minor >= 9);
        }

        /// <summary>pip 安装单个白名单包（缺模块自愈用）：worker subprocess 执行，清华→阿里→默认源兜底，失败退 --user。</summary>
        private bool InstallPythonModule(string pythonPath, string pipPackage, out string message)
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

        /// <summary>pip 安装 xlwings+pywin32：worker subprocess 执行（环境健康实证），清华→阿里→默认源兜底，装后 import 自检。</summary>
        private (bool Ok, string Message) InstallDeps(string pythonPath)
        {
            var probeScript = "import importlib.util,sys;sys.exit(0 if (importlib.util.find_spec('xlwings') and importlib.util.find_spec('win32com')) else 1)";
            var mirrors = new[] { "https://pypi.tuna.tsinghua.edu.cn/simple", "https://mirrors.aliyun.com/pypi/simple/", null };
            var lastCombined = "";
            foreach (var mirror in mirrors)
            {
                if (PythonWorkerHost.TryPipInstall(pythonPath, "xlwings pywin32", mirror, false, 280000, out var outText))
                {
                    AppLog.Info("python-env", $"InstallDeps ok=true mirror={(mirror ?? "default")}");
                    // 装后自检：import 两个库均成功才算装好。
                    var (f2, e2, o2, r2) = RunCapture(pythonPath, new[] { "-c", probeScript }, 20000);
                    if (f2 && e2 == 0)
                    {
                        return (true, "已通过 pip 安装 xlwings/pywin32 并自检通过（源：" + (mirror ?? "默认 PyPI") + "）。");
                    }

                    // 自检失败不立即返回（可能是该源包损坏/环境导入异常），记录原因并继续尝试下一镜像。
                    lastCombined = "pip 安装成功但 import 自检未通过：" + o2 + r2;
                    continue;
                }

                lastCombined = outText ?? "";
                if ((lastCombined.IndexOf("permission", StringComparison.OrdinalIgnoreCase) >= 0
                     || lastCombined.IndexOf("access is denied", StringComparison.OrdinalIgnoreCase) >= 0)
                    && !lastCombined.Contains("--user"))
                {
                    if (PythonWorkerHost.TryPipInstall(pythonPath, "xlwings pywin32", mirror, true, 280000, out var outUser))
                    {
                        var (f4, e4, o4, r4) = RunCapture(pythonPath, new[] { "-c", probeScript }, 20000);
                        if (f4 && e4 == 0)
                        {
                            return (true, "已通过 pip --user 安装 xlwings/pywin32 并自检通过（源：" + (mirror ?? "默认 PyPI") + "）。");
                        }

                        lastCombined = "pip --user 安装成功但 import 自检未通过：" + o4 + r4;
                    }
                }
            }

            return (false, DescribeInstallFailure(ClassifyInstallFailure(lastCombined), lastCombined));
        }

        private static string ClassifyInstallFailure(string combined)
        {
            if (combined.IndexOf("No module named pip", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "no_pip";
            }

            if (combined.IndexOf("Permission denied", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("Access is denied", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("is not writable", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "permission";
            }

            if (combined.IndexOf("Connection", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("TLS", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("Could not find a version", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "network";
            }

            return "unknown";
        }

        private static string DescribeInstallFailure(string reason, string combined)
        {
            var tail = combined.Length > 800 ? combined.Substring(combined.Length - 800) : combined;
            switch (reason)
            {
                case "no_pip":
                    return "当前 Python 缺少 pip 模块，无法自动安装依赖。请安装带 pip 的 Python 3.10/3.11（勾选 Add to PATH）后重试，或手动执行：python -m pip install xlwings pywin32。";
                case "permission":
                    return "安装依赖遇到权限/目录写入失败（--user 重试也未成功）。请以管理员权限重试，或手动执行：python -m pip install --user xlwings pywin32。\n" + tail;
                case "network":
                    return "网络无法访问 pip 源（含清华/阿里镜像），依赖安装失败。请检查网络后重试，或手动执行：python -m pip install xlwings pywin32。\n" + tail;
                default:
                    return "依赖安装失败（未知原因）：\n" + tail;
            }
        }

        /// <summary>把 worker 返回的连接类错误转为可读区分（Excel 未运行/实例不可 COM 访问）。</summary>
        internal static string ClassifyConnectionError(string output)
        {
            if (string.IsNullOrWhiteSpace(output))
            {
                return "";
            }

            if (output.IndexOf("Make sure to have", StringComparison.OrdinalIgnoreCase) >= 0
                && output.IndexOf("pywin32", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return " 提示：xlwings 的 Excel 引擎未加载——通常是依赖刚安装完成、旧 worker 缓存了失败状态。本次已重置 Python worker，请重试一次；若仍失败请检查 Python 环境。";
            }

            if (output.IndexOf("App not found", StringComparison.OrdinalIgnoreCase) >= 0
                || output.IndexOf("Could not connect", StringComparison.OrdinalIgnoreCase) >= 0
                || output.IndexOf("Excel is not running", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return " 提示：无法连接当前 Excel 实例——请确认 Excel 正在运行且为活动窗口后重试；若为位数不匹配场景请按上面指引重装 Python。";
            }

            if (output.IndexOf("books", StringComparison.OrdinalIgnoreCase) >= 0
                && output.IndexOf("active", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return " 提示：未能定位活动工作簿（xw.books.active），请确认至少打开一个工作簿后重试。";
            }

            return "";
        }

        /// <summary>运行外部进程并捕获 stdout/stderr（UTF-8），found=false 表示进程无法启动。</summary>
        private (bool Found, int Exit, string Output, string Error) RunCapture(string fileName, string[] args, int timeoutMs)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                PythonWorkerHost.PreparePythonChildEnv(psi);
                if (args != null && args.Length > 0)
                {
                    psi.Arguments = string.Join(" ", Array.ConvertAll(args, QuoteArg));
                }

                using var proc = new Process { StartInfo = psi };
                if (!proc.Start())
                {
                    return (false, -1, "", "进程无法启动。");
                }

                var stdout = new StringBuilder();
                var stderr = new StringBuilder();
                proc.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
                proc.ErrorDataReceived += (_, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                if (!proc.WaitForExit(Math.Max(timeoutMs, 1000)))
                {
                    try { proc.Kill(); } catch { }
                    return (true, -2, stdout.ToString(), stderr + "执行超时（limit=" + timeoutMs + "ms），进程已终止。");
                }

                proc.WaitForExit();
                return (true, proc.ExitCode, stdout.ToString().Trim(), stderr.ToString().Trim());
            }
            catch (Exception ex)
            {
                AppLog.Warn("PythonExcelService", $"[LogCover] PythonExcelService.cs:694  {ex.GetType().Name}: {ex.Message}");
                return (false, -1, "", ex.Message);
            }
        }

        internal static string QuoteArg(string arg)
        {
            return "\"" + arg.Replace("\"", "\"\"") + "\"";
        }

        private static int Clamp(int ms)
        {
            return Math.Max(1000, Math.Min(ms, 300000));
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (args != null && args.TryGetValue(key, out var raw) && raw != null)
            {
                try { return Convert.ToInt32(raw); } catch { }
            }

            return defaultValue;
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.TryGetValue(key, out var raw) || raw == null)
            {
                return defaultValue;
            }

            try { return Convert.ToBoolean(raw); } catch { }

            return defaultValue;
        }

        private static string OptionalString(Dictionary<string, object> args, string key)
        {
            return args != null && args.TryGetValue(key, out var raw) && raw != null ? Convert.ToString(raw) : null;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(object value)
        {
            if (value == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(value))
                {
                    Marshal.FinalReleaseComObject(value);
                }
            }
            catch (Exception)
            {
                AppLog.Warn("PythonExcelService", "[LogCover] PythonExcelService.cs:756  silent-catch covered (no exception variable)");
            }
        }
    }
}
