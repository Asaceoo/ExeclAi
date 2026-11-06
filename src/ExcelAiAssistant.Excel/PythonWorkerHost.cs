using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// Python 常驻 worker：单个长期存活的解释器进程，通过 stdin/stdout 一行式 JSON-RPC 复用执行，
    /// 替代 run_python_code 每轮「冷启动 python + 写临时脚本 + 进程启停」的高开销路径。
    ///
    /// 设计要点（成熟可靠）：
    /// 1. worker 脚本首次写入 %TEMP%\xai_python_worker.py 并复用；进程常驻，直到崩溃/超时/程序集卸载才重建。
    /// 2. 独立读线程持续消费 stdout，把〔请求 id → 响应 json〕按 id 分发到对应 TaskCompletionSource，
    ///    实现「写入一行请求 → await 结果(带超时)」，规避同步 ReadLine 无法超时导致的死锁。
    /// 3. worker 用 exec 执行代码片段，并把 print 输出重定向进内存缓冲，保证 stdout 只输出一行合法 JSON。
    /// 4. 错误信息经 JSON error 字段返回（不占 stdout），避免管道填满；进程异常/超时自动 kill 并重建。
    /// </summary>
    public static class PythonWorkerHost
    {
        private static readonly object Gate = new object();
        private static Process process;
        private static StreamWriter stdin;
        private static long counter;
        private static string currentPythonPath;
        private static volatile bool readerRunning;
        private static volatile int generation;
        // 读线程（ReaderLoop）与主线程（TryExecute/EnsureWorker）并发读写，须用线程安全字典，
        // 普通 Dictionary 在并发 Add/Remove + TryGetValue 下可能抛异常或损坏桶结构 → 请求悬挂超时。
        private static readonly ConcurrentDictionary<string, TaskCompletionSource<Reply>> Inflight =
            new ConcurrentDictionary<string, TaskCompletionSource<Reply>>(StringComparer.Ordinal);

        private const string WorkerFile = @"xai_python_worker.py";

        /// <summary>worker 脚本（标准库自包含，无需额外依赖/文件），只回单行 JSON。
        /// P1-B/C：注入 xai_result(obj) 结构化回传与 input_data 入参变量。
        /// internal 供测试断言截断标记等关键字符串与 C# 侧常量一致。</summary>
        internal const string WorkerMain = @"
import sys, io, json, base64, traceback, contextlib

# 双保险：即使上游某处仍写出了 UTF-8 BOM 头（.NET Framework 的 StandardInput 默认会带），
# utf-8-sig 也会在 readline 时剥离它，保证 json.loads 不会因 BOM 抛错（否则整条请求被读线程丢弃→超时）。
def _init_io():
    try:
        sys.stdin.reconfigure(encoding='utf-8-sig', errors='backslashreplace')
        sys.stdout.reconfigure(encoding='utf-8', errors='backslashreplace')
        sys.stderr.reconfigure(encoding='utf-8', errors='backslashreplace')
    except Exception:
        pass

_init_io()

def _no_input(*a, **k):
    raise RuntimeError('input() 被禁用：run_python_code 不能交互式等待输入，否则必然超时。请在代码内直接计算，不要读取键盘。')

class _Capped:
    # 捕获 print 输出但限制总量（按字符），防止失控打印把内存打爆直到超时。超限后仅保留前 limit 个字符。
    def __init__(self, limit):
        self._limit = max(int(limit), 8192)
        self._parts = []
        self._used = 0
        self._n = 0
    def write(self, s):
        try:
            s = str(s)
        except Exception:
            s = repr(s)
        if not s:
            return
        self._n += len(s)
        take = self._limit - self._used
        if take > 0:
            self._parts.append(s[:take])
            self._used += len(self._parts[-1])
    def flush(self):
        pass
    def getvalue(self):
        txt = ''.join(self._parts)
        return (txt + '\n...（已截断）') if self._n > self._limit else txt

def _json_default(o):
    # xai_result 序列化兜底：numpy/pandas 常见类型转原生，其余退化为 str。
    try:
        import numpy as np
        if isinstance(o, np.ndarray):
            return o.tolist()
        if isinstance(o, np.generic):
            return o.item()
    except Exception:
        pass
    try:
        import pandas as pd
        if isinstance(o, pd.DataFrame):
            return o.to_dict(orient='records')
        if isinstance(o, pd.Series):
            return o.tolist()
    except Exception:
        pass
    try:
        import datetime as _dt
        if isinstance(o, (_dt.datetime, _dt.date)):
            return o.isoformat()
    except Exception:
        pass
    if hasattr(o, 'tolist'):
        try:
            return o.tolist()
        except Exception:
            pass
    return str(o)

# P3-D1/D4 持久状态：_xai_state 在 worker 进程存活期间跨调用常驻（persist 与否都提供）；
# _persist_globals 仅当请求 persist=true 时被复用为执行命名空间（默认隔离，每次全新）。
# worker 崩溃/超时被重建后两者自然归零：首次 persist 调用会回 state_reset=true 供 AI
# 感知状态已重置。状态仅存内存，不落盘。
_xai_state = {}
_persist_globals = None

def _main():
    global _xai_state, _persist_globals
    while True:
        line = sys.stdin.readline()
        if not line:
            return
        line = line.strip()
        if not line:
            continue
        # 剥离可能残留的 UTF-8 BOM(U+FEFF)：.NET Framework 的 StandardInput 默认可能带 BOM 头，
        # strict utf-8 解码后表现为行首 \ufeff，json.loads 遇之必抛 → 整条请求被读线程丢弃 → 超时。
        line = line.lstrip('\ufeff').strip()
        payload = {}
        reply = {}
        try:
            payload = json.loads(line)
            code = base64.b64decode(payload.get('code') or '').decode('utf-8')
            out_limit = int(payload.get('out') or 262144)
            result_limit = int(payload.get('result_limit') or 1048576)
            buf = _Capped(out_limit)
            persist = bool(payload.get('persist'))
            state_reset = False
            if bool(payload.get('state_clear')):
                _xai_state.clear()
                _persist_globals = None
                # D4 契约：显式清除也是「旧状态已丢失」，无论 persist 与否都回 state_reset=true
                # （工具描述承诺「被清空→state_reset=true」，persist=false 时也须兑现）。
                state_reset = True
            if persist:
                if _persist_globals is None:
                    _persist_globals = {'__name__': '__main__', 'input': _no_input, 'xai_state': _xai_state}
                    state_reset = True
                g = _persist_globals
            else:
                g = {'__name__': '__main__', 'input': _no_input, 'xai_state': _xai_state}
            if 'input_data' in payload:
                g['input_data'] = payload['input_data']
            state = {'result': None, 'truncated': False, 'fallback': False}
            def xai_result(obj):
                state['result'] = obj
            g['xai_result'] = xai_result
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                exec(compile(code, '<run_python_code>', 'exec'), g)
            out = buf.getvalue()
            reply = {'id': payload.get('id'), 'ok': True, 'exit': 0, 'output': out}
            if state_reset:
                reply['state_reset'] = True
            if state['result'] is not None:
                try:
                    result_json = json.dumps(state['result'], ensure_ascii=False, default=_json_default)
                except Exception:
                    result_json = json.dumps({'__xai_unserializable__': repr(state['result'])}, ensure_ascii=False)
                    state['fallback'] = True
                if len(result_json) > result_limit:
                    reply['result_truncated'] = True
                    reply['result'] = result_json[:result_limit]
                else:
                    reply['result'] = json.loads(result_json)
                if state['fallback']:
                    reply['result_fallback'] = True
        except SystemExit as e:
            code = e.code
            reply = {'id': payload.get('id'), 'ok': True, 'exit': int(code) if code is not None else 0, 'output': ''}
        except Exception:
            reply = {'id': payload.get('id'), 'ok': False, 'exit': 1, 'error': traceback.format_exc()}
        sys.stdout.write(json.dumps(reply, ensure_ascii=False) + '\n')
        sys.stdout.flush()

if __name__ == '__main__':
    _main()
";

        /// <summary>经常驻 worker 的 subprocess 执行 pip 安装：worker 子进程环境已实证健康
        /// （沙箱 bash 派生环境下，.NET 直派的 pip 子进程会因环境块问题 WinError 10106，worker 路径不受影响）。
        /// packages 可为空格分隔的多个包名；返回是否安装成功；output 载入 pip 输出尾部（失败原因）。</summary>
        internal static bool TryPipInstall(string pythonPath, string packages, string mirrorUrl, bool withUserFlag, int timeoutMs, out string output)
        {
            output = null;
            try
            {
                var pkgArgs = string.Join(",", packages
                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(p => "'" + p.Replace("'", "\\'") + "'"));
                var code = new StringBuilder();
                code.Append("import subprocess,sys,traceback\n");
                code.Append("args=[sys.executable,'-m','pip','install','--disable-pip-version-check','--no-input'");
                if (withUserFlag)
                {
                    code.Append(",'--user'");
                }

                if (!string.IsNullOrEmpty(mirrorUrl))
                {
                    code.Append(",'-i','").Append(mirrorUrl.Replace("'", "\\'")).Append("'");
                }

                code.Append(",").Append(pkgArgs).Append("]\n");
                code.Append("def _dec(v):\n");
                code.Append("    return v.decode('utf-8','replace')[-1500:] if isinstance(v,bytes) else (v or '')[-1500:]\n");
                code.Append("try:\n");
                code.Append("    r=subprocess.run(args,capture_output=True,timeout=220)\n");
                code.Append("    xai_result({'code':r.returncode,'out':_dec(r.stdout),'err':_dec(r.stderr)})\n");
                code.Append("except Exception:\n");
                code.Append("    xai_result({'code':-999,'out':'','err':traceback.format_exc()[-1500:]})\n");

                var workerOk = TryExecute(
                    code.ToString(),
                    pythonPath,
                    Math.Max(timeoutMs, 1000),
                    256 * 1024,
                    null,
                    256 * 1024,
                    new ExecuteOptions(),
                    out var ok,
                    out var outputOrError,
                    out var result,
                    out _,
                    out _,
                    out _);

                if (!workerOk || !ok)
                {
                    output = outputOrError;
                    return false;
                }

                if (result is Newtonsoft.Json.Linq.JObject jo)
                {
                    var exitCode = jo.Value<int?>("code") ?? -1;
                    output = (Convert.ToString(jo["err"]) + "\n" + Convert.ToString(jo["out"])).Trim();
                    return exitCode == 0;
                }

                output = "worker 返回结构异常";
                return false;
            }
            catch (Exception ex)
            {
                AppLog.Warn("PythonWorkerHost", $"[LogCover] PythonWorkerHost.cs:264  {ex.GetType().Name}: {ex.Message}");
                output = ex.Message;
                return false;
            }
        }

        /// <summary>强制终止当前常驻 worker，下次 TryExecute 自动重建。用于依赖自动安装后丢弃旧 worker 可能损坏的导入状态
        /// （旧 worker 在 pywin32 刚装完的瞬间 import xlwings 会因引擎未注册而缓存失败态，导致后续一直报误导性的缺 pywin32）。</summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetEnvironmentStrings();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeEnvironmentStrings(IntPtr block);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool KernelSetEnvironmentVariable(string name, string value);

        /// <summary>进程环境去重（原生直读，不依赖会抛大小写重复异常的托管枚举）：
        /// 宿主环境块可能同时含 PROGRAMW6432/ProgramW6432 等大小写重复键（沙箱 bash 实证），
        /// .NET 构建 ProcessStartInfo.EnvironmentVariables（大小写不敏感字典）时直接抛「已添加项」。
        /// 删除后出现的重复键（Windows 环境块本就不区分大小写，语义无损）。幂等。</summary>
        private static bool envSanitized;
        internal static void SanitizeOwnEnvironment()
        {
            if (envSanitized)
            {
                return;
            }

            envSanitized = true;
            try
            {
                var block = GetEnvironmentStrings();
                if (block == IntPtr.Zero)
                {
                    return;
                }

                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var duplicates = new List<string>();
                try
                {
                    var cursor = block;
                    string entry;
                    while (!(entry = Marshal.PtrToStringUni(cursor)).Equals(string.Empty))
                    {
                        var eq = entry.IndexOf('=');
                        var name = eq > 0 ? entry.Substring(0, eq) : null;
                        if (!string.IsNullOrEmpty(name) && !seen.Add(name))
                        {
                            duplicates.Add(name);
                        }

                        cursor = (IntPtr)((long)cursor + (entry.Length + 1) * 2);
                    }
                }
                finally
                {
                    FreeEnvironmentStrings(block);
                }

                foreach (var dup in duplicates)
                {
                    KernelSetEnvironmentVariable(dup, null);
                }

                // 规范化自身进程环境键大小写：msys bash 会把 SystemRoot 等归一成全大写，CRT getenv 大小写敏感，
                // 孙进程（如 pip 的 winsock/asyncio）初始化会失败 WinError 10106（真机实证）。
                // 重写为规范键名后，任意子进程都拿到健康环境块。
                CanonicalizeOwn("SystemRoot", "C:\\WINDOWS");
                CanonicalizeOwn("windir", "C:\\WINDOWS");
                CanonicalizeOwn("SystemDrive", "C:");
                CanonicalizeOwn("ComSpec", "C:\\WINDOWS\\system32\\cmd.exe");
                CanonicalizeOwn("ProgramFiles", "C:\\Program Files");
                CanonicalizeOwn("ProgramData", "C:\\ProgramData");
                CanonicalizeOwn("ProgramW6432", "C:\\Program Files");
                CanonicalizeOwn("ALLUSERSPROFILE", "C:\\ProgramData");
            }
            catch (Exception ex)
            {
                AppLog.Warn("python-env", "SanitizeOwnEnvironment 异常（环境保持原样）：" + ex.Message);
            }
        }

        private static void CanonicalizeOwn(string name, string fallback)
        {
            var value = Environment.GetEnvironmentVariable(name);
            if (string.IsNullOrEmpty(value))
            {
                value = fallback;
            }

            Environment.SetEnvironmentVariable(name, value);
        }

        /// <summary>python 子进程环境统一入口：先去重宿主环境块，再强制规范化核心变量键名（CRT getenv 大小写敏感，
        /// msys bash 会把 SystemRoot 归一成大写 SYSTEMROOT，导致 winsock/asyncio 初始化失败 WinError 10106——真机实证），
        /// 最后设置 Python UTF-8 输出。全程 best-effort：环境定制失败不阻断调用（子进程仍可继承原样环境块启动）。</summary>
        internal static void PreparePythonChildEnv(System.Diagnostics.ProcessStartInfo psi)
        {
            SanitizeOwnEnvironment();
            try
            {
                var env = psi.EnvironmentVariables;
                ForceCanonical(env, "SystemRoot", "C:\\WINDOWS");
                ForceCanonical(env, "windir", "C:\\WINDOWS");
                ForceCanonical(env, "SystemDrive", "C:");
                ForceCanonical(env, "ComSpec", "C:\\WINDOWS\\system32\\cmd.exe");
                ForceCanonical(env, "ProgramFiles", "C:\\Program Files");
                ForceCanonical(env, "ProgramData", "C:\\ProgramData");
                ForceCanonical(env, "ALLUSERSPROFILE", "C:\\ProgramData");
                ForceCanonical(env, "APPDATA", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                ForceCanonical(env, "LOCALAPPDATA", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
                ForceCanonical(env, "TEMP", System.IO.Path.GetTempPath());
                ForceCanonical(env, "TMP", System.IO.Path.GetTempPath());
                env["PYTHONIOENCODING"] = "utf-8";
                env["PYTHONUTF8"] = "1";
            }
            catch (Exception xaiEx0) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:PreparePythonChildEnv 静默兜底触发: " + xaiEx0.Message);
            }
        }

        /// <summary>把环境键强制重写为规范大小写：StringDictionary 读值大小写不敏感，但写入时若原键大小写不同
        /// 会保留旧键——必须先 Remove（大小写不敏感）再重加，子进程才拿得到规范键名。</summary>
        private static void ForceCanonical(System.Collections.Specialized.StringDictionary env, string canonical, string fallback)
        {
            var value = env[canonical];
            if (string.IsNullOrEmpty(value))
            {
                value = string.IsNullOrEmpty(fallback) ? null : fallback;
            }

            if (value == null)
            {
                return;
            }

            env.Remove(canonical);
            env[canonical] = value;
        }

        public static void ResetWorker()
        {
            lock (Gate)
            {
                KillWorker();
            }
        }

        /// <summary>常规执行：向常驻 worker 提交代码并等待结果。成功且 ok 返回 true，outputOrError 载入输出/错误。
        /// inputData 为注入命名空间的入参（可空）；result 载入 xai_result 结构化结果（可空），resultTruncated 标记超限截断；
        /// stateReset 标记本调用（重）建了持久上下文（首次 persist / 显式清除 / worker 崩溃超时重建后首次 persist），AI 应据此丢弃旧状态认知。</summary>
        public static bool TryExecute(
            string code,
            string pythonPath,
            int timeoutMs,
            int outputLimitBytes,
            JToken inputData,
            int resultLimitBytes,
            ExecuteOptions options,
            out bool ok,
            out string outputOrError,
            out JToken result,
            out bool resultTruncated,
            out bool stateReset,
            out int exit)
        {
            ok = false;
            outputOrError = null;
            result = null;
            resultTruncated = false;
            stateReset = false;
            exit = -1;
            if (string.IsNullOrWhiteSpace(code))
            {
                outputOrError = "缺少执行代码。";
                return false;
            }

            lock (Gate)
            {
                try
                {
                    EnsureWorker(pythonPath);
                    var id = "req" + Interlocked.Increment(ref counter);
                    var tcs = new TaskCompletionSource<Reply>(TaskCreationOptions.RunContinuationsAsynchronously);
                    Inflight[id] = tcs;

                    var payload = JsonConvert.SerializeObject(new
                    {
                        id,
                        code = Base64(code),
                        @out = outputLimitBytes,
                        result_limit = Math.Max(resultLimitBytes, 1024),
                        input_data = inputData,
                        persist = options != null && options.Persist,
                        state_clear = options != null && options.StateClear
                    });
                    stdin.WriteLine(payload);
                    stdin.Flush();

                    var completed = tcs.Task.Wait(TimeSpan.FromMilliseconds(Math.Max(timeoutMs, 1000)));
                    Inflight.TryRemove(id, out _);
                    if (!completed)
                    {
                        KillWorker(); // worker 卡死：终止后下次自动重建
                        outputOrError = "Python 执行超时（limit=" + timeoutMs + "ms），进程已终止并自动重启。";
                        AppLog.Error("python", "执行超时(limit=" + timeoutMs + "ms)，worker 已终止重建");
                        return false;
                    }

                    var reply = tcs.Task.Result;
                    exit = reply.exit;
                    if (!reply.ok)
                    {
                        outputOrError = string.IsNullOrWhiteSpace(reply.error) ? "Python 执行失败。" : reply.error.Trim();
                        AppLog.Error("python", "执行失败 exit=" + exit + ": " + NoCr(outputOrError));
                        return false;
                    }

                    ok = true;
                    outputOrError = Truncate(reply.output ?? string.Empty, outputLimitBytes);
                    result = reply.result;
                    resultTruncated = reply.result_truncated;
                    stateReset = reply.state_reset;
                    return true;
                }
                catch (Exception ex)
                {
                    KillWorker();
                    outputOrError = "Python worker 异常：" + ex.Message;
                    AppLog.Error("python", "worker 异常：" + ex.Message);
                    return false;
                }
            }
        }

        private static void EnsureWorker(string pythonPath)
        {
            var effective = string.IsNullOrWhiteSpace(pythonPath) ? "python" : pythonPath;
            if (process != null && !process.HasExited && currentPythonPath == effective && readerRunning)
            {
                return;
            }

            KillWorker();
            var script = EnsureWorkerFile();
            var psi = new ProcessStartInfo
            {
                FileName = effective,
                Arguments = Quote("-u") + " " + Quote(script),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                CreateNoWindow = true,
                // 关键：.NET Framework 的 Process.StandardInput 可能带 UTF-8 BOM 头，而 Python 严格 utf-8 解码
                // 遇到 BOM 会抛 JSONDecodeError → worker 回 "id:null" 被读线程丢弃 → 每次调用卡到超时。
                // 注意：此处不依赖 ProcessStartInfo.StandardInputEncoding（不同框架版本/引用解析下未必存在，
                // 且会破坏跨框架编译）；BOM 由 worker 脚本内 _init_io() 用 utf-8-sig 重配置 + readline 后
                // 显式 lstrip('\ufeff') 双重剥除兜底，C# 侧无需也不能透出编码设置。
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };
            try
            {
                // 强制 Python 子进程以 UTF-8 写 stdout/stderr，与 C# 端 UTF8 解码对齐，避免中文输出在
                // GBK 控制台代码页下被读成乱码、进而破坏逐行 JSON 协议解析（worker 卡死超时的诱因之一）。
                PreparePythonChildEnv(psi);
            }
            catch (Exception xaiEx1) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:EnsureWorker 静默兜底触发: " + xaiEx1.Message);
            }
            process = Process.Start(psi);
            currentPythonPath = effective;
            stdin = process.StandardInput;
            ExcelTrace.Info("python_worker", $"worker 启动 pid={process.Id} python={effective}");

            DrainStderr(process);
            // 启动前清空旧请求表，避免读线程分发给已失效的 TCS；用代际号隔离新旧读线程，
            // 旧线程在进程重建后退出时不得误清新 worker 的在途请求。
            var myGen = unchecked(++generation);
            Inflight.Clear();
            var cts = new CancellationTokenSource();
            // 用专用后台线程读 stdout，而非 Task.Run：独立 exe/Addin 进程里 TPL 线程池在同一线程
            // 做同步等待时可能调度异常导致读线程迟迟不消费管道 → 逐行 JSON 永远不投递 → 超时。
            var readerThread = new Thread(() => ReaderLoop(cts.Token, myGen)) { IsBackground = true };
            readerThread.Start();
            readerRunning = true;
        }

        private static void DrainStderr(Process proc)
        {
            try
            {
                // Python worker 的 stderr（解释器启动告警/未捕获 traceback）此前被静默吞掉，
                // 一旦出现「worker 启动即崩/导入失败」就无从定位。现逐行汇入运行日志，覆盖调试需要。
                proc.ErrorDataReceived += (_, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(e.Data))
                    {
                        AppLog.Warn("python", "worker stderr: " + e.Data);
                    }
                };
                proc.BeginErrorReadLine();
            }
            catch (Exception xaiEx2) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:DrainStderr 静默兜底触发: " + xaiEx2.Message);
            }
        }

        /// <summary>持续读 stdout：解析 JSON 行，按 id 分发；收到 EOF（进程退出）时把本代在途请求标记失败。</summary>
        private static void ReaderLoop(CancellationToken token, int gen)
        {
            try
            {
                var reader = process.StandardOutput;
                while (!token.IsCancellationRequested)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                    {
                        break; // EOF：进程退出或流关闭
                    }

                    line = line.Trim();
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    Reply reply;
                    try
                    {
                        reply = JsonConvert.DeserializeObject<Reply>(line);
                    }
                    catch (Exception xaiEx3) {
                        AppLog.Warn("python-worker", "PythonWorkerHost.cs:ReaderLoop 静默兜底触发: " + xaiEx3.Message);
                        continue; // 忽略杂散行
                    }

                    if (reply == null || string.IsNullOrEmpty(reply.id))
                    {
                        continue;
                    }

                    if (Inflight.TryGetValue(reply.id, out var tcs))
                    {
                        Inflight.TryRemove(reply.id, out _);
                        tcs.TrySetResult(reply);
                    }
                }
            }
            catch (Exception xaiEx4) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:ReaderLoop 静默兜底触发: " + xaiEx4.Message);
            }
            finally
            {
                // 仅当仍是本代读线程时放行在途等待者；已换代则旧线程对 Inflight 不再负责。
                if (gen == generation)
                {
                    readerRunning = false;
                    var stale = new List<KeyValuePair<string, TaskCompletionSource<Reply>>>(Inflight);
                    Inflight.Clear();
                    foreach (var pair in stale)
                    {
                        pair.Value.TrySetResult(new Reply { id = pair.Key, ok = false, exit = -1, error = "Python 进程已退出。" });
                    }
                }
            }
        }

        private static string EnsureWorkerFile()
        {
            var dir = Path.GetTempPath();
            var path = Path.Combine(dir, WorkerFile);
            try
            {
                // 每次都以嵌入式最新脚本重写，确保 addin 升级后 worker 语义同步（承载 _Capped 输出上限与
                // input 禁用守卫），避免旧版本残留脚本仍在 %TEMP% 被复用。仅当进程未持有时调用，安全覆盖。
                File.WriteAllText(path, WorkerMain, new UTF8Encoding(false));
            }
            catch
            {
                // 写文件失败时退化为 -c 内联（较弱，但保底可用）。
                // 走到这里说明正常路径失败，交给 KillWorker 后外层捕获。
                throw new InvalidOperationException("无法写入 Python worker 脚本到临时目录。");
            }

            return path;
        }

        private static void KillWorker()
        {
            try
            {
                if (process != null && !process.HasExited)
                {
                    ExcelTrace.Info("python_worker", $"worker 终止 pid={process.Id}");
                    process.Kill();
                }
            }
            catch (Exception xaiEx5) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:KillWorker 静默兜底触发: " + xaiEx5.Message);
            }

            try
            {
                process?.WaitForExit(2000);
            }
            catch (Exception xaiEx6) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:KillWorker 静默兜底触发: " + xaiEx6.Message);
            }

            try
            {
                process?.Dispose();
            }
            catch (Exception xaiEx7) {
                AppLog.Warn("python-worker", "PythonWorkerHost.cs:KillWorker 静默兜底触发: " + xaiEx7.Message);
            }

            process = null;
            stdin = null;
            currentPythonPath = null;
            readerRunning = false;
            // 让旧读线程自行退出（EOF）。
        }

        private static string Base64(string text)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
        }

        private static string Truncate(string text, int limit)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            if (text.Length > limit)
            {
                return text.Substring(0, limit) + "\n" + ScriptRunnerService.OutputTruncatedMarker;
            }

            return text;
        }

        /// <summary>日志行内折叠换行，避免多行 traceback 撑爆单行日志。长度超限时截断。internal 供测试断言。</summary>
        internal static string NoCr(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var flattened = text.Replace("\r", " ").Replace("\n", " ").Trim();
            return flattened.Length > 300 ? flattened.Substring(0, 300) + "…" : flattened;
        }

        internal static string Quote(string arg)
        {
            // Windows 命令行（CreateProcess）内嵌引号用 "" 转义，而非 Unix 的 \"。
            return "\"" + arg.Replace("\"", "\"\"") + "\"";
        }

        /// <summary>执行选项：Persist 保留执行命名空间跨调用（默认隔离）；StateClear 先清空持久状态（xai_state + 命名空间）。</summary>
        public sealed class ExecuteOptions
        {
            public bool Persist;
            public bool StateClear;
        }

        private sealed class Reply
        {
            public string id { get; set; }
            public bool ok { get; set; }
            public int exit { get; set; }
            public string output { get; set; }
            public string error { get; set; }
            public JToken result { get; set; }
            public bool result_truncated { get; set; }
            public bool result_fallback { get; set; }
            public bool state_reset { get; set; }
        }
    }
}