using System.Diagnostics;
using System.Threading;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>常驻 Python worker 的单元/集成测试：验证纯计算成功路径、input() 被禁用（不阻塞超时）、
    /// 失控打印被输出上限截断（防止内存打爆）、P3-D1/D4 持久状态（persist 跨调用、默认隔离、xai_state 常驻、
    /// state_clear 显式清空、state_reset 重建提示）。依赖本机 python；不可用时自动跳过。</summary>
    public sealed class PythonWorkerHostTests
    {
        private static string PythonOrSkip()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "python",
                    Arguments = "-c \"print(1)\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                using var p = Process.Start(psi);
                if (p != null && p.WaitForExit(15000) && p.ExitCode == 0)
                {
                    return "python";
                }
            }
            catch
            {
            }

            return null;
        }

        [Fact]
        public void Simple_computation_returns_ok()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            var okExec = PythonWorkerHost.TryExecute(
                "print(6 * 7)\nprint('中文输出')", py, 15000, 65536, null, 1048576, null,
                out var ok, out var output, out var result, out var resultTruncated, out var stateReset, out var exit);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.Equal(0, exit);
            Assert.False(stateReset);
            Assert.Contains("42", output);
            Assert.Contains("中文输出", output);
            Assert.Null(result);
            Assert.False(resultTruncated);
        }

        [Fact]
        public void Input_is_disabled_does_not_block_until_timeout()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            // input() 被禁用守卫兜底：即便逃过静态校验，也不应把 worker 钉死到超时。
            var sw = Stopwatch.StartNew();
            var okExec = PythonWorkerHost.TryExecute(
                "x = input('prompt')", py, 8000, 4096, null, 1048576, null,
                out var ok, out var output, out _, out _, out _, out _);
            sw.Stop();

            // 0.8 秒内返回（远小于 8s 超时），且不是“超时”语义：要么失败要么成功，但绝不卡满 8s。
            Assert.True(sw.ElapsedMilliseconds < 8000);
            Assert.False(ok);    // 代码因 input() 被禁用而失败（worker 已返回错误答复而非超时）
            Assert.Contains("input", output);
        }

        [Fact]
        public void Runaway_print_is_truncated_not_overtime()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            // 大循环打印：超出 4KB 输出上限即被 _Capped 截断，返回成功且输出带截断标记。
            var okExec = PythonWorkerHost.TryExecute(
                "for i in range(20000):\n    print(i)", py, 10000, 4096, null, 1048576, null,
                out var ok, out var output, out _, out _, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.Contains("（已截断）", output);
            Assert.True(output.Length <= 4096 + 32); // 截断后不超过上限 + 标记
        }

        [Fact]
        public void Xai_result_returns_structured_json()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            var okExec = PythonWorkerHost.TryExecute(
                "xai_result({'a': 1, 'b': [1, 2, 3], 'c': '中文'})", py, 15000, 65536, null, 1048576, null,
                out var ok, out var output, out var result, out var resultTruncated, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.False(resultTruncated);
            Assert.NotNull(result);
            Assert.Equal(1, (int)result["a"]);
            Assert.Equal(3, ((JArray)result["b"]).Count);
            Assert.Equal("中文", (string)result["c"]);
        }

        [Fact]
        public void Xai_result_scalar_and_list_supported()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            var okExec = PythonWorkerHost.TryExecute(
                "xai_result([i * i for i in range(5)])", py, 15000, 65536, null, 1048576, null,
                out var ok, out _, out var result, out _, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.Equal(4, (int)result[2]);
        }

        [Fact]
        public void Input_data_is_injected_into_namespace()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            var input = JToken.Parse("{\"column\": \"A\", \"threshold\": 100}");
            var okExec = PythonWorkerHost.TryExecute(
                "xai_result({'col': input_data['column'], 'thr': input_data['threshold']})",
                py, 15000, 65536, input, 1048576, null,
                out var ok, out _, out var result, out _, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.Equal("A", (string)result["col"]);
            Assert.Equal(100, (int)result["thr"]);
        }

        [Fact]
        public void Input_data_2d_array_roundtrips()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            var input = JToken.Parse("[[1, 2], [3, 4]]");
            var okExec = PythonWorkerHost.TryExecute(
                "xai_result([[row[0] + row[1] for row in input_data]])",
                py, 15000, 65536, input, 1048576, null,
                out var ok, out _, out var result, out _, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.Equal(3, (int)result[0][0]);
            Assert.Equal(7, (int)result[0][1]);
        }

        [Fact]
        public void Oversized_result_is_truncated_not_failed()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            // 结果超 1KB 上限：worker 截断并标记 result_truncated，不抛异常。
            var okExec = PythonWorkerHost.TryExecute(
                "xai_result('x' * 5000)", py, 15000, 65536, null, 1024, null,
                out var ok, out _, out var result, out var resultTruncated, out _, out _);

            Assert.True(okExec);
            Assert.True(ok);
            Assert.True(resultTruncated);
            Assert.NotNull(result);
            Assert.True(result.ToString().Length <= 1024);
        }

        [Fact]
        public void Persist_retains_namespace_across_calls()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            // 干净基线：重建 worker 后首次 persist 必回 state_reset=true。
            PythonWorkerHost.ResetWorker();

            var ok1 = PythonWorkerHost.TryExecute(
                "x = 42", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok1ok, out _, out _, out _, out var reset1, out _);
            Assert.True(ok1 && ok1ok);
            Assert.True(reset1); // 首次 persist 重建命名空间

            var ok2 = PythonWorkerHost.TryExecute(
                "xai_result({'x': x})", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok2ok, out _, out var result2, out _, out var reset2, out _);
            Assert.True(ok2 && ok2ok);
            Assert.False(reset2); // 命名空间已复用，不重置
            Assert.Equal(42, (int)result2["x"]);
        }

        [Fact]
        public void Non_persist_calls_are_isolated()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            PythonWorkerHost.ResetWorker();

            var ok1 = PythonWorkerHost.TryExecute(
                "x = 42", py, 15000, 65536, null, 1048576, null,
                out var ok1ok, out _, out _, out _, out _, out _);
            Assert.True(ok1 && ok1ok);

            // 未开 persist 的下一次调用看不到上一轮普通变量。
            var ok2 = PythonWorkerHost.TryExecute(
                "xai_result('x' in globals())", py, 15000, 65536, null, 1048576, null,
                out var ok2ok, out _, out var result, out _, out _, out _);
            Assert.True(ok2 && ok2ok);
            Assert.False((bool)result);
        }

        [Fact]
        public void Xai_state_persists_even_without_persist()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            PythonWorkerHost.ResetWorker();

            var ok1 = PythonWorkerHost.TryExecute(
                "xai_state['x'] = 1", py, 15000, 65536, null, 1048576, null,
                out var ok1ok, out _, out _, out _, out _, out _);
            Assert.True(ok1 && ok1ok);

            // xai_state 是持久 dict：即使未开 persist，下次调用仍可读到。
            var ok2 = PythonWorkerHost.TryExecute(
                "xai_result(xai_state.get('x'))", py, 15000, 65536, null, 1048576, null,
                out var ok2ok, out _, out var result, out _, out _, out _);
            Assert.True(ok2 && ok2ok);
            Assert.Equal(1, (int)result);
        }

        [Fact]
        public void State_clear_wipes_namespace_and_xai_state()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            PythonWorkerHost.ResetWorker();

            // 第一次 persist 存下状态与命名空间变量。
            var ok1 = PythonWorkerHost.TryExecute(
                "xai_state['x'] = 1\nkeep = 'ns'", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok1ok, out _, out _, out _, out _, out _);
            Assert.True(ok1 && ok1ok);

            // state_clear=true：xai_state 与命名空间均清空，并回 state_reset=true。
            var ok2 = PythonWorkerHost.TryExecute(
                "xai_result({'x': xai_state.get('x'), 'ns': 'keep' in globals()})", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true, StateClear = true },
                out var ok2ok, out _, out var result, out _, out var reset2, out _);
            Assert.True(ok2 && ok2ok);
            Assert.True(reset2);
            Assert.Equal(JTokenType.Null, result["x"].Type);
            Assert.False((bool)result["ns"]);

            // 清空后再次 persist 读取，状态确为空。
            var ok3 = PythonWorkerHost.TryExecute(
                "xai_result({'x': xai_state.get('x')})", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok3ok, out _, out var result3, out _, out _, out _);
            Assert.True(ok3 && ok3ok);
            Assert.Equal(JTokenType.Null, result3["x"].Type);
        }

        [Fact]
        public void State_reset_after_worker_rebuild_is_reported()
        {
            var py = PythonOrSkip();
            if (py == null)
            {
                return;
            }

            // 干净基线：重建 worker 保证首次 persist 回 state_reset=true（隔离前序测试共享的静态 worker）。
            PythonWorkerHost.ResetWorker();

            // 首次 persist 建立状态。
            var ok1 = PythonWorkerHost.TryExecute(
                "xai_state['x'] = 1", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok1ok, out _, out _, out _, out var reset1, out _);
            Assert.True(ok1 && ok1ok);
            Assert.True(reset1);

            // 模拟 worker 崩溃/超时重建：ResetWorker 杀进程，持久状态丢失。
            PythonWorkerHost.ResetWorker();

            // 重建后首次 persist 调用必须回 state_reset=true，且 xai_state 为空。
            var ok2 = PythonWorkerHost.TryExecute(
                "xai_result({'x': xai_state.get('x')})", py, 15000, 65536, null, 1048576,
                new PythonWorkerHost.ExecuteOptions { Persist = true },
                out var ok2ok, out _, out var result, out _, out var reset2, out _);
            Assert.True(ok2 && ok2ok);
            Assert.True(reset2);
            Assert.Equal(JTokenType.Null, result["x"].Type);
        }

        [Fact]
        public void Quote_uses_windows_double_quote_escaping()
        {
            // F12 修复：Windows 命令行（CreateProcess）内嵌引号用 "" 转义，而非 Unix 的 \"。
            Assert.Equal("\"plain\"", PythonWorkerHost.Quote("plain"));
            Assert.Equal("\"a\"\"b\"", PythonWorkerHost.Quote("a\"b"));
        }

        [Fact]
        public void NoCr_folds_multiline_traceback_into_single_line()
        {
            // 多行 traceback（含 \n 与 \r\n）必须折叠为单行日志；缩进空格保留（仅折行，不压缩空白）。
            Assert.Equal("Traceback:     ZeroDivisionError", PythonWorkerHost.NoCr("Traceback:\n    ZeroDivisionError"));
            Assert.Equal("line1  line2", PythonWorkerHost.NoCr("line1\r\nline2"));
            Assert.Equal("padded", PythonWorkerHost.NoCr("   padded  "));
        }

        [Fact]
        public void NoCr_truncates_overlong_line_with_ellipsis()
        {
            var longLine = new string('x', 500);
            // 长度折叠后仍 > 300 时截断到 300 并追加省略号，保证单行日志不会撑爆/打爆日志 UI。
            var result = PythonWorkerHost.NoCr(longLine);
            Assert.Equal(301, result.Length); // 300 内容 + 1 省略号
            Assert.EndsWith("…", result);
            Assert.Equal(new string('x', 299) + "x…", result); // 截断内容为前 300 字符
        }

        [Fact]
        public void NoCr_handles_null_and_empty_safely()
        {
            // 空/只空白输入不抛异常：null 返回 null，空串保持空串。
            Assert.Null(PythonWorkerHost.NoCr(null));
            Assert.Equal(string.Empty, PythonWorkerHost.NoCr(string.Empty));
            Assert.Equal(string.Empty, PythonWorkerHost.NoCr("   \r\n   "));
        }
    }
}
