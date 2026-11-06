using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 一键自检（设置页「运行日志」）：在临时工作簿上用与 AI 相同的真实工具链实机跑
    /// 「写入(含形状自适应) / 图表 / VBA 执行」三项，逐项回报成败并写入运行日志，
    /// 用于核对 AI 生成的数据、图表、VBA 代码是否正确。临时工作簿用毕即关闭不保存，不污染用户数据。
    /// 注意：本类必须**无状态**——工具在 finally 里对解析出的 range/sheet/workbook 调 FinalReleaseComObject，
    /// 会连带销毁同一底层对象的所有 RCW，故任何跨工具调用的 worksheet/workbook 引用都会立刻失效。
    /// 因此每次建簿/写入/回读/关簿都按「重新解析 → 用完即放」，绝不把对象跨越工具调用持有多余一帧。
    /// </summary>
    internal sealed class DiagnosticSmokeRunner
    {
        private readonly dynamic app;

        public DiagnosticSmokeRunner(dynamic app)
        {
            this.app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public List<object> Run()
        {
            var results = new List<object>();
            dynamic prevAlerts = null;
            try
            {
                try
                {
                    prevAlerts = app.DisplayAlerts;
                }
                catch (Exception)
                {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:36  silent-catch covered (no exception variable)");
                    // 忽略读取失败
                }

                app.DisplayAlerts = false;

                // 建临时簿（创建即活动簿），随后一律按活动簿重新解析，用完即放。
                dynamic book = app.Workbooks.Add();
                try
                {
                    book.Activate();
                    WriteSetup();
                }
                finally
                {
                    Release((object)book);
                }

                var host = new ExcelToolHost((object)app, null, ExcelToolHost.CreateWorkbookHash("自检"), "smoke-run", "自检");
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                CheckWrite(registry, results);
                CheckChart(registry, results);
                CheckVba(registry, results);
                CheckVbaDataChannel(registry, results);
                CheckVbaDataChannelLargeInt(registry, results);
                CheckVbaDataChannelLargeString(registry, results);
                CheckPython(registry, results);
                CheckPythonDataChannel(registry, results);
                CheckPythonExcel(registry, results);
                CheckPythonExcelRollback(registry, results);
            }
            catch (Exception ex)
            {
                var msg = Safe(ex);
                results.Add(new { name = "自检框架", ok = false, detail = "异常：" + msg });
                DiagnosticsLog.Error("自检框架异常：" + msg);
            }
            finally
            {
                CloseTempBook();
                if (prevAlerts != null)
                {
                    try
                    {
                        app.DisplayAlerts = prevAlerts;
                    }
                    catch (Exception)
                    {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:85  silent-catch covered (no exception variable)");
                        // 忽略恢复失败
                    }
                }
            }

            return results;
        }

        private void WriteSetup()
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveSheet;
                ws.Name = "pf";
                ws.Range["A1"].Value2 = "x";
                ws.Range["B1"].Value2 = "y";
                for (var i = 2; i <= 5; i++)
                {
                    ws.Range["A" + i].Value2 = i;
                    ws.Range["B" + i].Value2 = i * i;
                }
            }
            finally
            {
                Release((object)ws);
            }
        }

        private void CloseTempBook()
        {
            dynamic activeBook = null;
            try
            {
                app.DisplayAlerts = false;
                activeBook = app.ActiveWorkbook;
                if (activeBook != null)
                {
                    activeBook.Close(false);
                }
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:128  silent-catch covered (no exception variable)");
                // 忽略关闭失败
            }
            finally
            {
                Release((object)activeBook);
            }
        }

        private void CheckWrite(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("write_excel_range", out var tool))
            {
                results.Add(new { name = "写入工具", ok = false, detail = "write_excel_range 未注册" });
                DiagnosticsLog.Error("自检-写入工具未注册");
                return;
            }

            try
            {
                // 用扁平列表写 3 格（G1:I1），专门验证「形状自适应重排」是否生效。
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["address"] = "G1:I1",
                    ["values"] = new object[] { "w", "x", "y" }
                });
                var g1 = SafeCell("G1");
                var i1 = SafeCell("I1");
                var ok = res.Ok && string.Equals(g1, "w", StringComparison.Ordinal) && string.Equals(i1, "y", StringComparison.Ordinal);
                results.Add(new { name = "写入工具（含形状自适应）", ok, detail = ok ? "G1:I1=w/x/y" : "失败 err=" + (res.Ok ? "回读不符(G1=" + g1 + ",I1=" + i1 + ")" : res.ErrorCode) });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-写入工具 ok=" + ok + " (G1=" + g1 + ",I1=" + i1 + ") " + res.ErrorCode);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "写入工具", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-写入工具异常：" + Safe(ex));
            }
        }

        private void CheckChart(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("create_excel_chart", out var tool))
            {
                results.Add(new { name = "图表生成", ok = false, detail = "create_excel_chart 未注册" });
                DiagnosticsLog.Error("自检-图表工具未注册");
                return;
            }

            try
            {
                var before = CountCharts();
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["source_address"] = "A1:B5",
                    ["chart_type"] = "scatter",
                    ["anchor"] = "E3",
                    ["title"] = "自检散点图"
                });
                var after = CountCharts();
                var chartType = ChartTypeOf();
                var ok = res.Ok && after > before;
                var typeOk = chartType == -4169; // xlXYScatter
                results.Add(new { name = "图表生成", ok, detail = ok ? "已生成散点图(count=" + after + ",type=" + chartType + (typeOk ? ") 类型正确" : "，类型偏差!") : "失败 err=" + res.ErrorCode });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-图表 ok=" + ok + " count=" + before + "->" + after + " chartType=" + chartType + " " + res.ErrorCode);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "图表生成", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-图表异常：" + Safe(ex));
            }
        }

        private void CheckVba(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                if (string.Equals(ExcelAiAddIn.HostKind, "WPS", StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new { name = "VBA 执行", ok = true, detail = "WPS 无 VBA 环境已过滤（预期跳过）" });
                    DiagnosticsLog.Info("自检-VBA 在 WPS 无 VBA 环境下已过滤（预期跳过）");
                    return;
                }

                results.Add(new { name = "VBA 执行", ok = false, detail = "run_vba_snippet 未注册" });
                DiagnosticsLog.Error("自检-VBA 工具未注册");
                return;
            }

            try
            {
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["snippet"] = "Sub TT自检()\n    Range(\"H1\").Value = \"vba\"\nEnd Sub",
                    ["macro_name"] = "TT自检"
                });
                var h1 = SafeCell("H1");
                var ok = res.Ok && string.Equals(h1, "vba", StringComparison.Ordinal);
                results.Add(new { name = "VBA 执行", ok, detail = ok ? "宏已执行，H1=vba" : "失败 " + (res.Ok ? "回读不符(H1=" + h1 + ")" : res.ErrorCode + "/" + res.Data) });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-VBA 执行 ok=" + ok + " H1=" + h1 + " " + (res.Ok ? "" : res.ErrorCode));
            }
            catch (Exception ex)
            {
                results.Add(new { name = "VBA 执行", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-VBA 异常：" + Safe(ex));
            }
        }

        private void CheckVbaDataChannel(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                if (string.Equals(ExcelAiAddIn.HostKind, "WPS", StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new { name = "VBA 数据通道", ok = true, detail = "WPS 无 VBA 环境已过滤（预期跳过）" });
                    DiagnosticsLog.Info("自检-VBA 数据通道在 WPS 无 VBA 环境下已过滤（预期跳过）");
                    return;
                }

                results.Add(new { name = "VBA 数据通道", ok = false, detail = "run_vba_snippet 未注册" });
                DiagnosticsLog.Error("自检-VBA 数据通道工具未注册");
                return;
            }

            try
            {
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["snippet"] = "Sub TT通道()\n    Dim v As Variant\n    v = XaiArg(\"column\")\n    XaiReturn Array(v, 42)\nEnd Sub",
                    ["macro_name"] = "TT通道",
                    ["args"] = new Dictionary<string, object> { ["column"] = "A" }
                });
                var result = ResultOf(res.Data);
                var ok = res.Ok && result != null && result.Type == Newtonsoft.Json.Linq.JTokenType.Array
                         && string.Equals((string)result[0], "A", StringComparison.Ordinal) && (int)result[1] == 42;
                var detail = ok
                    ? "XaiArg 读到 column=A，XaiReturn 回传 [\"A\",42]"
                    : "失败 " + (res.Ok ? "回传不符 result=" + (result?.ToString() ?? "null") : res.ErrorCode + "/" + res.Data);
                results.Add(new { name = "VBA 数据通道", ok, detail });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-VBA 数据通道 ok=" + ok + " " + detail);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "VBA 数据通道", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-VBA 数据通道异常：" + Safe(ex));
            }
        }

        // 大整数入参：id=3000000000（>2^31-1），验证 XaiParseNum 的 CLng 溢出回退 Double 路径，
        // 以及 VBA 侧不再因 XaiArg 读大整数触发错误 6。
        private void CheckVbaDataChannelLargeInt(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                if (string.Equals(ExcelAiAddIn.HostKind, "WPS", StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new { name = "VBA 数据通道(大整数)", ok = true, detail = "WPS 无 VBA 环境已过滤（预期跳过）" });
                    DiagnosticsLog.Info("自检-VBA 大整数在 WPS 无 VBA 环境下已过滤（预期跳过）");
                    return;
                }

                results.Add(new { name = "VBA 数据通道(大整数)", ok = false, detail = "run_vba_snippet 未注册" });
                DiagnosticsLog.Error("自检-VBA 数据通道(大整数)工具未注册");
                return;
            }

            try
            {
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["snippet"] = "Sub TT大整数()\n    Dim id As Variant\n    id = XaiArg(\"id\")\n    XaiReturn Array(id, 42)\nEnd Sub",
                    ["macro_name"] = "TT大整数",
                    ["args"] = new Dictionary<string, object> { ["id"] = 3000000000L }
                });
                var result = ResultOf(res.Data);
                var ok = res.Ok && result != null && result.Type == Newtonsoft.Json.Linq.JTokenType.Array
                         && (long)result[0] == 3000000000L && (int)result[1] == 42;
                var detail = ok
                    ? "XaiArg 读到 id=3000000000（>2^31 无溢出），XaiReturn 回传 [3000000000,42]"
                    : "失败 " + (res.Ok ? "回传不符 result=" + (result?.ToString() ?? "null") : res.ErrorCode + "/" + res.Data);
                results.Add(new { name = "VBA 数据通道(大整数)", ok, detail });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-VBA 数据通道(大整数) ok=" + ok + " " + detail);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "VBA 数据通道(大整数)", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-VBA 数据通道(大整数)异常：" + Safe(ex));
            }
        }

        // 超长且内嵌双引号/中文的入参：验证 VbaStringLiteral 按原始字符分块、逐块独立转义，
        // 不切断 "" 转义对导致字面量未闭合（此前该缺陷会让模块编译失败、工具挂死）。
        private void CheckVbaDataChannelLargeString(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                if (string.Equals(ExcelAiAddIn.HostKind, "WPS", StringComparison.OrdinalIgnoreCase))
                {
                    results.Add(new { name = "VBA 数据通道(长串分块)", ok = true, detail = "WPS 无 VBA 环境已过滤（预期跳过）" });
                    DiagnosticsLog.Info("自检-VBA 长串分块在 WPS 无 VBA 环境下已过滤（预期跳过）");
                    return;
                }

                results.Add(new { name = "VBA 数据通道(长串分块)", ok = false, detail = "run_vba_snippet 未注册" });
                DiagnosticsLog.Error("自检-VBA 数据通道(长串分块)工具未注册");
                return;
            }

            try
            {
                var big = new string('字', 60) + "\"引号\"" + new string('号', 60);
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["snippet"] = "Sub TT长串()\n    Dim v As Variant\n    v = XaiArg(\"s\")\n    XaiReturn Array(v)\nEnd Sub",
                    ["macro_name"] = "TT长串",
                    ["args"] = new Dictionary<string, object> { ["s"] = big }
                });
                var result = ResultOf(res.Data);
                var ok = res.Ok && result != null && result.Type == Newtonsoft.Json.Linq.JTokenType.Array
                         && string.Equals((string)result[0], big, StringComparison.Ordinal);
                var detail = ok
                    ? "超长含引号串经分块注入后往返一致（len=" + big.Length + "）"
                    : "失败 " + (res.Ok ? "回传不符 result=" + (result?.ToString() ?? "null") : res.ErrorCode + "/" + res.Data);
                results.Add(new { name = "VBA 数据通道(长串分块)", ok, detail });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-VBA 数据通道(长串分块) ok=" + ok + " " + detail);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "VBA 数据通道(长串分块)", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-VBA 数据通道(长串分块)异常：" + Safe(ex));
            }
        }

        private void CheckPythonDataChannel(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_python_code", out var tool))
            {
                results.Add(new { name = "Python 数据通道", ok = false, detail = "run_python_code 未注册" });
                DiagnosticsLog.Error("自检-Python 数据通道工具未注册");
                return;
            }

            try
            {
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "xai_result({'got': input_data['column'], 'sum': input_data['nums'][0] + input_data['nums'][1]})",
                    ["input_data"] = new Dictionary<string, object>
                    {
                        ["column"] = "B",
                        ["nums"] = new object[] { 1, 2 }
                    }
                });
                var result = ResultOf(res.Data);
                var ok = res.Ok && result != null && result.Type == Newtonsoft.Json.Linq.JTokenType.Object
                         && string.Equals((string)result["got"], "B", StringComparison.Ordinal) && (int)result["sum"] == 3;
                var detail = ok
                    ? "input_data 注入成功，xai_result 回传 {got:B,sum:3}"
                    : "失败 " + (res.Ok ? "回传不符 result=" + (result?.ToString() ?? "null") : res.ErrorCode + "/" + res.Data);
                results.Add(new { name = "Python 数据通道", ok, detail });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-Python 数据通道 ok=" + ok + " " + detail);
            }
            catch (Exception ex)
            {
                results.Add(new { name = "Python 数据通道", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-Python 数据通道异常：" + Safe(ex));
            }
        }

        private void CheckPythonExcel(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_python_excel", out var tool))
            {
                results.Add(new { name = "Python 直连Excel", ok = false, detail = "run_python_excel 未注册" });
                DiagnosticsLog.Error("自检-Python 直连Excel 工具未注册");
                return;
            }

            try
            {
                // 首调会触发依赖自动安装（xlwings/pywin32 缺失时），耗时较长属正常。
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "xw.books.active.sheets.active.range(\"J1\").value = 42",
                    ["write_ranges"] = new object[] { "J1" }
                });
                var j1 = SafeCell("J1");
                var ok = res.Ok && string.Equals(j1, "42", StringComparison.Ordinal);
                results.Add(new { name = "Python 直连Excel", ok, detail = ok ? "已通过 xlwings 写 J1=42" : "失败 " + (res.Ok ? "回读不符(J1=" + j1 + ")" : res.ErrorCode + "/" + res.Data) });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-Python 直连Excel ok=" + ok + " J1=" + j1 + " " + (res.Ok ? "" : res.ErrorCode));
            }
            catch (Exception ex)
            {
                results.Add(new { name = "Python 直连Excel", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-Python 直连Excel 异常：" + Safe(ex));
            }
        }

        private void CheckPythonExcelRollback(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_python_excel", out var tool))
            {
                results.Add(new { name = "Python 直连回滚", ok = false, detail = "run_python_excel 未注册" });
                DiagnosticsLog.Error("自检-Python 直连回滚 工具未注册");
                return;
            }

            try
            {
                // 预置 K1=keep，声明 write_ranges 后代码先写脏值再抛异常 → 应自动回滚到 keep。
                WriteSetupValue("K1", "keep");
                var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "xw.books.active.sheets.active.range(\"K1\").value = \"dirty\"\nraise RuntimeError(\"boom\")",
                    ["write_ranges"] = new object[] { "K1" }
                });
                var k1 = SafeCell("K1");
                var ok = !res.Ok && string.Equals(k1, "keep", StringComparison.Ordinal);
                results.Add(new { name = "Python 直连回滚", ok, detail = ok ? "失败后已自动回滚 K1=keep" : "失败 " + (res.Ok ? "未按预期报错" : "K1=" + k1 + " err=" + res.ErrorCode) });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error, "自检-Python 直连回滚 ok=" + ok + " K1=" + k1 + " " + (res.Ok ? "" : res.ErrorCode));
            }
            catch (Exception ex)
            {
                results.Add(new { name = "Python 直连回滚", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-Python 直连回滚 异常：" + Safe(ex));
            }
        }

        private static Newtonsoft.Json.Linq.JToken ResultOf(object data)
        {
            try
            {
                if (data == null)
                {
                    return null;
                }

                var jo = Newtonsoft.Json.Linq.JObject.FromObject(data);
                return jo["result"];
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:468  silent-catch covered (no exception variable)");
                return null;
            }
        }

        private void CheckPython(ToolRegistry registry, List<object> results)
        {
            if (!registry.TryGet("run_python_code", out var tool))
            {
                results.Add(new { name = "Python 执行", ok = false, detail = "run_python_code 未注册" });
                DiagnosticsLog.Error("自检-Python 工具未注册");
                return;
            }

            try
            {
                results.Add(new { name = "Python 执行", ok = true, detail = "开始常驻 worker 计时" });

                var sw = System.Diagnostics.Stopwatch.StartNew();
                var r1 = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "print('v1', 1+1)"
                });
                sw.Stop();
                var firstMs = sw.ElapsedMilliseconds;

                sw.Restart();
                var r2 = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["code"] = "print('v2', 2*2)"
                });
                sw.Stop();
                var secondMs = sw.ElapsedMilliseconds;

                var ok = r1.Ok && r2.Ok;
                results.Add(new
                {
                    name = "Python 执行(复用)",
                    ok,
                    detail = ok
                        ? "首次=" + firstMs + "ms, 复用后=" + secondMs + "ms（复用应远小于首次）"
                        : "失败 " + (r1.Ok ? r2.ErrorCode + "/" + r2.Data : r1.ErrorCode + "/" + r1.Data)
                });
                DiagnosticsLog.Write(ok ? LogLevel.Info : LogLevel.Error,
                    "自检-Python worker 首次=" + firstMs + "ms 复用后=" + secondMs + "ms ok=" + ok + " " +
                    (ok ? "" : (r1.Ok ? r2.ErrorCode : r1.ErrorCode)));
            }
            catch (Exception ex)
            {
                results.Add(new { name = "Python 执行", ok = false, detail = "异常：" + Safe(ex) });
                DiagnosticsLog.Error("自检-Python 异常：" + Safe(ex));
            }
        }

        private string SafeCell(string address)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveSheet;
                return Convert.ToString(ws.Range[address].Value2);
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:531  silent-catch covered (no exception variable)");
                return "<读失败>";
            }
            finally
            {
                Release((object)ws);
            }
        }

        private void WriteSetupValue(string address, string value)
        {
            dynamic ws = null;
            try
            {
                ws = app.ActiveSheet;
                ws.Range[address].Value2 = value;
            }
            finally
            {
                Release((object)ws);
            }
        }

        private int CountCharts()
        {
            dynamic ws = null;
            dynamic charts = null;
            try
            {
                ws = app.ActiveSheet;
                charts = ws.ChartObjects();
                return Convert.ToInt32(charts.Count);
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:565  silent-catch covered (no exception variable)");
                return 0;
            }
            finally
            {
                Release((object)charts);
                Release((object)ws);
            }
        }

        private int ChartTypeOf()
        {
            dynamic ws = null;
            dynamic charts = null;
            dynamic chartObject = null;
            dynamic chart = null;
            try
            {
                ws = app.ActiveSheet;
                charts = ws.ChartObjects();
                if (Convert.ToInt32(charts.Count) > 0)
                {
                    chartObject = charts.Item(1);
                    chart = chartObject.Chart;
                    return Convert.ToInt32(chart.ChartType);
                }
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:593  silent-catch covered (no exception variable)");
                // 忽略
            }
            finally
            {
                Release((object)chart);
                Release((object)chartObject);
                Release((object)charts);
                Release((object)ws);
            }

            return 0;
        }

        private static string Safe(Exception exception)
        {
            return exception?.Message ?? "未知错误";
        }

        private static void Release(object value)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value))
                {
                    Marshal.FinalReleaseComObject(value);
                }
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] DiagnosticSmokeRunner.cs:622  silent-catch covered (no exception variable)");
                // 忽略释放失败
            }
        }
    }
}