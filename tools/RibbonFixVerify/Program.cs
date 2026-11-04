using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Excel;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace RibbonFixVerify
{
    /// <summary>
    /// 针对「JYYJ快捷功能按钮不能用」的真机修复验证（独立 out-of-proc Excel 实例 + 用户测试文件，不改原文件）。
    /// 覆盖两层：
    ///   A) 选择解析层（RibbonCommandService.TryResolveSelection，私有方法用反射调用）：
    ///      · 单格空选区不再抛 IndexOutOfRange（原 RootCause 1）
    ///      · 范围选区矩阵读取正确
    ///      · 连续多次解析（触发 finally SafeRelease）不再抛 InvalidComObject（原 RootCause 2）
    ///   B) 执行层（ExcelToolHost 真实注册表在测试文件上跑）：
    ///      sort_range / dedup_range / clean_excel_data / text_split 四个日志里<Failed 的工具全部 Ok。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly HashSet<int> Spawned = new HashSet<int>();
        private const string TestFile = @"C:\Users\iamly\Desktop\新建文件夹\新建 Microsoft Excel 工作表.xlsx";

        private static int Main()
        {
            Console.WriteLine("== RibbonFixVerify：JYYJ快捷功能修复真机验证 ==");
            var before = new HashSet<int>();
            foreach (var p in Process.GetProcessesByName("EXCEL")) { before.Add(p.Id); p.Dispose(); }

            dynamic app = null;
            dynamic wb = null;
            dynamic ws = null;
            try
            {
                var t = Type.GetTypeFromProgID("Excel.Application");
                if (t == null) { Console.WriteLine("[fatal] 无 Excel.Application COM"); return 2; }
                app = Activator.CreateInstance(t);
                app.Visible = true;
                app.DisplayAlerts = false;
                app.AutomationSecurity = 1;

                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { Spawned.Add(p.Id); }
                    p.Dispose();
                }

                if (!System.IO.File.Exists(TestFile))
                {
                    Console.WriteLine("[fatal] 测试文件不存在：" + TestFile);
                    return 2;
                }
                wb = app.Workbooks.Open(TestFile);
                ws = app.ActiveSheet;

                // ===== P0）RCW共享性探针（前置、只读身份比对，不释放，不污染后续）：=====
                var gp1 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(ws);
                var gp2 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(app.ActiveSheet);
                dynamic probeWb = app.ActiveWorkbook;
                dynamic probeWs1 = probeWb.Worksheets[Convert.ToString(ws.Name)];
                dynamic probeWs2 = probeWb.Worksheets[Convert.ToString(ws.Name)];
                var gp3 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeWs1);
                var gp4 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeWs2);
                Console.WriteLine("[PROBE] ActiveSheet 重复务必同IUnknown=" + (gp1 == gp2)
                    + "；ActiveSheet vs Worksheets[x] 同=" + (gp1 == gp3)
                    + "；Worksheets[x] 两次调用同=" + (gp3 == gp4));
                dynamic probeR1 = probeWs1.Range["A1:B6"];
                dynamic probeR2 = probeWs1.Range["A1:B6"];
                var gr1 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeR1);
                var gr2 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeR2);
                Console.WriteLine("[PROBE] Range['A1:B6'] 两次调用同IUnknown=" + (gr1 == gr2));
                dynamic probeSel1 = app.Selection;
                dynamic probeSel2 = app.Selection;
                var gs1 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeSel1);
                var gs2 = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(probeSel2);
                Console.WriteLine("[PROBE] Application.Selection 两次调用同IUnknown=" + (gs1 == gs2)
                    + "；(类型 " + probeSel1.GetType().Name + ")");

                var service = new RibbonCommandService(app);
                var m = typeof(RibbonCommandService).GetMethod(
                    "TryResolveSelection",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                // ===== A1）单格空选区：不得再抛 IndexOutOfRange / InvalidComObject =====
                Select(app, ws, "Z1");
                RetVal(a => m.Invoke(service, a), "A1 单格空选区解析", out bool ok1, out SelectInfo s1);
                Expect(ok1 && s1.Status == 0 && s1.Matrix[0, 0] == null,
                    "A1 单格空选区解析成功（不再越界，值=null）",
                    $"ok={ok1} status={s1.Status}");

                // ===== A2）范围选区：矩阵与行数列数正确 =====
                WriteTable(app, ws, "A1:B6", new string[,] {
                    { "品名", "数量" }, { "苹果", "3" }, { "苹果", "3" }, { "梨", "5" }, { "", "" }, { "蓝莓", "7" } });
                Select(app, ws, "A1:B6");
                try { Console.WriteLine("[A2] Select 后 Application.Selection.Address=" + app.Selection.Address); }
                catch (Exception se) { Console.WriteLine("[A2] 读 Selection.Address 失败 " + se.GetType().Name + " " + se.Message); }
                RetVal(a => m.Invoke(service, a), "A2 范围选区解析", out bool ok2, out SelectInfo s2);
                Expect(ok2 && s2.Matrix != null && s2.Matrix.GetLength(0) == 6 && s2.Matrix.GetLength(1) == 2
                        && Convert.ToString(s2.Matrix[1, 0]) == "苹果",
                    "A2 范围选区矩阵正确（6行2列，[1,0]=苹果）",
                    ok2 ? $"rows={s2.Matrix.GetLength(0)} cols={s2.Matrix.GetLength(1)}" : $"ok={ok2}");

                // ===== A3）连续多次解析（触发 finally SafeRelease 静态绑定）：不得抛 InvalidComObject =====
                var a3Broken = false;
                for (var i = 0; i < 5 && !a3Broken; i++)
                {
                    Select(app, ws, i % 2 == 0 ? "Z1" : "A1:B6");
                    var okn = m.Invoke(service, new object[] { null, null, null, null, null });
                    if (!(okn is bool b)) { a3Broken = true; break; }
                }
                Expect(!a3Broken, "A3 连续5次选区解析（SafeRelease 静态绑定）无 InvalidComObject", a3Broken ? "异常" : "全程无异常");

                // ===== B0) 直调隔离：dedup_range 自身逻辑（noop writer，绕开注册表 writer 链）=====
                var host = new ExcelToolHost(app);
                ToolRegistry registry = host.CreateRegistry();
                WriteTable(app, ws, "A1:B6", new string[,] {
                    { "品名", "数量" }, { "苹果", "3" }, { "苹果", "3" }, { "梨", "5" }, { "", "" }, { "蓝莓", "7" } });
                try
                {
                    var svc = new ExcelDataToolsService(
                        (object)app,
                        (Func<Dictionary<string, object>, object[,], string, string, ToolResult>)((a, m, t, tn) => new ToolResult { Ok = true }));
                    var dr = svc.DedupRange(Range(
                        ("address", "$A$1:$B$6"), ("sheet", SheetName(ws))));
                    Expect(dr.Ok, "B0 dedup_range 自身逻辑（noop writer）", dr.Ok ? "ok" : dr.ErrorCode + " " + dr.Data);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("      [B0-CATCH] dedup_range 自身逻辑：\n" + ex);
                    Fail("dedup_range 自身逻辑抛异常 " + ex.GetType().Name + " " + ex.Message);
                }

                // ===== B1) 通过注册表执行层：sort_range 先于 dedup，检验共享 RCW 污染时序 =====
                WriteTable(app, ws, "A1:B6", new string[,] {
                    { "品名", "数量" }, { "苹果", "3" }, { "苹果", "3" }, { "梨", "5" }, { "", "" }, { "蓝莓", "7" } });
                B(registry, "sort_range", Range(
                    ("address", "$A$1:$B$6"), ("sheet", SheetName(ws)), ("sort_column", "2"), ("direction", "desc"), ("header_rows", "1")),
                    "排序 sort_range");

                WriteTable(app, ws, "A1:B6", new string[,] {
                    { "品名", "数量" }, { "苹果", "3" }, { "苹果", "3" }, { "梨", "5" }, { "", "" }, { "蓝莓", "7" } });
                B(registry, "dedup_range", Range(
                    ("address", "$A$1:$B$6"), ("sheet", SheetName(ws))),
                    "去重 dedup_range");

                WriteTable(app, ws, "A1:B6", new string[,] {
                    { "品名", "数量" }, { "苹果", "3" }, { "苹果", "3" }, { "梨", "5" }, { "", "" }, { "蓝莓", "7" } });
                B(registry, "clean_excel_data", Range(
                    ("address", "$A$1:$B$6"), ("sheet", SheetName(ws)),
                    ("operations", new Dictionary<string, object>[] { new Dictionary<string, object> { ["tool"] = "remove_blank_rows" } })),
                    "清洗 clean_excel_data");

                WriteTable(app, ws, "D1:D3", new string[,] { { "x,y" }, { "m,n" }, { "p,q" } });
                B(registry, "text_split", Range(
                    ("address", "$D$1:$D$3"), ("sheet", SheetName(ws)), ("column", "1"), ("delimiter", ",")),
                    "文本拆分 text_split");
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[fatal] harness 异常：" + ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                foreach (var id in Spawned)
                {
                    try { Process.GetProcessById(id).Kill(); } catch { }
                }
                Console.WriteLine("\n== 结果：通过 " + _passed + "，失败 " + _failed + " ==");
            }
            return _failed == 0 ? 0 : 1;
        }

        private static string SheetName(dynamic ws)
        {
            try { return Convert.ToString(ws.Name); } catch { return ""; }
        }

        private static void Select(dynamic app, dynamic ws, string addr)
        {
            app.ScreenUpdating = true;
            try { ws.Activate(); } catch { }
            try { ws.Range(addr).Select(); }
            catch { try { app.Range(addr).Select(); } catch { try { app.ActiveSheet.Range(addr).Activate(); } catch { } } }
        }

        private static void WriteTable(dynamic app, dynamic ws, string range, string[,] cells)
        {
            try { ws.Range(range).ClearContents(); } catch { }
            dynamic rng = ws.Range(range);
            dynamic target = rng.Resize[cells.GetLength(0), cells.GetLength(1)];
            var arr = new object[cells.GetLength(0), cells.GetLength(1)];
            for (var r = 0; r < cells.GetLength(0); r++)
                for (var c = 0; c < cells.GetLength(1); c++)
                    arr[r, c] = cells[r, c] == "" ? null : cells[r, c];
            target.Value2 = arr;
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void TryExec(ToolRegistry registry, string tool, Dictionary<string, object> args, string label)
        {
            if (!registry.TryGet(tool, out var d)) { Fail(label + "：工具未注册"); return; }
            ToolResult res = null;
            try { res = d.Execute(args); }
            catch (Exception ex)
            {
                Console.WriteLine("      [TRY-CATCH] " + label + "：\n" + ex);
                Fail(label + "：抛异常 " + ex.GetType().Name + " " + ex.Message);
                return;
            }
            Expect(res != null && res.Ok, label, res.Ok ? (res.ErrorCode ?? "ok") : res.ErrorCode + " " + res.Data);
        }

        private static void B(ToolRegistry registry, string tool, Dictionary<string, object> args, string label)
        {
            try
            {
                if (!registry.TryGet(tool, out var d))
                {
                    Fail(label + "：工具未注册");
                    return;
                }
                ToolResult res;
                try
                {
                    res = d.Execute(args);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("      [RAW-THROW] " + label + " 执行抛异常：\n" + ex);
                    Fail(label + "：执行抛异常 " + ex.GetType().Name + " " + ex.Message);
                    return;
                }
                Expect(res.Ok, label, res.Ok ? (res.ErrorCode ?? "ok") : res.ErrorCode + " " + res.Data);
                if (!res.Ok && res.ErrorCode == "HARNESS_EXCEPTION") { Console.WriteLine("      " + res.Data); }
            }
            catch (Exception ex)
            {
                Console.WriteLine("      [OUTER-THROW] " + label + " 抛异常：\n" + ex);
                Fail(label + "：外层抛异常 " + ex.GetType().Name + " " + ex.Message);
            }
        }

        private static void RetVal(Func<object[], object> invoker, string tag, out bool ok, out SelectInfo info)
        {
            info = new SelectInfo();
            try
            {
                var args = new object[] { null, null, null, null, null };
                var r = invoker(args);
                ok = r is bool b && b;
                info.Matrix = args[1] as object[,];
                info.Rows = args[2] == null ? 0 : Convert.ToInt32(args[2]);
                info.Cols = args[3] == null ? 0 : Convert.ToInt32(args[3]);
                info.Status = args[4] == null ? -1 : Convert.ToInt32(args[4]);
            }
            catch (Exception ex)
            {
                Console.WriteLine("      [raw] " + tag + " 抛异常: " + ex.GetType().Name + " " + ex.Message);
                ok = false;
            }
        }

        private sealed class SelectInfo { public object[,] Matrix; public int Rows; public int Cols; public int Status; }

        private static void Expect(bool c, string name, string detail)
        {
            if (c) { _passed++; Console.WriteLine("[PASS] " + name + " → " + detail); }
            else { _failed++; Console.WriteLine("[FAIL] " + name + " → " + detail); }
        }

        private static void Fail(string msg) { _failed++; Console.WriteLine("[FAIL] " + msg); }

        private static Dictionary<string, object> Range(params (string k, object v)[] pairs)
        {
            var m = new Dictionary<string, object>();
            foreach (var (k, v) in pairs) { m[k] = v; }
            return m;
        }
    }
}