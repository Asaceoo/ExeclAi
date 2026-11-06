// P3-D2 挂起定位诊断：把候选 VBA 写法作为「首个 run_vba_snippet 调用」在全新工作簿上执行，
// 逐个验证是 XaiStateSet 特定挂起还是序列累积问题。用法：VerifyP3Diag <caseKey>。
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace VerifyP3Diag
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            // 支持逗号分隔的顺序用例（同一 Excel 会话连续执行，用于跨调用状态验证）。
            var keys = args.Length > 0 ? args[0].Split(',') : new[] { "set" };

            var result = 1;
            var runner = new Thread(() => result = RunSequence(keys)) { IsBackground = true };
            runner.SetApartmentState(ApartmentState.STA);
            runner.Start();
            if (!runner.Join(120000))
            {
                runner.Abort();
                Console.WriteLine("FAIL  120s 超时（某写法挂起）");
                return 1;
            }

            return result;
        }

        private static int RunSequence(string[] keys)
        {
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                wb = app.Workbooks.Add();

                var host = new ExcelToolHost((object)app);
                host.ScriptOptions.Enabled = true;
                var registry = host.CreateRegistry();

                if (!registry.TryGet("run_vba_snippet", out var tool))
                {
                    Console.WriteLine("FAIL  run_vba_snippet 未注册");
                    return 1;
                }

                var rc = 0;
                foreach (var key in keys)
                {
                    var snippet = Snippet(key);
                    Console.WriteLine($"== VerifyP3Diag [{key}] ==");
                    Console.WriteLine("snippet: " + snippet.Replace("\n", " ⏎ "));

                    if (key == "moddump")
                    {
                        DumpHelperModule(app, wb);
                        continue;
                    }

                    if (key == "bisect")
                    {
                        BisectReadCall(app, wb);
                        continue;
                    }

                    if (key == "manual")
                    {
                        ManualFreshModule(app, wb);
                        continue;
                    }

                    var callArgs = new Dictionary<string, object> { ["snippet"] = snippet };
                    if (key == "arg")
                    {
                        callArgs["args"] = Newtonsoft.Json.Linq.JToken.Parse("{\"total\": 42}");
                    }
                    if (key.StartsWith("sclear_"))
                    {
                        callArgs["state_clear"] = true;
                    }

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    var r = tool.Execute(callArgs);
                    sw.Stop();
                    var stateReset = "(?)";
                    try
                    {
                        var t = r.Data.GetType();
                        var p = t.GetProperty("state_reset");
                        stateReset = p == null ? "(n/a)" : Convert.ToString(p.GetValue(r.Data, null));
                    }
                    catch
                    {
                    }

                    Console.WriteLine($"{(r.Ok ? "PASS" : "FAIL")}  {r.ErrorCode}: {Short(r.Data as string)} [state_reset={stateReset}] [{sw.ElapsedMilliseconds}ms]");
                    if (r.Ok && (key.StartsWith("read") || key.StartsWith("cross")))
                    {
                        dynamic ws = wb.Worksheets.Item[1];
                        var raw = ws.Cells.Item[1, 1].Value2;
                        Console.WriteLine("  A1=" + (raw == null ? "(null)" : Convert.ToString(raw)));
                    }

                    if (!r.Ok) rc = 1;
                }

                return rc;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL  异常: " + ex.Message);
                return 1;
            }
            finally
            {
                try { if (wb != null) wb.Close(false); } catch { }
                try { if (app != null) { app.Quit(); } } catch { }
                try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch { }
            }
        }

        private static string Snippet(string key)
        {
            switch (key)
            {
                case "clear":
                    return "Sub T()\n XaiStateClear\n Range(\"A1\").Value = 1\nEnd Sub";
                case "set":
                    return "Sub T()\n XaiStateSet \"k\", 7\n Range(\"A1\").Value = 1\nEnd Sub";
                case "write":
                    return "Sub T()\n XaiStateSet \"k\", 7\nEnd Sub";
                case "read":
                    return "Sub T()\n Range(\"A1\").Value = XaiStateGet(\"k\")\nEnd Sub";
                case "cross": // 同一次编译会话内 Set+Get 直连
                    return "Sub T()\n XaiStateSet \"k\", 7\n Range(\"A1\").Value = XaiStateGet(\"k\")\nEnd Sub";
                case "crossstr": // 字符串状态同会话直连
                    return "Sub T()\n XaiStateSet \"k\", \"abc\"\n Range(\"A1\").Value = XaiStateGet(\"k\")\nEnd Sub";
                case "readraw": // 直接读原始状态串（不含解析），区分「没写入」与「解析失败」
                    return "Sub T()\n Range(\"A1\").Value = XaiStateRaw\nEnd Sub";
                case "dict": // 仅 CreateObject + 赋值，不调 XaiStateRawSet/XaiJson
                    return "Sub T()\n Dim d As Object\n Set d = CreateObject(\"Scripting.Dictionary\")\n d(\"k\") = 7\n Range(\"A1\").Value = d(\"k\")\nEnd Sub";
                case "rawset": // 仅调 XaiStateRawSet（跨模块写），不解析不序列化；VBA 字符串内嵌引号须用 "" 转义
                    return "Sub T()\n XaiStateRawSet \"{\"\"k\"\":7}\"\n Range(\"A1\").Value = 1\nEnd Sub";
                case "get": // 仅 XaiStateGet 读空状态
                    return "Sub T()\n Range(\"A1\").Value = XaiStateGet(\"k\")\nEnd Sub";
                case "set2": // 先手动 rawset 再 Set
                    return "Sub T()\n XaiStateRawSet \"{\"\"k\"\":1}\"\n XaiStateSet \"k\", 7\n Range(\"A1\").Value = 1\nEnd Sub";
                case "arg": // args 通道 + XaiArg 读（隔离 XaiState 是否才是问题）
                    return "Sub T()\n Range(\"A1\").Value = XaiArg(\"total\")\nEnd Sub";
                case "parse": // 仅 XaiParse(XaiStateRaw)
                    return "Sub T()\n Dim obj As Variant\n obj = XaiParse(XaiStateRaw)\n Range(\"A1\").Value = 1\nEnd Sub";
                case "json": // 仅 XaiJson(d,0) 序列化 Dictionary
                    return "Sub T()\n Dim d As Object\n Set d = CreateObject(\"Scripting.Dictionary\")\n d(\"k\") = 7\n Range(\"A1\").Value = Len(XaiJson(d, 0))\nEnd Sub";
                default:
                    return "Sub T()\n Range(\"A1\").Value = 1\nEnd Sub";
            }
        }

        private static string Short(string s)
        {
            if (s == null) return "(null)";
            return s.Length <= 700 ? s : s.Substring(0, 700) + "…";
        }

        /// <summary>手动复刻「每轮新建运行模块」生产流程：新建模块 → 声明区 + 辅助块 + 片段 → 解析入口 →
        /// 手动 Run 探针，逐步打印，定位 449「参数不可选」的确切来源。</summary>
        private static void ManualFreshModule(dynamic app, dynamic wb)
        {
            dynamic vbProject = null;
            dynamic components = null;
            dynamic module = null;
            dynamic cm = null;
            try
            {
                vbProject = wb.VBProject;
                components = vbProject.VBComponents;
                module = components.Add(1); // vbext_ct_StdModule
                try { module.Name = "JYYJRun_manual"; } catch { }

                cm = module.CodeModule;
                cm.AddFromString(ScriptRunnerService.VbaHelperDeclarations);
                var helperBlock = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024 * 1024);
                cm.AddFromString(helperBlock);
                cm.AddFromString("Sub T()\n Range(\"A1\").Value = 42\nEnd Sub");

                var all = Convert.ToString(cm.Lines(1, Convert.ToInt32(cm.CountOfLines)));
                var entry = ScriptRunnerService.FindEntryProcedure(all);
                Console.WriteLine("  [manual] FindEntryProcedure=" + (entry ?? "(null)"));
                Console.WriteLine("  [manual] CountOfLines=" + Convert.ToInt32(cm.CountOfLines) + "  all.Split 行数=" + all.Split('\n').Length);
                Console.WriteLine("  [manual] 含 Sub T=" + all.Contains("Sub T"));
                var idxT = all.IndexOf("Sub T", StringComparison.Ordinal);
                Console.WriteLine("  [manual] Sub T 所在行号=" + (idxT < 0 ? "-1" : (all.Substring(0, idxT).Split('\n').Length + 1).ToString()));
                // 打印末尾 12 行，确认最后的过程到底是什么。
                var lines = all.Split('\n');
                for (var li = Math.Max(0, lines.Length - 12); li < lines.Length; li++)
                {
                    Console.WriteLine("    |" + lines[li]);
                }
                // 打印 Sub T 前后 3 行，确认其位置与形态。
                if (idxT >= 0)
                {
                    var startLine = all.Substring(0, idxT).Split('\n').Length;
                    for (var li = Math.Max(0, startLine - 2); li <= startLine + 2 && li < lines.Length; li++)
                    {
                        Console.WriteLine("   T|" + lines[li]);
                    }
                }

                var moduleName = Convert.ToString(module.Name);
                var probeFQ = moduleName + ".XaiProbe9F2E";
                var proc = moduleName + "." + entry;
                Console.WriteLine("  [manual] probeFQ=" + probeFQ + "  proc=" + proc);
                Console.WriteLine("  [manual] helperBlock=" + helperBlock.Split('\n').Length + " 行");

                // 手动执行探针（与生产同路径），捕获任何异常。
                try
                {
                    var r = app.Run(probeFQ, proc);
                    Console.WriteLine("  [manual] 探针返回: " + (r == null ? "(null)" : Convert.ToString(r)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  [manual] 探针 Run 异常: " + ex.Message + "  HRESULT=0x" + ex.HResult.ToString("X8"));
                }

                dynamic ws = wb.Worksheets.Item[1];
                var raw = ws.Cells.Item[1, 1].Value2;
                Console.WriteLine("  [manual] A1=" + (raw == null ? "(null)" : Convert.ToString(raw)));
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL  manual 异常: " + ex.Message);
            }
            finally
            {
                try { if (components != null && module != null) components.Remove(module); } catch { }
                Release((object)cm);
                Release((object)module);
                Release((object)components);
                Release((object)vbProject);
            }
        }

        /// <summary>转储草稿模块 JYYJHelper 结构：总行数、声明区行数、声明区文本、JYYJXaiState 是否就绪，
        /// 用于判断「保留声明区」路径是否真的未触碰声明区。</summary>
        private static void DumpHelperModule(dynamic app, dynamic wb)
        {
            try
            {
                dynamic vbProject = wb.VBProject;
                dynamic components = vbProject.VBComponents;
                try
                {
                    dynamic m = components.Item("JYYJHelper");
                    dynamic cm = m.CodeModule;
                    var total = Convert.ToInt32(cm.CountOfLines);
                    var decl = Convert.ToInt32(cm.CountOfDeclarationLines);
                    Console.WriteLine($"  JYYJHelper 总行数={total} 声明区行数={decl}");
                    if (decl > 0)
                    {
                        Console.WriteLine("  声明区文本:\n" + Convert.ToString(cm.Lines(1, decl)));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  JYYJHelper 模块读取失败: " + ex.Message);
                }

                try
                {
                    dynamic s = components.Item("JYYJXaiState");
                    dynamic scm = s.CodeModule;
                    Console.WriteLine($"  JYYJXaiState 存在 总行数={Convert.ToInt32(scm.CountOfLines)}");
                    // 直读原始状态值（运行未编辑模块的过程，不触发重编译）：判断值是否在 write 调用后仍存活。
                    try
                    {
                        var raw = app.Run("JYYJXaiState.XaiStateRaw");
                        Console.WriteLine("  JYYJXaiState.gXaiState=(" + (raw == null ? "null" : Convert.ToString(raw)) + ")");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("  直读 gXaiState 失败: " + ex.Message);
                    }
                }
                catch
                {
                    Console.WriteLine("  JYYJXaiState 不存在");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  VBProject 读取失败: " + ex.Message);
            }
        }

        /// <summary>二分定位「read 调用过程中 gXaiState 何时被清零」：先经真实 run_vba_snippet 写状态，
        /// 再手工复刻 read 调用的每个操作（判静态→删过程区→加片段→补声明→补辅助块→Run 探针），
        /// 每阶段后直读 gXaiState，定位归零的确切操作，判断是「大块 DeleteLines 触发整工程重编译」还是其它因素。</summary>
        private static void BisectReadCall(dynamic app, dynamic wb)
        {
            var host = new ExcelToolHost((object)app);
            host.ScriptOptions.Enabled = true;
            var registry = host.CreateRegistry();
            if (!registry.TryGet("run_vba_snippet", out var tool))
            {
                Console.WriteLine("FAIL run_vba_snippet 未注册");
                return;
            }

            // 阶段0：真实 write 调用写状态。
            var w = tool.Execute(new Dictionary<string, object> { ["snippet"] = Snippet("write") });
            Console.WriteLine("write 调用: " + (w.Ok ? "PASS" : "FAIL " + w.ErrorCode + " " + Short(w.Data as string)));
            DumpState(app, "阶段0 write 后");

            dynamic vbProject = null;
            dynamic components = null;
            dynamic module = null;
            dynamic cm = null;
            try
            {
                vbProject = wb.VBProject;
                components = vbProject.VBComponents;

                // 阶段A：EnsureStateModule 等价——读状态模块代码判断 ready（不回写、不重建）。
                try
                {
                    dynamic sm = components.Item("JYYJXaiState");
                    dynamic scm = sm.CodeModule;
                    var scnt = Convert.ToInt32(scm.CountOfLines);
                    var scode = scnt > 0 ? Convert.ToString(scm.Lines(1, scnt)) : "";
                    Console.WriteLine("  阶段A EnsureStateModule: 存在=" + (scnt > 0) + " ready=" + (scode.IndexOf("XaiStateRaw", StringComparison.Ordinal) >= 0));
                    Release((object)scm);
                    Release((object)sm);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  阶段A 状态模块读取失败: " + ex.Message);
                }
                DumpState(app, "阶段A EnsureStateModule 后");

                // 阶段B：FindOrAddScratchModule + ReplaceModuleCode 的「读声明区判静态」。
                try { module = components.Item("JYYJHelper"); } catch { module = null; }
                Console.WriteLine("  阶段B JYYJHelper 存在=" + (module != null));
                cm = module.CodeModule;
                var total = Convert.ToInt32(cm.CountOfLines);
                var decl = Convert.ToInt32(cm.CountOfDeclarationLines);
                var current = decl > 0 ? Convert.ToString(cm.Lines(1, decl)) : "";
                var isStatic = ScriptRunnerService.IsStaticDeclarationSection(current);
                Console.WriteLine($"  阶段B ReplaceModuleCode 读取: total={total} decl={decl} 静态声明={isStatic}");

                // 阶段C：DeleteLines 删过程区（生产每轮大删 ~200 行）。
                if (isStatic)
                {
                    if (total > decl) cm.DeleteLines(decl + 1, total - decl);
                }
                else
                {
                    cm.DeleteLines(1, total);
                }
                DumpState(app, "阶段C 大删过程区后");

                // 阶段D：AddFromString 片段。
                cm.AddFromString("Sub T()\n Range(\"A1\").Value = XaiStateGet(\"k\")\nEnd Sub");
                DumpState(app, "阶段D 加片段后");

                // 阶段E：AppendVbaProbe 的 EnsureStaticDeclarations 等价。
                var decl2 = Convert.ToInt32(cm.CountOfDeclarationLines);
                var current2 = decl2 > 0 ? Convert.ToString(cm.Lines(1, decl2)) : "";
                if (ScriptRunnerService.IsStaticDeclarationSection(current2))
                {
                    Console.WriteLine("  阶段E EnsureStaticDeclarations: 声明区静态，跳过");
                }
                else
                {
                    Console.WriteLine("  阶段E EnsureStaticDeclarations: 声明区非静态，重写");
                    if (decl2 > 0) cm.DeleteLines(1, decl2);
                    cm.InsertLines(1, ScriptRunnerService.VbaHelperDeclarations);
                }
                DumpState(app, "阶段E 补声明后");

                // 阶段F：AppendVbaProbe 的「探针已存在则跳过」+ AddFromString(helperBlock)。
                var hasProbe = false;
                var cnt2 = Convert.ToInt32(cm.CountOfLines);
                if (cnt2 > 0)
                {
                    var all2 = Convert.ToString(cm.Lines(1, cnt2));
                    hasProbe = all2.IndexOf("XaiProbe9F2E", StringComparison.Ordinal) >= 0;
                }
                if (!hasProbe)
                {
                    var helperBlock = ScriptRunnerService.BuildVbaHelperBlock("{}", 1024 * 1024);
                    cm.AddFromString(helperBlock);
                    Console.WriteLine("  阶段F 补辅助块（" + helperBlock.Split('\n').Length + " 行）");
                }
                else
                {
                    Console.WriteLine("  阶段F 辅助块已含探针，跳过");
                }
                DumpState(app, "阶段F 补辅助块后");

                // 阶段G：Run 探针（真实工具路径的入口）。
                try
                {
                    var probeRaw = app.Run("JYYJHelper.XaiProbe9F2E", "T");
                    Console.WriteLine("  阶段G 探针返回: " + (probeRaw == null ? "(null)" : Convert.ToString(probeRaw)));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  阶段G 探针 Run 失败: " + ex.Message);
                }
                DumpState(app, "阶段G Run 探针后");

                dynamic ws = wb.Worksheets.Item[1];
                var raw = ws.Cells.Item[1, 1].Value2;
                Console.WriteLine("  A1=" + (raw == null ? "(null)" : Convert.ToString(raw)));
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL bisect 异常: " + ex.Message);
            }
            finally
            {
                Release((object)cm);
                Release((object)module);
                Release((object)components);
                Release((object)vbProject);
            }
        }

        /// <summary>直读 JYYJXaiState.gXaiState（运行未编辑模块的过程，不触发重编译），判断状态是否仍存活。</summary>
        private static void DumpState(dynamic app, string label)
        {
            try
            {
                var raw = app.Run("JYYJXaiState.XaiStateRaw");
                Console.WriteLine("  [" + label + "] gXaiState=" + (raw == null ? "null" : Convert.ToString(raw)));
            }
            catch (Exception ex)
            {
                Console.WriteLine("  [" + label + "] 读 gXaiState 失败: " + ex.Message);
            }
        }

        private static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o))
            {
                try { Marshal.FinalReleaseComObject(o); } catch { }
            }
        }
    }
}
