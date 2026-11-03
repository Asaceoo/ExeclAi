using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Addin;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Workflow;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace VerifyVerifyTools
{
    /// <summary>
    /// 真机验证 A+B 新增自检能力的全链路（与侧栏桥 CreateRegistry 同一条链路）：
    ///   A：verify_write_result / verify_data_clean / verify_pivot / verify_marker 四新工具
    ///      —— 用真实操作（写入/清洗/透视/标记）产出产物，再以「期望命中/期望给错」双向断言；
    ///   B：run_workflow 步骤自检注入 —— WorkflowExecutor 对带 Verify 的步骤自动执行自检，
    ///      达标则放行、未达标则停在当前步。
    /// 故意给错期望值确认自检能报 issue（而不是永远 pass）。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try
            {
                RunAll();
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            if (Failures.Count > 0)
            {
                foreach (var f in Failures) Console.WriteLine("FAILED: " + f);
                return 1;
            }

            Console.WriteLine("ALL VERIFY-TOOLS REAL-MACHINE CHECKS OK");
            return 0;
        }

        private static void RunAll()
        {
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();
                dynamic sheet1 = wb.Worksheets.Item[1];
                sheet1.Name = "vvt";
                Release((object)sheet1);

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyverifytools-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost(
                    (object)app,
                    repo,
                    ExcelToolHost.CreateWorkbookHash(workbookPath),
                    "vvt-session",
                    "自检");
                var registry = host.CreateRegistry();

                SetupData(registry);            // 源数据：写入/清洗/透视/标记 四类真实产物
                VerifyWrite(registry);          // VVT-A1 写入结果自检
                VerifyClean(registry);          // VVT-A2 数据清洗自检
                VerifyPivot(registry);          // VVT-A3 透视表自检
                VerifyMarker(registry);         // VVT-A4 标记自检
                VerifyWorkflowInjection(registry); // VVT-B 流程步骤自检注入
            }
            finally
            {
                try { dynamic active = app.ActiveWorkbook; if (active != null) { active.Close(false); Release((object)active); } }
                catch (Exception ex) { Console.WriteLine("NOTE  workbook-close: " + Safe(ex)); }
                try { if (app != null) { app.ScreenUpdating = true; app.Quit(); Release((object)app); } }
                catch (Exception ex) { Console.WriteLine("NOTE  app-quit: " + Safe(ex)); }
                KillExcel();
            }
        }

        // ===================== 数据准备 =====================

        private static void SetupData(ToolRegistry registry)
        {
            // 1) 透视数据源（含表头 + 4 行数据）
            WriteRange(registry, "vvt", "A1:B5", new object[,]
            {
                { "品类", "金额" },
                { "华东", 120D }, { "华东", 80D }, { "华北", 90D }, { "华南", 60D }
            });

            // 2) 清洗脏数据：A10 起放表头 + 含空行/文本数字/坏日期的脏数据
            WriteRange(registry, "vvt", "A10:C14", new object[,]
            {
                { "name", "date", "qty" },
                { "A", "2026-01-01", "12" },
                { null, null, null },
                { "B", "2026-2-3", 34 },
                { "C", "not-a-date", "abc" }
            });
            WriteRange(registry, "vvt", "D10", new object[,] { { "文本数字" } });

            // 3) 标记数据：A20 起含空值 + 离群值（数值主体 10~13 聚类，9999 为 IQR 离群点）
            WriteRange(registry, "vvt", "A20:B26", new object[,]
            {
                { "item", "val" },
                { "a", 10D }, { "b", 12D }, { "c", 11D }, { "d", 13D }, { "e", null }, { "f", 9999D }
            });
        }

        // ===================== A1 写入结果自检 =====================

        private static void VerifyWrite(ToolRegistry registry)
        {
            try
            {
                var address = "F2:G3";
                WriteRange(registry, "vvt", address, new object[,] { { "甲", 10D }, { "乙", 20D } });

                var hit = RunVerify(registry, "verify_write_result", new Dictionary<string, object>
                {
                    ["address"] = address,
                    ["expected_values"] = JArray.FromObject(new object[,] { { "甲", 10D }, { "乙", 20D } }),
                    ["expected_count"] = 4
                });
                Assert("vvt-a1-write-hit", hit != null && hit.Ok && V(hit, "verdict") == "ok",
                    "verdict=" + V(hit, "verdict") + " issues=" + V(hit, "issues"));

                var missValue = RunVerify(registry, "verify_write_result", new Dictionary<string, object>
                {
                    ["address"] = address,
                    ["expected_values"] = JArray.FromObject(new object[,] { { "甲", 999D }, { "乙", 20D } })
                });
                var missValueOk = missValue != null && missValue.Ok
                    && V(missValue, "verdict") != null && V(missValue, "verdict").Contains("value")
                    && V(missValue, "issues") != null && V(missValue, "issues").Contains("G2");
                Assert("vvt-a1-write-value-miss", missValueOk,
                    "verdict=" + V(missValue, "verdict") + " issues=" + V(missValue, "issues"));

                var missCount = RunVerify(registry, "verify_write_result", new Dictionary<string, object>
                {
                    ["address"] = address,
                    ["expected_count"] = 5
                });
                Assert("vvt-a1-write-count-miss", missCount != null && missCount.Ok
                    && V(missCount, "verdict") != null && V(missCount, "verdict").Contains("count"),
                    "verdict=" + V(missCount, "verdict") + " issues=" + V(missCount, "issues"));
            }
            catch (Exception ex) { Fail("vvt-a1-write", "HARNESS: " + Safe(ex)); }
        }

        // ===================== A2 数据清洗自检 =====================

        private static void VerifyClean(ToolRegistry registry)
        {
            try
            {
                // 先真实清洗：去空行（脏数据 A10:C14 → remove_blank_rows），产出应无整行空白
                if (!registry.TryGet("clean_excel_data", out var cleanTool))
                {
                    Fail("vvt-a2-clean-run", "clean_excel_data 未注册");
                    return;
                }

                var cleanRes = cleanTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = "vvt",
                    ["address"] = "A10:C14",
                    ["operations"] = new List<object>
                    {
                        new Dictionary<string, object> { ["tool"] = "remove_blank_rows" }
                    }
                });
                Assert("vvt-a2-clean-run", cleanRes != null && cleanRes.Ok,
                    "ok=" + (cleanRes?.Ok) + " err=" + (cleanRes == null || cleanRes.Ok ? "-" : cleanRes.ErrorCode));

                var hit = RunVerify(registry, "verify_data_clean", new Dictionary<string, object>
                {
                    ["address"] = "A10:C14",
                    ["header_rows"] = 1,
                    ["checks"] = JArray.FromObject(new[]
                    {
                        new { type = "no_blank_rows" }
                    })
                });
                Assert("vvt-a2-clean-hit", hit != null && hit.Ok && V(hit, "verdict") == "ok",
                    "verdict=" + V(hit, "verdict") + " issues=" + V(hit, "issues"));

                // 未清洗的脏列（date 仍为 "not-a-date"）→ dates_valid 应命中
                var missDate = RunVerify(registry, "verify_data_clean", new Dictionary<string, object>
                {
                    ["address"] = "A10:C14",
                    ["header_rows"] = 1,
                    ["checks"] = JArray.FromObject(new[]
                    {
                        new { type = "dates_valid", columns = new[] { 2 } }
                    })
                });
                Assert("vvt-a2-clean-date-miss", missDate != null && missDate.Ok
                    && V(missDate, "verdict") != null && V(missDate, "verdict").Contains("date")
                    && V(missDate, "issues") != null && V(missDate, "issues").Contains("日期"),
                    "verdict=" + V(missDate, "verdict") + " issues=" + V(missDate, "issues"));
            }
            catch (Exception ex) { Fail("vvt-a2-clean", "HARNESS: " + Safe(ex)); }
        }

        // ===================== A3 透视表自检 =====================

        private static void VerifyPivot(ToolRegistry registry)
        {
            try
            {
                if (!registry.TryGet("create_excel_pivot_summary", out var pivotTool))
                {
                    Fail("vvt-a3-pivot-run", "create_excel_pivot_summary 未注册");
                    return;
                }

                var created = pivotTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = "vvt",
                    ["source_address"] = "A1:B5",
                    ["target_address"] = "F6"
                });
                Assert("vvt-a3-pivot-run", created != null && created.Ok,
                    "ok=" + (created?.Ok) + " err=" + (created == null || created.Ok ? "-" : created.ErrorCode));

                var target = V(created, "target_address");
                var rowCount = V(created, "row_count");
                // 数据：华东 200、华北 90、华南 60 → row_count=3
                var hit = RunVerify(registry, "verify_pivot", new Dictionary<string, object>
                {
                    ["target_address"] = target,
                    ["expected_count"] = Convert.ToInt32(rowCount),
                    ["expected_rows"] = JArray.FromObject(new[]
                    {
                        new { label = "华东", value = 200D },
                        new { label = "华北", value = 90D }
                    })
                });
                Assert("vvt-a3-pivot-hit", hit != null && hit.Ok && V(hit, "verdict") == "ok"
                    && string.Equals(V(hit, "target_address"), target, StringComparison.OrdinalIgnoreCase),
                    "target=" + target + " row_count=" + rowCount + " verdict=" + V(hit, "verdict") + " issues=" + V(hit, "issues"));

                var missCount = RunVerify(registry, "verify_pivot", new Dictionary<string, object>
                {
                    ["target_address"] = target,
                    ["expected_count"] = 99
                });
                Assert("vvt-a3-pivot-count-miss", missCount != null && missCount.Ok
                    && V(missCount, "verdict") != null && V(missCount, "verdict").Contains("count"),
                    "verdict=" + V(missCount, "verdict") + " issues=" + V(missCount, "issues"));

                var missLabel = RunVerify(registry, "verify_pivot", new Dictionary<string, object>
                {
                    ["target_address"] = target,
                    ["expected_rows"] = JArray.FromObject(new[] { new { label = "不存在区域", value = 1D } })
                });
                Assert("vvt-a3-pivot-missing-miss", missLabel != null && missLabel.Ok
                    && V(missLabel, "verdict") != null && V(missLabel, "verdict").Contains("missing"),
                    "verdict=" + V(missLabel, "verdict") + " issues=" + V(missLabel, "issues"));
            }
            catch (Exception ex) { Fail("vvt-a3-pivot", "HARNESS: " + Safe(ex)); }
        }

        // ===================== A4 标记自检 =====================

        private static void VerifyMarker(ToolRegistry registry)
        {
            try
            {
                if (!registry.TryGet("flag_cells", out var flagTool))
                {
                    Fail("vvt-a4-marker-run", "flag_cells 未注册");
                    return;
                }

                // 指定高亮色：outlier=红 #FF0000、empty=蓝 #0000FF，保证可预期
                var flagged = flagTool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = "vvt",
                    ["address"] = "A20:B26",
                    ["flag_types"] = new List<object> { "outlier", "empty" },
                    ["fill_colors"] = new Dictionary<string, object> { ["outlier"] = "#FF0000", ["empty"] = "#0000FF" }
                });
                Assert("vvt-a4-marker-run", flagged != null && flagged.Ok,
                    "ok=" + (flagged?.Ok) + " err=" + (flagged == null || flagged.Ok ? "-" : flagged.ErrorCode + "|" + flagged.Data));

                // 实际标记数 = outlier + empty 计数之和
                var flaggedJson = V(flagged, "flagged");
                var expectedCount = 0;
                if (!string.IsNullOrEmpty(flaggedJson))
                {
                    var jo = JObject.Parse(flaggedJson);
                    foreach (var kv in jo) expectedCount += kv.Value.Value<int>();
                }

                var hit = RunVerify(registry, "verify_marker", new Dictionary<string, object>
                {
                    ["address"] = "A20:B26",
                    ["expected_count"] = expectedCount,
                    ["expected_fill_colors"] = new Dictionary<string, object> { ["red"] = "#FF0000", ["blue"] = "#0000FF" },
                    ["expected_marked"] = new List<object> { "B25", "B26" }
                });
                Assert("vvt-a4-marker-hit", hit != null && hit.Ok && V(hit, "verdict") == "ok",
                    "flagged=" + flaggedJson + " expected_count=" + expectedCount + " verdict=" + V(hit, "verdict") + " issues=" + V(hit, "issues"));

                var missCount = RunVerify(registry, "verify_marker", new Dictionary<string, object>
                {
                    ["address"] = "A20:B26",
                    ["expected_count"] = expectedCount + 1
                });
                Assert("vvt-a4-marker-count-miss", missCount != null && missCount.Ok
                    && V(missCount, "verdict") != null && V(missCount, "verdict").Contains("count"),
                    "verdict=" + V(missCount, "verdict") + " issues=" + V(missCount, "issues"));

                // 期望颜色给错 → color 标（实际只有红/蓝，期望绿）
                var missColor = RunVerify(registry, "verify_marker", new Dictionary<string, object>
                {
                    ["address"] = "A20:B26",
                    ["expected_fill_colors"] = new Dictionary<string, object> { ["green"] = "#00FF00" }
                });
                Assert("vvt-a4-marker-color-miss", missColor != null && missColor.Ok
                    && V(missColor, "verdict") != null && V(missColor, "verdict").Contains("color"),
                    "verdict=" + V(missColor, "verdict") + " issues=" + V(missColor, "issues"));
            }
            catch (Exception ex) { Fail("vvt-a4-marker", "HARNESS: " + Safe(ex)); }
        }

        // ===================== B 流程步骤自检注入 =====================

        private static void VerifyWorkflowInjection(ToolRegistry registry)
        {
            try
            {
                // 1) 达标自检：主步骤成功后 verify_pivot 注入并通过 → 流程完成
                //    经 {selection_sheet} 占位符注入目标表（与内置模板同路径），并引用 {step0.*} 前一步产物。
                var passOverrides = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["selection"] = "A1:B5",
                    ["selection_sheet"] = "vvt"
                };
                var passTemplate = new WorkflowTemplate
                {
                    Name = "vvt-透视自检通过",
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep
                        {
                            Tool = "create_excel_pivot_summary",
                            Description = "生成透视",
                            Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" },
                            Verify = new WorkflowVerify
                            {
                                Tool = "verify_pivot",
                                Description = "透视汇总自检",
                                Arguments = new Dictionary<string, object>
                                {
                                    ["target_address"] = "{step0.target_address}",
                                    ["expected_count"] = "{step0.row_count}",
                                    ["sheet"] = "{selection_sheet}"
                                }
                            }
                        }
                    }
                };

                var executor = new WorkflowExecutor(registry, (_, __) => Task.FromResult(true));
                var pass = executor.RunAsync(passTemplate, passOverrides, null, 0, CancellationToken.None).GetAwaiter().GetResult();
                var passOk = pass.Completed
                    && pass.Steps.Count == 1
                    && pass.Steps[0].Verify != null
                    && pass.Steps[0].Verify.Passed
                    && pass.Steps[0].Verify.Verdict == "ok";
                Assert("vvt-b-workflow-verify-pass", passOk,
                    "Completed=" + pass.Completed + " Passed=" + (pass.Steps.Count > 0 && pass.Steps[0].Verify != null ? pass.Steps[0].Verify.Passed.ToString() : "n/a")
                    + " verdict=" + (pass.Steps.Count > 0 && pass.Steps[0].Verify != null ? pass.Steps[0].Verify.Verdict : "n/a"));

                // 2) 未达标自检：expected_count 给错 → 停在当前步，后续步骤不执行
                var neverCalls = 0;
                registry.Register(new ToolDescriptor("vvt_never", "not reached", RiskLevel.ReadOnly,
                    _ => { neverCalls++; return new ToolResult { Ok = true }; }));

                var failTemplate = new WorkflowTemplate
                {
                    Name = "vvt-透视自检未达标",
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep
                        {
                            Tool = "create_excel_pivot_summary",
                            Description = "生成透视",
                            Arguments = new Dictionary<string, object> { ["source_address"] = "{selection}", ["sheet"] = "{selection_sheet}" },
                            Verify = new WorkflowVerify
                            {
                                Tool = "verify_pivot",
                                Description = "透视汇总自检",
                                Arguments = new Dictionary<string, object>
                                {
                                    ["target_address"] = "{step0.target_address}",
                                    ["expected_count"] = 99,
                                    ["sheet"] = "{selection_sheet}"
                                }
                            }
                        },
                        new WorkflowStep { Tool = "vvt_never", Description = "不应执行" }
                    }
                };
                var fail = executor.RunAsync(failTemplate, passOverrides, null, 0, CancellationToken.None).GetAwaiter().GetResult();
                var failOk = !fail.Completed
                    && fail.FailedIndex == 0
                    && fail.Steps.Count == 1
                    && fail.Steps[0].Verify != null
                    && !fail.Steps[0].Verify.Passed
                    && neverCalls == 0;
                Assert("vvt-b-workflow-verify-block", failOk,
                    "Completed=" + fail.Completed + " FailedIndex=" + (fail.FailedIndex?.ToString() ?? "null")
                    + " steps=" + fail.Steps.Count + " neverCalls=" + neverCalls
                    + " verdict=" + (fail.Steps[0].Verify?.Verdict ?? "n/a"));

                // 3) 空 Tool 自检 → fail-closed 阻断，不静默放行
                var emptyToolTemplate = new WorkflowTemplate
                {
                    Name = "vvt-自检未配置工具",
                    Steps = new List<WorkflowStep>
                    {
                        new WorkflowStep
                        {
                            Tool = "vvt_noop",
                            Description = "任意主步骤",
                            Verify = new WorkflowVerify { Description = "未指定自检工具" }
                        }
                    }
                };
                registry.Register(new ToolDescriptor("vvt_noop", "noop", RiskLevel.ReadOnly,
                    _ => new ToolResult { Ok = true }));
                var emptyToolRun = executor.RunAsync(emptyToolTemplate, null, null, 0, CancellationToken.None).GetAwaiter().GetResult();
                var emptyToolOk = !emptyToolRun.Completed
                    && emptyToolRun.FailedIndex == 0
                    && emptyToolRun.Steps[0].Verify != null
                    && !emptyToolRun.Steps[0].Verify.Passed
                    && emptyToolRun.Steps[0].Verify.Result?.ErrorCode == "VERIFY_NOT_CONFIGURED";
                Assert("vvt-b-workflow-verify-fail-closed", emptyToolOk,
                    "Completed=" + emptyToolRun.Completed + " FailedIndex=" + (emptyToolRun.FailedIndex?.ToString() ?? "null")
                    + " err=" + (emptyToolRun.Steps[0].Verify?.Result?.ErrorCode ?? "n/a"));
            }
            catch (Exception ex) { Fail("vvt-b-workflow", "HARNESS: " + Safe(ex)); }
        }

        // ===================== 辅助 =====================

        private static void WriteRange(ToolRegistry registry, string sheet, string address, object[,] values)
        {
            if (!registry.TryGet("write_excel_range", out var tool))
            {
                throw new InvalidOperationException("write_excel_range 未注册");
            }

            var res = tool.Execute(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = sheet,
                ["address"] = address,
                ["values"] = values
            });
            if (res == null || !res.Ok)
            {
                throw new InvalidOperationException("write_excel_range 失败: " + (res?.Data ?? res?.ErrorCode));
            }
        }

        private static ToolResult RunVerify(ToolRegistry registry, string name, Dictionary<string, object> extra)
        {
            if (!registry.TryGet(name, out var tool))
            {
                Fail(name + "-notfound", name + " 未注册");
                return null;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { ["sheet"] = "vvt" };
            foreach (var kv in extra) args[kv.Key] = kv.Value;
            return tool.Execute(args);
        }

        private static string V(ToolResult r, string key)
        {
            if (r?.Data == null || r.Data is string) return null;
            try
            {
                var jo = JObject.FromObject(r.Data);
                var token = jo[key];
                return token == null ? null : token.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static string Safe(Exception ex)
        {
            try
            {
                var inner = ex;
                while (inner.InnerException != null) inner = inner.InnerException;
                return inner.GetType().Name + ": " + (inner.Message ?? string.Empty);
            }
            catch { return "unknown"; }
        }

        private static void Assert(string name, bool ok, string detail)
        {
            if (ok) Pass(name, detail);
            else Fail(name, detail);
        }

        private static void Pass(string name, string detail)
        {
            _passed++;
            Console.WriteLine("PASS  " + name + "  [" + detail + "]");
        }

        private static void Fail(string name, string detail)
        {
            _failed++;
            Failures.Add(name + " -> " + detail);
            Console.WriteLine("FAIL  " + name + "  [" + detail + "]");
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }

        private static void KillExcel()
        {
            try
            {
                foreach (var p in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {
                    try { p.Kill(); } catch { }
                }
            }
            catch { }
        }
    }
}
