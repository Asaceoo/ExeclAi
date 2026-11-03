using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace RealMachineVerify
{
    /// <summary>
    /// 真机验证 harness（踩坑手册 §3.9/§5 稳妥做法）：
    /// 独立 out-of-proc Excel 实例 + 新建未存盘内存工作簿 + AutomationSecurity=1，
    /// 用真实的 ExcelToolHost.CreateRegistry() 组装工具表，跑通「主工具 → AutoVerifyPolicy → verify_* 回读判定」
    /// 完整自动自检链路（含负向用例与进程无泄漏检查）。不依赖模型 API，不触碰已安装的加载项。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static HashSet<int> _excelPidsBefore = new HashSet<int>();
        private static HashSet<int> _spawnedPids = new HashSet<int>();

        /// <summary>harness 启动 Excel 前记录已有进程，启动后记录新建的 PID（踩坑手册 §5.2 泄漏兜底用）。</summary>
        private static HashSet<int> SnapshotNewExcelPids()
        {
            var before = _excelPidsBefore;
            var spawned = new HashSet<int>();
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    if (!before.Contains(process.Id))
                    {
                        spawned.Add(process.Id);
                    }
                }
                finally
                {
                    process.Dispose();
                }
            }

            return spawned;
        }

        /// <summary>强杀本 harness 启动的残留 EXCEL.EXE，返回被杀的 PID 列表。</summary>
        private static List<int> KillSpawnedExcel()
        {
            var killed = new List<int>();
            var spawned = _spawnedPids;
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    if (spawned.Contains(process.Id))
                    {
                        process.Kill();
                        killed.Add(process.Id);
                    }
                }
                catch
                {
                }
                finally
                {
                    process.Dispose();
                }
            }

            return killed;
        }

        private static int Main()
        {
            Console.WriteLine("== 真机验证：执行层自动自检链路 ==");
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                _excelPidsBefore.Add(process.Id);
                process.Dispose();
            }

            Console.WriteLine("[env] 运行前 EXCEL.EXE 进程数：" + _excelPidsBefore.Count);

            dynamic app = null;
            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    Console.WriteLine("[fatal] 未找到 Excel.Application COM 组件（本机未安装 Excel？）");
                    return 2;
                }

                app = Activator.CreateInstance(excelType);
                app.Visible = false;
                app.DisplayAlerts = false;
                app.AutomationSecurity = 1; // msoAutomationSecurityLow：未存盘工作簿内允许宏

                _spawnedPids = SnapshotNewExcelPids();
                app.Workbooks.Add(); // 新建未存盘工作簿，全程不 SaveAs
                var host = new ExcelToolHost(app);
                var registry = host.CreateRegistry();
                Console.WriteLine("[env] 工具注册表就绪，共 " + registry.Tools.Count + " 个工具。\n");

                // 用例 1（正向）：区域值写入 → verify_write_result 逐格回读核对。
                var writeArgs = Range(
                    ("address", "$A$1:$B$3"),
                    ("values", new JArray
                    {
                        new JArray { "品名", "数量" },
                        new JArray { "苹果", 3 },
                        new JArray { "梨", 5 }
                    }));
                var writeResult = ExecuteTool(registry, "write_excel_range", writeArgs);
                Expect(writeResult.Ok, "用例1 准备：write_excel_range 写入", Describe(writeResult));
                RunAutoVerifyChain(registry, "write_excel_range", writeArgs, writeResult,
                    expectOk: true, label: "用例1 值写入自动自检");

                // 用例 2（负向）：实际写入 hello，但把期望篡改为 world —— 自检必须抓到回读不符。
                var cellArgs = Range(("address", "C1"), ("value", "hello"));
                var cellResult = ExecuteTool(registry, "write_excel_cell", cellArgs);
                Expect(cellResult.Ok, "用例2 准备：write_excel_cell 写入", Describe(cellResult));
                RunAutoVerifyChain(registry, "write_excel_cell", cellArgs, cellResult,
                    expectOk: false, label: "用例2 负向：期望与实际不符必须报错",
                    mutate: args => args["expected_values"] = new JArray { new JArray { "world" } });

                // 用例 3：图表生成 → verify_chart（存在性/数据点/类型交叉核对）。
                var chartArgs = Range(
                    ("source_address", "A1:B3"),
                    ("chart_type", "ColumnClustered"),
                    ("title", "销量图"));
                var chartResult = ExecuteTool(registry, "create_excel_chart", chartArgs);
                Expect(chartResult.Ok, "用例3 准备：create_excel_chart", Describe(chartResult));
                RunAutoVerifyChain(registry, "create_excel_chart", chartArgs, chartResult,
                    expectOk: true, label: "用例3 图表自动自检");

                // 用例 4：排序 → verify_sort_filter（排序列单调性回读）。
                var sortArgs = Range(
                    ("address", "A1:B3"),
                    ("sort_column", 2),
                    ("direction", "desc"),
                    ("header_rows", 1));
                var sortResult = ExecuteTool(registry, "sort_range", sortArgs);
                Expect(sortResult.Ok, "用例4 准备：sort_range 排序", Describe(sortResult));
                RunAutoVerifyChain(registry, "sort_range", sortArgs, sortResult,
                    expectOk: true, label: "用例4 排序自动自检");

                // 用例 5：格式设置 → verify_cell_format（属性逐项回读）。
                var formatArgs = Range(
                    ("address", "$A$1:$B$1"),
                    ("bold", true),
                    ("number_format", "0.00"));
                var formatResult = ExecuteTool(registry, "format_excel_range", formatArgs);
                Expect(formatResult.Ok, "用例5 准备：format_excel_range", Describe(formatResult));
                RunAutoVerifyChain(registry, "format_excel_range", formatArgs, formatResult,
                    expectOk: true, label: "用例5 格式自动自检");

                // 用例 6：透视汇总（显式 target_address）→ 结果富化 → verify_pivot 行数断言。
                var pivotArgs = Range(
                    ("source_address", "A1:B3"),
                    ("label_column", 1),
                    ("value_column", 2),
                    ("target_address", "$E$1"));
                var pivotResult = ExecuteTool(registry, "create_excel_pivot_summary", pivotArgs);
                Expect(pivotResult.Ok, "用例6 准备：create_excel_pivot_summary", Describe(pivotResult));
                RunAutoVerifyChain(registry, "create_excel_pivot_summary", pivotArgs, pivotResult,
                    expectOk: true, label: "用例6 透视自动自检（显式输出地址）");

                // 用例 7：透视汇总不传 target_address（自动落点）→ 从结果富化出地址与行数。
                var pivotArgs2 = Range(
                    ("source_address", "A1:B3"),
                    ("label_column", 1),
                    ("value_column", 2));
                var pivotResult2 = ExecuteTool(registry, "create_excel_pivot_summary", pivotArgs2);
                Expect(pivotResult2.Ok, "用例7 准备：create_excel_pivot_summary（自动落点）", Describe(pivotResult2));
                RunAutoVerifyChain(registry, "create_excel_pivot_summary", pivotArgs2, pivotResult2,
                    expectOk: true, label: "用例7 透视自动自检（自动落点地址富化）");

                // 踩坑库对真实错误文本的识别（纯逻辑，用真实日志样例）。
                var hits = PitfallLibrary.Match("write_excel_range 失败：0x800A9C68 cannot run the macro");
                Expect(hits.Count > 0, "用例8 踩坑库签名识别", "命中 " + hits.Count + " 条");

                // 用例 9：批量回写（值格 + 公式格混合）→ 逐格自动自检（2 条计划）。
                var batchArgs = Range(("cells", new JArray
                {
                    new JObject { ["address"] = "D1", ["value"] = "标记" },
                    new JObject { ["address"] = "D2", ["value"] = "=D3*2", ["value_mode"] = "formula" }
                }));
                var batchResult = ExecuteTool(registry, "write_excel_cells_batch", batchArgs);
                Expect(batchResult.Ok, "用例9 准备：write_excel_cells_batch", Describe(batchResult));
                RunAutoVerifyChain(registry, "write_excel_cells_batch", batchArgs, batchResult,
                    expectOk: true, label: "用例9 批量回写自动自检（逐格）");

                // 用例 10：追加行 → 目标地址从结果 AffectedRanges 富化 → verify_write_result。
                var appendArgs = Range(
                    ("start_cell", "A1"),
                    ("values", new JArray { new JArray { "橙子", 7 } }));
                var appendResult = ExecuteTool(registry, "append_excel_rows", appendArgs);
                Expect(appendResult.Ok, "用例10 准备：append_excel_rows", Describe(appendResult));
                RunAutoVerifyChain(registry, "append_excel_rows", appendArgs, appendResult,
                    expectOk: true, label: "用例10 追加行自动自检（地址富化）");

                // 用例 11：新增工作表 → verify_structure 存在性；删除后 → 不存在。
                var addArgs = Range(("name", "自检测试表"));
                var addResult = ExecuteTool(registry, "add_excel_worksheet", addArgs);
                Expect(addResult.Ok, "用例11 准备：add_excel_worksheet", Describe(addResult));
                RunAutoVerifyChain(registry, "add_excel_worksheet", addArgs, addResult,
                    expectOk: true, label: "用例11 新增表结构自检");
                var deleteArgs = Range(("sheet", "自检测试表"));
                var deleteResult = ExecuteTool(registry, "delete_excel_worksheet", deleteArgs);
                Expect(deleteResult.Ok, "用例11 准备：delete_excel_worksheet", Describe(deleteResult));
                RunAutoVerifyChain(registry, "delete_excel_worksheet", deleteArgs, deleteResult,
                    expectOk: true, label: "用例11 删除表结构自检（缺席断言）");

                // 用例 12：透视数量守恒对账（明细列合计 == 汇总列合计）。
                ExecuteTool(registry, "write_excel_range", Range(
                    ("address", "$G$1:$H$3"),
                    ("values", new JArray { new JArray { "品名", "金额" }, new JArray { "甲", 4 }, new JArray { "乙", 6 } })));
                var pivot2Args = Range(
                    ("source_address", "G1:H3"),
                    ("label_column", 1),
                    ("value_column", 2),
                    ("target_address", "$G$5"));
                var pivot2Result = ExecuteTool(registry, "create_excel_pivot_summary", pivot2Args);
                Expect(pivot2Result.Ok, "用例12 准备：create_excel_pivot_summary", Describe(pivot2Result));
                RunAutoVerifyChain(registry, "create_excel_pivot_summary", pivot2Args, pivot2Result,
                    expectOk: true, label: "用例12 透视守恒对账（明细 10 == 汇总 10）");

                // 用例 13（负向）：超大区域写入必须被确定性拦截（单次上限 400,000 格）。
                var oversized = new JArray();
                for (var row = 0; row < 401; row++)
                {
                    var cells = new JArray();
                    for (var col = 0; col < 1000; col++)
                    {
                        cells.Add(0);
                    }

                    oversized.Add(cells);
                }

                var oversizedResult = ExecuteTool(registry, "write_excel_range", Range(
                    ("address", "$J$1:$K$402"), ("values", oversized)));
                Expect(!oversizedResult.Ok, "用例13 负向：401,000 格写入必须被拒绝", Describe(oversizedResult));

                // 用例 14（负向）：文本型数字必须被 type_aware 抓到——单元格设为文本格式后写入 "3"，
                // 期望值按数值 3 断言，verdict 应为 type（常规比对会判等，只有类型感知能发现）。
                var textFormatResult = ExecuteTool(registry, "format_excel_range", Range(
                    ("address", "B7"), ("number_format", "@")));
                Expect(textFormatResult.Ok, "用例14 准备：B7 设为文本格式", Describe(textFormatResult));
                var textNumArgs = Range(("address", "B7"), ("value", "3"));
                var textNumResult = ExecuteTool(registry, "write_excel_cell", textNumArgs);
                Expect(textNumResult.Ok, "用例14 准备：B7 写入文本型数字 \"3\"", Describe(textNumResult));
                RunAutoVerifyChain(registry, "write_excel_cell", textNumArgs, textNumResult,
                    expectOk: false, label: "用例14 负向：文本型数字被 type_aware 抓到",
                    mutate: args => args["expected_values"] = new JArray { new JArray { 3 } });

                // 用例 15（负向）：公式产生 #DIV/0! 必须被 verify_formula 抓到。
                var div0Args = Range(("address", "C7"), ("value", "=1/0"));
                var div0Result = ExecuteTool(registry, "write_excel_cell", div0Args);
                Expect(div0Result.Ok, "用例15 准备：C7 写入 =1/0", Describe(div0Result));
                RunAutoVerifyChain(registry, "write_excel_cell", div0Args, div0Result,
                    expectOk: false, label: "用例15 负向：#DIV/0! 被 verify_formula 抓到");

                // 用例 16（负向）：保护表写入必须被拒绝（结构护栏）。
                var protectAdd = ExecuteTool(registry, "add_excel_worksheet", Range(("name", "保护测试")));
                Expect(protectAdd.Ok, "用例16 准备：新增工作表", Describe(protectAdd));
                dynamic protectSheet = app.ActiveWorkbook.Worksheets.Item["保护测试"];
                try
                {
                    protectSheet.Protect(); // harness 直接设置保护态（工具层无保护开关，属真机环境构造）
                }
                finally
                {
                    // 手册 §5.1 无状态纪律：COM 对象用完即放，即使本进程最终会被强杀。
                    try
                    {
                        if (System.Runtime.InteropServices.Marshal.IsComObject(protectSheet))
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(protectSheet);
                        }
                    }
                    catch
                    {
                    }
                }

                var protectedWrite = ExecuteTool(registry, "write_excel_range", Range(
                    ("sheet", "保护测试"), ("address", "A1"), ("values", new JArray { new JArray { "x" } })));
                Expect(!protectedWrite.Ok && protectedWrite.ErrorCode == "SHEET_PROTECTED",
                    "用例16 负向：保护表写入必须返回 SHEET_PROTECTED（M17 集中守卫）", Describe(protectedWrite));

                // M17：第二类写工具同口径拦截（flag_cells 直接写 Interior）。
                var protectedFlag = ExecuteTool(registry, "flag_cells", Range(("sheet", "保护测试"), ("address", "A1:B2")));
                Expect(!protectedFlag.Ok && protectedFlag.ErrorCode == "SHEET_PROTECTED",
                    "用例16b 负向：保护表标记必须返回 SHEET_PROTECTED", Describe(protectedFlag));

                // M17：豁免工具不受拦截——freeze_panes 为窗口视图操作，保护态下原生可用。
                var protectedFreeze = ExecuteTool(registry, "freeze_panes", Range(("sheet", "保护测试"), ("cell", "B2")));
                Expect(protectedFreeze.Ok, "用例16c 豁免：保护表冻结窗格不受守卫拦截", Describe(protectedFreeze));

                // 清理保护表（Unprotect + 删除），否则活动表停留在受保护表上，后续写入用例全部被拒。
                try
                {
                    dynamic ps = app.ActiveWorkbook.Worksheets.Item["保护测试"];
                    try { ps.Unprotect(); } finally
                    {
                        try
                        {
                            if (System.Runtime.InteropServices.Marshal.IsComObject(ps))
                            {
                                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ps);
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // M17：解保后同一写入恢复 OK——守卫不误伤正常路径。
                var unprotectWrite = ExecuteTool(registry, "write_excel_range", Range(
                    ("sheet", "保护测试"), ("address", "A1"), ("values", new JArray { new JArray { "x" } })));
                Expect(unprotectWrite.Ok, "用例16d 收尾：解保后写入恢复（守卫不误伤）", Describe(unprotectWrite));
                var deleteProtectSheet = ExecuteTool(registry, "delete_excel_worksheet", Range(("sheet", "保护测试")));
                Expect(deleteProtectSheet.Ok, "用例16 收尾：删除保护测试表", Describe(deleteProtectSheet));

                // 用例 17：自检覆盖审计——所有写能力工具必须有自动自检映射或带理由的豁免，
                // 防止新工具带着「永不自检」上线。用真实注册表逐一核对。
                var unmapped = new List<string>();
                foreach (var descriptor in registry.Tools)
                {
                    if (descriptor.RiskLevel == RiskLevel.ReadOnly ||
                        AutoVerifyPolicy.MappedPrimaryTools.Contains(descriptor.Name) ||
                        AutoVerifyPolicy.IsExemptedFromAutoVerify(descriptor.Name))
                    {
                        continue;
                    }

                    unmapped.Add(descriptor.Name + "(" + descriptor.RiskLevel + ")");
                }

                Expect(unmapped.Count == 0, "用例17 自检覆盖审计：写能力工具全部有映射或豁免",
                    unmapped.Count == 0 ? "覆盖完整" : "未覆盖：" + string.Join(", ", unmapped));

                // 用例 18：回滚端到端——挂接操作历史的宿主里「写入 → 覆写 → undo → 断言恢复为第一次写入」。
                var rollbackDb = Path.Combine(Path.GetTempPath(), "xai-rollback-" + Guid.NewGuid().ToString("N") + ".db");
                var rollbackHost = new ExcelToolHost(
                    app, new OperationRepository(rollbackDb), "harness-hash", "rollback-session");
                var rollbackRegistry = rollbackHost.CreateRegistry();
                try
                {
                    var firstValues = new JArray { new JArray { "第", 1 }, new JArray { "次", 11 } };
                    var secondValues = new JArray { new JArray { "第", 2 }, new JArray { "次", 22 } };
                    var write1 = ExecuteTool(rollbackRegistry, "write_excel_range", Range(
                        ("address", "$N$1:$O$2"), ("values", firstValues)));
                    Expect(write1.Ok, "用例18 准备：第一次写入", Describe(write1));
                    var write2 = ExecuteTool(rollbackRegistry, "write_excel_range", Range(
                        ("address", "$N$1:$O$2"), ("values", secondValues)));
                    Expect(write2.Ok, "用例18 准备：第二次写入（覆盖）", Describe(write2));

                    var opId = ExtractOperationId(write2);
                    Expect(!string.IsNullOrWhiteSpace(opId), "用例18 准备：写结果带 operation_id", opId ?? "(null)");

                    var undoResult = ExecuteTool(rollbackRegistry, "undo_excel_operation", Range(("operation_id", opId)));
                    Expect(undoResult.Ok, "用例18 回滚端到端：undo 执行", Describe(undoResult));

                    // 回读断言：区域应恢复为第一次写入的值。
                    var verifyAfterUndo = ExecuteTool(rollbackRegistry, "verify_write_result", Range(
                        ("address", "$N$1:$O$2"), ("expected_values", firstValues)));
                    var verdictAfterUndo = ExtractVerdict(verifyAfterUndo);
                    Expect(verifyAfterUndo.Ok && verdictAfterUndo == "ok",
                        "用例18 回滚端到端：回滚后区域恢复为第一次写入的值",
                        verdictAfterUndo + "；issues=" + DescribeIssues(verifyAfterUndo));

                    // 用例 19：全功能健康巡检（真实 Excel + 临时工作簿，18 个子系统逐项体检）。
                    var sweep = new ExcelAiAssistant.Addin.FeatureHealthSweep(app);
                    var sweepResults = sweep.Run();
                    foreach (var item in sweepResults)
                    {
                        Expect(item.Ok, "用例19 巡检[" + item.Name + "]", item.Detail);
                    }
                    var failedCount = sweepResults.Count(item => !item.Ok);
                    Expect(failedCount == 0, "用例19 全功能巡检汇总", failedCount + " 项失败 / 共 " + sweepResults.Count + " 项");

                    // 用例 20：功能区守卫/自检/样本链路真机端到端（弹窗本身需人工点击，此处直接走
                    // 服务层同款决策逻辑 + 真实 COM 数据，验证「守卫拦截判定 / 自检回读断言 / 样本锚点」成立）。
                    // 20a：text_merge 目标列占用守卫——目标矩形预算 → 空放行 → 有数据拦截。
                    ExecuteTool(registry, "write_excel_range", Range(
                        ("address", "$A$10:$B$12"),
                        ("values", new JArray { new JArray { "品名", "数量" }, new JArray { "苹果", 3 }, new JArray { "梨", 5 } })));
                    var mergeSel = new ParamSelection { SheetName = "Sheet1", Address = "$A$10:$B$12", Rows = 3, Cols = 2 };
                    var mergeArgs = Range(("target_column", 3));
                    var guardRectOk = RibbonCommandPlanner.TryGetTargetRect("text_merge", mergeSel, null, mergeArgs, 10, 1,
                        out var guardSheet, out var guardTop, out var guardLeft, out var guardRows, out var guardCols);
                    Expect(guardRectOk && guardTop == 10 && guardLeft == 3 && guardRows == 3 && guardCols == 1,
                        "用例20a 目标矩形预算（text_merge → C10:C12）",
                        guardRectOk ? $"{guardSheet}!{guardTop},{guardLeft} {guardRows}x{guardCols}" : "TryGetTargetRect=false");

                    var emptyGuardBlock = ReadBlockValues(app, "Sheet1", 10, 3, 3, 1);
                    Expect(!RibbonCommandPlanner.MatrixHasData(emptyGuardBlock),
                        "用例20a 目标区为空 → 守卫放行（IsTargetOccupied=false 前提）", "MatrixHasData=false");
                    ExecuteTool(registry, "write_excel_cell", Range(("address", "C11"), ("value", "占用")));
                    var occupiedGuardBlock = ReadBlockValues(app, "Sheet1", 10, 3, 3, 1);
                    Expect(RibbonCommandPlanner.MatrixHasData(occupiedGuardBlock),
                        "用例20a 目标区已有数据 → 守卫拦截（IsTargetOccupied=true 前提）", "MatrixHasData=true");

                    // 20b：fill_random_numbers 执行 → 回读值域自检通过 → 样本锚点提取；篡改越界值必须被抓到。
                    var randomResult = ExecuteTool(registry, "fill_random_numbers", Range(
                        ("address", "$E$10:$F$12"), ("min", 1), ("max", 10), ("decimals", 1)));
                    Expect(randomResult.Ok, "用例20b 准备：fill_random_numbers", Describe(randomResult));
                    var randomBlock = ReadBlockValues(app, "Sheet1", 10, 5, 3, 2);
                    var rangeFail = RibbonCommandPlanner.VerifyRandomRange(randomBlock, 1, 10, 1);
                    Expect(rangeFail == null, "用例20b 随机数回读值域 ∈ [1,10]（自检通过）", rangeFail ?? "通过");
                    var samples = RibbonCommandPlanner.PickSampleValues(randomBlock, 3);
                    Expect(samples.Count > 0 && samples.Count <= 3, "用例20b 样本锚点提取 1-3 个",
                        samples.Count + " 个：" + string.Join("，", samples));
                    ExecuteTool(registry, "write_excel_cell", Range(("address", "E10"), ("value", 99)));
                    var tamperedBlock = ReadBlockValues(app, "Sheet1", 10, 5, 3, 2);
                    var tamperFail = RibbonCommandPlanner.VerifyRandomRange(tamperedBlock, 1, 10, 1);
                    Expect(tamperFail != null, "用例20b 负向：篡改值 99 被自检抓到", tamperFail ?? "(未抓到)");

                    // 用例 21：sort_range 回滚端到端——排序结果必须带 operation_id（回归锚点），
                    // 且 undo 后区域恢复排序前原状。曾因 ExcelDataToolsService 丢弃 writer 的
                    // operation_id 而不可回滚，此用例锁死该链路。
                    var sortBlockArgs = Range(
                        ("address", "$A$20:$B$22"),
                        ("values", new JArray
                        {
                            new JArray { "丙", 3 },
                            new JArray { "甲", 1 },
                            new JArray { "乙", 2 }
                        }));
                    var sortBlockResult = ExecuteTool(rollbackRegistry, "write_excel_range", sortBlockArgs);
                    Expect(sortBlockResult.Ok, "用例21 准备：写入排序测试块", Describe(sortBlockResult));
                    var sortBlockBefore = ReadBlockValues(app, "Sheet1", 20, 1, 3, 2);

                    var sort21Args = Range(
                        ("address", "$A$20:$B$22"),
                        ("sort_column", 2),
                        ("direction", "desc"),
                        ("header_rows", 0));
                    var sort21Result = ExecuteTool(rollbackRegistry, "sort_range", sort21Args);
                    Expect(sort21Result.Ok, "用例21 准备：sort_range 按数量降序", Describe(sort21Result));
                    var sort21OpId = ExtractOperationId(sort21Result);
                    Expect(!string.IsNullOrWhiteSpace(sort21OpId),
                        "用例21 排序结果必须带 operation_id（回归锚点）", sort21OpId ?? "(null)");

                    var undoSort21 = ExecuteTool(rollbackRegistry, "undo_excel_operation",
                        Range(("operation_id", sort21OpId)));
                    Expect(undoSort21.Ok, "用例21 回滚 sort_range", Describe(undoSort21));
                    var sort21After = ReadBlockValues(app, "Sheet1", 20, 1, 3, 2);
                    var sort21Restored = CompareMatrices(sort21After, sortBlockBefore);
                    Expect(sort21Restored, "用例21 回滚后排序测试块恢复原状",
                        sort21Restored ? "矩阵一致" : "矩阵不一致：" + MatrixText(sort21After) + " vs " + MatrixText(sortBlockBefore));

                    // 用例 22：函数库「完整占位公式」插入链路——宿主 InsertFormulaToRange 正是侧边栏
                    // InsertFormula（SidebarBridge）的内部落点（设计-函数库面板.md §6）。写入完整占位公式
                    // 必须被 Excel 接受（半截骨架会被拒），且带 operation_id 可回滚。
                    var fxArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["address"] = "$S$1",
                        ["formula"] = "=SUM(A1:A9)",
                        ["sheet"] = "Sheet1"
                    };
                    var fxResult = rollbackHost.InsertFormulaToRange(fxArgs);
                    Expect(fxResult.Ok, "用例22 函数库完整占位公式插入（InsertFormulaToRange）", Describe(fxResult));
                    var fxInserted = ReadFormula(app, "Sheet1", 1, 19);
                    Expect(fxInserted == "=SUM(A1:A9)",
                        "用例22 公式以公式形态落盘（非半截/非文本）", fxInserted ?? "(null)");
                    var fxOpId = ExtractOperationId(fxResult);
                    Expect(!string.IsNullOrWhiteSpace(fxOpId),
                        "用例22 插入结果带 operation_id（回滚锚点）", fxOpId ?? "(null)");
                    var fxUndo = ExecuteTool(rollbackRegistry, "undo_excel_operation",
                        Range(("operation_id", fxOpId)));
                    Expect(fxUndo.Ok, "用例22 回滚公式插入", Describe(fxUndo));
                    var fxAfterUndo = ReadFormula(app, "Sheet1", 1, 19);
                    Expect(string.IsNullOrEmpty(fxAfterUndo),
                        "用例22 回滚后 S1 恢复为空", fxAfterUndo ?? "(非空)");

                    // 用例 23：图片插入 + verify_image 自动自检链路（M13：insert_image）。
                    var m13Png = Path.Combine(Path.GetTempPath(), "rmv_m13.png");
                    File.WriteAllBytes(m13Png, Convert.FromBase64String(
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="));
                    var imgArgs = Range(
                        ("sheet", "Sheet1"),
                        ("image_path", m13Png),
                        ("anchor", "$E$10"),
                        ("width", 40),
                        ("height", 30));
                    var imgResult = ExecuteTool(registry, "insert_image", imgArgs);
                    Expect(imgResult.Ok, "用例23 准备：insert_image 锚点 E10 插入图片", Describe(imgResult));
                    RunAutoVerifyChain(registry, "insert_image", imgArgs, imgResult,
                        expectOk: true, label: "用例23 图片自动自检（锚点/尺寸）");

                    // 用例 23 负路径（M13-R 审查闭环）：真实报错必须是显式错误码而非裸 COM 异常。
                    var negZero = ExecuteTool(registry, "insert_image", Range(
                        ("sheet", "Sheet1"), ("image_path", m13Png), ("anchor", "A20"), ("width", 0)));
                    Expect(!negZero.Ok && Convert.ToString(negZero.ErrorCode) == "INVALID_ARGS",
                        "用例23 负路径：width=0 显式报 INVALID_ARGS（不可见图片防线）", Describe(negZero));
                    var negNonNum = ExecuteTool(registry, "insert_image", Range(
                        ("sheet", "Sheet1"), ("image_path", m13Png), ("anchor", "A20"), ("height", "abc")));
                    Expect(!negNonNum.Ok && Convert.ToString(negNonNum.ErrorCode) == "INVALID_ARGS",
                        "用例23 负路径：height=非数字 显式报 INVALID_ARGS（而非 EXCEL_COM_ERROR）", Describe(negNonNum));
                    var negNotFound = ExecuteTool(registry, "insert_image", Range(
                        ("sheet", "Sheet1"), ("image_path", Path.Combine(Path.GetTempPath(), "rmv_no_such_m13.png")), ("anchor", "A20")));
                    Expect(!negNotFound.Ok && Convert.ToString(negNotFound.ErrorCode) == "IMAGE_NOT_FOUND",
                        "用例23 负路径：文件不存在 显式报 IMAGE_NOT_FOUND", Describe(negNotFound));
                    var corruptPng = Path.Combine(Path.GetTempPath(), "rmv_corrupt_m13.png");
                    File.WriteAllBytes(corruptPng, new byte[] { 0x00, 0x01, 0x02, 0x03, 0xFF });
                    var negCorrupt = ExecuteTool(registry, "insert_image", Range(
                        ("sheet", "Sheet1"), ("image_path", corruptPng), ("anchor", "A20")));
                    Expect(!negCorrupt.Ok && Convert.ToString(negCorrupt.ErrorCode) == "IMAGE_INVALID",
                        "用例23 负路径：损坏图片文件 显式报 IMAGE_INVALID（而非裸 E_FAIL）", Describe(negCorrupt));

                    try { if (File.Exists(m13Png)) File.Delete(m13Png); } catch { }

                    // 用例 24：冻结窗格 + verify_freeze_panes 自动自检链路（M13：freeze_panes）。
                    var freezeArgs = Range(("sheet", "Sheet1"), ("cell", "B2"));
                    var freezeResult = ExecuteTool(registry, "freeze_panes", freezeArgs);
                    Expect(freezeResult.Ok, "用例24 准备：freeze_panes 冻结在 B2", Describe(freezeResult));
                    RunAutoVerifyChain(registry, "freeze_panes", freezeArgs, freezeResult,
                        expectOk: true, label: "用例24 冻结自动自检（状态+位置）");
                    var unfreezeArgs = Range(("sheet", "Sheet1"), ("unfreeze", true));
                    var unfreezeResult = ExecuteTool(registry, "freeze_panes", unfreezeArgs);
                    Expect(unfreezeResult.Ok, "用例24 准备：解除冻结", Describe(unfreezeResult));
                    RunAutoVerifyChain(registry, "freeze_panes", unfreezeArgs, unfreezeResult,
                        expectOk: true, label: "用例24 解冻自动自检");

                    // 用例 25：列质量画像（M14：analyze_column_profile）——数值+错误+空值混合列的三联单面板。
                    var seed25 = ExecuteTool(registry, "write_excel_range", Range(
                        ("address", "$A$26:$A$31"),
                        ("values", new JArray
                        {
                            new JArray { "成绩" },
                            new JArray { 10 },
                            new JArray { 20 },
                            new JArray { 20 },
                            new JArray { "#N/A" },
                            new JArray { "" }
                        })));
                    Expect(seed25.Ok, "用例25 准备：写入画像测试列（3 数值 + 1 错误 + 1 空值）", Describe(seed25));
                    var profile25 = ExecuteTool(registry, "analyze_column_profile", Range(
                        ("address", "$A$26:$A$31"), ("column", 1)));
                    Expect(profile25.Ok, "用例25 执行：analyze_column_profile", Describe(profile25));
                    var p25 = profile25.Data == null
                        ? null
                        : JObject.FromObject(profile25.Data);
                    Expect(p25 != null && p25["quality"] != null && Convert.ToInt32(p25["quality"]["blank_count"]) == 1,
                        "用例25 质量：空值计数 = 1", p25?["quality"]?["blank_count"]?.ToString() ?? "(null)");
                    Expect(Convert.ToInt32(p25["quality"]["error_count"]) == 1,
                        "用例25 质量：错误计数 = 1（#N/A）", p25?["quality"]?["error_count"]?.ToString() ?? "(null)");
                    Expect(Convert.ToInt32(p25["quality"]["distinct_count"]) == 3,
                        "用例25 质量：唯一值计数 = 3（10/20/#N/A）", p25?["quality"]?["distinct_count"]?.ToString() ?? "(null)");
                    Expect(p25?["numbers"] != null && Math.Abs(Convert.ToDouble(p25["numbers"]["mean"]) - 16.6667) < 0.001,
                        "用例25 剖析：数值均值 = 16.6667", p25?["numbers"]?["mean"]?.ToString() ?? "(null)");
                    var flags25 = p25?["quality_flags"]?.Select(t => t.ToString()).ToList() ?? new List<string>();
                    Expect(flags25.Contains("has_errors"),
                        "用例25 提示：has_errors 质量标记", string.Join(",", flags25));
                    var p25neg = ExecuteTool(registry, "analyze_column_profile", Range(
                        ("address", "$A$26:$A$31"), ("column", 9)));
                    Expect(!p25neg.Ok && Convert.ToString(p25neg.ErrorCode) == "INVALID_COLUMN",
                        "用例25 负路径：列号越界显式报 INVALID_COLUMN", Describe(p25neg));

                    // 用例 26：超链接管理 + verify_hyperlink 自动自检链路（M15：manage_hyperlinks）。
                    var seed26 = ExecuteTool(registry, "write_excel_range", Range(
                        ("address", "$C$28:$C$30"),
                        ("values", new JArray
                        {
                            new JArray { "官网" },
                            new JArray { "跳转" },
                            new JArray { "空位" }
                        })));
                    Expect(seed26.Ok, "用例26 准备：写入超链接宿主单元格", Describe(seed26));

                    var addExtArgs = Range(("sheet", "Sheet1"), ("action", "add"), ("address", "$C$28"),
                        ("link_url", "https://example.com/jyyj"), ("display_text", "官网"));
                    var addExt = ExecuteTool(registry, "manage_hyperlinks", addExtArgs);
                    Expect(addExt.Ok, "用例26 准备：add 外部超链接", Describe(addExt));
                    RunAutoVerifyChain(registry, "manage_hyperlinks", addExtArgs, addExt,
                        expectOk: true, label: "用例26 外链自动自检（url+显示文本）");

                    var addIntArgs = Range(("sheet", "Sheet1"), ("action", "add"), ("address", "$C$29"),
                        ("link_subaddress", "Sheet1!A1"));
                    var addInt = ExecuteTool(registry, "manage_hyperlinks", addIntArgs);
                    Expect(addInt.Ok, "用例26 准备：add 内部超链接", Describe(addInt));
                    RunAutoVerifyChain(registry, "manage_hyperlinks", addIntArgs, addInt,
                        expectOk: true, label: "用例26 内链自动自检（subaddress）");

                    var list26 = ExecuteTool(registry, "manage_hyperlinks", Range(("sheet", "Sheet1"), ("action", "list")));
                    var list26Data = list26.Data == null ? null : JObject.FromObject(list26.Data);
                    Expect(list26.Ok && list26Data != null && Convert.ToInt32(list26Data["hyperlink_count"]) == 2,
                        "用例26 list：整表超链接计数 = 2", list26Data?["hyperlink_count"]?.ToString() ?? "(null)");

                    var negDup = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "add"), ("address", "C28"), ("link_url", "https://example.com/2")));
                    Expect(!negDup.Ok && Convert.ToString(negDup.ErrorCode) == "HYPERLINK_EXISTS",
                        "用例26 负路径：重复添加显式报 HYPERLINK_EXISTS", Describe(negDup));
                    var negBoth = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "add"), ("address", "C30"),
                        ("link_url", "https://a.com"), ("link_subaddress", "Sheet1!A1")));
                    Expect(!negBoth.Ok && Convert.ToString(negBoth.ErrorCode) == "INVALID_ARGS",
                        "用例26 负路径：link_url 与 link_subaddress 互斥显式报 INVALID_ARGS", Describe(negBoth));
                    var negUpdNoTarget = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "update"), ("address", "C30"), ("link_url", "https://example.com/x")));
                    Expect(!negUpdNoTarget.Ok && Convert.ToString(negUpdNoTarget.ErrorCode) == "HYPERLINK_NOT_FOUND",
                        "用例26 负路径：update 无链接单元格显式报 HYPERLINK_NOT_FOUND", Describe(negUpdNoTarget));
                    var negRmNoTarget = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "remove"), ("address", "C30")));
                    Expect(!negRmNoTarget.Ok && Convert.ToString(negRmNoTarget.ErrorCode) == "HYPERLINK_NOT_FOUND",
                        "用例26 负路径：remove 无链接单元格显式报 HYPERLINK_NOT_FOUND", Describe(negRmNoTarget));

                    var updArgs = Range(("sheet", "Sheet1"), ("action", "update"), ("address", "$C$28"),
                        ("link_url", "https://example.com/updated"));
                    var upd = ExecuteTool(registry, "manage_hyperlinks", updArgs);
                    Expect(upd.Ok, "用例26 执行：update 外链地址", Describe(upd));
                    RunAutoVerifyChain(registry, "manage_hyperlinks", updArgs, upd,
                        expectOk: true, label: "用例26 更新自动自检");

                    var rmArgs = Range(("sheet", "Sheet1"), ("action", "remove"), ("address", "$C$28:$C$29"));
                    var rm = ExecuteTool(registry, "manage_hyperlinks", rmArgs);
                    Expect(rm.Ok, "用例26 执行：remove 区域删除 2 个链接", Describe(rm));
                    RunAutoVerifyChain(registry, "manage_hyperlinks", rmArgs, rm,
                        expectOk: true, label: "用例26 删除自动自检（expected_present=false）");

                    // 用例 26 准备（F6 序列修正）：remove 已清空整表，先补种一条链接再验证 remove_all。
                    var reseed26 = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "add"), ("address", "$C$29"), ("link_url", "https://example.com/reseed")));
                    Expect(reseed26.Ok, "用例26 准备：补种一条链接供清空", Describe(reseed26));
                    var rmAllArgs = Range(("sheet", "Sheet1"), ("action", "remove_all"));
                    var rmAll = ExecuteTool(registry, "manage_hyperlinks", rmAllArgs);
                    Expect(rmAll.Ok, "用例26 执行：remove_all 清空整表", Describe(rmAll));
                    RunAutoVerifyChain(registry, "manage_hyperlinks", rmAllArgs, rmAll,
                        expectOk: true, label: "用例26 清空自动自检（expected_count=0）");
                    var negRmAllEmpty = ExecuteTool(registry, "manage_hyperlinks", Range(
                        ("sheet", "Sheet1"), ("action", "remove_all")));
                    Expect(!negRmAllEmpty.Ok && Convert.ToString(negRmAllEmpty.ErrorCode) == "HYPERLINK_NOT_FOUND",
                        "用例26 负路径：空表 remove_all 显式报 HYPERLINK_NOT_FOUND", Describe(negRmAllEmpty));
                }
                finally
                {
                    try { if (File.Exists(rollbackDb)) File.Delete(rollbackDb); } catch { }
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[fatal] harness 异常：" + ex.Message + "\n" + ex.StackTrace);
            }
            finally
            {
                // 不走 workbook.Close/app.Quit：out-of-proc 动态 COM 调用在退出路径会触发进程级硬故障
                // （fail-fast，托管侧不可捕获）。工作簿是未存盘草稿、无保存价值，直接按 PID 强杀本
                // harness 启动的实例（踩坑手册 §5.2 的泄漏兜底思路），绝不触碰用户自己的 Excel 进程。
                Console.WriteLine("[cleanup] 结束本 harness 启动的 EXCEL.EXE...");
                var killed = KillSpawnedExcel();
                Console.WriteLine("[cleanup] 已结束 " + killed.Count + " 个进程（PID：" + string.Join(",", killed) + "）。");
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.WriteLine("\n== 结果：通过 " + _passed + "，失败 " + _failed + "，无进程泄漏（草稿实例已按 PID 结束）==");
            return _failed == 0 ? 0 : 1;
        }

        private static ToolResult ExecuteTool(ToolRegistry registry, string tool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(tool, out var descriptor))
            {
                return new ToolResult { Ok = false, ErrorCode = "UNKNOWN_TOOL", Data = tool };
            }

            try
            {
                return descriptor.Execute(args);
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, ErrorCode = "HARNESS_EXCEPTION", Data = ex.Message };
            }
        }

        /// <summary>复刻 AgentRunner.RunAutoVerify 的编排：TryPlanAll 出全部计划 → 结果富化 → 逐条执行只读 verify。</summary>
        private static void RunAutoVerifyChain(
            ToolRegistry registry,
            string tool,
            Dictionary<string, object> args,
            ToolResult executed,
            bool expectOk,
            string label,
            Action<Dictionary<string, object>> mutate = null)
        {
            var plans = new List<AutoVerifyPlan>();
            if (!AutoVerifyPolicy.TryPlanAll(tool, args, plans) || plans.Count == 0)
            {
                Fail(label + "：未生成自检计划（预期必须有）");
                return;
            }

            var checkedPlans = 0;
            foreach (var plan in plans)
            {
                if (!AutoVerifyPolicy.EnrichFromResult(plan, tool, executed?.Data, executed?.AffectedRanges))
                {
                    Fail(label + "：结果富化失败（计划被丢弃）");
                    continue;
                }

                mutate?.Invoke(plan.Args);

                if (!registry.TryGet(plan.VerifyTool, out var verifyTool) ||
                    verifyTool.RiskLevel != RiskLevel.ReadOnly)
                {
                    Fail(label + "：verify 工具 " + plan.VerifyTool + " 不可用或非只读");
                    continue;
                }

                ToolResult verifyResult;
                try
                {
                    verifyResult = verifyTool.Execute(plan.Args);
                }
                catch (Exception ex)
                {
                    Fail(label + "：verify 执行异常 " + ex.Message);
                    continue;
                }

                checkedPlans++;
                var verdict = ExtractVerdict(verifyResult);
                var detail = verdict + "；issues=" + DescribeIssues(verifyResult);
                if (verifyResult.Ok)
                {
                    detail = detail + "；data=" + Describe(verifyResult);
                }
                else
                {
                    detail = detail + "；err=" + Describe(verifyResult);
                }

                if (expectOk)
                {
                    Expect(verifyResult.Ok && verdict == "ok", label + "（计划 " + checkedPlans + "，期望 verdict=ok）", detail);
                }
                else
                {
                    Expect(verifyResult.Ok && !string.IsNullOrEmpty(verdict) && verdict != "ok",
                        label + "（计划 " + checkedPlans + "，期望 verdict≠ok）", detail);
                }
            }

            if (checkedPlans == 0)
            {
                Fail(label + "：没有任何计划被执行");
            }
        }

        private static string DescribeIssues(ToolResult result)
        {
            if (result?.Data == null)
            {
                return "(无)";
            }

            if (!(JToken.FromObject(result.Data) is JObject data) || !(data["issues"] is JArray issues))
            {
                return "(无)";
            }

            return issues.Count == 0 ? "(空)" : string.Join(" | ", issues);
        }

        /// <summary>从真实 Excel 读回矩阵（0 基归一化：COM 多格 Value2 为 1 基数组），供守卫/自检链路断言。</summary>
        private static object[,] ReadBlockValues(dynamic app, string sheetName, int top, int left, int rows, int cols)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic r1 = null;
            dynamic r2 = null;
            dynamic block = null;
            try
            {
                workbook = app.ActiveWorkbook;
                sheet = workbook.Worksheets.Item[sheetName];
                r1 = sheet.Cells[top, left];
                r2 = sheet.Cells[top + rows - 1, left + cols - 1];
                block = sheet.Range[r1, r2];
                var obj = block.Value2;
                var result = new object[rows, cols];
                if (obj is object[,] m)
                {
                    var lo0 = m.GetLowerBound(0);
                    var lo1 = m.GetLowerBound(1);
                    for (var r = 0; r < rows; r++)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            var v = (r < m.GetLength(0) && c < m.GetLength(1)) ? m[lo0 + r, lo1 + c] : DBNull.Value;
                            result[r, c] = v is DBNull ? null : v;
                        }
                    }
                }
                else
                {
                    result[0, 0] = obj;
                }

                return result;
            }
            finally
            {
                if (Marshal.IsComObject(block)) { try { Marshal.FinalReleaseComObject(block); } catch { } }
                if (Marshal.IsComObject(r2)) { try { Marshal.FinalReleaseComObject(r2); } catch { } }
                if (Marshal.IsComObject(r1)) { try { Marshal.FinalReleaseComObject(r1); } catch { } }
                if (Marshal.IsComObject(sheet)) { try { Marshal.FinalReleaseComObject(sheet); } catch { } }
                if (Marshal.IsComObject(workbook)) { try { Marshal.FinalReleaseComObject(workbook); } catch { } }
            }
        }

        private static string ExtractOperationId(ToolResult result)
        {
            if (result?.Data == null || !(JToken.FromObject(result.Data) is JObject data))
            {
                return null;
            }

            return data.Value<string>("operation_id");
        }

        /// <summary>读回单个单元格的 .Formula（A1 样式），核验公式确以公式形态落盘、回滚后恢复为空。</summary>
        private static string ReadFormula(dynamic app, string sheetName, int top, int left)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic cell = null;
            try
            {
                workbook = app.ActiveWorkbook;
                sheet = workbook.Worksheets.Item[sheetName];
                cell = sheet.Cells[top, left];
                var formula = cell.Formula;
                return formula == null ? null : Convert.ToString(formula);
            }
            finally
            {
                if (Marshal.IsComObject(cell)) { try { Marshal.FinalReleaseComObject(cell); } catch { } }
                if (Marshal.IsComObject(sheet)) { try { Marshal.FinalReleaseComObject(sheet); } catch { } }
                if (Marshal.IsComObject(workbook)) { try { Marshal.FinalReleaseComObject(workbook); } catch { } }
            }
        }

        private static string ExtractVerdict(ToolResult result)
        {
            if (result?.Data == null || !(JToken.FromObject(result.Data) is JObject data))
            {
                return null;
            }

            return data.Value<string>("verdict");
        }

        private static bool CompareMatrices(object[,] a, object[,] b)
        {
            if (a.GetLength(0) != b.GetLength(0) || a.GetLength(1) != b.GetLength(1))
            {
                return false;
            }

            for (var r = 0; r < a.GetLength(0); r++)
            {
                for (var c = 0; c < a.GetLength(1); c++)
                {
                    var va = a[r, c];
                    var vb = b[r, c];
                    if (va is DBNull) { va = null; }
                    if (vb is DBNull) { vb = null; }
                    if (!object.Equals(va, vb))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static string MatrixText(object[,] matrix)
        {
            var parts = new List<string>();
            for (var r = 0; r < matrix.GetLength(0); r++)
            {
                for (var c = 0; c < matrix.GetLength(1); c++)
                {
                    var v = matrix[r, c];
                    parts.Add(v is DBNull || v == null ? "(空)" : Convert.ToString(v));
                }
            }

            return string.Join("|", parts);
        }

        private static void Expect(bool condition, string name, string detail)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("[PASS] " + name + "　→ " + detail);
            }
            else
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name + "　→ " + detail);
            }
        }

        private static void Fail(string message)
        {
            _failed++;
            Console.WriteLine("[FAIL] " + message);
        }

        private static string Describe(ToolResult result)
        {
            if (result == null)
            {
                return "null";
            }

            return result.Ok ? "ok" : (result.ErrorCode ?? "failed") + " " + result.Data;
        }

        private static Dictionary<string, object> Range(params (string key, object value)[] pairs)
        {
            var map = new Dictionary<string, object>();
            foreach (var (key, value) in pairs)
            {
                map[key] = value;
            }

            return map;
        }
    }
}
