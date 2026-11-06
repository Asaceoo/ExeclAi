using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace SlicerProbe
{
    /// <summary>
    /// M-D0 切片器原生支持探测（只读探测，设计文档 docs/设计-看板切片器探测harness.md）。
    /// 独立 out-of-proc Excel 实例 + 动态内存工作簿，逐项探测 Slicer COM 能力，
    /// 产出 slicer_native_ready 判定与 HRESULT 错误目录。
    /// 探测失败不是程序失败：每项独立记录，结论看报告。
    /// 清理纪律沿用 RealMachineVerify 踩坑手册 §5.2：不 Close/Quit，按 PID 强杀本 harness 实例。
    /// </summary>
    internal static class Program
    {
        private static dynamic workbook;
        private static readonly List<ProbeItem> Items = new List<ProbeItem>();
        private static readonly List<string> Errors = new List<string>();
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();

        private static int Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "md1")
            {
                return Md1Probe.Run(args);
            }
            if (args.Length > 0 && args[0] == "sb0")
            {
                return SuperBoardProbe.Run();
            }
            if (args.Length > 1 && args[0] == "sb5")
            {
                return SuperBoardProbe.RunWizard(args[1]);
            }
            if (args.Length > 1 && args[0] == "sb3")
            {
                return SuperBoardProbe.RunFullV2(args[1]);
            }
            if (args.Length > 1 && args[0] == "sb2")
            {
                return SuperBoardProbe.RunUserFile(args[1]);
            }
            if (args.Length > 1 && args[0] == "sb1")
            {
                return SuperBoardProbe.RunFull(args[1]);
            }
            if (args.Length > 1 && args[0] == "ud")
            {
                return Md1Probe.RunUserDataSuite(args[1]);
            }
            if (args.Length > 2 && args[0] == "expw")
            {
                return ExpWorkbookProbe.Run(args[1], args[2]);
            }
            if (args.Length > 1 && args[0] == "expn")
            {
                return ExpNativeProbe.Run(args[1]);
            }
            if (args.Length > 1 && args[0] == "udchart")
            {
                return UserChartProbe.Run(args[1]);
            }
            if (args.Length > 1 && args[0] == "udchart2")
            {
                return ChartTestProbe.Run(args[1]);
            }
            if (args.Length > 1 && args[0] == "exam")
            {
                return ExamProbe.Run(args[1], args.Length > 2 ? args[2] : null);
            }
            if (args.Length > 1 && args[0] == "macrochain")
            {
                return MacroChainProbe.Run(args[1], args.Length > 2 ? args[2] : null);
            }
            if (args.Length > 0 && args[0] == "slicerprobe")
            {
                return ChartAllProbe.RunSlicerProbe();
            }
            if (args.Length > 0 && args[0] == "enumprobe")
            {
                return ChartAllProbe.RunEnumProbe();
            }
            if (args.Length > 2 && args[0] == "charts12")
            {
                return ChartAllProbe.Run(args[1], args[2]);
            }
            if (args.Length > 2 && args[0] == "full")
            {
                return FullFeatureProbe.Run(args[1], args[2]);
            }
            if (args.Length > 0 && args[0] == "pyfix")
            {
                return PyFixProbe.Run(args.Length > 1 ? args[1] : "python");
            }
            if (args.Length > 0 && args[0] == "dl")
            {
                return DataLabelProbe.Run();
            }
            if (args[0] == "make")
            {
                if (args.Length > 8)
                {
                    return Md1Probe.MakeDashboardEx(args[1], args[2], args[3],
                        int.Parse(args[4]), int.Parse(args[5]), int.Parse(args[6]), int.Parse(args[7]), int.Parse(args[8]));
                }
                return Md1Probe.MakeDashboard(args[1], args[2]);
            }

            Console.WriteLine("== M-D0 切片器原生支持探测（SlicerProbe）==");

            dynamic app = null;
            var probeTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var excelVersion = "unknown";
            var excelBuild = "unknown";

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
                app.AutomationSecurity = 1;

                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (!SpawnedPids.Contains(process.Id))
                        {
                            SpawnedPids.Add(process.Id);
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }

                workbook = app.Workbooks.Add();
                try
                {
                    var tmpPath = Path.Combine(Path.GetTempPath(), "SlicerProbe_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx");
                    workbook.SaveAs(tmpPath, 51); // FileFormat=51(xlOpenXMLWorkbook)，脱离兼容模式，切片器可用
                    Console.WriteLine("[env] 工作簿已转存 xlsx：" + tmpPath);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("[warn] SaveAs 失败（" + saveEx.Message + "），继续以内存工作簿探测——若 P2 报兼容模式错误即为此因");
                }
                excelVersion = Convert.ToString(app.Version);
                excelBuild = Convert.ToString(app.Build);
                Console.WriteLine("[env] Excel Version=" + excelVersion + " Build=" + excelBuild
                    + " 实例数=" + SpawnedPids.Count + "\n");

                // P1: API 存在性
                dynamic slicerCaches = null;
                var p1 = Probe("P1", "SlicerCaches API 存在性", () =>
                {
                    slicerCaches = workbook.SlicerCaches;
                    return "SlicerCaches 属性可访问";
                });
                if (p1 == null)
                {
                    // P1 失败：后续项无意义，直接出报告
                    Finish(app, probeTime, excelVersion, excelBuild);
                    return 0;
                }

                // P2: 构造透视表 + Add2 创建 SlicerCache
                dynamic pivotTable = null;
                dynamic sheetData = workbook.Worksheets.Item[1];
                dynamic sheetPivot = workbook.Worksheets.Add();
                var p2 = Probe("P2", "SlicerCaches.Add2 创建（基于透视表）", () =>
                {
                    BuildTestData(sheetData);
                    try { Console.WriteLine("[diag] CompatibilityMode=" + workbook.CompatibilityMode); } catch (Exception cmEx) { Console.WriteLine("[diag] CompatibilityMode 读取失败: " + cmEx.Message); }
                    dynamic pivotCache = workbook.PivotCaches().Create(
                        SourceType: 1, SourceData: sheetData.Range("A1:D19")); // 18行数据+表头=19，xlDatabase=1，Create=新式API
                    pivotTable = sheetPivot.PivotTables().Add(
                        PivotCache: pivotCache, TableDestination: sheetPivot.Range("A3"));
                    pivotTable.PivotFields("地区").Orientation = 1; // xlRowField
                    pivotTable.PivotFields("销售额").Orientation = 4; // xlDataField
                    dynamic cache = workbook.SlicerCaches.Add2(pivotTable, "地区");
                    BindFirstCache(() => cache.SlicerItems);
                    return "SlicerCache 创建成功 name=" + Convert.ToString(cache.Name);
                });

                // P3: Slicers.Add 放置
                dynamic sheetBoard = workbook.Worksheets.Add();
                var p3 = Probe("P3", "Slicers.Add 放置切片器", () =>
                {
                    dynamic cache = workbook.SlicerCaches.Item(1);
                    dynamic slicer = cache.Slicers().Add(
                        SlicerDestination: sheetBoard,
                        Top: 24, Left: 24, Width: 144, Height: 192);
                    slicer.Caption = "地区";
                    return "切片器已放置于看板 sheet";
                });

                // P4: 选择驱动联动（关键项）
                var p4 = Probe("P4", "切片器选择驱动透视表联动", () =>
                {
                    dynamic pivot = sheetPivot.PivotTables().Item(1);
                    Console.WriteLine("[diag] pivot1 name=" + Convert.ToString(pivot.Name) + " source=" + Convert.ToString(pivot.SourceData));
                    object rawBefore = pivot.GetPivotData("销售额").Value2; // 未筛选总额（单选任一项后必变小，断言b的可靠基线）
                    double before = rawBefore is double d1 ? d1 : Convert.ToDouble(rawBefore, System.Globalization.CultureInfo.InvariantCulture);
                    dynamic items = slicerItemsOfFirstCache();
                    Console.WriteLine("[diag] SlicerItems.Count=" + items.Count);
                    for (int i = 1; i <= items.Count; i++)
                    {
                        dynamic it = items.Item(i);
                        bool sel = i == 1;
                        it.Selected = sel;
                        Console.WriteLine("[diag] item[" + i + "]=" + Convert.ToString(it.Name) + " selected=" + sel);
                    }
                    // 验证选中状态回读
                    dynamic it1 = items.Item(1);
                    Console.WriteLine("[diag] 回读 item1.Selected=" + it1.Selected + " item2.Selected=" + items.Item(2).Selected);
                    app.Calculate();
                    object rawAfter = pivot.TableRange1.Cells.Item(2, 2).Value2; // 筛选后第一数据行=华北值
                    double after = rawAfter is double d2 ? d2 : Convert.ToDouble(rawAfter, System.Globalization.CultureInfo.InvariantCulture);
                    Console.WriteLine("[diag] 筛选后第一数据行值=" + after);
                    // 三重断言（不依赖具体数值大小假设）：
                    // a) 选择状态回读正确（单选第1项）—— COM variant 需防御式转换（null/Empty 视为未选中）
                    object rawSel1 = items.Item(1).Selected;
                    object rawSel2 = items.Item(2).Selected;
                    Console.WriteLine("[diag] rawSel1=" + (rawSel1?.ToString() ?? "null") + " rawSel2=" + (rawSel2?.ToString() ?? "null"));
                    bool sel1 = rawSel1 != null && rawSel1 is bool b1 && b1;
                    bool sel2 = rawSel2 != null && rawSel2 is bool b2 && b2;
                    if (!sel1 || sel2)
                    {
                        throw new Exception("SlicerItems 选中状态回读与设定不符（rawSel1=" + (rawSel1?.ToString() ?? "null") + ", rawSel2=" + (rawSel2?.ToString() ?? "null") + "）（联动控制失效）");
                    }
                    // b) 透视值发生变化（联动确实驱动了重算）
                    if (Math.Abs(after - before) < 0.0001)
                    {
                        throw new Exception("透视值未随切片器选择变化（联动闭环未成立）");
                    }
                    // c) 变化后的值必须等于某单项之和（防"值变了但不是筛选驱动"的伪联动）
                    string selectedName;
                    try
                    {
                        selectedName = Convert.ToString(items.Item(1).Name);
                    }
                    catch (Exception nEx)
                    {
                        throw new Exception("读取 SlicerItem.Name 失败: " + nEx.GetType().Name + " " + nEx.Message);
                    }
                    var expectedSums = FirstItemExpectedSums(sheetData);
                    double itemSum = 0;
                    bool found = !string.IsNullOrEmpty(selectedName) && expectedSums.TryGetValue(selectedName, out itemSum);
                    if (!found || Math.Abs(after - itemSum) > 0.0001)
                    {
                        throw new Exception("筛选后值=" + after + " 与选中项「" + selectedName + "」聚合值不匹配（联动语义不正确）");
                    }
                    return "选中单项后透视值变化 " + before + " -> " + after + "（联动成立，值与单项聚合吻合）";
                });

                // P5: 多透视表关联（Report Connections 等效）
                var p5 = Probe("P5", "SlicerCache 关联多透视表", () =>
                {
                    // 切片器只能关联共享同一 PivotCache 的透视表 —— 直接从 P2 的透视表取其 cache
                    dynamic pivotCache1 = pivotTable.PivotCache();
                    dynamic pivot2 = sheetPivot.PivotTables().Add(
                        pivotCache1, TableDestination: sheetPivot.Range("F3"));
                    pivot2.PivotFields("品类").Orientation = 1;
                    pivot2.PivotFields("销售额").Orientation = 4;
                    Console.WriteLine("[diag] pivot2 created: " + Convert.ToString(pivot2.Name));
                    dynamic cache = workbook.SlicerCaches.Item(1);
                    dynamic cachePivots = cache.PivotTables;
                    Console.WriteLine("[diag] cachePivots type=" + cachePivots.GetType().Name + " count=" + cachePivots.Count);
                    try
                    {
                        cachePivots.GetType().InvokeMember("AddPivotTable",
                            System.Reflection.BindingFlags.InvokeMethod, null, cachePivots, new object[] { pivot2 });
                    }
                    catch (System.Reflection.TargetInvocationException tie)
                    {
                        // 解包真实 COM 异常（HRESULT 是错误目录的关键信息）
                        var inner = tie.InnerException ?? tie;
                        throw new COMException("AddPivotTable 反射调用失败: " + inner.Message, inner.HResult);
                    }
                    return "第二透视表已挂接，cache.PivotTables.Count=" + cachePivots.Count;
                });

                // P6: 只读回读现有切片器
                Probe("P6", "SlicerCaches/Slicers 只读回读", () =>
                {
                    int caches = workbook.SlicerCaches().Count;
                    int slicers = 0;
                    for (int i = 1; i <= caches; i++)
                    {
                        slicers += workbook.SlicerCaches().Item(i).Slicers().Count;
                    }
                    return "caches=" + caches + " slicers=" + slicers;
                });

                // P7: 删除清理
                var p7 = Probe("P7", "Slicer/Cache 删除清理", () =>
                {
                    dynamic cache = workbook.SlicerCaches.Item(1);
                    try { cache.ClearAllFilters(); } catch (Exception cfEx) { Console.WriteLine("[diag] ClearAllFilters: " + cfEx.Message); }
                    dynamic slicer = cache.Slicers().Item(1);
                    try
                    {
                        slicer.Delete();
                        Console.WriteLine("[diag] slicer.Delete 完成");
                    }
                    catch (Exception sdEx)
                    {
                        Console.WriteLine("[diag] slicer.Delete 失败: " + sdEx.Message + "（尝试直接删 cache）");
                    }
                    // 坑：slicer.Delete 后旧 cache RCW 句柄失效（0x800A01A8）；重新解析——若 caches 已空说明 slicer.Delete 级联清了 cache
                    int remaining = workbook.SlicerCaches().Count;
                    if (remaining > 0)
                    {
                        dynamic cacheFresh = workbook.SlicerCaches.Item(1);
                        cacheFresh.Delete();
                    }
                    return "slicer.Delete 后 caches=" + remaining + "，清理完成 caches.Count=" + workbook.SlicerCaches().Count;
                });

                // P8: 跨 sheet 放置已在 P3 完成（切片器在 sheetBoard，透视在 sheetPivot）
                if (p3 != null && p3.Passed)
                {
                    Items.Add(new ProbeItem
                    {
                        Id = "P8",
                        Name = "切片器跨 sheet 放置",
                        Passed = true,
                        Detail = "P3 已在独立看板 sheet 放置成功，本项由 P3 覆盖"
                    });
                }
            }
            catch (Exception ex)
            {
                Errors.Add("fatal: " + ex.Message);
                Console.WriteLine("[fatal] " + ex.Message);
            }
            finally
            {
                Finish(app, probeTime, excelVersion, excelBuild);
            }

            return 0;
        }

        /// <summary>单项探测：异常捕获 + HRESULT 记录，失败不中断后续项。</summary>
        private static ProbeItem Probe(string id, string name, Func<string> action)
        {
            var item = new ProbeItem { Id = id, Name = name };
            try
            {
                item.Detail = action();
                item.Passed = true;
                Console.WriteLine("[PASS] " + id + " " + name + " → " + item.Detail);
            }
            catch (System.Runtime.InteropServices.COMException cex)
            {
                item.Passed = false;
                item.HResult = "0x" + cex.HResult.ToString("X8");
                item.Error = cex.Message;
                Errors.Add(id + " " + name + " HRESULT=" + item.HResult + " " + cex.Message);
                Console.WriteLine("[FAIL] " + id + " " + name + " → HRESULT=" + item.HResult + " " + cex.Message);
            }
            catch (Exception ex)
            {
                item.Passed = false;
                item.Error = ex.GetType().Name + ": " + ex.Message;
                Errors.Add(id + " " + name + " " + item.Error);
                Console.WriteLine("[FAIL] " + id + " " + name + " → " + item.Error);
            }
            Items.Add(item);
            return item;
        }

        /// <summary>计算 SlicerItems 第一项（透视行序第一项）的预期聚合值：明细中该地区全部行之和。</summary>
        private static double FirstItemExpectedSum(dynamic sheetData)
        {
#pragma warning disable CS0162 // 未可达代码：保留占位说明
            // 透视行序由数据首现序决定（PivotTable 默认按数据源顺序排列自定义列表外的项）。
            // 为避免顺序假设脆弱化，改为：全部 3 个地区各自聚合值中，寻找与 after 匹配的一个即可。
            // 本函数返回"最小的那个地区聚合值"，P4 中改用集合匹配（见 FirstItemExpectedSums）。
            throw new NotImplementedException("由 FirstItemExpectedSums 替代");
        }

        /// <summary>返回 3 个地区各自的聚合值集合（按明细顺序）。</summary>
        private static Dictionary<string, double> FirstItemExpectedSums(dynamic sheetData)
        {
            var sums = new Dictionary<string, double>();
            for (int r = 2; r <= 19; r++) // 18 行数据
            {
                object rawRegion = sheetData.Cells.Item[r, 1].Value2;
                object rawVal = sheetData.Cells.Item[r, 4].Value2;
                if (rawRegion == null || rawVal == null)
                {
                    Console.WriteLine("[diag] 聚合扫描：行 " + r + " 出现 null（region=" + (rawRegion == null) + ", val=" + (rawVal == null) + "），跳过");
                    continue;
                }
                var region = Convert.ToString(rawRegion);
                double val = rawVal is double dv ? dv : Convert.ToDouble(rawVal, System.Globalization.CultureInfo.InvariantCulture);
                if (!sums.ContainsKey(region)) { sums[region] = 0; }
                sums[region] += val;
            }
            return sums;
        }

        /// <summary>取第一个 SlicerCache 的 SlicerItems（延迟到 P2 成功后才有值）。</summary>
        private static Func<dynamic> slicerItemsOfFirstCache = () =>
        {
            throw new InvalidOperationException("slicer cache 未创建");
        };

        /// <summary>P2 成功后回填缓存引用，供 P4 使用。</summary>
        private static void BindFirstCache(Func<dynamic> getter)
        {
            slicerItemsOfFirstCache = getter;
        }

        /// <summary>写 18 行模拟明细（地区3×品类3×月份2，数值可区分），范围 A1:D19。</summary>
        private static void BuildTestData(dynamic sheet)
        {
            var regions = new[] { "华东", "华北", "华南" };
            var cats = new[] { "办公设备", "家用电器", "数码产品" };
            var months = new[] { "1月", "2月" };
            sheet.Range("A1").Value2 = "地区";
            sheet.Range("B1").Value2 = "品类";
            sheet.Range("C1").Value2 = "月份";
            sheet.Range("D1").Value2 = "销售额";
            int row = 2;
            foreach (var r in regions)
            {
                foreach (var c in cats)
                {
                    foreach (var m in months)
                    {
                        sheet.Cells.Item[row, 1].Value2 = r;
                        sheet.Cells.Item[row, 2].Value2 = c;
                        sheet.Cells.Item[row, 3].Value2 = m;
                        sheet.Cells.Item[row, 4].Value2 = row * 100 + 7; // 行间可区分
                        row++;
                    }
                }
            }
        }

        /// <summary>出报告 + 清理（按 PID 强杀本 harness 实例，沿用踩坑手册 §5.2）。</summary>
        private static void Finish(dynamic app, string probeTime, string version, string build)
        {
            try
            {
                var p1 = Items.Find(i => i.Id == "P1");
                var p4 = Items.Find(i => i.Id == "P4");
                var ready = p1 != null && p1.Passed && p4 != null && p4.Passed;
                var report = new ProbeReport
                {
                    ProbeTime = probeTime,
                    ExcelVersion = version,
                    ExcelBuild = build,
                    SlicerNativeReady = ready,
                    Conclusion = ready
                        ? "slicer_native_ready=true：P3 可开发 native 联动模式"
                        : (p1 != null && !p1.Passed
                            ? "slicer_native_ready=false：SlicerCaches API 不可用，固定走 cell 模拟方案"
                            : "slicer_native_ready=false：切片器创建/驱动存在缺陷，固定走 cell 模拟方案"),
                    Items = Items,
                    ErrorCatalog = Errors.Count > 0 ? string.Join("; ", Errors) : "无"
                };
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (SpawnedPids.Contains(process.Id))
                        {
                            report.ExcelPid = process.Id;
                        }
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
                ReportWriter.Write(report, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "reports"));
                Console.WriteLine("\n== 结论：" + report.Conclusion);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal] 报告生成失败：" + ex.Message);
            }
            finally
            {
                // 清理临时探测文件（SaveAs 到 %TEMP%，强杀进程后手工删除）
                try
                {
                    if (workbook != null)
                    {
                        string fullName = Convert.ToString(workbook.FullName);
                        workbook.Close(false);
                        if (File.Exists(fullName)) { File.Delete(fullName); }
                    }
                }
                catch { }
                Console.WriteLine("[cleanup] 结束本 probe 启动的 EXCEL.EXE...");
                var killed = new List<int>();
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (SpawnedPids.Contains(process.Id))
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
                Console.WriteLine("[cleanup] 已结束 " + killed.Count + " 个进程（PID：" + string.Join(",", killed) + "）。");
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
