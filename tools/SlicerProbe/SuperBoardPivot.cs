using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;

namespace SlicerProbe
{
    internal static partial class SuperBoardProbe
    {
        /// <summary>
        /// SB3 超级看板透视/切片器配置验收（FR-SB-10 Q3-A/Q4-A）：
        /// pivot_rows/pivot_measures 生成真·可拖拽数据透视表 + slicer_enabled native 切片器多选挂载。
        /// </summary>
        public static int RunFullV2(string sourcePath)
        {
            Console.WriteLine("== SB3 超级看板透视/切片器全链 ==");
            if (!System.IO.File.Exists(sourcePath)) { Console.WriteLine("FILE_NOT_FOUND"); return 2; }
            string workPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "SB3_user_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".xlsx");
            System.IO.File.Copy(sourcePath, workPath, true);

            dynamic excel = null;
            dynamic workbook = null;
            try
            {
                var before = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    if (!before.Contains(p.Id)) { SpawnedPids.Add(p.Id); }
                    p.Dispose();
                }
                appLocalStatic = excel;
                workbook = excel.Workbooks.Open(workPath);
                dynamic ws = workbook.Worksheets.Item["Sheet1"];

                double amountSum = 0, profitSum = 0, quantitySum = 0;
                for (int r = 11; r <= 27; r++)
                {
                    amountSum += ToD(ws.Cells.Item[r, 11].Value2);
                    profitSum += ToD(ws.Cells.Item[r, 13].Value2);
                    quantitySum += ToD(ws.Cells.Item[r, 9].Value2);
                }

                var dashboard = new ExcelDashboardService(appLocalStatic);
                var pivotChart = new ExcelPivotChartService(appLocalStatic, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, appLocalStatic, null, null);

                // ---- 生成：透视字段（行=品类/销售员，值=金额/利润）+ 切片器 ----
                var gen = ExecSuper(registry, "render_super_dashboard", new Dictionary<string, object>
                {
                    ["source_address"] = "B10:S27",
                    ["sheet"] = "Sheet1",
                    ["dashboard_sheet"] = "超级看板",
                    ["label_column"] = 14,
                    ["category_column"] = 5,
                    ["month_column"] = 1,
                    ["value_column"] = 10,
                    ["quantity_column"] = 8,
                    ["profit_column"] = 12,
                    ["pivot_analysis"] = "true",
                    ["pivot_rows"] = "品类,地区",
                    ["pivot_measures"] = "金额,利润",
                    ["slicer_enabled"] = "true"
                });
                string genJson = Newtonsoft.Json.JsonConvert.SerializeObject(gen.Data);
                Expect(gen.Ok, "S1 生成（透视+切片器）", gen.Ok ? genJson.Substring(0, Math.Min(300, genJson.Length)) : gen.ErrorCode + " " + gen.Data);
                if (!gen.Ok) { return 1; }

                // ---- S2 真透视表断言：sheet 存在 + PivotTables≥1 + 行字段/值字段 ----
                dynamic pivotSheet = null;
                bool pivotOk = false, rowFieldsOk = false, dataFieldsOk = false;
                string pivotDetail = "";
                foreach (dynamic s in workbook.Worksheets)
                {
                    if (Convert.ToString(s.Name).StartsWith("透视分析-"))
                    {
                        pivotSheet = s;
                        int ptCount = Convert.ToInt32(s.PivotTables().Count);
                        pivotOk = ptCount >= 1;
                        if (pivotOk)
                        {
                            dynamic pt = s.PivotTables().Item(1);
                            var rowNames = new List<string>();
                            for (int i = 1; i <= Convert.ToInt32(pt.RowFields().Count); i++)
                            {
                                rowNames.Add(Convert.ToString(pt.RowFields().Item(i).Name));
                            }
                            rowFieldsOk = rowNames.Any(n => n.Contains("品类") || n.Contains("产品类别"))
                                       && rowNames.Any(n => n.Contains("销售员"));
                            var dataNames = new List<string>();
                            for (int i = 1; i <= Convert.ToInt32(pt.DataFields().Count); i++)
                            {
                                dataNames.Add(Convert.ToString(pt.DataFields().Item(i).SourceName));
                            }
                            dataFieldsOk = dataNames.Any(n => n.Contains("金额")) && dataNames.Any(n => n.Contains("利润"));
                            pivotDetail = "pt=" + ptCount + " rows=[" + string.Join("/", rowNames) + "] data=[" + string.Join("/", dataNames) + "]";
                            Release((object)pt);
                        }
                        else { pivotDetail = "PivotTables=0"; }
                    }
                }
                Expect(pivotSheet != null && pivotOk && rowFieldsOk && dataFieldsOk, "S2 真透视表（可拖拽）", pivotDetail);

                // ---- S3 透视表值与源直和对账（总计行）----
                bool totalOk = false; string totalDetail = "";
                if (pivotSheet != null && pivotOk)
                {
                    // 扫描透视表区域找「总计」行：金额汇总总计==amountSum
                    dynamic pt = pivotSheet.PivotTables().Item(1);
                    dynamic tr = pt.TableRange1;
                    int tRows = Convert.ToInt32(tr.Rows.Count);
                    int tCols = Convert.ToInt32(tr.Columns.Count);
                    for (int rr = 1; rr <= tRows && !totalOk; rr++)
                    {
                        for (int cc = 1; cc <= tCols; cc++)
                        {
                            var label = Convert.ToString(sheetCell(tr, rr, cc).Value2);
                            if (label != null && label.Contains("总计"))
                            {
                                // 总计行同一行向后找数值（金额总计）
                                for (int cc2 = cc + 1; cc2 <= tCols; cc2++)
                                {
                                    var v = sheetCell(tr, rr, cc2).Value2;
                                    if (v is double d && Math.Abs(d - amountSum) < 0.01) { totalOk = true; break; }
                                    var v2 = sheetCell(tr, rr + 1, cc2).Value2;
                                    if (v2 is double d2 && Math.Abs(d2 - amountSum) < 0.01) { totalOk = true; break; }
                                }
                                totalDetail = "总计@" + rr + "," + cc;
                                break;
                            }
                        }
                    }
                    Release((object)tr);
                    Release((object)pt);
                }
                Expect(totalOk, "S3 透视表总计对账", totalDetail + " 期望金额总计=" + amountSum);

                // ---- S4 切片器挂载（slicer_enabled）----
                dynamic board = workbook.Worksheets.Item["超级看板"];
                int slicerCaches = 0;
                try { slicerCaches = Convert.ToInt32(workbook.SlicerCaches().Count); } catch { }
                string slicerInResult = genJson.Contains("native") ? "native" : "无";
                Expect(slicerCaches >= 1 || slicerInResult == "native", "S4 切片器挂载", "caches=" + slicerCaches + " result=" + slicerInResult);

                // ---- S5 看板守恒回归（透视配置不干扰主链）----
                dynamic staging = workbook.Worksheets.Item["_超级看板_calc"];
                double zSales = ToD(staging.Range["Z2"].Value2);
                Expect(Math.Abs(zSales - amountSum) < 0.01, "S5 主链守恒回归", "Z2=" + zSales + " 期望=" + amountSum);

                Console.WriteLine("== SB3 结果：通过 " + _passed + "，失败 " + _failed + " ==");
                workbook.SaveAs(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SB3_super_board_pivot.xlsx"));
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][SB3] " + ex.GetType().Name + " " + ex.Message);
                return 3;
            }
            finally
            {
                foreach (var pid in SpawnedPids)
                {
                    try { Process.GetProcessById(pid)?.Kill(); } catch { }
                }
                SpawnedPids.Clear();
            }
        }

        private static dynamic sheetCell(dynamic range, int r, int c)
        {
            return range.Cells.Item[r, c];
        }
    }
}
