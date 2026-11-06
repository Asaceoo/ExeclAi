using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using ExcelAiAssistant.Addin;

namespace SlicerProbe
{
    internal static partial class SuperBoardProbe
    {
        /// <summary>
        /// SB5 超级看板向导分步全链验收（v3.7.1.49）：
        /// Step1 create_super_table → Step2 build_pivot_sheet（新分步工具）→ Step3 render_super_dashboard（底座复用+不删 Step2 透视）；
        /// + 向导表单三页离屏渲染（DrawToBitmap）→ AI 看图验收。
        /// </summary>
        public static int RunWizard(string sourcePath)
        {
            Console.WriteLine("== SB5 超级看板向导分步全链 ==");
            if (!System.IO.File.Exists(sourcePath)) { Console.WriteLine("FILE_NOT_FOUND"); return 2; }
            string workPath = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "SB5_wizard_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".xlsx");
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

                double amountSum = 0, profitSum = 0;
                for (int r = 11; r <= 27; r++)
                {
                    amountSum += ToD(ws.Cells.Item[r, 11].Value2);
                    profitSum += ToD(ws.Cells.Item[r, 13].Value2);
                }
                Console.WriteLine("[diag] 金额和=" + amountSum + " 利润和=" + profitSum);

                var dashboard = new ExcelDashboardService(appLocalStatic);
                var pivotChart = new ExcelPivotChartService(appLocalStatic, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stub = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stub, stub, null, pivotChart, dashboard);
                ExcelThirdPhaseRegistration.RegisterThirdPhaseTools(registry, appLocalStatic, null, null);

                // ---- W1 Step1：create_super_table ----
                var w1 = ExecSuper(registry, "create_super_table", new Dictionary<string, object>
                {
                    ["source_address"] = "B10:S27", ["sheet"] = "Sheet1", ["super_sheet"] = "超级数据表",
                    ["label_column"] = 14, ["category_column"] = 5, ["month_column"] = 1,
                    ["value_column"] = 10, ["quantity_column"] = 8
                });
                Expect(w1.Ok, "W1 分步①数据表", w1.Ok ? "超级数据表已生成" : w1.ErrorCode + " " + w1.Data);

                // ---- W2 Step2：build_pivot_sheet（新分步工具）----
                var w2 = ExecSuper(registry, "build_pivot_sheet", new Dictionary<string, object>
                {
                    ["source_address"] = "B10:S27", ["sheet"] = "Sheet1",
                    ["label_column"] = 14, ["category_column"] = 5, ["month_column"] = 1,
                    ["value_column"] = 10, ["quantity_column"] = 8, ["profit_column"] = 12,
                    ["pivot_rows"] = "品类,销售员", ["pivot_measures"] = "金额,利润",
                    ["pivot_sheet_name"] = "透视分析"
                });
                Expect(w2.Ok, "W2 分步②透视表（build_pivot_sheet）", w2.Ok ? Convert.ToString(w2.Data).Substring(0, Math.Min(200, Convert.ToString(w2.Data).Length)) : w2.ErrorCode + " " + w2.Data);

                // 断言透视 sheet 存在且有透视表
                bool pivotSheetOk = false; string pivotDetail = "";
                foreach (dynamic s in workbook.Worksheets)
                {
                    if (Convert.ToString(s.Name) == "透视分析")
                    {
                        int ptCount = Convert.ToInt32(s.PivotTables().Count);
                        pivotSheetOk = ptCount >= 1;
                        pivotDetail = "PivotTables=" + ptCount;
                    }
                }
                Expect(pivotSheetOk, "W2b 透视表实体断言", pivotDetail);

                // ---- W3 Step3：render_super_dashboard（底座复用；不带 pivot_analysis）----
                var w3 = ExecSuper(registry, "render_super_dashboard", new Dictionary<string, object>
                {
                    ["source_address"] = "B10:S27", ["sheet"] = "Sheet1", ["dashboard_sheet"] = "超级看板",
                    ["label_column"] = 14, ["category_column"] = 5, ["month_column"] = 1,
                    ["value_column"] = 10, ["quantity_column"] = 8, ["profit_column"] = 12,
                    ["slicer_enabled"] = "true"
                });
                Expect(w3.Ok, "W3 分步③超级看板", w3.Ok ? "created" : w3.ErrorCode + " " + w3.Data);

                // ---- W4 主链守恒 + Step2 透视保留（未被幂等删除）----
                dynamic staging = workbook.Worksheets.Item["_超级看板_calc"];
                double zSales = ToD(staging.Range["Z2"].Value2);
                bool pivotKept = false;
                foreach (dynamic s in workbook.Worksheets)
                {
                    if (Convert.ToString(s.Name) == "透视分析") { pivotKept = Convert.ToInt32(s.PivotTables().Count) >= 1; }
                }
                Expect(Math.Abs(zSales - amountSum) < 0.01 && pivotKept,
                    "W4 主链守恒+Step2 透视保留", "Z2=" + zSales + "(期望" + amountSum + ") 透视保留=" + pivotKept);

                // ---- W5 向导表单三页离屏渲染（STA 线程）----
                string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SB5_wizard_form");
                System.IO.Directory.CreateDirectory(dir);
                Exception threadError = null;
                var cols = new List<string> { "月份", "品类", "地区", "金额", "数量", "利润" };
                var thread = new Thread(() =>
                {
                    try
                    {
                        using (var wizard = new SuperBoardWizardForm("Sheet1", "$B$10:$S$27", cols,
                            (tool, toolArgs) => ExecSuper(registry, tool, toolArgs)))
                        {
                            wizard.StartPosition = FormStartPosition.Manual;
                            wizard.Location = new Point(10, 10);
                            wizard.Show();
                            for (int step = 0; step < 3; step++)
                            {
                                wizard.WizardShowStep(step);
                                Application.DoEvents();
                                using (var bmp = new Bitmap(wizard.Width, wizard.Height))
                                {
                                    wizard.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                                    bmp.Save(System.IO.Path.Combine(dir, "wizard_step" + (step + 1) + ".png"),
                                        System.Drawing.Imaging.ImageFormat.Png);
                                }
                            }
                            wizard.Close();
                        }
                    }
                    catch (Exception ex) { threadError = ex; }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                if (threadError != null)
                {
                    Console.WriteLine("[fatal][SB5 表单渲染] " + threadError.GetType().Name + " " + threadError.Message);
                    return 3;
                }
                bool allShots = System.IO.File.Exists(System.IO.Path.Combine(dir, "wizard_step1.png"))
                             && System.IO.File.Exists(System.IO.Path.Combine(dir, "wizard_step2.png"))
                             && System.IO.File.Exists(System.IO.Path.Combine(dir, "wizard_step3.png"));
                Expect(allShots, "W5 向导三页离屏渲染", dir);

                Console.WriteLine("== SB5 结果：通过 " + _passed + "，失败 " + _failed + " ==");
                Console.WriteLine("[out] " + dir);
                var savePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SB5_wizard_board.xlsx");
                if (System.IO.File.Exists(savePath)) { System.IO.File.Delete(savePath); }
                workbook.SaveAs(savePath);
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][SB5] " + ex.GetType().Name + " " + ex.Message);
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
    }
}
