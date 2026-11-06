using System;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;

namespace SlicerProbe
{
    /// <summary>
    /// SB0 超级看板先导探针（2026-09-12，FR-SB-9/FR-SB-7 实现前定生死）：
    /// ① Shapes.AddShape 圆角矩形 + 阴影（视频 KPI 卡观感的 COM 可达性）；
    /// ② 文本框绑定单元格公式三条 COM 路径（TextBoxes 集合 / DrawingObject.Formula / OLEFormat）
    ///    ——历史教训：dynamic 下 Shape.DrawingObject 解析不出 .Chart，.Formula 可达性未知，探针先行；
    /// ③ 指标选择器镜像联动：数据验证下拉 + CHOOSE(MATCH()) 镜像公式，改下拉格→镜像区重算。
    /// 独立 out-of-proc Excel + PID 强杀清理（踩坑手册 §5.2）。
    /// </summary>
    internal static partial class SuperBoardProbe
    {
        private static readonly HashSet<int> SpawnedPids = new HashSet<int>();
        private static int _passed;
        private static int _failed;
        private static dynamic appLocalStatic;

        private static void Expect(bool ok, string id, string detail)
        {
            Console.WriteLine((ok ? "[PASS] " : "[FAIL] ") + id + (ok ? "" : " → " + detail));
            if (ok) { _passed++; } else { _failed++; }
        }

        public static int Run()
        {
            Console.WriteLine("== SB0 超级看板先导探针 ==");
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
                workbook = excel.Workbooks.Add();

                // ---- 测试数据：月份/金额/数量/利润 12 行 ----
                dynamic ws = workbook.Worksheets.Item[1];
                ws.Name = "数据";
                ws.Range["A1"].Value2 = "月份"; ws.Range["B1"].Value2 = "金额"; ws.Range["C1"].Value2 = "数量"; ws.Range["D1"].Value2 = "利润";
                for (int m = 1; m <= 12; m++)
                {
                    ws.Cells.Item[m + 1, 1].Value2 = m + "月";
                    ws.Cells.Item[m + 1, 2].Value2 = 1000.0 * m;
                    ws.Cells.Item[m + 1, 3].Value2 = 10 * m;
                    ws.Cells.Item[m + 1, 4].Value2 = 300.0 * m;
                }

                // ---- P1 圆角矩形 + 阴影 ----
                dynamic shape = null;
                try
                {
                    shape = ws.Shapes.AddShape(5, 8, 40, 140, 64); // msoShapeRoundedRectangle
                    shape.Fill.ForeColor.RGB = 0x00D7EDAD;
                    shape.Line.Visible = 0;
                    dynamic sh = shape.Shadow;
                    // 2026-09-12 SB0 迭代：Type=25 报错根因未知，先只设 Transparency/Blur/Offset（视频观感三要素）
                    sh.ForeColor.RGB = 0x000000;
                    sh.Transparency = 0.65;
                    sh.Blur = 12;
                    sh.OffsetX = 0; sh.OffsetY = 4;
                    double tRead = Convert.ToDouble(sh.Transparency);
                    double bRead = Convert.ToDouble(sh.Blur);
                    Release(sh);
                    bool ok = Math.Abs(tRead - 0.65) < 0.001 && Math.Abs(bRead - 12) < 0.001;
                    Release(shape); shape = null;
                    Expect(ok, "P1 圆角矩形+阴影", "Transparency=" + tRead + " Blur=" + bRead);
                }
                catch (Exception ex)
                {
                    Release(shape);
                    Expect(false, "P1 圆角矩形+阴影", ex.GetType().Name + " " + ex.Message);
                }

                // ---- P2 文本框绑定单元格公式（赋值与读回分离；读回走 TextFrame2）----
                // SB0#3 修正：VBA 官方语义 TextBox.Formula="A1" 是【纯单元格引用】（不带=、非函数表达式），
                // 之前传 "=SUM(...)" 被 Excel 拒绝属误用；绑定镜像单元格（staging 模式）本就是设计形态
                bool bindSet = false; string bindRead = null; string bindBy = "DrawingObject";
                dynamic tb = null;
                try
                {
                    tb = ws.Shapes.AddTextbox(1, 150, 40, 140, 64);
                    dynamic drawing = tb.DrawingObject;
                    drawing.Formula = "数据!B2";   // 纯引用（无=）：VBA TextBox.Formula 官方语义
                    bindSet = true;
                    Release(drawing); drawing = null;
                    dynamic tf2 = tb.TextFrame2;
                    bindRead = Convert.ToString(tf2.TextRange.Text);
                    Release(tf2); tf2 = null;
                }
                catch (Exception exa)
                {
                    Console.WriteLine("[diag] 路径a 纯引用 失败: " + exa.GetType().Name + " " + exa.Message);
                    Release(tb); tb = null;
                }
                if (!bindSet)
                {
                    try
                    {
                        dynamic tbs = ws.TextBoxes();
                        dynamic tb2 = tbs.Add(150, 40, 140, 64);
                        tb2.Formula = "=数据!B2";   // 变体：带等号引用
                        bindSet = true; bindBy = "TextBoxes(=ref)";
                        bindRead = Convert.ToString(tb2.TextFrame2.TextRange.Text);
                        Release(tb2); Release(tbs); tb2 = null; tbs = null;
                    }
                    catch (Exception exb)
                    {
                        Console.WriteLine("[diag] 路径b =引用 失败: " + exb.GetType().Name + " " + exb.Message);
                    }
                }
                // 联动验证：绑定成功后改源值，看文本框显示值是否跟随（=公式绑定的直接证据），验证后还原现场
                bool linkAlive = false;
                if (bindSet)
                {
                    try
                    {
                        ws.Range["B2"].Value2 = 999999.0;
                        dynamic tf2v = null; dynamic tbv = ws.Shapes.Item(2);
                        tf2v = tbv.TextFrame2;
                        string after = Convert.ToString(tf2v.TextRange.Text);
                        Release(tf2v); Release(tbv); tf2v = null; tbv = null;
                        linkAlive = after.Contains("999999");   // 绑定的是数据!B2 单格，改值后文本应实时变
                        ws.Range["B2"].Value2 = 1000.0;         // 还原现场（P3 期望值按原始数据计算）
                        Console.WriteLine("[diag] P2 联动后文本=" + after + "（已还原 B2=1000）");
                    }
                    catch (Exception exl) { Console.WriteLine("[diag] P2 联动读失败: " + exl.Message); }
                }
                Expect(bindSet && bindRead != null && linkAlive, "P2 文本框绑定单元格公式+联动",
                    "set=" + bindSet + " read=" + (bindRead ?? "null") + " linkAlive=" + linkAlive + " path=" + bindBy);

                // ---- P3 指标选择器镜像联动（下拉 + CHOOSE(MATCH) 镜像公式）----
                try
                {
                    // 视图名清单 A1:A2；两视图 KPI 值 C 列（视图1）与 D 列（视图2）第 2 行
                    ws.Range["F1"].Value2 = "销售视图"; ws.Range["F2"].Value2 = "利润视图";
                    ws.Range["G2"].Formula = "=SUM(数据!B2:B13)";   // 销售视图 KPI
                    ws.Range["H2"].Formula = "=SUM(数据!D2:D13)";   // 利润视图 KPI
                    // 看板侧控制格 I1（下拉）+ 镜像格 I2（CHOOSE/MATCH）
                    dynamic v = ws.Range["I1"].Validation;
                    v.Delete();
                    v.Add(Type: 3, Formula1: "=$F$1:$F$2");
                    Release(v); v = null;
                    ws.Range["I1"].Value2 = "销售视图";
                    ws.Range["I2"].Formula = "=CHOOSE(MATCH($I$1,$F$1:$F$2,0),$G$2,$H$2)";
                    double v1 = Convert.ToDouble(ws.Range["I2"].Value2);
                    ws.Range["I1"].Value2 = "利润视图";
                    double v2 = Convert.ToDouble(ws.Range["I2"].Value2);
                    // 期望：v1=金额和(78000) v2=利润和(23400)
                    bool ok = Math.Abs(v1 - 78000) < 0.01 && Math.Abs(v2 - 23400) < 0.01;
                    Expect(ok, "P3 下拉+CHOOSE 镜像联动", "v1=" + v1 + " v2=" + v2 + " 期望 78000/23400");
                }
                catch (Exception ex)
                {
                    Expect(false, "P3 下拉+CHOOSE 镜像联动", ex.GetType().Name + " " + ex.Message);
                }

                Console.WriteLine("== SB0 结果：通过 " + _passed + "，失败 " + _failed + " ==");
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[fatal][SB0] " + ex.GetType().Name + " " + ex.Message);
                return 3;
            }
            finally
            {
                // PID 强杀清理（不 Close/Quit——out-of-proc 动态 COM 退出红线）
                foreach (var pid in SpawnedPids)
                {
                    try { Process.GetProcessById(pid)?.Kill(); } catch { }
                }
                SpawnedPids.Clear();
            }
        }

        private static void Release(object o)
        {
            if (o != null && System.Runtime.InteropServices.Marshal.IsComObject(o))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
        }
    }
}
