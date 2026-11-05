using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace VerifySpotlightProbe
{
    /// <summary>
    /// 真机探针 v2：验证「单表达式定义名 JYYJ_SEL」方案——
    /// 1) late-binding 能否写 Names.Item.RefersTo（长表达式）；
    /// 2) 写 JYYJ_SEL 后不 Calculate 是否跟随；
    /// 3) 1 次写 + Calculate 的耗时 vs 4 次写；
    /// 4) 多选（Ctrl 多区域）逐区域判定是否正确（A1:A5 + C1:C5 → A/C 列+行1-5，B 列不亮）；
    /// 5) 整列/整行/全选收敛行为。
    /// </summary>
    internal static class Program
    {
        private const string SentinelName = "JYYJ_VIEW_MODE";
        private const string SelName = "JYYJ_SEL";
        private const int FullRows = 1048576;
        private const int FullCols = 16384;
        private static int _passed;
        private static int _failed;

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            dynamic app = null;
            dynamic wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                wb = app.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                ws.Name = "small";
                for (int r = 1; r <= 8; r++)
                    for (int c = 1; c <= 8; c++)
                        ws.Cells.Item[r, c].Value2 = r * 10 + c;

                wb.Names.Add(SentinelName, "=1");
                wb.Names.Add(SelName, "=FALSE");
                wb.Names.Add("JYYJ_SEL_R1", "=1");
                wb.Names.Add("JYYJ_SEL_R2", "=1");
                wb.Names.Add("JYYJ_SEL_C1", "=1");
                wb.Names.Add("JYYJ_SEL_C2", "=1");
                dynamic fc = ws.Cells.FormatConditions.Add(2, 0, "=AND(JYYJ_VIEW_MODE,JYYJ_SEL)");
                fc.Interior.ColorIndex = 6;
                fc.Interior.Pattern = 1;
                app.Range["B2"].Select();
                ws.Calculate();

                // ===== 1) late-binding 写 Names.Item.RefersTo（短值） =====
                bool refOk = true;
                try { wb.Names.Item(SelName).RefersTo = "=OR(AND(ROW()>=2,ROW()<=2),AND(COLUMN()>=2,COLUMN()<=2))"; }
                catch (Exception ex) { refOk = false; Console.WriteLine("  [ERR] RefersTo 写入: " + ex.Message); }
                Assert(refOk, "late-binding 写 Names.Item.RefersTo（表达式）成功");

                // ===== 2) 写 JYYJ_SEL 后不 Calculate 是否跟随 =====
                wb.Names.Item(SelName).RefersTo = "=OR(AND(ROW()>=4,ROW()<=4),AND(COLUMN()>=4,COLUMN()<=4))";
                Thread.Sleep(300);
                int cFollow = ColorOf(ws, 4, 1);
                Assert(cFollow == 6, "单写 RefersTo(无Calculate) D4 行内 A4 黄=" + cFollow + " 期望6");
                int cDiag = ColorOf(ws, 1, 1);
                Assert(cDiag != 6, "单写 RefersTo(无Calculate) A1 非黄=" + cDiag);

                // ===== 3) 单格耗时：1写 vs 4写（各+Calculate） =====
                Measure(10, "1-write+Calc (JYYJ_SEL)", () =>
                {
                    wb.Names.Item(SelName).RefersTo = "=OR(AND(ROW()>=5,ROW()<=5),AND(COLUMN()>=5,COLUMN()<=5))";
                    ws.Calculate();
                });
                Measure(10, "4-write+Calc (old R1R2C1C2)", () =>
                {
                    wb.Names.Item("JYYJ_SEL_R1").RefersTo = "=5";
                    wb.Names.Item("JYYJ_SEL_R2").RefersTo = "=5";
                    wb.Names.Item("JYYJ_SEL_C1").RefersTo = "=5";
                    wb.Names.Item("JYYJ_SEL_C2").RefersTo = "=5";
                    ws.Calculate();
                });
                Measure(10, "Calc-only", () => ws.Calculate());

                // ===== 4) 多选 A1:A5 + C1:C5 逐区域判定 =====
                app.Range["A1:A5,C1:C5"].Select();
                dynamic sel = app.Selection;
                int areas = Convert.ToInt32(sel.Areas.Count);
                Assert(areas == 2, "多选区域数=" + areas + " 期望2");
                string predicate = BuildPredicate(sel);
                Console.WriteLine("  [DUMP] 多选 predicate=" + predicate);
                wb.Names.Item(SelName).RefersTo = "=" + predicate;
                ws.Calculate();
                Thread.Sleep(300);

                Assert(ColorOf(ws, 2, 1) == 6, "多选 A 列(2,1) 黄=6");
                Assert(ColorOf(ws, 4, 1) == 6, "多选 A 列(4,1) 黄=6");
                Assert(ColorOf(ws, 2, 3) == 6, "多选 C 列(2,3) 黄=6");
                Assert(ColorOf(ws, 1, 3) == 6, "多选 C 列(1,3) 黄=6");
                Assert(ColorOf(ws, 2, 2) != 6, "多选 B 列(2,2) 非黄（未选B列）");
                Assert(ColorOf(ws, 6, 1) != 6, "多选 第6行(6,1) 非黄（只选到第5行）");
                Assert(ColorOf(ws, 3, 4) != 6, "多选 D3 非黄（C 列外）");

                // ===== 5) 连续多格选区 A1:B4（单区域）→ 行1-4 ∪ 列A-B =====
                app.Range["A1:B4"].Select();
                sel = app.Selection;
                predicate = BuildPredicate(sel);
                Console.WriteLine("  [DUMP] 连续区 predicate=" + predicate);
                wb.Names.Item(SelName).RefersTo = "=" + predicate;
                ws.Calculate();
                Thread.Sleep(300);
                Assert(ColorOf(ws, 3, 1) == 6, "连续区 A3 黄=6（行3在选区）");
                Assert(ColorOf(ws, 1, 2) == 6, "连续区 B1 黄=6（B列在选区）");
                Assert(ColorOf(ws, 5, 1) != 6, "连续区 A5 非黄（行5不在选区）");
                Assert(ColorOf(ws, 3, 3) != 6, "连续区 C3 非黄（C列不在选区）");

                // ===== 6) 整列/整行/全选收敛 =====
                app.Columns["A"].Select();
                sel = app.Selection;
                predicate = BuildPredicate(sel);
                Console.WriteLine("  [DUMP] 整列A predicate=" + predicate);
                wb.Names.Item(SelName).RefersTo = "=" + predicate;
                ws.Calculate();
                Thread.Sleep(300);
                Assert(ColorOf(ws, 1, 1) == 6, "整列A A1 黄=6");
                Assert(ColorOf(ws, 8, 1) == 6, "整列A A8 黄=6");
                Assert(ColorOf(ws, 1, 2) != 6, "整列A B1 非黄");

                app.Rows["1"].Select();
                sel = app.Selection;
                predicate = BuildPredicate(sel);
                Console.WriteLine("  [DUMP] 整行1 predicate=" + predicate);
                wb.Names.Item(SelName).RefersTo = "=" + predicate;
                ws.Calculate();
                Thread.Sleep(300);
                Assert(ColorOf(ws, 1, 2) == 6, "整行1 B1 黄=6");
                Assert(ColorOf(ws, 2, 1) != 6, "整行1 A2 非黄");

                // ===== 7) 指针变体：边框框选区 =====
                app.Range["B2:D3"].Select();
                sel = app.Selection;
                predicate = BuildPointerPredicate(sel);
                Console.WriteLine("  [DUMP] 指针 predicate=" + predicate);
                // 建指针规则（边框），聚光灯规则先删
                for (int i = ws.Cells.FormatConditions.Count; i >= 1; i--)
                {
                    dynamic f = ws.Cells.FormatConditions.Item[i];
                    string f1 = "";
                    try { f1 = Convert.ToString(f.Formula1); } catch { }
                    if (f1.IndexOf(SentinelName, StringComparison.Ordinal) >= 0) f.Delete();
                }
                dynamic fp = ws.Cells.FormatConditions.Add(2, 0, "=AND(JYYJ_VIEW_MODE," + predicate + ")");
                fp.Borders.ColorIndex = 3;
                fp.Borders.LineStyle = 1;
                wb.Names.Item(SelName).RefersTo = "=" + predicate;
                ws.Calculate();
                Thread.Sleep(300);
                Assert(BorderOf(ws, 2, 2, 7) == 3, "指针 B2 左边框红=3");
                Assert(BorderOf(ws, 3, 3, 7) == 3, "指针 D3 左边框红=3");
                Assert(BorderOf(ws, 1, 1, 7) != 3, "指针 A1 无红框");

                // ===== 8) 大表耗时 =====
                dynamic ws2 = wb.Worksheets.Add();
                ws2.Name = "big";
                object[,] big = new object[1000, 100];
                for (int r = 0; r < 1000; r++)
                    for (int c = 0; c < 100; c++)
                        big[r, c] = r * 100 + c;
                ws2.Range["A1:CJ1000"].Value2 = big;
                dynamic fc2 = ws2.Cells.FormatConditions.Add(2, 0, "=AND(JYYJ_VIEW_MODE,JYYJ_SEL)");
                fc2.Interior.ColorIndex = 6;
                fc2.Interior.Pattern = 1;
                ws2.Activate();
                app.Range["B2"].Select();
                ws2.Calculate();
                Measure(10, "BIG 1-write+Calc", () =>
                {
                    wb.Names.Item(SelName).RefersTo = "=OR(AND(ROW()>=30,ROW()<=30),AND(COLUMN()>=30,COLUMN()<=30))";
                    ws2.Calculate();
                });

                wb.Close(false);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("HARNESS: " + ex);
            }
            finally
            {
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    Marshal.FinalReleaseComObject(app);
                }
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            return _failed > 0 ? 1 : 0;
        }

        /// <summary>聚光灯：逐区域「行∪列」十字；整列→只该列；整行→只该行；全选→无。</summary>
        private static string BuildPredicate(dynamic selection)
        {
            dynamic areas = selection.Areas;
            int count = Convert.ToInt32(areas.Count);
            var clauses = new System.Collections.Generic.List<string>();
            for (int i = 1; i <= count; i++)
            {
                dynamic area = areas[i];
                int r1 = Convert.ToInt32(area.Row);
                int c1 = Convert.ToInt32(area.Column);
                int r2 = r1 + Convert.ToInt32(area.Rows.Count) - 1;
                int c2 = c1 + Convert.ToInt32(area.Columns.Count) - 1;
                bool fullRow = (r2 - r1 + 1) >= FullRows;
                bool fullCol = (c2 - c1 + 1) >= FullCols;
                if (fullRow && fullCol) continue;
                if (fullRow) clauses.Add("AND(COLUMN()>=" + c1 + ",COLUMN()<=" + c2 + ")");
                else if (fullCol) clauses.Add("AND(ROW()>=" + r1 + ",ROW()<=" + r2 + ")");
                else clauses.Add("OR(AND(ROW()>=" + r1 + ",ROW()<=" + r2 + "),AND(COLUMN()>=" + c1 + ",COLUMN()<=" + c2 + "))");
            }
            if (clauses.Count == 0) return "FALSE";
            return "OR(" + string.Join(",", clauses.Select(c => "(" + c + ")")) + ")";
        }

        /// <summary>指针：逐区域包围盒；整行/整列收敛到单格框。</summary>
        private static string BuildPointerPredicate(dynamic selection)
        {
            dynamic areas = selection.Areas;
            int count = Convert.ToInt32(areas.Count);
            var clauses = new System.Collections.Generic.List<string>();
            for (int i = 1; i <= count; i++)
            {
                dynamic area = areas[i];
                int r1 = Convert.ToInt32(area.Row);
                int c1 = Convert.ToInt32(area.Column);
                int r2 = r1 + Convert.ToInt32(area.Rows.Count) - 1;
                int c2 = c1 + Convert.ToInt32(area.Columns.Count) - 1;
                bool fullRow = (r2 - r1 + 1) >= FullRows;
                bool fullCol = (c2 - c1 + 1) >= FullCols;
                if (fullRow && fullCol) continue;
                if (fullRow || fullCol)
                {
                    // 收敛到活动格单格框
                    continue;
                }
                clauses.Add("AND(ROW()>=" + r1 + ",ROW()<=" + r2 + ",COLUMN()>=" + c1 + ",COLUMN()<=" + c2 + ")");
            }
            if (clauses.Count == 0) return "FALSE";
            return "OR(" + string.Join(",", clauses.Select(c => "(" + c + ")")) + ")";
        }

        private static int ColorOf(dynamic ws, int r, int c)
        {
            for (int tryN = 0; tryN < 10; tryN++)
            {
                try { return Convert.ToInt32(ws.Cells.Item[r, c].DisplayFormat.Interior.ColorIndex); }
                catch { Thread.Sleep(100); }
            }
            return -9999;
        }

        private static int BorderOf(dynamic ws, int r, int c, int edge)
        {
            for (int tryN = 0; tryN < 10; tryN++)
            {
                try { return Convert.ToInt32(ws.Cells.Item[r, c].DisplayFormat.Borders.Item[edge].ColorIndex); }
                catch { Thread.Sleep(100); }
            }
            return -9999;
        }

        private static void Measure(int n, string label, Action action)
        {
            var times = new long[n];
            for (int k = 0; k < n; k++)
            {
                var sw = Stopwatch.StartNew();
                action();
                sw.Stop();
                times[k] = sw.ElapsedMilliseconds;
            }
            long sum = 0, max = 0;
            foreach (var t in times) { sum += t; if (t > max) max = t; }
            Console.WriteLine("  " + label + " x" + n + " avg=" + (sum / (double)n).ToString("0.0") + "ms max=" + max + "ms");
        }

        private static void Assert(bool cond, string msg)
        {
            if (cond) { _passed++; Console.WriteLine("[PASS] " + msg); }
            else { _failed++; Console.WriteLine("[FAIL] " + msg); }
        }
    }
}
