using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ProbeSpotlightCleanup
{
    internal static class Program
    {
        private const int XlExpression = 2;
        private const int XlEqual = 0;
        private const string SentinelName = "JYYJ_VIEW_MODE";
        private const string SentinelToken = "JYYJ_VIEW_MODE";

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            dynamic app = null;
            try
            {
                app = Marshal.GetActiveObject("Excel.Application");
                Console.WriteLine("Connected to interactive Excel: " + app.Name);
            }
            catch (Exception ex)
            {
                Console.WriteLine("NO INTERACTIVE EXCEL: " + ex.GetType().Name + " " + ex.Message);
                return 2;
            }

            dynamic wb = null;
            dynamic ws = null;
            try
            {
                wb = app.Workbooks.Add();
                ws = wb.Worksheets.Item[1];
                for (int r = 1; r <= 6; r++)
                    for (int c = 1; c <= 6; c++)
                        ws.Cells.Item[r, c].Value2 = r * 10 + c;

                wb.Names.Add(SentinelName, "=1");
                // 注入 3 条哨兵表达式规则（模拟 churn 累积后的真实状态），验证「删除后再访问同一集合」是否炸
                dynamic area = ws.Range[ws.Cells[1, 1], ws.Cells[6, 6]];
                for (int k = 0; k < 3; k++)
                {
                    dynamic fc = area.FormatConditions.Add(XlExpression, XlEqual,
                        "=AND(JYYJ_VIEW_MODE,OR(CELL(\"row\")=ROW(),CELL(\"col\")=COLUMN()))");
                    fc.Interior.ColorIndex = 6;
                    fc.Interior.Pattern = 1;
                }
                Console.WriteLine("Injected 3 sentinel rules + name.");

                int before = SentinelRuleCount(ws);
                Console.WriteLine("Before cleanup, sentinel rule count = " + before);

                // 方式1：与 Addin 完全一致（dynamic + finally ReleaseComObject）
                Console.WriteLine("\n=== 方式1: dynamic + finally ReleaseComObject ===");
                RemoveMarkedRulesV1(ws);
                Console.WriteLine("After V1, sentinel rule count = " + SentinelRuleCount(ws));

                // 方式2：dynamic + 不 ReleaseComObject（避免 finally 覆盖真实异常）
                Console.WriteLine("\n=== 方式2: dynamic 无 ReleaseComObject ===");
                RemoveMarkedRulesV2(ws);
                Console.WriteLine("After V2, sentinel rule count = " + SentinelRuleCount(ws));

                // 方式3：用 UsedRange 而非 Cells
                Console.WriteLine("\n=== 方式3: UsedRange.FormatConditions ===");
                RemoveMarkedRulesV3(ws);
                Console.WriteLine("After V3, sentinel rule count = " + SentinelRuleCount(ws));

                // 方式4：仅删除我们持有的单条引用
                Console.WriteLine("\n=== 方式4: 删除方式1创建的引用 ===");
                Console.WriteLine("After V4, sentinel rule count = " + SentinelRuleCount(ws));
            }
            catch (Exception ex)
            {
                Console.WriteLine("HARNESS EXCEPTION: " + ex.GetType().Name + " " + ex.Message);
                if (ex is COMException ce) Console.WriteLine("  HRESULT=0x" + ce.ErrorCode.ToString("X8") + " HResult=0x" + ce.HResult.ToString("X8"));
                return 1;
            }
            finally
            {
                try { wb?.Close(false); } catch { }
                try { app?.Quit(); } catch { }
                try { Release(wb); Release(ws); Release(app); } catch { }
            }

            Console.WriteLine("\nDONE");
            return 0;
        }

        private static int SentinelRuleCount(dynamic ws)
        {
            int n = 0;
            try
            {
                dynamic fcs = ws.Cells.FormatConditions;
                int count = Convert.ToInt32(fcs.Count);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic f = fcs[i];
                        string fml = Convert.ToString(f.Formula1);
                        if (fml.IndexOf(SentinelToken, StringComparison.Ordinal) >= 0) n++;
                    }
                    catch (Exception ex) { Console.WriteLine("  count-enum item " + i + " fail: " + Describe(ex)); }
                }
            }
            catch (Exception ex) { Console.WriteLine("  count-enum fail: " + Describe(ex)); }
            return n;
        }

        private static void RemoveMarkedRulesV1(dynamic sheet)
        {
            try
            {
                dynamic fcs = sheet.Cells.FormatConditions;
                int count = Convert.ToInt32(fcs.Count);
                for (int i = count; i >= 1; i--)
                {
                    dynamic fc = null;
                    try
                    {
                        fc = fcs[i];
                        string formula1 = Convert.ToString(fc.Formula1);
                        if (formula1.IndexOf(SentinelToken, StringComparison.Ordinal) >= 0)
                        {
                            fc.Delete();
                            Console.WriteLine("  V1 deleted item " + i);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("  V1 item " + i + " fail: " + Describe(ex));
                    }
                    finally
                    {
                        if (fc != null)
                        {
                            try { Marshal.ReleaseComObject(fc); }
                            catch (Exception ex) { Console.WriteLine("  V1 finally release fail: " + Describe(ex)); }
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("  V1 outer fail: " + Describe(ex)); }
        }

        private static void RemoveMarkedRulesV2(dynamic sheet)
        {
            try
            {
                dynamic fcs = sheet.Cells.FormatConditions;
                int count = Convert.ToInt32(fcs.Count);
                for (int i = count; i >= 1; i--)
                {
                    dynamic fc = null;
                    try
                    {
                        fc = fcs[i];
                        string formula1 = Convert.ToString(fc.Formula1);
                        if (formula1.IndexOf(SentinelToken, StringComparison.Ordinal) >= 0)
                        {
                            fc.Delete();
                            Console.WriteLine("  V2 deleted item " + i);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("  V2 item " + i + " fail: " + Describe(ex));
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("  V2 outer fail: " + Describe(ex)); }
        }

        private static void RemoveMarkedRulesV3(dynamic sheet)
        {
            try
            {
                dynamic fcs = sheet.UsedRange.FormatConditions;
                int count = Convert.ToInt32(fcs.Count);
                for (int i = count; i >= 1; i--)
                {
                    dynamic fc = null;
                    try
                    {
                        fc = fcs[i];
                        string formula1 = Convert.ToString(fc.Formula1);
                        if (formula1.IndexOf(SentinelToken, StringComparison.Ordinal) >= 0)
                        {
                            fc.Delete();
                            Console.WriteLine("  V3 deleted item " + i);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("  V3 item " + i + " fail: " + Describe(ex));
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("  V3 outer fail: " + Describe(ex)); }
        }

        private static string Describe(Exception ex)
        {
            if (ex is COMException ce)
                return $"{ex.GetType().Name} 0x{ce.ErrorCode:X8} {ce.Message}";
            return ex.GetType().Name + " " + ex.Message;
        }

        private static void Release(object o)
        {
            try { if (o != null) Marshal.ReleaseComObject(o); } catch { }
        }
    }
}
