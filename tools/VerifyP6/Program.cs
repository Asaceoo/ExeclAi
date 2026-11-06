using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace VerifyP6
{
    /// <summary>
    /// 真机冒烟（第六期 P6-A/P6-B）：explain_excel_formula、flag_cells。
    /// 走与生产完全相同的 ExcelToolHost.CreateRegistry 链路，在真实 Excel 执行并断言。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            try { RunAll(); }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("HARNESS: " + Safe(ex));
                Console.WriteLine("HARNESS EXCEPTION: " + Safe(ex));
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine();
            Console.WriteLine("== PASSED=" + _passed + " FAILED=" + _failed + " ==");
            foreach (var f in Failures) Console.WriteLine("FAILED: " + f);
            return _failed == 0 ? 0 : 1;
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
                dynamic ws = wb.Worksheets.Item[1];
                ws.Name = "dt";
                Release((object)ws); ws = null; // 无状态 harness：生产工具自行重解析，不跨调用持有 RCW

                var workbookPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyp6-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(workbookPath), "verifyp6", "第 1 轮");
                var registry = host.CreateRegistry();

                CheckExplain(registry, app);
                CheckFlag(registry, app);
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); Release((object)wb); } } catch { }
                try { app.Quit(); } catch { }
                Release((object)app);
            }
        }

        // ===================== A6 公式解释 =====================
        private static void CheckExplain(ToolRegistry registry, dynamic app)
        {
            var t = Need(registry, "explain_excel_formula");
            Write(app, "A1", 10D); Write(app, "B1", 20D);

            // 1) CellRef 引用值来自真机：=SUM(A1,B1) 的 A1/B1 应读到 10/20
            // 无状态铁律：不得跨工具调用持有 worksheet RCW——工具内部 finally 对同名 sheet 调
            // FinalReleaseComObject 会连带销毁 harness 持有的 RCW。所有写入均用一次性辅助函数自解析、用完即放。
            SetFormula(app, "F1", "=SUM(A1,B1)");
            var j = Run(t, new Dictionary<string, object> { ["address"] = "F1", ["sheet"] = "dt" });
            Assert("explain-ok", Ok(j), "explain_excel_formula 执行成功");
            var args = Args(j);
            Assert("explain-sum-name", Str(j, "func_name") == "SUM", $"函数名识别 SUM，实际={Str(j, "func_name")}");
            Assert("explain-cellref-kinds", Str(args[0], "kind") == "cellref" && Str(args[1], "kind") == "cellref", "两个参数识别为单元格引用");
            Assert("explain-cellref-values", DNum(args[0], "value") == 10 && DNum(args[1], "value") == 20, "引用实际值读取自真机(10/20)");

            // 2) IF 三参数分类
            SetFormula(app, "E2", "=IF(E2=\"\",\"\",D2*E2)"); Write(app, "D2", 100D);
            var j2 = Run(t, new Dictionary<string, object> { ["address"] = "E2", ["sheet"] = "dt" });
            var a2 = Args(j2);
            Assert("explain-ifname", Str(j2, "func_name") == "IF", "IF 函数名识别");
            Assert("explain-ifargs", a2.Count == 3, "IF 拆分出 3 个参数");
            Assert("explain-argkinds",
                Str(a2[0], "kind") == "expression" && Str(a2[1], "kind") == "constant" && Str(a2[2], "kind") == "expression",
                "IF 参数分类 expression/constant/expression");

            // 3) 命名区域 → namedrange
            AddNamedRange(app, "Sales", "=dt!$A$1:$A$3");
            SetFormula(app, "C1", "=SUM(Sales)");
            var j3 = Run(t, new Dictionary<string, object> { ["address"] = "C1", ["sheet"] = "dt" });
            var a3 = Args(j3);
            Assert("explain-namedrange", Ok(j3) && a3.Count > 0 && Str(a3[0], "kind") == "namedrange", "命名区域参数识别为 namedrange");

            // 4) 跨表引用无法取值 → degraded 降级
            SetFormula(app, "C2", "=SUM(Sheet2!A1:B2)");
            var j4 = Run(t, new Dictionary<string, object> { ["address"] = "C2", ["sheet"] = "dt" });
            Assert("explain-degrade", Ok(j4) && Len(j4, "degraded") > 0, $"无法真机取值标记降级，degraded数={Len(j4, "degraded")}");

            // 5) 只读：explain 后 A1 仍 10、B1 仍 20
            Assert("explain-readonly", Cell(app, "A1") == "10" && Cell(app, "B1") == "20", "explain 只读，不修改源值");
        }

        // ===================== B6 标记式清洗 =====================
        private static void CheckFlag(ToolRegistry registry, dynamic app)
        {
            var t = Need(registry, "flag_cells");
            // 种子须精确填满 A1:D4（4 行），不得留空行，否则空值检测会算入多余空行。
            // C 列含日期字符串：Excel 的 Value2 会把 ISO 日期自动转成数值（OADate）→ 回读判为 numeric，
            // 坏日期(format)分支永不触发。须先把 C 列设为文本格式再写入，使日期串保持文本形态。
            dynamic colC = app.ActiveSheet.Range["C1:C4"];
            colC.NumberFormat = "@";
            Release((object)colC);
            dynamic rng = app.ActiveSheet.Range["A1:D4"];
            rng.Value2 = new object[,]
            {
                { "编号", 100D, "2026-01-01", "ok" },
                { "编号", 200D, "2026-01-05", "" },
                { "编号", 9999D, "not-a-date", "x" },
                { "编号", 150D, "2026-02-01", "ok" }
            };
            Release((object)rng);

            var j = Run(t, new Dictionary<string, object>
            {
                ["address"] = "A1:D4", ["sheet"] = "dt",
                ["flag_types"] = new JArray { "outlier", "format", "empty" }
            });
            Assert("flag-ok", Ok(j), "flag_cells 执行成功");
            if (!Ok(j)) { Console.WriteLine("  flag detail: " + Dump(j)); return; }

            Assert("flag-counts", Num(j, "flagged.outlier") == 1 && Num(j, "flagged.format") == 1 && Num(j, "flagged.empty") == 1,
                "异常/格式/空值各识别 1 处，实际=" + Dump(Pick(j, "flagged")));

            // flag_cells 内部建诊断表（Worksheets.Add）后 ActiveSheet 已切到诊断表，
            // 源数据断言必须显式按 sheet="dt" 定位，不能再依赖 app.ActiveSheet。
            Assert("flag-readonly", CellIn(app, "dt", "B3") == "9999", "flag 只标记，不修改源数据（9999 仍在 dt!B3）");

            var cOut = InteriorColorIn(app, "dt", "B3");
            var cNorm = InteriorColorIn(app, "dt", "B2");
            Assert("flag-highlight", cOut != cNorm && cOut != 16777215, $"离群格已上色且区别于正常格，B3色={cOut} B2色={cNorm}");

            var diag = Str(j, "diagnosis_sheet");
            var diagRows = Num(j, "diagnosis_rows");
            Assert("flag-diag", !string.IsNullOrEmpty(diag) && SheetExists(app, diag), "诊断 sheet 已创建，名称=" + diag + " 行=" + diagRows);
            if (!string.IsNullOrEmpty(diag))
            {
                // 诊断表前几行被 empty 占据（D 列为类型/ E 列为值），转储整表判断是否含 outlier/9999
                var dump = DumpRangeIn(app, diag, "A1:F" + diagRows);
                Assert("flag-diag-content", dump.Contains("outlier") && dump.Contains("9999"), "诊断表含 类型=outlier 与 值=9999，内容=" + dump);

                DeleteSheet(app, diag);
                Assert("flag-diag-delete", !SheetExists(app, diag), "诊断 sheet 可整体删除");
            }
        }

        // ===================== 工具 / 断言辅助 =====================
        private static JObject Run(ToolDescriptor t, Dictionary<string, object> args)
        {
            var r = t.Execute(args);
            var o = new JObject { ["ok"] = r.Ok, ["code"] = r.ErrorCode ?? "" };
            if (r.Data != null) o["data"] = JToken.Parse(JsonConvert.SerializeObject(r.Data));
            return o;
        }

        private static JArray Args(JObject j)
        {
            var d = j["data"] as JObject;
            return d?["args"] as JArray ?? new JArray();
        }

        private static ToolDescriptor Need(ToolRegistry registry, string name)
        {
            if (registry.TryGet(name, out var tt)) return tt;
            throw new InvalidOperationException("工具未注册：" + name);
        }

        private static void Assert(string id, bool condition, string message)
        {
            if (condition) { _passed++; Console.WriteLine("PASS " + id + " — " + message); }
            else { _failed++; Failures.Add(id + ": " + message); Console.WriteLine("FAIL " + id + " — " + message); }
        }

        private static bool Ok(JObject j) => j.Value<bool>("ok");
        private static string Str(JObject j, string p) => TStr(Pick(j, p));
        private static string Str(JToken t, string name) => t is JObject o ? TStr(o[name]) : null;
        private static int Num(JObject j, string p)
        {
            var t = Pick(j, p);
            return t == null ? 0 : (int)Math.Round(DVal(t));
        }
        private static int Len(JObject j, string p)
        {
            var a = Pick(j, p) as JArray;
            return a == null ? 0 : a.Count;
        }
        private static double DNum(JToken t, string name)
        {
            var v = t?[name];
            return v == null ? double.NaN : DVal(v);
        }
        private static double DVal(JToken t)
        {
            if (t == null || t.Type == JTokenType.Null) return double.NaN;
            try { return t.Value<double>(); }
            catch { double d; return double.TryParse(t.ToString(), out d) ? d : double.NaN; }
        }
        private static JToken Pick(JObject j, string p)
        {
            var cur = (JToken)j["data"];
            if (string.IsNullOrEmpty(p)) return cur;
            foreach (var seg in p.Split('.'))
            {
                if (cur == null) return null;
                cur = (cur as JObject)?[seg];
            }
            return cur;
        }
        private static string TStr(JToken t)
        {
            if (t == null || t.Type == JTokenType.Null) return null;
            return t is JValue v ? v.ToString() : t.ToString();
        }
        private static string Dump(object o)
        {
            try { return o == null ? "<null>" : Convert.ToString(o); }
            catch (Exception ex) { return "<dump: " + ex.Message + ">"; }
        }

        // ===================== Excel 辅助 =====================
        private static void Write(dynamic app, string addr, object value)
        {
            dynamic ws = app.ActiveSheet;
            try { var r = ws.Range[addr]; r.Value2 = value; Release((object)r); }
            finally { Release((object)ws); }
        }

        private static string Cell(dynamic app, string addr)
        {
            dynamic ws = app.ActiveSheet;
            try { var v = ws.Range[addr].Value2; return v == null ? string.Empty : Convert.ToString(v); }
            finally { Release((object)ws); }
        }

        private static string CellIn(dynamic app, string sheet, string addr)
        {
            dynamic wb = null; dynamic ws = null;
            try
            {
                wb = app.ActiveWorkbook;
                ws = wb.Worksheets[sheet];
                var v = ws.Range[addr].Value2;
                return v == null ? string.Empty : Convert.ToString(v);
            }
            catch (Exception ex) { return "<err:" + ex.Message + ">"; }
            finally { Release((object)ws); Release((object)wb); }
        }

        private static bool SheetExists(dynamic app, string name)
        {
            dynamic wb = null;
            try
            {
                wb = app.ActiveWorkbook;
                var c = Convert.ToInt32(wb.Worksheets.Count);
                for (var i = 1; i <= c; i++)
                    if (Convert.ToString(wb.Worksheets.Item[i].Name) == name) return true;
                return false;
            }
            catch { return false; }
            finally { Release((object)wb); }
        }

        private static void DeleteSheet(dynamic app, string name)
        {
            dynamic wb = null; dynamic ws = null;
            try { wb = app.ActiveWorkbook; ws = wb.Worksheets[name]; ws.Delete(); }
            catch { }
            finally { Release((object)ws); Release((object)wb); }
        }

        private static long InteriorColorIn(dynamic app, string sheet, string addr)
        {
            dynamic wb = null; dynamic ws = null; dynamic r = null;
            try
            {
                wb = app.ActiveWorkbook; ws = wb.Worksheets[sheet]; r = ws.Range[addr];
                var c = r.Interior.Color;
                return c == null ? 16777215 : Convert.ToInt64(c);
            }
            catch { return -1; }
            finally { Release((object)r); Release((object)ws); Release((object)wb); }
        }

        private static string DumpRangeIn(dynamic app, string sheet, string addr)
        {
            dynamic wb = null; dynamic ws = null; dynamic rng = null;
            try
            {
                wb = app.ActiveWorkbook; ws = wb.Worksheets[sheet]; rng = ws.Range[addr];
                object v = rng.Value2;
                var parts = new List<string>();
                if (v is object[,] arr)
                {
                    var lb0 = arr.GetLowerBound(0); var lb1 = arr.GetLowerBound(1);
                    var hi0 = arr.GetUpperBound(0); var hi1 = arr.GetUpperBound(1);
                    for (var r = lb0; r <= hi0; r++)
                        for (var c = lb1; c <= hi1; c++)
                        {
                            var cel = arr[r, c];
                            if (cel != null) parts.Add(Convert.ToString(cel));
                        }
                }
                else if (v != null)
                {
                    parts.Add(Convert.ToString(v));
                }
                return string.Join("|", parts);
            }
            catch { return "<err>"; }
            finally { Release((object)rng); Release((object)ws); Release((object)wb); }
        }

        // 一次性写入辅助：自解析 worksheet、用完即放，绝不跨工具调用持有 RCW
        private static void SetFormula(dynamic app, string addr, string formula)
        {
            dynamic ws = null; dynamic r = null;
            try { ws = app.ActiveSheet; r = ws.Range[addr]; r.Formula = formula; }
            finally { Release((object)r); Release((object)ws); }
        }

        private static void AddNamedRange(dynamic app, string name, string refersTo)
        {
            dynamic wb = null;
            try { wb = app.ActiveWorkbook; wb.Names.Add(name, refersTo); }
            finally { Release((object)wb); }
        }

        private static string Safe(Exception ex) => ex?.Message ?? ex?.GetType().FullName ?? "未知异常";

        private static void Release(object com)
        {
            if (com != null)
            {
                try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(com); } catch { }
            }
        }
    }
}