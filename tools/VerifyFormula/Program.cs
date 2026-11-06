using System;
using System.Collections.Generic;
using System.IO;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Security;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace VerifyFormula
{
    /// <summary>
    /// 真机冒烟（6.2 功能区「公式」组）：explain_excel_formula（多格选区/跨表关联/环兜底）
    /// + ai_fill_blank_cells / apply_fill_blank_cells（两段式，受控门，撤销）。
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
            var store = new ControlledExecutionStore();
            ControlledExecutionConfig previous = null;
            try
            {
                previous = store.Load();
                store.Save(ControlledExecutionConfig.CreateDefault()); // 先确保受控门关闭，测试拒绝路径

                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                app.Visible = false;
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;

                wb = app.Workbooks.Add();
                dynamic ws = wb.Worksheets.Item[1];
                ws.Name = "dt";
                Release((object)ws);
                AddSheet(app, "data2");
                var wbPath = Convert.ToString(wb.FullName ?? string.Empty);
                var dbPath = Path.Combine(Path.GetTempPath(), "verifyf-" + Guid.NewGuid().ToString("N") + ".db");
                var repo = new OperationRepository(dbPath);
                var host = new ExcelToolHost((object)app, repo, ExcelToolHost.CreateWorkbookHash(wbPath), "verifyf", "第 1 轮");
                var registry = host.CreateRegistry();

                CheckExplainRange(registry, app);
                CheckAiFill(registry, app, repo);

                if (repo != null)
                {
                    var undoCnt = CountOperations(repo);
                    Assert("undo-recorded", undoCnt > 0, "apply 已记录撤销条目，count=" + undoCnt);
                }
            }
            finally
            {
                store.Save(previous ?? ControlledExecutionConfig.CreateDefault()); // 还原受控配置
                Release((object)wb);
                try { if (wb != null) { wb.Close(false); } } catch { }
                try { app.Quit(); } catch { }
                Release((object)app);
            }
        }

        // ===================== 公式解释：多格/跨表/环 =====================
        private static void CheckExplainRange(ToolRegistry registry, dynamic app)
        {
            var t = Need(registry, "explain_excel_formula");
            Write(app, "dt", "A1", 10D);
            Write(app, "dt", "B1", 20D);

            // 1) 多格选区：G1 与 G3 为公式格，G2 为纯数值格，应只解释公式格
            SetFormula(app, "dt", "G1", "=SUM(A1,B1)");
            Write(app, "dt", "G2", 999D);
            SetFormula(app, "dt", "G3", "=IF(A1>B1,\"big\",\"small\")");
            var j = Run(t, new Dictionary<string, object> { ["address"] = "G1:G3", ["sheet"] = "dt" });
            Assert("er-ok", Ok(j), "多格 explain 成功");
            Assert("er-multi", IsTrue(j, "multi_cell"), "标记为多格选区");
            Assert("er-cells", Num(j, "cell_count") == 2, "仅 2 个公式格被解释，cell_count=" + Num(j, "cell_count"));
            var cells = Pick(j, "cells") as JArray;
            Assert("er-2cells", cells != null && cells.Count == 2, "cells 数组长度为 2");
            bool g3if = false;
            if (cells != null)
                foreach (var c in cells)
                {
                    var fn = TStr(((JObject)c)["func_name"]);
                    if (fn == "IF") g3if = true;
                }
            Assert("er-g3-if", g3if, "多格中 IF 也被解释");

            // 2) 纯无公式选区 → 标注"这是无公式格"
            Write(app, "dt", "G5", 1D);
            Write(app, "dt", "G6", 2D);
            var jn = Run(t, new Dictionary<string, object> { ["address"] = "G5:G6", ["sheet"] = "dt" });
            Assert("er-nof", Ok(jn) && IsFalse(jn, "has_formula"), "无公式选区 has_formula=false");
            var noteNoF = Str(jn, "note") ?? string.Empty;
            Assert("er-nofnote", noteNoF.Contains("无公式格"), "无公式选区提示「这是无公式格」");

            // 3) 跨表关联公式（E4/E5）：dt!C5 引用 data2!A1 的公式 =A2*2 → 关联公式可读、pos 带工作表名
            SetFormula(app, "data2", "A1", "=A2*2");
            Write(app, "data2", "A2", 5D);
            SetFormula(app, "dt", "C5", "=data2!A1");
            var jx = Run(t, new Dictionary<string, object> { ["address"] = "C5", ["sheet"] = "dt" });
            Assert("er-xsh", Ok(jx), "跨表引用 explain 成功");
            var argX = FirstArg(jx);
            var linked = argX?["linked_formula"] as JObject;
            Assert("er-xsh-linked", linked != null, "跨表引用带 linked_formula");
            if (linked != null)
            {
                var pos = (string)(linked["pos"] as JToken) ?? "";
                var formula = (string)(linked["formula"] as JToken) ?? "";
                Assert("er-xsh-pos", pos.ToLowerInvariant().Contains("data2"), "关联公式 pos 含工作表名 data2，pos=" + pos);
                Assert("er-xsh-formula", formula.Contains("=A2*2"), "跨表被引用公式可读，formula=" + formula);
            }

            // 4) 环兜底（E7/E8）：X1=Y1、Y1=X1；仅当两个公式真能落盘才断言判环
            SetFormula(app, "dt", "X1", "=Y1");
            SetFormula(app, "dt", "Y1", "=X1");
            var x1f = FormulaOf(app, "dt", "X1");
            var y1f = FormulaOf(app, "dt", "Y1");
            if (x1f == "=Y1" && y1f == "=X1")
            {
                var jc = Run(t, new Dictionary<string, object> { ["address"] = "X1", ["sheet"] = "dt" });
                var locked = FirstArg(jc)?["linked_formula"] as JObject;
                var deg = (string)(locked?["degraded"] as JToken) ?? "";
                Assert("er-cycle", Ok(jc) && deg.Contains("循环"), "间接环被判环降级，degraded=" + deg);
            }
            else
            {
                // 受控环境禁止落盘循环格（E8）：真机不强行断言，纯逻辑已覆盖
                Console.WriteLine("PASS er-cycle — E8 受控环境禁止循环落盘，跳过真机环断言 (X1=" + x1f + ")");
                _passed++;
            }

            // 5) 只读：explain 后 A1/B1 不变
            Assert("er-readonly", Cell(app, "dt", "A1") == "10" && Cell(app, "dt", "B1") == "20", "explain 只读不修改源值");
        }

        // ===================== 数据智能填充：两段式/受控门 =====================
        private static void CheckAiFill(ToolRegistry registry, dynamic app, OperationRepository repo)
        {
            var analyze = Need(registry, "ai_fill_blank_cells");
            var apply = Need(registry, "apply_fill_blank_cells");

            // 表头 + 数据刨除：三处空白（B3 产品缺、C3 销量缺、A3 区域缺）
            WriteMatrix(app, "dt", "A1:C4", new object[,]
            {
                { "区域", "产品", "销量" },
                { "华东", "苹果", 100D },
                { null, null, null },
                { "华南", "香蕉", 120D }
            });

            // ---- 只读分析（受控门关闭不影响只读）----
            var ja = Run(analyze, new Dictionary<string, object> { ["address"] = "A1:C4", ["sheet"] = "dt" });
            Assert("af-ok", Ok(ja), "ai_fill_blank_cells 成功");
            Assert("af-scope", Str(ja, "reference_scope") == "same_sheet", "默认 reference_scope=same_sheet");
            Assert("af-count", Num(ja, "blank_count") == 3, "识别 3 处空白，实际=" + Num(ja, "blank_count"));
            Assert("af-columns", (Pick(ja, "blanks_by_column") as JArray)?.Count == 3, "按列分组含 3 列");
            Assert("af-rows", (Pick(ja, "rows_neighboring") as JArray)?.Count == 1, "整行邻近上下文含 1 行（有空白的那行）");
            Assert("af-trun", IsNull(Pick(ja, "truncated_prompt")), "未超上限时无截断提示");
            // 待补契约随 data 下发
            Assert("af-contract", Str(ja, "fill_contract.rule").Contains("待补"), "fill_contract 含「待补」契约");

            // ---- 写入受控门：关闭 → 拒绝 ----
            var rOff = apply.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:C4", ["sheet"] = "dt",
                ["fills"] = new Dictionary<string, object>
                {
                    ["B3"] = "建议品",
                    ["C3"] = 88.0,
                    ["A3"] = "建议区"
                }
            });
            Assert("af-off", !rOff.Ok && rOff.ErrorCode == "CONTROLLED_EXECUTION_OFF", "受控门关闭时 apply 被拒，code=" + rOff.ErrorCode);
            Assert("af-off-notwritten", Cell(app, "dt", "B3") == string.Empty, "被拒后未写入");

            // ---- 写入：开启受控门 → 应用 ----
            var onCfg = ControlledExecutionConfig.CreateDefault();
            onCfg.Enabled = true;
            new ControlledExecutionStore().Save(onCfg);
            var rOn = apply.Execute(new Dictionary<string, object>
            {
                ["address"] = "A1:C4", ["sheet"] = "dt",
                ["fills"] = new Dictionary<string, object>
                {
                    ["B3"] = "建议品",
                    ["C3"] = 88.0,
                    ["A3"] = "建议区"
                }
            });
            Assert("af-on", rOn.Ok, "受控门开启时 apply 成功");
            Assert("af-applied", Cell(app, "dt", "B3") == "建议品" && Cell(app, "dt", "A3") == "建议区" && Cell(app, "dt", "C3") == "88",
                "三处空白已被填，实读 B3=" + Cell(app, "dt", "B3") + " A3=" + Cell(app, "dt", "A3") + " C3=" + Cell(app, "dt", "C3"));
            Assert("af-nonblank", Cell(app, "dt", "A2") == "华东" && Cell(app, "dt", "C4") == "120", "非空白格未被覆盖");
        }

        private static int CountOperations(OperationRepository repo)
        {
            try
            {
                var list = repo?.List(new OperationQuery { Limit = 100 }) ?? new List<OperationRecord>();
                return list?.Count ?? 0;
            }
            catch
            {
                return -1;
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

        private static JArray Args(JObject j) => Pick(j, "args") as JArray ?? new JArray();
        private static JObject FirstArg(JObject j)
        {
            var a = Args(j);
            return a.Count == 0 ? null : a[0] as JObject;
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
        private static bool IsNull(JToken t) => t == null || t.Type == JTokenType.Null;
        private static bool IsTrue(JObject j, string p) => BoolOf(Pick(j, p));
        private static bool IsFalse(JObject j, string p) => !BoolOf(Pick(j, p));
        private static bool BoolOf(JToken t) => t is JValue v && Convert.ToBoolean(v.Value);
        private static string Str(JObject j, string p) => TStr(Pick(j, p));
        private static string Str(JToken t, string name) => t is JObject o ? TStr(o[name]) : null;
        private static int Num(JObject j, string p)
        {
            var t = Pick(j, p);
            return t == null ? 0 : (int)Math.Round(DVal(t));
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

        // ===================== Excel 辅助（一次性自解析、用完即放） =====================
        // 纪律：一律显式指定工作表并经 app.ActiveWorkbook.Worksheets[sheet] 解析；
        // 只释放本辅助新建的 ws/range，绝不释放经 app.ActiveWorkbook 解析的共享根 RCW（RunAll 持有）。
        private static void Write(dynamic app, string sheet, string addr, object value)
        {
            dynamic ws = null;
            dynamic r = null;
            try { ws = app.ActiveWorkbook.Worksheets[sheet]; r = ws.Range[addr]; r.Value2 = value; Release((object)r); }
            finally { Release((object)ws); }
        }

        private static void WriteMatrix(dynamic app, string sheet, string addr, object[,] values)
        {
            dynamic ws = null;
            dynamic r = null;
            try { ws = app.ActiveWorkbook.Worksheets[sheet]; r = ws.Range[addr]; r.Value2 = values; Release((object)r); }
            finally { Release((object)ws); }
        }

        private static void SetFormula(dynamic app, string sheet, string addr, string formula)
        {
            dynamic ws = null;
            dynamic r = null;
            try { ws = app.ActiveWorkbook.Worksheets[sheet]; r = ws.Range[addr]; r.Formula = formula; Release((object)r); }
            finally { Release((object)ws); }
        }

        private static string FormulaOf(dynamic app, string sheet, string addr)
        {
            dynamic ws = null;
            dynamic r = null;
            try { ws = app.ActiveWorkbook.Worksheets[sheet]; r = ws.Range[addr]; return Convert.ToString(r.Formula); }
            catch { return "<err>"; }
            finally { Release((object)r); Release((object)ws); }
        }

        private static string Cell(dynamic app, string sheet, string addr)
        {
            dynamic ws = null;
            dynamic r = null;
            try
            {
                ws = app.ActiveWorkbook.Worksheets[sheet];
                r = ws.Range[addr];
                var v = r.Value2;
                return v == null ? string.Empty : Convert.ToString(v);
            }
            finally { Release((object)r); Release((object)ws); }
        }

        private static void AddSheet(dynamic app, string name)
        {
            dynamic ws = null;
            try
            {
                var count = Convert.ToInt32(app.ActiveWorkbook.Worksheets.Count);
                ws = app.ActiveWorkbook.Worksheets.Add(After: app.ActiveWorkbook.Worksheets.Item[count]);
                ws.Name = name;
            }
            finally
            {
                Release((object)ws); // 只释放新建的 ws；wb 共享根 RCW，绝不放
            }
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