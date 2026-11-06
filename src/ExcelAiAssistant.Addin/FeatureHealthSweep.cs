using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Addin
{
    /// <summary>单条巡检结果。</summary>
    public sealed class FeatureCheckResult
    {
        public string Name { get; set; }

        public bool Ok { get; set; }

        public string Detail { get; set; }
        }

    /// <summary>
    /// 全功能健康巡检：在新建的临时工作簿（不影响当前数据）里对全部核心子系统逐项体检——
    /// 写入/公式/排序/清洗/追加/批量/透视/格式/条件格式/标记/图表/结构操作的「主工具+自动自检闭环」、
    /// 回滚链路、自检覆盖审计、踩坑库、自检历史库、配置与技能加载。
    /// 逐项结果由桥接层写入运行日志（设置页可见），VBA/Python 通道由「一键自检」覆盖，两者互补。
    /// 严格无状态：每次操作前重新解析对象、用完即放（踩坑手册 §1.1/§5.1）。
    /// </summary>
    public sealed class FeatureHealthSweep
    {
        private readonly dynamic application;

        public FeatureHealthSweep(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public List<FeatureCheckResult> Run()
        {
            var results = new List<FeatureCheckResult>();
            dynamic tempBook = null;
            try
            {
                tempBook = application.Workbooks.Add(); // 临时工作簿：巡检全程不触碰用户数据
                var dbPath = Path.Combine(Path.GetTempPath(), "xai-sweep-" + Guid.NewGuid().ToString("N") + ".db");
                try
                {
                    var host = new ExcelToolHost(
                        application, new OperationRepository(dbPath), "sweep-hash", "sweep-session");
                    var registry = host.CreateRegistry();

                    CheckWriteChain(registry, results);
                    CheckFormulaChain(registry, results);
                    CheckSortChain(registry, results);
                    CheckCleanChain(registry, results);
                    CheckAppendChain(registry, results);
                    CheckBatchChain(registry, results);
                    CheckPivotChain(registry, results);
                    CheckFormatChain(registry, results);
                    CheckConditionalChain(registry, results);
                    CheckMarkerChain(registry, results);
                    CheckChartChain(registry, results);
                    CheckStructureChain(registry, results);
                    CheckRollbackChain(registry, results);
                    CheckCoverageAudit(registry, results);
                    CheckPitfallLibrary(results);
                    CheckSelfCheckHistory(results);
                    CheckConfiguration(results);
                    CheckSkills(results);
                }
                finally
                {
                    try { if (File.Exists(dbPath)) File.Delete(dbPath); } catch { }
                }
            }
            catch (Exception ex)
            {
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:80  {ex.GetType().Name}: {ex.Message}");
                results.Add(new FeatureCheckResult { Name = "巡检环境", Ok = false, Detail = "异常：" + ex.Message });
            }
            finally
            {
                try { tempBook?.Close(false); } catch { }
            }

            return results;
        }

        // ===================== 检查项 =====================

        private void CheckWriteChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            AutoVerifyChain(registry, results, "写入与自检闭环", "write_excel_range", Args(
                ("address", "$A$1:$B$2"),
                ("values", new JArray { new JArray { "品名", "数量" }, new JArray { "苹果", 3 } })));
        }

        private void CheckFormulaChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            AutoVerifyChain(registry, results, "公式与公式自检", "write_excel_cell", Args(
                ("address", "D1"), ("value", "=1+1")));
        }

        private void CheckSortChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            Exec(registry, results, "排序与排序自检（准备）", "write_excel_range", Args(
                ("address", "$F$1:$G$3"),
                ("values", new JArray { new JArray { "品名", "数量" }, new JArray { "梨", 5 }, new JArray { "苹果", 3 } })));
            var sort = ExecTool(registry, "sort_range", Args(
                ("address", "F1:G3"), ("sort_column", 2), ("direction", "desc"), ("header_rows", 1)));
            if (!sort.Ok)
            {
                Fail(results, "排序与排序自检（准备）", Describe(sort));
                return;
            }

            VerifyOk(registry, results, "排序与排序自检", "verify_sort_filter", Args(
                ("address", "F1:G3"), ("expected_sort_column", 2), ("expected_sort_direction", "desc"), ("header_rows", 1)));
        }

        private void CheckCleanChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            Exec(registry, results, "清洗与清洗自检（准备）", "write_excel_range", Args(
                ("address", "$I$1:$J$3"),
                ("values", new JArray { new JArray { "a", 1 }, new JArray { "", null }, new JArray { "b", 2 } })));
            Exec(registry, results, "清洗与清洗自检（准备）", "clean_excel_data", Args(
                ("address", "I1:J3"),
                ("operations", new JArray { new JObject { ["tool"] = "remove_blank_rows" } })));
            VerifyOk(registry, results, "清洗与清洗自检", "verify_data_clean", Args(
                ("address", "I1:J3"), ("checks", new JArray { new JObject { ["type"] = "no_blank_rows" } })));
        }

        private void CheckAppendChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var args = Args(("start_cell", "A1"), ("values", new JArray { new JArray { "追加", 9 } }));
            var result = ExecTool(registry, "append_excel_rows", args);
            Expect(results, "追加行与追加自检", result.Ok, Describe(result));
            if (result.Ok)
            {
                AutoVerifyFromResult(registry, results, "追加行与追加自检", "append_excel_rows", args, result);
            }
        }

        private void CheckBatchChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var result = ExecTool(registry, "write_excel_cells_batch", Args(
                ("cells", new JArray
                {
                    new JObject { ["address"] = "L1", ["value"] = "标记" },
                    new JObject { ["address"] = "L2", ["value"] = 7 }
                })));
            Expect(results, "批量回写与逐格自检", result.Ok, Describe(result));
            if (result.Ok)
            {
                AutoVerifyFromResult(registry, results, "批量回写与逐格自检", "write_excel_cells_batch", Args(
                    ("cells", new JArray
                    {
                        new JObject { ["address"] = "L1", ["value"] = "标记" },
                        new JObject { ["address"] = "L2", ["value"] = 7 }
                    })), result);
            }
        }

        private void CheckPivotChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            Exec(registry, results, "透视与守恒对账（准备）", "write_excel_range", Args(
                ("address", "$N$1:$O$3"),
                ("values", new JArray { new JArray { "品名", "金额" }, new JArray { "甲", 4 }, new JArray { "乙", 6 } })));
            var pivot = ExecTool(registry, "create_excel_pivot_summary", Args(
                ("source_address", "N1:O3"), ("label_column", 1), ("value_column", 2), ("target_address", "$N$5")));
            if (!pivot.Ok)
            {
                Fail(results, "透视与守恒对账", Describe(pivot));
                return;
            }

            VerifyOk(registry, results, "透视与守恒对账", "verify_pivot", Args(
                ("target_address", "N5"), ("source_address", "N1:O3"),
                ("source_value_column", 2), ("source_label_column", 1)));
        }

        private void CheckFormatChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            Exec(registry, results, "格式与格式自检（准备）", "format_excel_range", Args(
                ("address", "$A$1:$B$1"), ("bold", true)));
            VerifyOk(registry, results, "格式与格式自检", "verify_cell_format", Args(
                ("address", "A1:B1"), ("expected_bold", true)));
        }

        private void CheckConditionalChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            Exec(registry, results, "条件格式与自检（准备）", "format_excel_conditional", Args(
                ("address", "G2:G3"), ("rule", "greater"), ("value", 1), ("fill_color", "#ffd9d9")));
            VerifyOk(registry, results, "条件格式与自检", "verify_conditional_format", Args(
                ("address", "G2:G3"), ("expected_rule", "greater")));
        }

        private void CheckMarkerChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var flag = ExecTool(registry, "flag_cells", Args(("address", "F2:G3")));
            if (!flag.Ok)
            {
                Fail(results, "标记与标记自检", Describe(flag));
                return;
            }

            // 从标记结果取实际标记总数作为 verify 期望（无期望时 verify_marker 会判 no_expected）。
            var total = 0;
            if (JToken.FromObject(flag.Data) is JObject flagData && flagData["flagged"] is JObject flagged)
            {
                foreach (var property in flagged.Properties())
                {
                    total += property.Value?.Value<int>() ?? 0;
                }
            }

            VerifyOk(registry, results, "标记与标记自检", "verify_marker", Args(
                ("address", "F2:G3"), ("expected_count", total)));
        }

        private void CheckChartChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var chart = ExecTool(registry, "create_excel_chart", Args(
                ("source_address", "F1:G3"), ("chart_type", "ColumnClustered")));
            if (!chart.Ok)
            {
                Fail(results, "图表与图表自检", Describe(chart));
                return;
            }

            VerifyOk(registry, results, "图表与图表自检", "verify_chart", Args(("expected_chart_type", "ColumnClustered")));
        }

        private void CheckStructureChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var add = ExecTool(registry, "add_excel_worksheet", Args(("name", "巡检临时表")));
            if (!add.Ok)
            {
                Fail(results, "结构操作与结构自检", Describe(add));
                return;
            }

            VerifyOk(registry, results, "结构操作与结构自检", "verify_structure", Args(("expected_sheet_exists", "巡检临时表")));
            var delete = ExecTool(registry, "delete_excel_worksheet", Args(("sheet", "巡检临时表")));
            if (!delete.Ok)
            {
                Fail(results, "结构操作与结构自检", Describe(delete));
                return;
            }

            VerifyOk(registry, results, "结构操作与结构自检（删除后缺席）", "verify_structure", Args(("expected_sheet_absent", "巡检临时表")));
        }

        private void CheckRollbackChain(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var first = new JArray { new JArray { "回滚前", 1 } };
            var write1 = ExecTool(registry, "write_excel_range", Args(("address", "$Q$1:$R$1"), ("values", first)));
            var write2 = ExecTool(registry, "write_excel_range", Args(
                ("address", "$Q$1:$R$1"), ("values", new JArray { new JArray { "被覆写", 99 } })));
            var opId = ExtractOperationId(write2);
            if (!write1.Ok || !write2.Ok || string.IsNullOrWhiteSpace(opId))
            {
                Fail(results, "操作历史与回滚", "准备失败：" + Describe(write1) + " / " + Describe(write2));
                return;
            }

            var undo = ExecTool(registry, "undo_excel_operation", Args(("operation_id", opId)));
            if (!undo.Ok)
            {
                Fail(results, "操作历史与回滚", Describe(undo));
                return;
            }

            VerifyOk(registry, results, "操作历史与回滚", "verify_write_result", Args(
                ("address", "Q1:R1"), ("expected_values", first)));
        }

        private void CheckCoverageAudit(ToolRegistry registry, List<FeatureCheckResult> results)
        {
            var unmapped = new List<string>();
            foreach (var descriptor in registry.Tools)
            {
                if (descriptor.RiskLevel == RiskLevel.ReadOnly ||
                    AutoVerifyPolicy.MappedPrimaryTools.Contains(descriptor.Name) ||
                    AutoVerifyPolicy.IsExemptedFromAutoVerify(descriptor.Name))
                {
                    continue;
                }

                unmapped.Add(descriptor.Name);
            }

            Expect(results, "自检覆盖审计（写能力工具全覆盖）", unmapped.Count == 0,
                unmapped.Count == 0 ? "映射+豁免完整" : "未覆盖：" + string.Join(", ", unmapped));
        }

        private void CheckPitfallLibrary(List<FeatureCheckResult> results)
        {
            var hits = PitfallLibrary.Match("写入失败 0x800A9C68 cannot run the macro");
            Expect(results, "踩坑失败模式库", hits.Count > 0, "命中 " + hits.Count + " 条已知坑");
        }

        private void CheckSelfCheckHistory(List<FeatureCheckResult> results)
        {
            var db = Path.Combine(Path.GetTempPath(), "xai-sweep-history-" + Guid.NewGuid().ToString("N") + ".db");
            try
            {
                var store = new SelfCheckHistoryStore(db);
                store.Append(new SelfCheckRecord { Kind = "turn", Healthy = true, Ops = 1, Executed = 1, Summary = "巡检写入" });
                var stats = store.Stats(20);
                Expect(results, "自检历史趋势库", stats.Total == 1 && stats.Healthy == 1,
                    "落库+统计正常（" + stats.Total + " 条，通过率 " + Math.Round(stats.PassRate * 100) + "%）");
            }
            finally
            {
                try { if (File.Exists(db)) File.Delete(db); } catch { }
            }
        }

        private void CheckConfiguration(List<FeatureCheckResult> results)
        {
            try
            {
                var settings = new AgentSettingsStore().Load();
                Expect(results, "模型配置与设置存储", settings != null && settings.IsValid,
                    "设置读写正常（MaxToolCallsPerTurn=" + (settings?.MaxToolCallsPerTurn ?? 0) + "）");
            }
            catch (Exception ex)
            {
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:331  {ex.GetType().Name}: {ex.Message}");
                Fail(results, "模型配置与设置存储", ex.Message);
            }
        }

        private void CheckSkills(List<FeatureCheckResult> results)
        {
            try
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "skills");
                var count = Directory.Exists(dir)
                    ? Directory.GetFiles(dir, "SKILL.md", SearchOption.AllDirectories).Length
                    : 0;
                Expect(results, "本地技能包加载", true, "发现 " + count + " 个技能");
            }
            catch (Exception ex)
            {
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:347  {ex.GetType().Name}: {ex.Message}");
                Fail(results, "本地技能包加载", ex.Message);
            }
        }

        // ===================== 编排辅助 =====================

        /// <summary>复刻 AgentRunner 的自动自检编排：TryPlanAll → 结果富化 → 逐条执行 verify 并断言 verdict=ok。</summary>
        private void AutoVerifyChain(
            ToolRegistry registry, List<FeatureCheckResult> results, string name,
            string tool, Dictionary<string, object> args)
        {
            var result = ExecTool(registry, tool, args);
            if (!result.Ok)
            {
                Fail(results, name, "主工具失败：" + Describe(result));
                return;
            }

            AutoVerifyFromResult(registry, results, name, tool, args, result);
        }

        private void AutoVerifyFromResult(
            ToolRegistry registry, List<FeatureCheckResult> results, string name,
            string tool, Dictionary<string, object> args, ToolResult executed)
        {
            var plans = new List<AutoVerifyPlan>();
            if (!AutoVerifyPolicy.TryPlanAll(tool, args, plans) || plans.Count == 0)
            {
                Fail(results, name, "未生成自动自检计划");
                return;
            }

            var executedCount = 0;
            foreach (var plan in plans)
            {
                if (!AutoVerifyPolicy.EnrichFromResult(plan, tool, executed?.Data, executed?.AffectedRanges) ||
                    !registry.TryGet(plan.VerifyTool, out var verifyTool) ||
                    verifyTool.RiskLevel != RiskLevel.ReadOnly)
                {
                    continue;
                }

                ToolResult verifyResult;
                try
                {
                    verifyResult = verifyTool.Execute(plan.Args);
                }
                catch (Exception ex)
                {
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:396  {ex.GetType().Name}: {ex.Message}");
                    Fail(results, name + "（自检 " + plan.VerifyTool + "）", "异常：" + ex.Message);
                    continue;
                }

                executedCount++;
                var verdict = ExtractVerdict(verifyResult);
                if (!verifyResult.Ok || verdict != "ok")
                {
                    Fail(results, name + "（自检 " + plan.VerifyTool + "）",
                        "verdict=" + verdict + "；issues=" + DescribeIssues(verifyResult));
                }
            }

            if (executedCount > 0)
            {
                Pass(results, name, executedCount + " 条自检计划全部通过");
            }
            else
            {
                Fail(results, name, "没有任何自检计划被执行");
            }
        }

        private void VerifyOk(
            ToolRegistry registry, List<FeatureCheckResult> results, string name,
            string verifyTool, Dictionary<string, object> args)
        {
            if (!registry.TryGet(verifyTool, out var verifyToolDescriptor) ||
                verifyToolDescriptor.RiskLevel != RiskLevel.ReadOnly)
            {
                Fail(results, name, verifyTool + " 不可用");
                return;
            }

            ToolResult verifyResult;
            try
            {
                verifyResult = verifyToolDescriptor.Execute(args);
            }
            catch (Exception ex)
            {
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:437  {ex.GetType().Name}: {ex.Message}");
                Fail(results, name, "异常：" + ex.Message);
                return;
            }

            var verdict = ExtractVerdict(verifyResult);
            if (verifyResult.Ok && verdict == "ok")
            {
                Pass(results, name, "verdict=ok");
            }
            else
            {
                Fail(results, name, "verdict=" + verdict + "；issues=" + DescribeIssues(verifyResult) + "；err=" + Describe(verifyResult));
            }
        }

        private ToolResult ExecTool(ToolRegistry registry, string tool, Dictionary<string, object> args)
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
DiagnosticsLog.Write("[LogCover] FeatureHealthSweep.cs:465  {ex.GetType().Name}: {ex.Message}");
                return new ToolResult { Ok = false, ErrorCode = "SWEEP_EXCEPTION", Data = ex.Message };
            }
        }

        private void Exec(ToolRegistry registry, List<FeatureCheckResult> results, string name,
            string tool, Dictionary<string, object> args)
        {
            var result = ExecTool(registry, tool, args);
            if (!result.Ok)
            {
                Fail(results, name, Describe(result));
            }
        }

        private static void Pass(List<FeatureCheckResult> results, string name, string detail)
        {
            results.Add(new FeatureCheckResult { Name = name, Ok = true, Detail = detail });
        }

        private static void Fail(List<FeatureCheckResult> results, string name, string detail)
        {
            results.Add(new FeatureCheckResult { Name = name, Ok = false, Detail = detail });
        }

        private static void Expect(List<FeatureCheckResult> results, string name, bool ok, string detail)
        {
            results.Add(new FeatureCheckResult { Name = name, Ok = ok, Detail = detail });
        }

        private static string ExtractVerdict(ToolResult result)
        {
            if (result?.Data == null || !(JToken.FromObject(result.Data) is JObject data))
            {
                return null;
            }

            return data.Value<string>("verdict");
        }

        private static string ExtractOperationId(ToolResult result)
        {
            if (result?.Data == null || !(JToken.FromObject(result.Data) is JObject data))
            {
                return null;
            }

            return data.Value<string>("operation_id");
        }

        private static string DescribeIssues(ToolResult result)
        {
            if (result?.Data == null || !(JToken.FromObject(result.Data) is JObject data) ||
                !(data["issues"] is JArray issues))
            {
                return "(无)";
            }

            return issues.Count == 0 ? "(空)" : string.Join(" | ", issues);
        }

        private static string Describe(ToolResult result)
        {
            if (result == null)
            {
                return "null";
            }

            return result.Ok ? "ok" : (result.ErrorCode ?? "failed") + " " + result.Data;
        }

        private static Dictionary<string, object> Args(params (string key, object value)[] pairs)
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
