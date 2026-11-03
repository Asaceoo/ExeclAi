using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// M27 审查 A6 回归锁：每个 mapped 工具产出的自检 plan，其参数键必须是对应 verify 工具
    /// 真实 schema 接受的参数（否则运行期 Required 抛错 → fail-closed 误拦合法写入）。
    /// </summary>
    public class AutoVerifyPlanContractTests
    {
        // verify 工具接受的参数白名单（与 ExcelVerifyService 的 Required/Optional 读取键一致）
        private static readonly Dictionary<string, HashSet<string>> VerifyParams = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal)
        {
            ["verify_write_result"] = new HashSet<string>(new[] { "sheet", "address", "expected_values", "expected_count", "tolerance", "type_aware" }, StringComparer.Ordinal),
            ["verify_formula"] = new HashSet<string>(new[] { "sheet", "address", "expected_formula", "expected_values", "recalc" }, StringComparer.Ordinal),
            ["verify_structure"] = new HashSet<string>(new[] { "sheet", "address", "expected_sheet_exists", "expected_sheet_absent", "expected_row_count", "expected_column_count" }, StringComparer.Ordinal),
        };

        private static ExcelToolHost NewHost()
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var app = Activator.CreateInstance(excelType);
            try { app.GetType().InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false }); } catch { }
            try { app.GetType().InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false }); } catch { }
            app.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, app, null)
               .GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null,
                   app.GetType().InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, app, null), null);
            return new ExcelToolHost(app, null, null, "contract-test");
        }

        [Fact]
        public void Set_element_plan_matches_verify_write_result_contract()
        {
            var map = new Dictionary<string, object> { ["address"] = "B7", ["value"] = "契约", ["sheet"] = "Sheet1" };
            Assert.True(AutoVerifyPolicy.TryPlan("set_element", map, out var plan), "set_element 应产出自检计划");
            Assert.True(VerifyParams.ContainsKey(plan.VerifyTool), "verify 工具需在本测试白名单内：" + plan.VerifyTool);
            var unknown = plan.Args.Keys.Where(k => !VerifyParams[plan.VerifyTool].Contains(k)).ToList();
            Assert.True(unknown.Count == 0, plan.VerifyTool + " 收到未知参数（运行期将被忽略或 Required 失败）：" + string.Join(",", unknown));
            Assert.True(plan.Args.ContainsKey("address"), "verify_write_result 必填 address");
        }

        [Fact]
        public void Import_csv_plan_matches_verify_structure_contract()
        {
            var map = new Dictionary<string, object> { ["file_path"] = "D:\\x.csv", ["sheet_name"] = "CSV_x" };
            Assert.True(AutoVerifyPolicy.TryPlan("import_csv", map, out var plan));
            Assert.Equal("verify_structure", plan.VerifyTool);
            var unknown = plan.Args.Keys.Where(k => !VerifyParams[plan.VerifyTool].Contains(k)).ToList();
            Assert.True(unknown.Count == 0, "未知参数：" + string.Join(",", unknown));
            Assert.True(plan.Args.ContainsKey("expected_sheet_exists"), "verify_structure 需要 expected_sheet_exists 才能核对导入结果");
        }

        [Fact]
        public void All_mapped_tools_plans_use_known_verify_params()
        {
            // 全量扫描：对每个 mapped 工具用最小可用参数试 plan，校验参数键合法性
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return; // 无 Excel 环境跳过（CI 兼容）
            var host = NewHost();
            if (host == null) return;
            try
            {
                var registry = host.CreateRegistry();
                var offenders = new List<string>();
                foreach (var toolName in AutoVerifyPolicy.MappedPrimaryTools)
                {
                    var descriptor = registry.Tools.FirstOrDefault(t => t.Name == toolName);
                    if (descriptor == null) continue;
                    var minimal = MinimalArgsFor(toolName);
                    if (minimal == null) continue;
                    if (!AutoVerifyPolicy.TryPlan(toolName, minimal, out var plan)) continue;
                    if (!VerifyParams.ContainsKey(plan.VerifyTool)) continue; // 其它 verify 族本测试不覆盖
                    var bad = plan.Args.Keys.Where(k => !VerifyParams[plan.VerifyTool].Contains(k)).ToList();
                    if (bad.Count > 0) offenders.Add(toolName + "→" + plan.VerifyTool + ": " + string.Join(",", bad));
                }
                Assert.True(offenders.Count == 0, "以下工具的自检参数不在 verify 契约内：" + string.Join(" | ", offenders));
            }
            finally
            {
                try
                {
                    var workbooks = host.GetType(); // no-op
                }
                catch { }
            }
        }

        private static Dictionary<string, object> MinimalArgsFor(string tool)
        {
            switch (tool)
            {
                case "write_excel_cell":
                case "set_element": return new Dictionary<string, object> { ["address"] = "A1", ["value"] = 1 };
                case "import_csv": return new Dictionary<string, object> { ["file_path"] = "D:\\t.csv", ["sheet_name"] = "CSV_t" };
                case "clear_excel_range": return new Dictionary<string, object> { ["address"] = "A1:B2" };
                default: return null;
            }
        }
    }
}
