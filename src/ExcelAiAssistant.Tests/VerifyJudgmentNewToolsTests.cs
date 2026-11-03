using System.Collections.Generic;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 新增自检维度的纯静态判据单测（2026-09-08 自检补强批次）：
    /// 清洗新检查项 no_text_matches / all_blank（删除、清空的确定性期望），
    /// 迷你图 / 数据验证 / 合并状态三个新 verify 维度，以及六条新自动映射的策略行为。
    /// </summary>
    public class VerifyJudgmentNewToolsTests
    {
        private static object[,] Matrix(params object[] cells)
        {
            var rows = 1;
            var cols = cells.Length;
            var m = new object[rows, cols];
            for (var c = 0; c < cols; c++)
            {
                m[0, c] = cells[c];
            }

            return m;
        }

        // ===================== no_text_matches / all_blank =====================

        [Fact]
        public void No_text_matches_passes_when_pattern_absent()
        {
            var values = Matrix("甲", "乙", "丙");
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_text_matches", Pattern = "临时" }
            };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 0, "A1:C1");
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void No_text_matches_fails_when_residual_text_found()
        {
            var values = Matrix("甲", "临时01", "丙");
            var checks = new List<ExcelVerifyJudgments.CleanCheck>
            {
                new ExcelVerifyJudgments.CleanCheck { Type = "no_text_matches", Pattern = "临时\\d+" }
            };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, checks, 0, "A1:C1");
            Assert.Contains("text_match", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void No_text_matches_fail_closed_on_missing_or_bad_pattern()
        {
            var values = Matrix("甲", "乙");
            var missing = new List<ExcelVerifyJudgments.CleanCheck> { new ExcelVerifyJudgments.CleanCheck { Type = "no_text_matches" } };
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, missing, 0, "A1:B1");
            Assert.Contains("no_pattern", v1);

            var bad = new List<ExcelVerifyJudgments.CleanCheck> { new ExcelVerifyJudgments.CleanCheck { Type = "no_text_matches", Pattern = "[" } };
            var (v2, i2, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(values, bad, 0, "A1:B1");
            Assert.Contains("bad_pattern", v2);
            Assert.Single(i2);
        }

        [Fact]
        public void All_blank_passes_on_empty_and_fails_on_residual()
        {
            var empty = Matrix(null, "", "  ");
            var pass = new List<ExcelVerifyJudgments.CleanCheck> { new ExcelVerifyJudgments.CleanCheck { Type = "all_blank" } };
            var (v1, _, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(empty, pass, 0, "A1:C1");
            Assert.Equal("ok", v1);

            var dirty = Matrix(null, "残留", null);
            var (v2, i2, _) = ExcelVerifyJudgments.ComputeDataCleanVerdict(dirty, pass, 0, "A1:C1");
            Assert.Contains("not_blank", v2);
            Assert.Single(i2);
        }

        // ===================== 迷你图 =====================

        [Fact]
        public void Sparkline_ok_when_type_and_source_match()
        {
            var groups = new List<ExcelVerifyJudgments.SparklineInfo>
            {
                new ExcelVerifyJudgments.SparklineInfo { Type = 2, Location = "$G$1", SourceData = "Sheet1!$A$1:$E$1" }
            };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeSparklineVerdict(groups, "G1", "column", "A1:E1");
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Sparkline_fails_on_missing_or_type_mismatch()
        {
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeSparklineVerdict(
                new List<ExcelVerifyJudgments.SparklineInfo>(), "G1", "column", null);
            Assert.Contains("missing", v1);

            var groups = new List<ExcelVerifyJudgments.SparklineInfo>
            {
                new ExcelVerifyJudgments.SparklineInfo { Type = 1, Location = "$G$1", SourceData = "A1:E1" }
            };
            var (v2, i2, _) = ExcelVerifyJudgments.ComputeSparklineVerdict(groups, "G1", "column", null);
            Assert.Contains("type", v2);
            Assert.Single(i2);
        }

        [Fact]
        public void Sparkline_fail_closed_without_expected_or_target()
        {
            var groups = new List<ExcelVerifyJudgments.SparklineInfo>
            {
                new ExcelVerifyJudgments.SparklineInfo { Type = 1, Location = "$G$1", SourceData = "A1:E1" }
            };
            var (v1, _, _) = ExcelVerifyJudgments.ComputeSparklineVerdict(groups, "G1", null, null);
            Assert.Contains("no_expected", v1);

            var (v2, _, _) = ExcelVerifyJudgments.ComputeSparklineVerdict(groups, null, "column", null);
            Assert.Contains("no_target", v2);
        }

        // ===================== 数据验证 =====================

        [Fact]
        public void Data_validation_ok_for_list_match()
        {
            var actual = new ExcelVerifyJudgments.DataValidationInfo { Type = 3, Formula1 = "\"甲,乙,丙\"" };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeDataValidationVerdict(actual, "list", "甲,乙,丙", null, null, null);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Data_validation_fails_when_rule_missing_or_type_mismatch()
        {
            var none = new ExcelVerifyJudgments.DataValidationInfo { Type = 0 };
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeDataValidationVerdict(none, "list", null, null, null, null);
            Assert.Contains("missing", v1);
            Assert.Single(i1);

            var interval = new ExcelVerifyJudgments.DataValidationInfo { Type = 1, Operator = 1, Formula1 = "10", Formula2 = "100" };
            var (v2, i2, _) = ExcelVerifyJudgments.ComputeDataValidationVerdict(interval, "list", null, null, null, null);
            Assert.Contains("type", v2);
        }

        [Fact]
        public void Data_validation_checks_operator_and_bounds()
        {
            var actual = new ExcelVerifyJudgments.DataValidationInfo { Type = 1, Operator = 5, Formula1 = "10" };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeDataValidationVerdict(
                actual, "interval", null, "between", "10", "100");
            Assert.Contains("operator", verdict);
            Assert.Equal(2, issues.Count); // operator 不符 + value2 边界缺失不符
        }

        [Fact]
        public void Data_validation_fail_closed_without_expectations()
        {
            var actual = new ExcelVerifyJudgments.DataValidationInfo { Type = 3, Formula1 = "\"甲\"" };
            var (verdict, _, _) = ExcelVerifyJudgments.ComputeDataValidationVerdict(actual, null, null, null, null, null);
            Assert.Contains("no_expected", verdict);
        }

        // ===================== 合并状态 =====================

        [Fact]
        public void Merge_ok_for_merged_and_mixed_when_expected_true()
        {
            var (v1, _, _) = ExcelVerifyJudgments.ComputeMergeVerdict(true, true, "A1:B2");
            Assert.Equal("ok", v1);

            var (v2, i2, _) = ExcelVerifyJudgments.ComputeMergeVerdict(null, true, "A1:B2");
            Assert.Equal("ok", v2);
            Assert.Single(i2); // 混合=通过但给提示
        }

        [Fact]
        public void Merge_fails_on_not_merged_when_expected_true()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeMergeVerdict(false, true, "A1:B2");
            Assert.Contains("not_merged", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Merge_fails_on_residual_merge_when_expected_false()
        {
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeMergeVerdict(true, false, "A1:B2");
            Assert.Contains("still_merged", v1);
            Assert.Single(i1);

            var (v2, _, _) = ExcelVerifyJudgments.ComputeMergeVerdict(null, false, "A1:B2");
            Assert.Contains("still_mixed", v2);
        }

        [Fact]
        public void Merge_fail_closed_without_expected_merged()
        {
            var (verdict, _, _) = ExcelVerifyJudgments.ComputeMergeVerdict(true, null, "A1:B2");
            Assert.Contains("no_expected", verdict);
        }

        // ===================== 策略映射 =====================

        private static List<AutoVerifyPlan> Plan(string tool, Dictionary<string, object> args)
        {
            var plans = new List<AutoVerifyPlan>();
            AutoVerifyPolicy.TryPlanAll(tool, args, plans);
            return plans;
        }

        [Fact]
        public void New_write_tools_produce_expected_verify_plans()
        {
            Assert.Equal("verify_data_clean", Plan("clear_excel_range", Dict("address", "A1:C10"))[0].VerifyTool);
            var delete = Plan("delete_text_matches", Dict("address", "A1:C10", "pattern", "临时\\d+"));
            Assert.Equal("verify_data_clean", delete[0].VerifyTool);

            var validation = Plan("data_validation", Dict("address", "A1:A5", "type", "list", "source", "A1:A3"));
            Assert.Equal("verify_data_validation", validation[0].VerifyTool);

            var merge = Plan("merge_or_split_cells", Dict("action", "merge", "address", "A1:B2"));
            Assert.Equal("verify_merge", merge[0].VerifyTool);
            var unmerge = Plan("merge_or_split_cells", Dict("action", "unmerge", "address", "A1:B2"));
            Assert.Equal("verify_merge", unmerge[0].VerifyTool);

            var dup = Plan("merge_duplicate_cells", Dict("address", "A1:A9"));
            Assert.Equal("verify_merge", dup[0].VerifyTool);

            var spark = Plan("create_excel_sparkline", Dict("data_range", "A1:E1", "dest_range", "G1", "type", "column"));
            Assert.Equal("verify_sparkline", spark[0].VerifyTool);
        }

        [Fact]
        public void Oversized_or_unbounded_targets_skip_auto_verify()
        {
            // 整列引用不可回读 → 不出计划（fail-safe，不做无依据断言）
            Assert.Empty(Plan("clear_excel_range", Dict("address", "A:C")));
            // 1,048,576 行 × 3 列远超上限 → 不出计划
            Assert.Empty(Plan("clear_excel_range", Dict("address", "A1:C1048576")));
            // 缺 pattern → 不出计划
            Assert.Empty(Plan("delete_text_matches", Dict("address", "A1:B2")));
        }

        [Fact]
        public void Delete_text_matches_verifies_target_address_when_present()
        {
            var plans = Plan("delete_text_matches", Dict("address", "A1:B2", "target_address", "D1:E2", "pattern", "x"));
            Assert.Equal("D1:E2", plans[0].Args["address"]);
        }

        private static Dictionary<string, object> Dict(params object[] keyValues)
        {
            var map = new Dictionary<string, object>();
            for (var i = 0; i < keyValues.Length; i += 2)
            {
                map[(string)keyValues[i]] = keyValues[i + 1];
            }

            return map;
        }

        // ===================== verify_table =====================

        private static ExcelVerifyJudgments.TableInfo Table(string name, string address, int rows, int cols, params string[] headers)
        {
            return new ExcelVerifyJudgments.TableInfo
            {
                Name = name,
                Address = address,
                Rows = rows,
                Columns = cols,
                Headers = headers
            };
        }

        [Fact]
        public void Table_verdict_fails_when_no_tables()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, "A1:C4", null, null, null,
                new List<ExcelVerifyJudgments.TableInfo>());
            Assert.Contains("no_tables", verdict);
            Assert.Contains(issues, i => i.Contains("没有任何智能表"));
        }

        [Fact]
        public void Table_verdict_fails_when_named_table_missing()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo> { Table("表1", "$A$1:$C$4", 4, 3, "姓名", "金额", "日期") };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict("不存在", null, null, null, null, tables);
            Assert.Contains("no_target", verdict);
            Assert.Contains(issues, i => i.Contains("未找到名为"));
        }

        [Fact]
        public void Table_verdict_passes_when_headers_and_dimensions_match()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo> { Table("tbl_x", "$A$1:$C$4", 4, 3, "姓名", "金额", "日期") };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(
                "tbl_x", null, new[] { "姓名", "金额", "日期" }, 4, 3, tables);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Table_verdict_fails_on_header_mismatch_and_dimension_drift()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo> { Table("tbl_x", "$A$1:$C$5", 5, 3, "姓名", "数量", "日期") };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(
                "tbl_x", null, new[] { "姓名", "金额", "日期" }, 4, 3, tables);
            Assert.Contains("headers", verdict);
            Assert.Contains(issues, i => i.Contains("表头第 2 列"));
            Assert.Contains(issues, i => i.Contains("行数"));
        }

        [Fact]
        public void Table_verdict_locates_by_address_without_name()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo>
            {
                Table("A", "$A$1:$C$4", 4, 3),
                Table("B", "$E$1:$G$4", 4, 3)
            };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, "E1:G4", null, null, null, tables);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Table_verdict_matches_subrange_inside_table()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo> { Table("A", "$A$1:$C$4", 4, 3) };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, "B2:C4", null, null, null, tables);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Table_verdict_rejects_outside_address_without_substring_false_positive()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo> { Table("A", "$A$1:$C$4", 4, 3) };
            // 完全在表外的地址不应命中
            var (verdictOutside, _, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, "F10:H12", null, null, null, tables);
            Assert.Contains("no_target", verdictOutside);
            // 反例锁：旧子串实现会把 "A1" 误配到任何含 "A1" 的地址；坐标包含语义下 "Z99" 不应命中
            var (verdictFar, _, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, "Z99", null, null, null, tables);
            Assert.Contains("no_target", verdictFar);
        }

        [Fact]
        public void Table_verdict_fail_closed_when_unlocatable_and_ambiguous()
        {
            var tables = new List<ExcelVerifyJudgments.TableInfo>
            {
                Table("A", "$A$1:$C$4", 4, 3),
                Table("B", "$E$1:$G$4", 4, 3)
            };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeTableVerdict(null, null, null, null, null, tables);
            Assert.Contains("no_target", verdict);
            Assert.Contains(issues, i => i.Contains("无法唯一定位"));
        }

        // ===================== verify_protection =====================

        private static ExcelVerifyJudgments.ProtectionInfo Protection(bool structure, bool sheet, bool? locked = null)
        {
            return new ExcelVerifyJudgments.ProtectionInfo
            {
                ProtectStructure = structure,
                SheetProtected = sheet,
                CellLocked = locked,
                CellAddress = locked.HasValue ? "A1" : null
            };
        }

        [Fact]
        public void Protection_verdict_fails_when_writable_expected_but_locked_on_protected_sheet()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeProtectionVerdict(
                Protection(false, true, true), null, true);
            Assert.Contains("write_blocked", verdict);
            Assert.Contains(issues, i => i.Contains("写入将被拒绝"));
        }

        [Fact]
        public void Protection_verdict_passes_when_cell_unlocked_on_protected_sheet()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeProtectionVerdict(
                Protection(false, true, false), null, true);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Protection_verdict_fails_when_protected_state_differs_from_expectation()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeProtectionVerdict(
                Protection(false, false, null), true, null);
            Assert.Contains("protect_state", verdict);
            Assert.Contains(issues, i => i.Contains("实际未保护"));
        }

        [Fact]
        public void Protection_verdict_is_informational_without_expectations()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeProtectionVerdict(
                Protection(true, true, true), null, null);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        // ===================== verify_image =====================

        private static ExcelVerifyJudgments.ImageInfo Pic(double left, double top, double width = 40, double height = 30, string name = "图片 1")
        {
            return new ExcelVerifyJudgments.ImageInfo { Name = name, Left = left, Top = top, Width = width, Height = height };
        }

        private static ExcelVerifyJudgments.CellBounds Cell(double left, double top, double right, double bottom)
        {
            return new ExcelVerifyJudgments.CellBounds { Left = left, Top = top, Right = right, Bottom = bottom };
        }

        [Fact]
        public void Image_ok_when_anchor_hit_and_size_matches()
        {
            var pics = new List<ExcelVerifyJudgments.ImageInfo> { Pic(310, 12) };
            var bounds = Cell(310, 12, 380, 32);
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeImageVerdict(1, "F2", bounds, 40, 30, pics);
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Image_fails_when_no_picture_within_anchor()
        {
            var pics = new List<ExcelVerifyJudgments.ImageInfo> { Pic(100, 12) };
            var bounds = Cell(310, 12, 380, 32);
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeImageVerdict(null, "F2", bounds, null, null, pics);
            Assert.Contains("no_target", verdict);
            Assert.Contains(issues, i => i.Contains("F2"));
        }

        [Fact]
        public void Image_fails_on_count_mismatch_and_size_mismatch()
        {
            var pics = new List<ExcelVerifyJudgments.ImageInfo> { Pic(310, 12), Pic(400, 50) };
            var bounds = Cell(310, 12, 380, 32);
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeImageVerdict(1, "F2", bounds, null, null, pics);
            Assert.Contains("count_mismatch", v1);
            Assert.Contains(i1, i => i.Contains("图片数量"));

            var (v2, i2, _) = ExcelVerifyJudgments.ComputeImageVerdict(2, "F2", bounds, 60, null, pics);
            Assert.Contains("size", v2);
            Assert.Contains(i2, i => i.Contains("宽度"));
        }

        [Fact]
        public void Image_fails_closed_when_size_expected_without_anchor()
        {
            var pics = new List<ExcelVerifyJudgments.ImageInfo> { Pic(310, 12) };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeImageVerdict(null, null, null, 40, null, pics);
            Assert.Contains("size_without_anchor", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Image_fails_closed_when_no_pictures_and_no_expectations()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeImageVerdict(null, null, null, null, null, new List<ExcelVerifyJudgments.ImageInfo>());
            Assert.Contains("no_pictures", verdict);
            Assert.Single(issues);
        }

        [Fact]
        public void Image_fails_when_anchor_unresolvable()
        {
            var pics = new List<ExcelVerifyJudgments.ImageInfo> { Pic(310, 12) };
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeImageVerdict(null, "F2", null, null, null, pics);
            Assert.Contains("invalid_anchor", verdict);
            Assert.Single(issues);
        }

        // ===================== verify_freeze_panes =====================

        private static ExcelVerifyJudgments.FreezePaneInfo Freeze(bool frozen, int row, int col)
        {
            return new ExcelVerifyJudgments.FreezePaneInfo { Frozen = frozen, SplitRow = row, SplitColumn = col };
        }

        [Fact]
        public void Freeze_ok_when_state_and_position_match()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(true, "B2", Freeze(true, 1, 1));
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }

        [Fact]
        public void Freeze_fails_when_position_differs()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(true, "C3", Freeze(true, 1, 1));
            Assert.Contains("freeze_position", verdict);
            Assert.Contains(issues, i => i.Contains("2 行 2 列"));
        }

        [Fact]
        public void Freeze_fails_when_expected_frozen_but_not_frozen()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(true, null, Freeze(false, 0, 0));
            Assert.Contains("freeze_state", verdict);
            Assert.Single(issues);

            var (v2, i2, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(false, null, Freeze(true, 1, 1));
            Assert.Contains("freeze_state", v2);
            Assert.Single(i2);
        }

        [Fact]
        public void Freeze_fails_when_cell_expected_but_not_frozen()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(null, "B2", Freeze(false, 0, 0));
            Assert.Contains("freeze_state", verdict);
            Assert.Contains(issues, i => i.Contains("实际未冻结"));
        }

        [Fact]
        public void Freeze_fails_closed_on_invalid_or_conflicting_expectations()
        {
            var (v1, i1, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(true, "A1", Freeze(true, 0, 0));
            Assert.Contains("invalid_cell", v1);
            Assert.Single(i1);

            var (v2, i2, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(true, "不是地址", Freeze(true, 1, 1));
            Assert.Contains("invalid_cell", v2);
            Assert.Single(i2);

            var (v3, i3, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(false, "B2", Freeze(false, 0, 0));
            Assert.Contains("invalid_expectation", v3);
            Assert.Single(i3);
        }

        [Fact]
        public void Freeze_is_informational_without_expectations()
        {
            var (verdict, issues, _) = ExcelVerifyJudgments.ComputeFreezeVerdict(null, null, Freeze(true, 1, 1));
            Assert.Equal("ok", verdict);
            Assert.Empty(issues);
        }
    }
}
