using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 「JYYJ快捷」规划层的纯逻辑单测（无 COM、无 UI）。
    /// 覆盖：安全分级判定（破坏性/覆盖写回/安全写回）、二次确认判定、参数框字段模型、
    /// 工具实参组装与校验、去重删行数预算、文本拆分开放列预算。
    /// </summary>
    public sealed class RibbonCommandPlannerTests
    {
        private static ParamSelection Selection(int cols = 3, params string[] columnNames)
        {
            var sel = new ParamSelection { SheetName = "Sheet1", Address = "$A$1:$C$4", Rows = 4, Cols = cols };
            for (var c = 0; c < cols; c++)
            {
                sel.ColumnNames.Add(c < columnNames.Length ? columnNames[c] : $"列{c + 1}");
            }
            return sel;
        }

        // ================= 安全分级判定 =================

        [Theory]
        [InlineData("dedup_range", CommandSafety.DeleteRows)]
        [InlineData("sort_range", CommandSafety.Overwrite)]
        [InlineData("filter_excel_data", CommandSafety.Overwrite)]
        [InlineData("transpose_range", CommandSafety.Overwrite)]
        [InlineData("clean_excel_data", CommandSafety.SafeWrite)]
        [InlineData("create_excel_pivot_summary", CommandSafety.SafeWrite)]
        [InlineData("text_split", CommandSafety.SafeWrite)]
        [InlineData("text_merge", CommandSafety.SafeWrite)]
        [InlineData("text_clean", CommandSafety.SafeWrite)]
        [InlineData("fill_formula", CommandSafety.SafeWrite)]
        public void Classify_maps_each_tool_to_locked_safety(string toolName, CommandSafety expected)
        {
            Assert.Equal(expected, RibbonCommandPlanner.Classify(toolName));
        }

        [Fact]
        public void Classify_throws_on_unknown_tool()
        {
            Assert.Throws<ArgumentException>(() => RibbonCommandPlanner.Classify("not_a_tool"));
        }

        // ================= 二次确认判定 =================

        [Fact]
        public void RequiresSecondaryConfirm_only_for_delete_rows()
        {
            Assert.True(RibbonCommandPlanner.RequiresSecondaryConfirm(CommandSafety.DeleteRows));
            Assert.False(RibbonCommandPlanner.RequiresSecondaryConfirm(CommandSafety.Overwrite));
            Assert.False(RibbonCommandPlanner.RequiresSecondaryConfirm(CommandSafety.SafeWrite));
        }

        [Fact]
        public void OverlayHint_present_only_for_overwrite()
        {
            var sel = Selection(2);
            Assert.Equal("将覆盖原区域（约 8 个单元格）", RibbonCommandPlanner.OverlayHint(CommandSafety.Overwrite, sel));
            Assert.Null(RibbonCommandPlanner.OverlayHint(CommandSafety.SafeWrite, sel));
            Assert.Null(RibbonCommandPlanner.OverlayHint(CommandSafety.DeleteRows, sel));
        }

        [Fact]
        public void DeleteWarning_injects_row_count()
        {
            Assert.Equal("将删除 3 行，确定？", RibbonCommandPlanner.DeleteWarning(3));
        }

        // ================= 参数框字段模型 =================

        [Fact]
        public void BuildFields_column_options_use_selected_header_names()
        {
            var sel = Selection(3, "姓名", "年龄", "城市");
            var fields = RibbonCommandPlanner.BuildFields("sort_range", sel);

            // select 下拉应注入真实列名（含序号前缀）
            var colField = fields.Single(f => f.Name == "sort_column");
            Assert.Equal("select", colField.Type);
            Assert.Equal(new[] { "1", "2", "3" }, colField.Options.Select(o => o.Value).ToArray());
            Assert.Equal(new[] { "1·姓名", "2·年龄", "3·城市" }, colField.Options.Select(o => o.Display).ToArray());
        }

        [Fact]
        public void BuildFields_clean_exposes_toggles_and_date_column_and_fill_value()
        {
            var fields = RibbonCommandPlanner.BuildFields("clean_excel_data", Selection(3, "a", "b", "c"));
            Assert.Equal(9, fields.Count);
            Assert.Contains(fields, f => f.Name == "remove_blank_rows" && f.Type == "checkbox");
            Assert.Contains(fields, f => f.Name == "fix_dates" && f.Type == "checkbox");
            Assert.Contains(fields, f => f.Name == "fill_blanks" && f.Type == "checkbox");
            Assert.Contains(fields, f => f.Name == "trim" && f.Type == "checkbox");
            Assert.Contains(fields, f => f.Name == "remove_non_printable" && f.Type == "checkbox");
            var dateCol = fields.Single(f => f.Name == "date_column");
            Assert.Equal("select", dateCol.Type);
            Assert.Equal("1", dateCol.Default);
            var dateFormat = fields.Single(f => f.Name == "date_format");
            Assert.Equal("text", dateFormat.Type);
            var fillColumn = fields.Single(f => f.Name == "fill_column");
            Assert.Equal("select", fillColumn.Type);
            Assert.Equal("0", fillColumn.Default);
            Assert.Equal("全部列", fillColumn.Options[0].Display);
            var fillValue = fields.Single(f => f.Name == "fill_value");
            Assert.Equal("text", fillValue.Type);
            Assert.Equal("待补", fillValue.Default);
        }

        [Fact]
        public void BuildFields_text_merge_offers_new_column_option()
        {
            var fields = RibbonCommandPlanner.BuildFields("text_merge", Selection(2, "A", "B"));
            var target = fields.Single(f => f.Name == "target_column");
            Assert.Equal("select", target.Type);
            Assert.Equal((2 + 1).ToString(), target.Default);
            Assert.Contains(target.Options, o => o.Value == "3" && o.Display == "选区右侧新列");
        }

        [Fact]
        public void BuildFields_all_tools_expose_hints_on_every_field()
        {
            // 「加入说明」：每个字段都要有 Hint，告诉用户这个配置是做什么的。
            var tools = new[]
            {
                "clean_excel_data", "dedup_range", "sort_range", "transpose_range", "filter_excel_data",
                "create_excel_pivot_summary", "text_split", "text_merge", "text_clean", "fill_formula",
                "insert_text_into_cell", "delete_text_matches", "fill_random_numbers", "fill_number_sequence",
                "insert_blank_rows_every_n", "delete_blank_rows_range", "fill_blank_cells_down", "merge_duplicate_cells"
            };
            foreach (var tool in tools)
            {
                var fields = RibbonCommandPlanner.BuildFields(tool, Selection(3, "a", "b", "c"));
                Assert.NotEmpty(fields);
                foreach (var f in fields)
                {
                    Assert.False(string.IsNullOrWhiteSpace(f.Hint), $"{tool}.{f.Name} 缺少 Hint 说明");
                }
            }
        }

        [Fact]
        public void BuildFields_dedup_exposes_case_sensitive_and_sort_offers_secondary_column()
        {
            var dedup = RibbonCommandPlanner.BuildFields("dedup_range", Selection(3));
            var cs = dedup.Single(f => f.Name == "case_sensitive");
            Assert.Equal("checkbox", cs.Type);
            Assert.Equal("false", cs.Default);

            var sort = RibbonCommandPlanner.BuildFields("sort_range", Selection(3, "A", "B", "C"));
            var c2 = sort.Single(f => f.Name == "sort_column2");
            Assert.Equal("select", c2.Type);
            Assert.Equal("0", c2.Default);
            Assert.Equal("无", c2.Options[0].Display);
            Assert.Equal(new[] { "1", "2", "3" }, c2.Options.Skip(1).Select(o => o.Value).ToArray());
            var d2 = sort.Single(f => f.Name == "direction2");
            Assert.Equal("asc", d2.Default);
        }

        // ================= 工具实参组装与校验 =================

        [Fact]
        public void BuildArgs_clean_assembles_operations_array_with_date_column_and_fill_value()
        {
            var args = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(3),
                new Dictionary<string, object>
                {
                    ["remove_blank_rows"] = "true",
                    ["fix_dates"] = "true",
                    ["date_column"] = "2",
                    ["fill_blanks"] = "true",
                    ["fill_value"] = "待补"
                }) as Dictionary<string, object>;

            Assert.NotNull(args);
            var ops = args["operations"] as List<Dictionary<string, object>>;
            Assert.Equal(3, ops.Count);
            Assert.Equal("remove_blank_rows", ops[0]["tool"]);
            Assert.Equal("fix_dates", ops[1]["tool"]);
            Assert.Equal(2, Convert.ToInt32(ops[1]["column"]));
            Assert.Equal("fill_blanks", ops[2]["tool"]);
            Assert.Equal("待补", ops[2]["value"]);
            Assert.Equal(Selection(3).SheetName, args["sheet"]);
        }

        [Fact]
        public void BuildArgs_clean_maps_trim_and_remove_non_printable_ops()
        {
            var args = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(2),
                new Dictionary<string, object>
                {
                    ["remove_blank_rows"] = "false",
                    ["fix_dates"] = "false",
                    ["fill_blanks"] = "false",
                    ["trim"] = "true",
                    ["remove_non_printable"] = "true"
                }) as Dictionary<string, object>;

            Assert.NotNull(args);
            var ops = args["operations"] as List<Dictionary<string, object>>;
            Assert.Equal(2, ops.Count);
            Assert.Equal("trim", ops[0]["tool"]);
            Assert.Equal("remove_non_printable", ops[1]["tool"]);
        }

        [Fact]
        public void BuildArgs_clean_rejects_invalid_date_column()
        {
            var build = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(2),
                new Dictionary<string, object>
                {
                    ["remove_blank_rows"] = "false",
                    ["fix_dates"] = "true",
                    ["date_column"] = "9",
                    ["fill_blanks"] = "false"
                });
            Assert.IsType<string>(build);
        }

        [Fact]
        public void BuildArgs_clean_rejects_empty_operation_set()
        {
            var build = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(2),
                new Dictionary<string, object> { ["remove_blank_rows"] = "false", ["fix_dates"] = "false", ["fill_blanks"] = "false" });

            Assert.IsType<string>(build);
        }

        [Fact]
        public void BuildArgs_clean_maps_date_format_and_fill_column_when_provided()
        {
            var args = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(3),
                new Dictionary<string, object>
                {
                    ["remove_blank_rows"] = "false",
                    ["fix_dates"] = "true",
                    ["date_column"] = "2",
                    ["date_format"] = "yyyy-MM-dd",
                    ["fill_blanks"] = "true",
                    ["fill_column"] = "3",
                    ["fill_value"] = "0"
                }) as Dictionary<string, object>;

            Assert.NotNull(args);
            var ops = args["operations"] as List<Dictionary<string, object>>;
            Assert.Equal("yyyy-MM-dd", ops[0]["date_format"]);
            Assert.Equal(3, Convert.ToInt32(ops[1]["column"]));
        }

        [Fact]
        public void BuildArgs_clean_omits_fill_column_zero_to_avoid_bounds_error()
        {
            // fill_column=0（全部列）不得注入 column：fill_blanks 对 column=0 会得 start=-1 越界
            var args = RibbonCommandPlanner.BuildArgs("clean_excel_data", Selection(3),
                new Dictionary<string, object>
                {
                    ["remove_blank_rows"] = "false", ["fix_dates"] = "false",
                    ["fill_blanks"] = "true", ["fill_column"] = "0", ["fill_value"] = "x"
                }) as Dictionary<string, object>;

            Assert.NotNull(args);
            var ops = args["operations"] as List<Dictionary<string, object>>;
            Assert.False(ops[0].ContainsKey("column"));
        }

        [Fact]
        public void BuildArgs_dedup_whole_row_leaves_columns_absent()
        {
            var args = RibbonCommandPlanner.BuildArgs("dedup_range", Selection(3),
                new Dictionary<string, object> { ["scope"] = "whole", ["keep"] = "first", ["has_header"] = "true" }) as Dictionary<string, object>;

            Assert.NotNull(args);
            Assert.False(args.ContainsKey("columns"));
            Assert.Equal("first", args["keep"]);
            Assert.Equal(1, Convert.ToInt32(args["header_rows"]));
        }

        [Fact]
        public void BuildArgs_dedup_columns_converts_and_validates()
        {
            var build = RibbonCommandPlanner.BuildArgs("dedup_range", Selection(3),
                new Dictionary<string, object> { ["scope"] = "columns", ["columns"] = new List<string> { "1", "3" }, ["has_header"] = "false" }) as Dictionary<string, object>;

            Assert.NotNull(build);
            Assert.Equal(new[] { 1, 3 }, (build["columns"] as List<int>).ToArray());
            Assert.Equal(0, Convert.ToInt32(build["header_rows"]));
        }

        [Fact]
        public void BuildArgs_dedup_rejects_out_of_range_columns()
        {
            var build = RibbonCommandPlanner.BuildArgs("dedup_range", Selection(2),
                new Dictionary<string, object> { ["scope"] = "columns", ["columns"] = new List<string> { "5" }, ["has_header"] = "true" });

            Assert.IsType<string>(build);
        }

        [Fact]
        public void BuildArgs_dedup_passes_case_sensitive()
        {
            var on = RibbonCommandPlanner.BuildArgs("dedup_range", Selection(3),
                new Dictionary<string, object> { ["scope"] = "whole", ["keep"] = "first", ["has_header"] = "true", ["case_sensitive"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(on);
            Assert.True((bool)on["case_sensitive"]);

            var def = RibbonCommandPlanner.BuildArgs("dedup_range", Selection(3),
                new Dictionary<string, object> { ["scope"] = "whole", ["keep"] = "first", ["has_header"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(def);
            Assert.False((bool)def["case_sensitive"]);
        }

        [Fact]
        public void BuildArgs_sort_validates_column_bounds()
        {
            var ok = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "3", ["direction"] = "desc", ["has_header"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(3, Convert.ToInt32(ok["sort_column"]));
            Assert.Equal("desc", ok["direction"]);

            var bad = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "9", ["direction"] = "asc", ["has_header"] = "true" });
            Assert.IsType<string>(bad);
        }

        [Fact]
        public void BuildArgs_sort_passes_secondary_column_only_when_selected()
        {
            var with2 = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "1", ["direction"] = "asc", ["sort_column2"] = "2", ["direction2"] = "desc", ["has_header"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(with2);
            Assert.Equal(2, Convert.ToInt32(with2["sort_column2"]));
            Assert.Equal("desc", with2["direction2"]);

            var without = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "1", ["direction"] = "asc", ["sort_column2"] = "0", ["has_header"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("sort_column2"));

            var bad = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "1", ["sort_column2"] = "9", ["has_header"] = "true" });
            Assert.IsType<string>(bad);
        }

        [Fact]
        public void BuildArgs_filter_requires_value_when_operator_needs_it()
        {
            var bad = RibbonCommandPlanner.BuildArgs("filter_excel_data", Selection(3),
                new Dictionary<string, object> { ["column"] = "2", ["operator"] = "equals", ["value"] = "" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("filter_excel_data", Selection(3),
                new Dictionary<string, object> { ["column"] = "2", ["operator"] = "not_empty", ["value"] = "" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(2, Convert.ToInt32(ok["column"]));
            Assert.Equal("not_empty", ok["operator"]);
        }

        [Fact]
        public void BuildArgs_filter_passes_target_address_when_filled()
        {
            var with = RibbonCommandPlanner.BuildArgs("filter_excel_data", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["operator"] = "not_empty", ["value"] = "", ["target_address"] = "Z9" }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal("Z9", with["target_address"]);

            var without = RibbonCommandPlanner.BuildArgs("filter_excel_data", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["operator"] = "not_empty", ["value"] = "" }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("target_address"));
        }

        [Fact]
        public void BuildArgs_transpose_injects_target_only_for_custom_mode()
        {
            var topLeft = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object> { ["target_mode"] = "top_left", ["target_address"] = "X1" }) as Dictionary<string, object>;
            Assert.NotNull(topLeft);
            Assert.False(topLeft.ContainsKey("target_address"));

            var custom = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object> { ["target_mode"] = "custom", ["target_address"] = "$F$1" }) as Dictionary<string, object>;
            Assert.NotNull(custom);
            Assert.Equal("$F$1", custom["target_address"]);

            var bad = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object> { ["target_mode"] = "custom", ["target_address"] = "" });
            Assert.IsType<string>(bad);
        }

        [Fact]
        public void BuildArgs_pivot_validates_columns_and_uses_auto_value()
        {
            var args = RibbonCommandPlanner.BuildArgs("create_excel_pivot_summary", Selection(4, "类", "a", "b", "c"),
                new Dictionary<string, object> { ["label_column"] = "1", ["value_column"] = "0" }) as Dictionary<string, object>;
            Assert.NotNull(args);
            Assert.Equal(1, Convert.ToInt32(args["label_column"]));
            Assert.Equal(0, Convert.ToInt32(args["value_column"]));
            Assert.Equal("$A$1:$C$4", args["source_address"]);
        }

        [Fact]
        public void BuildArgs_text_split_requires_delimiter()
        {
            var bad = RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = "", ["keep_source"] = "true" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "2", ["delimiter"] = ",", ["keep_source"] = "false" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(2, Convert.ToInt32(ok["column"]));
            Assert.False((bool)ok["keep_source"]);
        }

        [Fact]
        public void BuildArgs_text_merge_validates_columns_and_target()
        {
            var bad = RibbonCommandPlanner.BuildArgs("text_merge", Selection(3),
                new Dictionary<string, object> { ["columns"] = new List<string>(), ["target_column"] = "4", ["delimiter"] = " ", ["skip_empty"] = "true" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("text_merge", Selection(3, "A", "B", "C"),
                new Dictionary<string, object> { ["columns"] = new List<string> { "1", "2" }, ["target_column"] = "4", ["delimiter"] = "-", ["skip_empty"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(new[] { 1, 2 }, (ok["columns"] as List<int>).ToArray());
            Assert.Equal(4, Convert.ToInt32(ok["target_column"]));
        }

        [Fact]
        public void BuildArgs_text_clean_maps_toggles_and_case()
        {
            var args = RibbonCommandPlanner.BuildArgs("text_clean", Selection(2),
                new Dictionary<string, object> { ["trim"] = "true", ["collapse"] = "false", ["remove_non_printable"] = "true", ["case"] = "upper", ["remove_chars"] = "-" }) as Dictionary<string, object>;
            Assert.NotNull(args);
            Assert.True((bool)args["trim"]);
            Assert.False((bool)args["collapse_whitespace"]);
            Assert.True((bool)args["remove_non_printable"]);
            Assert.Equal("upper", args["case"]);
            Assert.Equal("-", args["remove_chars"]);
        }

        [Fact]
        public void BuildArgs_fill_formula_requires_and_parses_start_row()
        {
            var bad = RibbonCommandPlanner.BuildArgs("fill_formula", Selection(2),
                new Dictionary<string, object> { ["formula"] = "", ["start_row"] = "" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("fill_formula", Selection(2),
                new Dictionary<string, object> { ["formula"] = "=A{row}*B{row}", ["start_row"] = "3" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal("=A{row}*B{row}", ok["formula"]);
            Assert.Equal(3, Convert.ToInt32(ok["start_row"]));
        }

        // ================= 去重删行数预算（与工具 DedupKey 语义对齐） =================

        [Fact]
        public void CountRemovedRows_counts_whole_row_duplicates_excluding_header()
        {
            var matrix = new object[,]
            {
                { "A", "B" },
                { 1D, 2D },
                { 1D, 2D }, // 重复
                { 3D, 4D }
            };

            Assert.Equal(1, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new int[0], false, false));
        }

        [Fact]
        public void CountRemovedRows_respects_columns_and_keep_last()
        {
            var matrix = new object[,]
            {
                { "H1", "H2" },
                { 1D, "x" },
                { 1D, "y" }, // 第 1 列相同 → 按第 1 列判重，keep first 时删此行
                { 2D, "z" }
            };

            Assert.Equal(1, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new[] { 1 }, false, false));
            // keep last 与 keep first 都要删一行（首/末保留不同）
            Assert.Equal(1, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new[] { 1 }, true, false));
        }

        [Fact]
        public void CountRemovedRows_ignores_case_when_case_sensitive_false()
        {
            var matrix = new object[,]
            {
                { "A", "B" },
                { "foo", 1D },
                { "FOO", 1D } // 整行仅大小写不同 → 不区分大小写时视为重复
            };

            Assert.Equal(1, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new int[0], false, false));
            Assert.Equal(0, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new int[0], false, true));
        }

        [Fact]
        public void CountRemovedRows_empty_body_returns_zero()
        {
            var matrix = new object[,] { { "H", "H2" } };
            Assert.Equal(0, RibbonCommandPlanner.CountRemovedRows(matrix, 1, new int[0], false, false));
        }

        // ================= 文本拆分开放列预算（预留空列守卫） =================

        [Fact]
        public void SplitTargetColumns_computes_full_band_from_max_parts()
        {
            var matrix = new object[,]
            {
                { "name", "note" },
                { "a,b,c", "x" },
                { "d,e", "y" },
                { "f", "z" }
            };

            // 源列 1，分隔英文逗号，最大 3 段 → 写第 2..4 列
            var band = RibbonCommandPlanner.SplitTargetColumns(matrix, 1, ",");
            Assert.Equal(new[] { 2, 3, 4 }, band.ToArray());
        }

        [Fact]
        public void SplitTargetColumns_guards_invalid_source_column()
        {
            var matrix = new object[,] { { "a", "b" } };
            Assert.Empty(RibbonCommandPlanner.SplitTargetColumns(matrix, 5, ","));
            Assert.Empty(RibbonCommandPlanner.SplitTargetColumns(matrix, 1, string.Empty));
        }

        // 回归：工具层 Data 常用匿名对象（new { ... }），修复前 DataObject 只认 Dictionary/JObject，
        // 导致 BuildSummary 拿不到统计字段而全部退化到兜底文案。用匿名对象 + 各工具实际读取的键验证解析。
        [Fact]
        public void BuildSummary_reads_fields_from_anonymous_object_data_all_tools()
        {
            var sel = Selection(2, "A", "B");

            var dedup = new ToolResult { Ok = true, Data = new { deduped_rows_removed = 3 } };
            Assert.Contains("删除重复 3 行", RibbonCommandPlanner.BuildSummary("dedup_range", dedup, sel));

            var split = new ToolResult { Ok = true, Data = new { target_column = 4, split_columns = 3 } };
            Assert.Contains("从第 4 列", RibbonCommandPlanner.BuildSummary("text_split", split, sel));

            var fill = new ToolResult { Ok = true, Data = new { filled_cells = 2 } };
            Assert.Contains("共 2 格", RibbonCommandPlanner.BuildSummary("fill_formula", fill, sel));
        }

        // 回归：匿名对象 Data 的摘要不落在兜底文案上即为通过的健壮性信号。
        [Fact]
        public void BuildSummary_anonymous_data_does_not_fall_back()
        {
            var sel = Selection(2, "A", "B");
            var result = new ToolResult { Ok = true, Data = new { target_column = 4 } };

            var summary = RibbonCommandPlanner.BuildSummary("text_merge", result, sel);

            Assert.Contains("写入第 4 列", summary);
        }

        // ================= 常用组（ExcelJyyjtool）：安全分级 =================

        [Theory]
        [InlineData("insert_text_into_cell", CommandSafety.SafeWrite)]
        [InlineData("delete_text_matches", CommandSafety.SafeWrite)]
        [InlineData("fill_random_numbers", CommandSafety.SafeWrite)]
        [InlineData("fill_number_sequence", CommandSafety.SafeWrite)]
        [InlineData("insert_blank_rows_every_n", CommandSafety.Overwrite)]
        [InlineData("delete_blank_rows_range", CommandSafety.DeleteRows)]
        [InlineData("fill_blank_cells_down", CommandSafety.SafeWrite)]
        [InlineData("merge_duplicate_cells", CommandSafety.Overwrite)]
        public void Classify_maps_common_group_tools(string toolName, CommandSafety expected)
        {
            Assert.Equal(expected, RibbonCommandPlanner.Classify(toolName));
        }

        // ================= 常用组：参数框字段模型 =================

        [Fact]
        public void BuildFields_insert_text_requires_insertion_and_offers_positioning()
        {
            var fields = RibbonCommandPlanner.BuildFields("insert_text_into_cell", Selection(2, "A", "B"));
            Assert.Equal(5, fields.Count);
            var insertion = fields.Single(f => f.Name == "insertion");
            Assert.Equal("text", insertion.Type);
            Assert.True(insertion.Required);
            var when = fields.Single(f => f.Name == "when");
            Assert.Equal(new[] { "char", "separator" }, when.Options.Select(o => o.Value).ToArray());
            Assert.Equal("char", when.Default);
            var where = fields.Single(f => f.Name == "where");
            Assert.Equal(new[] { "before", "after" }, where.Options.Select(o => o.Value).ToArray());
            Assert.Equal("after", where.Default);
        }

        [Fact]
        public void BuildFields_delete_blank_exposes_only_fully_blank_toggle()
        {
            var fields = RibbonCommandPlanner.BuildFields("delete_blank_rows_range", Selection(3));
            var only = Assert.Single(fields);
            Assert.Equal("only_fully_blank", only.Name);
            Assert.Equal("checkbox", only.Type);
            Assert.Equal("false", only.Default);
        }

        [Fact]
        public void BuildFields_fill_sequence_exposes_start_step_count_direction()
        {
            var fields = RibbonCommandPlanner.BuildFields("fill_number_sequence", Selection(3));
            Assert.Equal(4, fields.Count);
            Assert.Equal("text", fields.Single(f => f.Name == "start").Type);
            Assert.Equal("1", fields.Single(f => f.Name == "step").Default);
            Assert.Equal("10", fields.Single(f => f.Name == "count").Default);
            Assert.Equal("down", fields.Single(f => f.Name == "direction").Default);
        }

        // ================= 常用组：实参组装与校验 =================

        [Fact]
        public void BuildArgs_insert_text_requires_insertion_and_positions()
        {
            var bad = RibbonCommandPlanner.BuildArgs("insert_text_into_cell", Selection(2),
                new Dictionary<string, object> { ["insertion"] = "", ["when"] = "char", ["where"] = "after", ["index"] = "0" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("insert_text_into_cell", Selection(2),
                new Dictionary<string, object> { ["insertion"] = "-", ["when"] = "char", ["where"] = "after", ["index"] = "2" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal("-", ok["insertion"]);
            Assert.Equal("char", ok["when"]);
            Assert.Equal("after", ok["where"]);
            Assert.Equal(2, Convert.ToInt32(ok["index"]));
            Assert.False(ok.ContainsKey("separator"));
        }

        [Fact]
        public void BuildArgs_insert_text_requires_separator_when_separator_mode()
        {
            var bad = RibbonCommandPlanner.BuildArgs("insert_text_into_cell", Selection(2),
                new Dictionary<string, object> { ["insertion"] = "-", ["when"] = "separator", ["where"] = "before", ["index"] = "2", ["separator"] = "" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("insert_text_into_cell", Selection(2),
                new Dictionary<string, object> { ["insertion"] = "-", ["when"] = "separator", ["where"] = "before", ["index"] = "2", ["separator"] = "," }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(",", ok["separator"]);
            Assert.Equal("separator", ok["when"]);
        }

        [Fact]
        public void BuildFields_camelCase_serialization_frontend_contract()
        {
            // 回归：前端 on() 按 f.type/f.label/f.name/f.options 读取，而 ParamField/ParamOption 是
            // PascalCase 属性。若序列化不做 camelCase，字段类型全失配 → 参数框 body 空白，
            // 而 collect() 空转 + BuildArgs 默认值兜底掩盖了"用户看不到参数"这个缺陷。
            var fields = RibbonCommandPlanner.BuildFields("clean_excel_data", Selection(3, "a", "b", "c"));
            var jsonSettings = new Newtonsoft.Json.JsonSerializerSettings
            {
                ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver()
            };
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(fields, Newtonsoft.Json.Formatting.None, jsonSettings);
            var arr = Newtonsoft.Json.Linq.JArray.Parse(json);
            Assert.True(arr.Count > 0);
            foreach (Newtonsoft.Json.Linq.JToken item in arr)
            {
                var j = (Newtonsoft.Json.Linq.JObject)item;
                Assert.True(j.ContainsKey("type"), "fields 必须含 camelCase 的 type 键，否则前端类型失配导致参数框空白");
                Assert.True(j.ContainsKey("name"));
                Assert.True(j.ContainsKey("label"));
                Assert.False(j.ContainsKey("Type"));
                // 「加入说明」：hint 必须序列化输出（camelCase），前端 hintHtml() 靠它渲染小字说明
                Assert.True(j.ContainsKey("hint"), "fields 必须含 camelCase 的 hint 键");
                Assert.False(string.IsNullOrWhiteSpace(j.Value<string>("hint")));

                // 「选项空白」根因锁定：序列化后选项键必须是 value/display，前端 o.value/o.display 才能读到。
                // 若键为 Value/Display（PascalCase），前端读到 undefined → 下拉/勾选全部空白。
                if (j["options"] is Newtonsoft.Json.Linq.JArray opts && opts.Count > 0)
                {
                    foreach (var o in opts)
                    {
                        var oj = (Newtonsoft.Json.Linq.JObject)o;
                        Assert.True(oj.ContainsKey("value"), "选项必须含 camelCase 的 value 键");
                        Assert.True(oj.ContainsKey("display"), "选项必须含 camelCase 的 display 键");
                        Assert.False(oj.ContainsKey("Value"));
                        Assert.False(oj.ContainsKey("Display"));
                    }
                }
            }
        }

        [Fact]
        public void BuildArgs_delete_text_requires_pattern()
        {
            var bad = RibbonCommandPlanner.BuildArgs("delete_text_matches", Selection(2),
                new Dictionary<string, object> { ["pattern"] = "  " });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("delete_text_matches", Selection(2),
                new Dictionary<string, object> { ["pattern"] = @"\d+" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(@"\d+", ok["pattern"]);
        }

        [Fact]
        public void BuildArgs_random_numbers_validates_range_and_decimals()
        {
            var badRange = RibbonCommandPlanner.BuildArgs("fill_random_numbers", Selection(2),
                new Dictionary<string, object> { ["min"] = "10", ["max"] = "1", ["decimals"] = "0" });
            Assert.IsType<string>(badRange);

            var badDecimals = RibbonCommandPlanner.BuildArgs("fill_random_numbers", Selection(2),
                new Dictionary<string, object> { ["min"] = "1", ["max"] = "10", ["decimals"] = "9" });
            Assert.IsType<string>(badDecimals);

            var ok = RibbonCommandPlanner.BuildArgs("fill_random_numbers", Selection(2),
                new Dictionary<string, object> { ["min"] = "1", ["max"] = "10", ["decimals"] = "2" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(2, Convert.ToInt32(ok["decimals"]));
        }

        [Fact]
        public void BuildArgs_number_sequence_validates_direction_and_caps_count()
        {
            var badDir = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                new Dictionary<string, object> { ["step"] = "1", ["count"] = "10", ["direction"] = "diag" });
            Assert.IsType<string>(badDir);

            var ok = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                new Dictionary<string, object> { ["step"] = "-2", ["count"] = "5", ["direction"] = "right" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(-2D, Convert.ToDouble(ok["step"]));
            Assert.Equal("right", ok["direction"]);
        }

        [Fact]
        public void BuildArgs_number_sequence_passes_start_when_filled()
        {
            var with = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                new Dictionary<string, object> { ["start"] = "100", ["step"] = "1", ["count"] = "5", ["direction"] = "down" }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal(100D, Convert.ToDouble(with["start"]));

            var without = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                new Dictionary<string, object> { ["step"] = "1", ["count"] = "5", ["direction"] = "down" }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("start"));

            var bad = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                new Dictionary<string, object> { ["start"] = "abc", ["step"] = "1", ["count"] = "5", ["direction"] = "down" });
            Assert.IsType<string>(bad);
        }

        [Fact]
        public void BuildArgs_insert_blank_validates_positive_interval_and_count()
        {
            var bad = RibbonCommandPlanner.BuildArgs("insert_blank_rows_every_n", Selection(2),
                new Dictionary<string, object> { ["interval"] = "0", ["count"] = "1" });
            Assert.IsType<string>(bad);

            var ok = RibbonCommandPlanner.BuildArgs("insert_blank_rows_every_n", Selection(2),
                new Dictionary<string, object> { ["interval"] = "2", ["count"] = "3" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(2, Convert.ToInt32(ok["interval"]));
            Assert.Equal(3, Convert.ToInt32(ok["count"]));
        }

        [Fact]
        public void BuildArgs_delete_blank_and_fill_down_and_merge_map_booleans()
        {
            var del = RibbonCommandPlanner.BuildArgs("delete_blank_rows_range", Selection(2),
                new Dictionary<string, object> { ["only_fully_blank"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(del);
            Assert.True((bool)del["only_fully_blank"]);

            var fill = RibbonCommandPlanner.BuildArgs("fill_blank_cells_down", Selection(2),
                new Dictionary<string, object> { ["copy_value"] = "false" }) as Dictionary<string, object>;
            Assert.NotNull(fill);
            Assert.False((bool)fill["copy_value"]);

            var merge = RibbonCommandPlanner.BuildArgs("merge_duplicate_cells", Selection(2),
                new Dictionary<string, object> { ["first_only"] = "true" }) as Dictionary<string, object>;
            Assert.NotNull(merge);
            Assert.True((bool)merge["first_only"]);
        }

        // ================= 删除空行行数预算（与 DeleteBlankRowsRange 语义对齐） =================

        [Fact]
        public void CountRemovedBlankRows_only_fully_blank_mode()
        {
            var matrix = new object[,]
            {
                { "A", "B" },
                { 1D, 2D },
                { null, null },       // 整行全空 → 删
                { 3D, DBNull.Value }  // 非全空 → 保留
            };

            Assert.Equal(1, RibbonCommandPlanner.CountRemovedBlankRows(matrix, true));
        }

        [Fact]
        public void CountRemovedBlankRows_first_column_empty_mode()
        {
            var matrix = new object[,]
            {
                { null, "x" }, // 首列为空 → 删
                { 1D, "y" },
                { DBNull.Value, "z" } // 首列为空（DBNull）→ 删
            };

            Assert.Equal(2, RibbonCommandPlanner.CountRemovedBlankRows(matrix, false));
        }

        [Fact]
        public void CountRemovedBlankRows_null_matrix_returns_zero()
        {
            Assert.Equal(0, RibbonCommandPlanner.CountRemovedBlankRows(null, true));
            Assert.Equal(0, RibbonCommandPlanner.CountRemovedBlankRows(null, false));
        }

        // ================= 常用组：结果摘要（读匿名对象字段） =================

        [Fact]
        public void BuildSummary_common_group_reads_tool_stats()
        {
            var sel = Selection(2, "A", "B");

            var delBlank = new ToolResult { Ok = true, Data = new { deleted_rows = 3 } };
            Assert.Contains("已删除 3 个空行", RibbonCommandPlanner.BuildSummary("delete_blank_rows_range", delBlank, sel));

            var insBlank = new ToolResult { Ok = true, Data = new { inserted_rows = 5, interval = 2 } };
            Assert.Contains("已每隔 2 行插入 5 个空行", RibbonCommandPlanner.BuildSummary("insert_blank_rows_every_n", insBlank, sel));

            var fillDown = new ToolResult { Ok = true, Data = new { filled_cells = 7 } };
            Assert.Contains("已向下填充 7 个空单元格", RibbonCommandPlanner.BuildSummary("fill_blank_cells_down", fillDown, sel));

            var delText = new ToolResult { Ok = true, Data = new { changed_cells = 4 } };
            Assert.Contains("共改动 4 格", RibbonCommandPlanner.BuildSummary("delete_text_matches", delText, sel));

            var rand = new ToolResult { Ok = true, Data = new { rows = 2, columns = 3, min = 1D, max = 5D } };
            Assert.Contains("已填充 2×3 个随机数（1–5）", RibbonCommandPlanner.BuildSummary("fill_random_numbers", rand, sel));

            var seq = new ToolResult { Ok = true, Data = new { count = 5, start = 10D, step = 2D, direction = "down" } };
            Assert.Contains("起 10，步长 2，共 5 个", RibbonCommandPlanner.BuildSummary("fill_number_sequence", seq, sel));

            var merge = new ToolResult { Ok = true, Data = new { merged = true } };
            Assert.Contains("已合并相邻相同单元格", RibbonCommandPlanner.BuildSummary("merge_duplicate_cells", merge, sel));
        }

        // ================= 对抗审查：参数框选项契约审计（2026-09-07） =================
        // 双向锁定：①每个弹窗字段名必须被 BuildArgs 消费；②BuildArgs 产出的每个键
        // 必须落在工具 schema 的真实参数名内——防止「弹窗给了选项但工具根本不读」的假选项。

        private static readonly Dictionary<string, string[]> RealToolParameterContract = new Dictionary<string, string[]>
        {
            ["clean_excel_data"] = new[] { "address", "sheet", "operations" },
            ["dedup_range"] = new[] { "address", "sheet", "columns", "keep", "header_rows", "case_sensitive" },
            ["sort_range"] = new[] { "address", "sheet", "sort_column", "direction", "sort_column2", "direction2", "header_rows" },
            ["transpose_range"] = new[] { "address", "sheet", "target_address", "target_sheet" },
            ["filter_excel_data"] = new[] { "address", "sheet", "column", "operator", "value", "target_address" },
            ["create_excel_pivot_summary"] = new[] { "source_address", "target_address", "label_column", "value_column" },
            ["text_split"] = new[] { "address", "sheet", "column", "delimiter", "target_column", "keep_source" },
            ["text_merge"] = new[] { "address", "sheet", "columns", "target_column", "delimiter", "skip_empty" },
            ["text_clean"] = new[] { "address", "sheet", "trim", "collapse_whitespace", "remove_non_printable", "case", "remove_chars", "find", "replace" },
            ["fill_formula"] = new[] { "address", "sheet", "formula", "start_row" },
            ["insert_text_into_cell"] = new[] { "address", "sheet", "insertion", "when", "where", "index", "separator" },
            ["delete_text_matches"] = new[] { "address", "sheet", "pattern" },
            ["fill_random_numbers"] = new[] { "address", "sheet", "min", "max", "decimals" },
            ["fill_number_sequence"] = new[] { "address", "sheet", "start", "step", "count", "direction" },
            ["insert_blank_rows_every_n"] = new[] { "address", "sheet", "interval", "count" },
            ["delete_blank_rows_range"] = new[] { "address", "sheet", "only_fully_blank" },
            ["fill_blank_cells_down"] = new[] { "address", "sheet", "copy_value" },
            ["merge_duplicate_cells"] = new[] { "address", "sheet", "first_only" }
        };

        // 弹窗 UI 字段名 → BuildArgs 实际消费名（含 UI 专用名：scope/has_header/target_mode/collapse，
        // 以及 clean 折叠进 operations 的开关/列字段）。方向①用本表校验「没有死字段」。
        private static readonly Dictionary<string, string[]> ConsumedUiFieldNames = new Dictionary<string, string[]>
        {
            ["clean_excel_data"] = new[] { "remove_blank_rows", "fix_dates", "date_column", "date_format", "fill_blanks", "fill_column", "fill_value", "trim", "remove_non_printable" },
            ["dedup_range"] = new[] { "scope", "columns", "keep", "has_header", "case_sensitive" },
            ["sort_range"] = new[] { "sort_column", "direction", "sort_column2", "direction2", "has_header" },
            ["transpose_range"] = new[] { "target_mode", "target_address", "target_sheet" },
            ["filter_excel_data"] = new[] { "column", "operator", "value", "target_address" },
            ["create_excel_pivot_summary"] = new[] { "label_column", "value_column", "target_address" },
            ["text_split"] = new[] { "column", "delimiter", "target_column", "keep_source" },
            ["text_merge"] = new[] { "columns", "target_column", "delimiter", "skip_empty" },
            ["text_clean"] = new[] { "trim", "collapse", "remove_non_printable", "case", "remove_chars", "find", "replace" },
            ["fill_formula"] = new[] { "formula", "start_row" },
            ["insert_text_into_cell"] = new[] { "insertion", "when", "where", "index", "separator" },
            ["delete_text_matches"] = new[] { "pattern" },
            ["fill_random_numbers"] = new[] { "min", "max", "decimals" },
            ["fill_number_sequence"] = new[] { "start", "step", "count", "direction" },
            ["insert_blank_rows_every_n"] = new[] { "interval", "count" },
            ["delete_blank_rows_range"] = new[] { "only_fully_blank" },
            ["fill_blank_cells_down"] = new[] { "copy_value" },
            ["merge_duplicate_cells"] = new[] { "first_only" }
        };

        [Fact]
        public void BuildFields_every_field_name_is_consumed_by_build_args()
        {
            // 方向①：弹窗暴露的每个字段，BuildArgs 必须按同名消费（否则选项是死的）。
            foreach (var kv in ConsumedUiFieldNames)
            {
                var fields = RibbonCommandPlanner.BuildFields(kv.Key, Selection(3, "a", "b", "c"));
                foreach (var f in fields)
                {
                    Assert.True(kv.Value.Contains(f.Name), $"{kv.Key} 字段 {f.Name} 不在消费契约中，BuildArgs 未消费该选项");
                }
            }
        }

        [Fact]
        public void BuildArgs_produced_keys_stay_within_real_tool_parameters()
        {
            // 方向②：用弹窗默认值 + 必需文本补样，BuildArgs 产出的每个键必须在真实 schema 参数内。
            var requiredFills = new Dictionary<string, string>
            {
                ["insert_text_into_cell"] = "-",
                ["delete_text_matches"] = "x",
                ["fill_formula"] = "=A{row}"
            };
            foreach (var kv in RealToolParameterContract)
            {
                var fields = RibbonCommandPlanner.BuildFields(kv.Key, Selection(3, "a", "b", "c"));
                var values = new Dictionary<string, object>();
                foreach (var f in fields)
                {
                    if (f.Type == "multiselect")
                    {
                        values[f.Name] = new List<string>(f.Defaults);
                    }
                    else if (requiredFills.TryGetValue(kv.Key, out var fill) && f.Name == "formula" && kv.Key == "fill_formula")
                    {
                        values[f.Name] = fill;
                    }
                    else
                    {
                        values[f.Name] = requiredFills.ContainsKey(kv.Key) && f.Required ? requiredFills[kv.Key] : f.Default;
                    }
                }

                var build = RibbonCommandPlanner.BuildArgs(kv.Key, Selection(3, "a", "b", "c"), values);
                var args = build as Dictionary<string, object>;
                Assert.NotNull(args);
                foreach (var key in args.Keys)
                {
                    Assert.True(kv.Value.Contains(Convert.ToString(key)), $"{kv.Key} 产出了 schema 之外的参数 {key}，属假选项或未接线");
                }
            }
        }

        [Fact]
        public void BuildFields_transpose_offers_target_sheet_and_keeps_address_same_sheet()
        {
            var fields = RibbonCommandPlanner.BuildFields("transpose_range", Selection(3, "a", "b", "c"));
            Assert.Equal(3, fields.Count);
            var sheetField = fields.Single(f => f.Name == "target_sheet");
            Assert.Equal("text", sheetField.Type);
            Assert.False(string.IsNullOrWhiteSpace(sheetField.Hint));
            // 回归：target_address 提示不得再承诺跨表地址（写路径 sheet.Range 不支持 ! 跨表）
            var addr = fields.Single(f => f.Name == "target_address");
            Assert.DoesNotContain("!", addr.Hint);
            Assert.Contains("目标工作表名", addr.Hint);
        }

        [Fact]
        public void BuildArgs_transpose_maps_target_sheet_when_filled()
        {
            var with = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object>
                {
                    ["target_mode"] = "custom",
                    ["target_address"] = "F1",
                    ["target_sheet"] = "Sheet2"
                }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal("F1", with["target_address"]);
            Assert.Equal("Sheet2", with["target_sheet"]);

            var without = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object> { ["target_mode"] = "top_left" }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("target_address"));
            Assert.False(without.ContainsKey("target_sheet"));
        }

        [Fact]
        public void BuildFields_text_split_offers_target_column()
        {
            var fields = RibbonCommandPlanner.BuildFields("text_split", Selection(3, "a", "b", "c"));
            Assert.Equal(4, fields.Count);
            var tc = fields.Single(f => f.Name == "target_column");
            Assert.Equal("text", tc.Type);
            Assert.False(string.IsNullOrWhiteSpace(tc.Hint));
        }

        [Fact]
        public void BuildArgs_text_split_maps_and_validates_target_column()
        {
            var with = RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object>
                {
                    ["column"] = "1",
                    ["delimiter"] = ",",
                    ["target_column"] = "5",
                    ["keep_source"] = "true"
                }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal(5, Convert.ToInt32(with["target_column"]));

            var without = RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = "," }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("target_column"));

            Assert.IsType<string>(RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = ",", ["target_column"] = "0" }));
            Assert.IsType<string>(RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = ",", ["target_column"] = "abc" }));
            Assert.IsType<string>(RibbonCommandPlanner.BuildArgs("text_split", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["delimiter"] = ",", ["target_column"] = "99999" }));
        }

        [Fact]
        public void BuildFields_text_clean_offers_find_and_replace()
        {
            var fields = RibbonCommandPlanner.BuildFields("text_clean", Selection(3, "a", "b", "c"));
            Assert.Equal(7, fields.Count);
            var find = fields.Single(f => f.Name == "find");
            Assert.Equal("text", find.Type);
            Assert.False(string.IsNullOrWhiteSpace(find.Hint));
            var replace = fields.Single(f => f.Name == "replace");
            Assert.Equal("text", replace.Type);
        }

        [Fact]
        public void BuildArgs_text_clean_maps_find_replace_only_when_filled()
        {
            var with = RibbonCommandPlanner.BuildArgs("text_clean", Selection(2),
                new Dictionary<string, object>
                {
                    ["find"] = "abc",
                    ["replace"] = "xyz"
                }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal("abc", with["find"]);
            Assert.Equal("xyz", with["replace"]);

            var findOnly = RibbonCommandPlanner.BuildArgs("text_clean", Selection(2),
                new Dictionary<string, object> { ["find"] = "abc" }) as Dictionary<string, object>;
            Assert.NotNull(findOnly);
            Assert.Equal("abc", findOnly["find"]);
            Assert.False(findOnly.ContainsKey("replace"));

            var none = RibbonCommandPlanner.BuildArgs("text_clean", Selection(2),
                new Dictionary<string, object> { ["replace"] = " " }) as Dictionary<string, object>;
            Assert.NotNull(none);
            Assert.False(none.ContainsKey("find"));
            Assert.False(none.ContainsKey("replace"));
        }

        [Fact]
        public void BuildFields_pivot_offers_target_address()
        {
            var fields = RibbonCommandPlanner.BuildFields("create_excel_pivot_summary", Selection(3, "a", "b", "c"));
            Assert.Equal(3, fields.Count);
            var ta = fields.Single(f => f.Name == "target_address");
            Assert.Equal("text", ta.Type);
            Assert.False(string.IsNullOrWhiteSpace(ta.Hint));
        }

        [Fact]
        public void BuildArgs_pivot_maps_target_address_when_filled()
        {
            var with = RibbonCommandPlanner.BuildArgs("create_excel_pivot_summary", Selection(2),
                new Dictionary<string, object> { ["label_column"] = "1", ["value_column"] = "2", ["target_address"] = "E5" }) as Dictionary<string, object>;
            Assert.NotNull(with);
            Assert.Equal("E5", with["target_address"]);

            var without = RibbonCommandPlanner.BuildArgs("create_excel_pivot_summary", Selection(2),
                new Dictionary<string, object> { ["label_column"] = "1", ["value_column"] = "0" }) as Dictionary<string, object>;
            Assert.NotNull(without);
            Assert.False(without.ContainsKey("target_address"));
        }

        [Fact]
        public void BuildFields_fill_random_min_default_matches_schema()
        {
            // 对齐工具 schema：min 默认 0（此前弹窗默认 1，与工具文档默认不一致）
            var fields = RibbonCommandPlanner.BuildFields("fill_random_numbers", Selection(2));
            Assert.Equal("0", fields.Single(f => f.Name == "min").Default);
        }

        // ================= 对抗审查第二轮（2026-09-07）：矛盾配置 / NaN 渗透 / 预算 / 枚举校验 =================

        [Fact]
        public void BuildArgs_transpose_rejects_target_sheet_in_top_left_mode()
        {
            // top_left 写回源区，填目标工作表名无意义且工具层行为未定义，规划层拦截
            var build = RibbonCommandPlanner.BuildArgs("transpose_range", Selection(2),
                new Dictionary<string, object> { ["target_mode"] = "top_left", ["target_sheet"] = "Sheet2" });
            Assert.IsType<string>(build);
            Assert.Contains("指定位置", build as string);

            // custom 模式 + target_sheet 合法（已由 BuildArgs_transpose_maps_target_sheet_when_filled 覆盖）
        }

        [Fact]
        public void BuildArgs_sort_rejects_invalid_direction()
        {
            var build = RibbonCommandPlanner.BuildArgs("sort_range", Selection(3),
                new Dictionary<string, object> { ["sort_column"] = "1", ["direction"] = "diag", ["has_header"] = "true" });
            Assert.IsType<string>(build);
        }

        [Fact]
        public void BuildArgs_filter_rejects_invalid_operator()
        {
            var build = RibbonCommandPlanner.BuildArgs("filter_excel_data", Selection(3),
                new Dictionary<string, object> { ["column"] = "1", ["operator"] = "like", ["value"] = "x" });
            Assert.IsType<string>(build);
        }

        [Fact]
        public void BuildArgs_delete_text_rejects_invalid_regex()
        {
            // 非法正则（"["）预检拦截，避免工具层静默不删除让用户误以为成功
            var build = RibbonCommandPlanner.BuildArgs("delete_text_matches", Selection(2),
                new Dictionary<string, object> { ["pattern"] = "[" });
            Assert.IsType<string>(build);
            Assert.Contains("正则表达式无效", build as string);
        }

        [Fact]
        public void BuildArgs_random_numbers_rejects_nan_and_infinity()
        {
            // .NET Framework 的 double.TryParse 能解析 NaN/Infinity 符号，必须显式拦截，
            // 否则 NaN/Infinity 渗透到工具层填格（Excel 显示 #NUM! 或写坏数据）。
            foreach (var bad in new[] { "NaN", "Infinity", "-Infinity", "1e999" })
            {
                var minBad = RibbonCommandPlanner.BuildArgs("fill_random_numbers", Selection(2),
                    new Dictionary<string, object> { ["min"] = bad, ["max"] = "100", ["decimals"] = "0" });
                Assert.IsType<string>(minBad);

                var maxBad = RibbonCommandPlanner.BuildArgs("fill_random_numbers", Selection(2),
                    new Dictionary<string, object> { ["min"] = "0", ["max"] = bad, ["decimals"] = "0" });
                Assert.IsType<string>(maxBad);
            }
        }

        [Fact]
        public void BuildArgs_number_sequence_rejects_nan_and_infinity()
        {
            foreach (var bad in new[] { "NaN", "Infinity", "-Infinity" })
            {
                var stepBad = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                    new Dictionary<string, object> { ["step"] = bad, ["count"] = "5", ["direction"] = "down" });
                Assert.IsType<string>(stepBad);

                var startBad = RibbonCommandPlanner.BuildArgs("fill_number_sequence", Selection(2),
                    new Dictionary<string, object> { ["start"] = bad, ["step"] = "1", ["count"] = "5", ["direction"] = "down" });
                Assert.IsType<string>(startBad);
            }
        }

        [Fact]
        public void BuildArgs_insert_blank_caps_insertion_budget()
        {
            // 工具层无上限：每隔 N 行插 count 行，按选区行数预算插入总量，超限拒绝。
            // Selection 固定 4 行：interval=1, count=50000 → 预算 (4/1+1)*50000=250000 > 100000
            var over = RibbonCommandPlanner.BuildArgs("insert_blank_rows_every_n", Selection(4),
                new Dictionary<string, object> { ["interval"] = "1", ["count"] = "50000" });
            Assert.IsType<string>(over);
            Assert.Contains("过多", over as string);

            var ok = RibbonCommandPlanner.BuildArgs("insert_blank_rows_every_n", Selection(1000),
                new Dictionary<string, object> { ["interval"] = "10", ["count"] = "5" }) as Dictionary<string, object>;
            Assert.NotNull(ok);
            Assert.Equal(10, Convert.ToInt32(ok["interval"]));
            Assert.Equal(5, Convert.ToInt32(ok["count"]));
        }

        // ================= 执行标准增强（2026-09-07）：目标守卫 / 地址解析 / 执行后自检纯函数 =================

        [Theory]
        [InlineData("F1", 1, 6)]
        [InlineData("$F$1", 1, 6)]
        [InlineData("f1", 1, 6)]
        [InlineData("A1", 1, 1)]
        [InlineData("XFD1048576", 1048576, 16384)]
        [InlineData("", 0, 0)]
        [InlineData("abc", 0, 0)]
        [InlineData("A0", 0, 0)]
        [InlineData("Sheet2!F1", 0, 0)]
        public void TryParseAddress_parses_absolute_addresses(string addr, int expectRow, int expectCol)
        {
            var ok = RibbonCommandPlanner.TryParseAddress(addr, out var row, out var col);
            if (expectRow == 0)
            {
                Assert.False(ok);
            }
            else
            {
                Assert.True(ok);
                Assert.Equal(expectRow, row);
                Assert.Equal(expectCol, col);
            }
        }

        [Fact]
        public void TryGetTargetRect_text_merge_returns_result_column_rect()
        {
            var sel = Selection(2);
            var args = new Dictionary<string, object> { ["target_column"] = 3 };
            var ok = RibbonCommandPlanner.TryGetTargetRect("text_merge", sel, null, args, 1, 1,
                out var sheet, out var top, out var left, out var rows, out var cols);
            Assert.True(ok);
            Assert.Equal("Sheet1", sheet);
            Assert.Equal(1, top);
            Assert.Equal(3, left);   // startCol(1) + target_column(3) - 1
            Assert.Equal(4, rows);   // 选区行数
            Assert.Equal(1, cols);
        }

        [Fact]
        public void TryGetTargetRect_text_split_returns_band_rect()
        {
            var sel = Selection(2);
            var matrix = new object[,] { { "a,b", "x" }, { "c,d", "y" }, { "e,f", "z" } };
            var args = new Dictionary<string, object> { ["column"] = 1, ["delimiter"] = "," };
            var ok = RibbonCommandPlanner.TryGetTargetRect("text_split", sel, matrix, args, 1, 1,
                out var sheet, out var top, out var left, out var rows, out var cols);
            Assert.True(ok);
            Assert.Equal(1, top);
            Assert.Equal(2, left);   // 源列右侧相邻列
            Assert.Equal(4, rows);
            Assert.Equal(2, cols);   // maxParts=2 → 列 2..3
        }

        [Fact]
        public void TryGetTargetRect_transpose_custom_returns_transposed_rect()
        {
            var sel = Selection(3); // 4 行 3 列
            var args = new Dictionary<string, object> { ["target_mode"] = "custom", ["target_address"] = "$F$1" };
            var ok = RibbonCommandPlanner.TryGetTargetRect("transpose_range", sel, null, args, 1, 1,
                out var sheet, out var top, out var left, out var rows, out var cols);
            Assert.True(ok);
            Assert.Equal(6, left);
            Assert.Equal(1, top);
            Assert.Equal(3, rows);   // 转置后 3 行
            Assert.Equal(4, cols);   // 转置后 4 列
        }

        [Fact]
        public void TryGetTargetRect_returns_false_for_non_guard_tools()
        {
            var sel = Selection(2);
            Assert.False(RibbonCommandPlanner.TryGetTargetRect("sort_range", sel, null, new Dictionary<string, object>(), 1, 1,
                out _, out _, out _, out _, out _));
            // transpose 原位模式不守卫
            Assert.False(RibbonCommandPlanner.TryGetTargetRect("transpose_range", sel, null,
                new Dictionary<string, object> { ["target_mode"] = "top_left" }, 1, 1, out _, out _, out _, out _, out _));
        }

        [Fact]
        public void VerifyTransposeResult_checks_dimension_and_anchors()
        {
            var source = new object[,] { { "a", "b", "c" }, { "d", "e", "f" } };
            var good = new object[,] { { "a", "d" }, { "b", "e" }, { "c", "f" } };
            Assert.Null(RibbonCommandPlanner.VerifyTransposeResult(source, good));

            var wrongDim = new object[,] { { "a", "d" }, { "b", "e" } };
            var fail = RibbonCommandPlanner.VerifyTransposeResult(source, wrongDim);
            Assert.NotNull(fail);
            Assert.Contains("维度不符", fail);

            var wrongValue = new object[,] { { "a", "d" }, { "b", "X" }, { "c", "f" } };
            var fail2 = RibbonCommandPlanner.VerifyTransposeResult(source, wrongValue);
            Assert.NotNull(fail2);
            Assert.Contains("锚点", fail2);
        }

        [Fact]
        public void VerifyRandomRange_accepts_in_range_and_rejects_out_of_range()
        {
            var values = new object[,] { { 1.0, 2.5 }, { 4.0, 5.0 } };
            Assert.Null(RibbonCommandPlanner.VerifyRandomRange(values, 1.0, 5.0, 0));

            var outOfRange = new object[,] { { 1.0, 6.0 } };
            var fail = RibbonCommandPlanner.VerifyRandomRange(outOfRange, 1.0, 5.0, 0);
            Assert.NotNull(fail);

            // 舍入容差：min=0.5/max=0.6、decimals=0 时 round 后得 1，仍应在容差内通过
            var rounded = new object[,] { { 1.0 } };
            Assert.Null(RibbonCommandPlanner.VerifyRandomRange(rounded, 0.5, 0.6, 0));

            var empty = new object[,] { { null } };
            Assert.NotNull(RibbonCommandPlanner.VerifyRandomRange(empty, 1.0, 5.0, 0));
        }

        // ================= 执行标准 5-8（2026-09-07）：审计注册表 / 错误码文案 =================

        [Fact]
        public void NeedsPostVerify_covers_high_risk_tools()
        {
            foreach (var name in new[] { "transpose_range", "dedup_range", "delete_blank_rows_range", "fill_random_numbers" })
            {
                Assert.True(RibbonCommandPlanner.NeedsPostVerify(name), $"{name} 应纳入执行后自检（新增高风险工具必须在此表态）");
            }
            foreach (var name in new[] { "sort_range", "filter_excel_data", "text_clean", "clean_excel_data" })
            {
                Assert.False(RibbonCommandPlanner.NeedsPostVerify(name), $"{name} 不应自检（低风险或动态目标）");
            }
        }

        [Fact]
        public void BuildSummary_invalid_pattern_shows_friendly_text()
        {
            // 第 7 条：INVALID_PATTERN/INVALID_CASE 带中文文案时展示文案而非裸码
            var result = new ToolResult { Ok = false, ErrorCode = "INVALID_PATTERN", Data = "正则表达式过长（≤256 字符），已取消执行。" };
            Assert.Equal("执行失败：正则表达式过长（≤256 字符），已取消执行。",
                RibbonCommandPlanner.BuildSummary("delete_text_matches", result, Selection(2)));

            var bare = new ToolResult { Ok = false, ErrorCode = "INVALID_CASE", Data = null };
            var summary = RibbonCommandPlanner.BuildSummary("text_clean", bare, Selection(2));
            Assert.Contains("INVALID_CASE", summary);
        }

        // ================= 纯函数抽取（2026-09-07）：目标占用判定 / 样本提取 =================

        [Fact]
        public void MatrixHasData_detects_any_nonblank_cell()
        {
            Assert.False(RibbonCommandPlanner.MatrixHasData(null), "null 矩阵不应视为占用");
            Assert.False(RibbonCommandPlanner.MatrixHasData(new object[,] { { null, DBNull.Value }, { string.Empty, null } }),
                "全空/空白矩阵不应视为占用");

            Assert.True(RibbonCommandPlanner.MatrixHasData(new object[,] { { "x", null } }), "任一非空文本即占用");
            Assert.True(RibbonCommandPlanner.MatrixHasData(new object[,] { { 0, null } }), "数值 0 也是数据（非空白判断，非非空判断）");
            Assert.True(RibbonCommandPlanner.MatrixHasData(new object[,] { { "  " } }) == false, "纯空白文本不算占用");
        }

        [Fact]
        public void PickSampleValues_returns_up_to_max_nonempty()
        {
            Assert.Empty(RibbonCommandPlanner.PickSampleValues(null, 3));
            Assert.Empty(RibbonCommandPlanner.PickSampleValues(new object[,] { { null, DBNull.Value } }, 3));
            Assert.Empty(RibbonCommandPlanner.PickSampleValues(new object[,] { { "a" } }, 0));

            var values = new object[,] { { 1.5, null, "x" }, { DBNull.Value, "y", 2 } };
            var samples = RibbonCommandPlanner.PickSampleValues(values, 3);
            Assert.Equal(3, samples.Count);
            Assert.Equal(new[] { "1.5", "x", "y" }, samples);

            var one = RibbonCommandPlanner.PickSampleValues(new object[,] { { "z", "w" } }, 1);
            Assert.Single(one);
            Assert.Equal("z", one[0]);
        }
    }
}