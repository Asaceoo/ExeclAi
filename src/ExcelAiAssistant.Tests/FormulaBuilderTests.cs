using System;
using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// FormulaBuilder 公式构造器锁定真值单测（模板文本/表头钳制/引用卫生/criteria 形态/绝对相对转换）。
    /// </summary>
    public class FormulaBuilderTests
    {
        private static FormulaBuilder.BuildOptions O(string template, string source)
        {
            return new FormulaBuilder.BuildOptions { Template = template, SourceAddress = source };
        }

        private static FormulaBuilder.BuildOptions With(FormulaBuilder.BuildOptions o, Action<FormulaBuilder.BuildOptions> set)
        {
            set(o);
            return o;
        }

        [Fact]
        public void Sum_plain_range_bare_address_same_sheet()
        {
            var f = FormulaBuilder.Build(With(O("sum", "$A$1:$F$21"), o => { o.ValueColumn = 6; }));
            Assert.Equal("=SUM($F$2:$F$21)", f);
        }

        [Fact]
        public void Header_rows_clamped_data_starts_after_header()
        {
            // 表头 10 行（测试2026 教训：表头不在第 1 行）
            var f = FormulaBuilder.Build(With(O("sum", "$B$10:$S$27"), o => { o.ValueColumn = 1; o.HeaderRows = 10; }));
            Assert.Equal("=SUM($B$20:$B$27)", f);
        }

        [Fact]
        public void Cross_sheet_qualified_and_quoted_when_needed()
        {
            var f = FormulaBuilder.Build(With(O("count", "$A$1:$B$5"), o => { o.ValueColumn = 2; o.SheetName = "My Sheet"; }));
            Assert.Equal("=COUNT('My Sheet'!$B$2:$B$5)", f);
            var f2 = FormulaBuilder.Build(With(O("count", "$A$1:$B$5"), o => { o.ValueColumn = 2; o.SheetName = "明细数据"; }));
            Assert.Equal("=COUNT(明细数据!$B$2:$B$5)", f2);
        }

        [Fact]
        public void Average_wrapped_in_iferror_three_piece()
        {
            var f = FormulaBuilder.Build(With(O("average", "$A$1:$B$5"), o => { o.ValueColumn = 2; }));
            Assert.Equal("=IFERROR(AVERAGE($B$2:$B$5),0)", f);
        }

        [Fact]
        public void Sumifs_criteria_number_and_text_and_operator_and_cellref()
        {
            var o = With(O("sumifs", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.CriteriaColumn = 1; });
            o.Criteria = "100";
            Assert.Equal("=SUMIFS($C$2:$C$21,$A$2:$A$21,100)", FormulaBuilder.Build(o));
            o.Criteria = "华东";
            Assert.Equal("=SUMIFS($C$2:$C$21,$A$2:$A$21,\"华东\")", FormulaBuilder.Build(o));
            o.Criteria = ">=100";
            Assert.Equal("=SUMIFS($C$2:$C$21,$A$2:$A$21,\">=100\")", FormulaBuilder.Build(o));
            o.Criteria = "$H$1";
            Assert.Equal("=SUMIFS($C$2:$C$21,$A$2:$A$21,$H$1)", FormulaBuilder.Build(o));
        }

        [Fact]
        public void Countifs_and_averageifs()
        {
            var o = With(O("countifs", "$A$1:$B$9"), x => { x.CriteriaColumn = 1; x.Criteria = "华东"; });
            Assert.Equal("=COUNTIFS($A$2:$A$9,\"华东\")", FormulaBuilder.Build(o));
            var o2 = With(O("averageifs", "$A$1:$C$9"), x => { x.ValueColumn = 3; x.CriteriaColumn = 1; x.Criteria = "华东"; });
            Assert.Equal("=IFERROR(AVERAGEIFS($C$2:$C$9,$A$2:$A$9,\"华东\"),0)", FormulaBuilder.Build(o2));
        }

        [Fact]
        public void Weighted_avg_requires_weight_column()
        {
            var o = With(O("weighted_avg", "$A$1:$C$9"), x => { x.ValueColumn = 2; x.WeightColumn = 3; });
            Assert.Equal("=IFERROR(SUMPRODUCT($B$2:$B$9,$C$2:$C$9)/SUM($C$2:$C$9),0)", FormulaBuilder.Build(o));
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(With(O("weighted_avg", "$A$1:$C$9"), o => { o.ValueColumn = 2; })));
        }

        [Fact]
        public void Distinct_count_uses_label_column()
        {
            var o = With(O("distinct_count", "$A$1:$B$9"), x => { x.CriteriaColumn = 1; });
            Assert.Equal("=IFERROR(SUMPRODUCT(1/COUNTIF($A$2:$A$9,$A$2:$A$9)),0)", FormulaBuilder.Build(o));
        }

        [Fact]
        public void Unknown_template_fails_closed()
        {
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(O("vlookup_magic", "$A$1:$B$9")));
        }

        [Fact]
        public void Wrap_iferror_idempotent()
        {
            Assert.Equal("=IFERROR(SUM(A1:A2),0)", FormulaBuilder.WrapIfError("=SUM(A1:A2)"));
            Assert.Equal("=IFERROR(SUM(A1:A2),0)", FormulaBuilder.WrapIfError("=IFERROR(SUM(A1:A2),0)"));
            Assert.Equal("=IFERROR(SUM(A1:A2),\"\")", FormulaBuilder.WrapIfError("=SUM(A1:A2)", "\"\""));
        }

        [Fact]
        public void ToAbsolute_and_ToRelative()
        {
            Assert.Equal("=SUM($A$2:$A$9)", FormulaBuilder.ToAbsolute("=SUM(A2:A9)"));
            Assert.Equal("=SUM(A2:A9)", FormulaBuilder.ToRelative("=SUM($A$2:$A$9)"));
            // 函数名内的字母数字不被误转
            Assert.Equal("=LOG10($A$1)", FormulaBuilder.ToAbsolute("=LOG10(A1)"));
        }

        [Fact]
        public void TryParseAddress_sheet_variants()
        {
            Assert.True(FormulaBuilder.TryParseAddress("'My Sheet'!$A$1:$B$2", out var sheet, out var c1, out var r1, out var c2, out var r2));
            Assert.Equal("My Sheet", sheet);
            Assert.Equal(1, c1); Assert.Equal(1, r1); Assert.Equal(2, c2); Assert.Equal(2, r2);
            Assert.True(FormulaBuilder.TryParseAddress("明细!C3", out var s2, out var cc1, out var rr1, out var cc2, out var rr2));
            Assert.Equal("明细", s2);
            Assert.Equal(cc2, cc1); Assert.Equal(rr2, rr1);
            Assert.False(FormulaBuilder.TryParseAddress("not an address", out _, out _, out _, out _, out _));
        }

        [Fact]
        public void ClampDataRows_pushes_first_row()
        {
            Assert.Equal("$A$2:$F$21", FormulaBuilder.ClampDataRows("$A$1:$F$21", 1));
            Assert.Equal("$A$11:$F$21", FormulaBuilder.ClampDataRows("$A$1:$F$21", 10));
        }

        [Fact]
        public void Rank_mixed_reference_and_order()
        {
            var f = FormulaBuilder.Build(With(O("rank", "$A$1:$C$21"), o => { o.ValueColumn = 3; }));
            Assert.Equal("=IFERROR(RANK.EQ($C2,$C$2:$C$21,0),0)", f); // $C2 混合引用：列绝对行相对，向下填充逐行联动
            var f2 = FormulaBuilder.Build(With(O("rank", "$A$1:$C$21"), o => { o.ValueColumn = 3; o.RankOrder = 1; }));
            Assert.Equal("=IFERROR(RANK.EQ($C2,$C$2:$C$21,1),0)", f2);
        }

        [Fact]
        public void Cumulative_sum_mixed_end_anchor()
        {
            var f = FormulaBuilder.Build(With(O("cumulative_sum", "$A$1:$C$21"), o => { o.ValueColumn = 3; }));
            Assert.Equal("=SUM($C$2:$C2)", f); // 首端绝对锁起点 + 尾端相对随填充联动 = 逐行累计
        }

        [Fact]
        public void Percent_of_total_row_over_sum()
        {
            var f = FormulaBuilder.Build(With(O("percent_of_total", "$A$1:$C$21"), o => { o.ValueColumn = 3; }));
            Assert.Equal("=IFERROR($C2/SUM($C$2:$C$21),0)", f); // 行相对：填充后逐行除以总和
        }

        [Fact]
        public void Median_and_stdev()
        {
            Assert.Equal("=MEDIAN($C$2:$C$21)", FormulaBuilder.Build(With(O("median", "$A$1:$C$21"), o => { o.ValueColumn = 3; })));
            Assert.Equal("=IFERROR(STDEV.P($C$2:$C$21),0)", FormulaBuilder.Build(With(O("stdev", "$A$1:$C$21"), o => { o.ValueColumn = 3; })));
        }

        [Fact]
        public void Growth_rate_diff_ratio_dual_range()
        {
            var o = With(O("growth_rate", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.SecondSourceAddress = "$B$1:$D$21"; x.SecondValueColumn = 2; });
            Assert.Equal("=IFERROR(($C2-$C2)/$C2,0)", FormulaBuilder.Build(o)); // 本期/上期均行相对，逐行同环比 // 第二区域 B 列起始 → 值列 C（同形）
            var o2 = With(O("diff", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.SecondSourceAddress = "$B$1:$D$21"; x.SecondValueColumn = 2; });
            Assert.Equal("=$C2-$C2", FormulaBuilder.Build(o2));
            var o3 = With(O("ratio", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.SecondSourceAddress = "$B$1:$D$21"; x.SecondValueColumn = 2; });
            Assert.Equal("=IFERROR($C2/$C2,0)", FormulaBuilder.Build(o3));
        }

        [Fact]
        public void Growth_rate_row_mismatch_fails_closed()
        {
            var o = With(O("growth_rate", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.SecondSourceAddress = "$B$1:$D$30"; x.SecondValueColumn = 2; });
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(o));
        }

        [Fact]
        public void Growth_rate_second_col_out_of_width_fails_closed()
        {
            var o = With(O("growth_rate", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.SecondSourceAddress = "$B$1:$D$21"; x.SecondValueColumn = 9; });
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(o));
        }

        [Fact]
        public void ParseIntent_sum_plain()
        {
            var intent = FormulaBuilder.ParseIntent("对金额求和", new[] { "月份", "品类", "金额" }, 3);
            Assert.NotNull(intent);
            Assert.Equal("sum", intent.Template);
        }

        [Fact]
        public void ParseIntent_sumifs_by_column_name()
        {
            var intent = FormulaBuilder.ParseIntent("按品类统计华东区的销售额", new[] { "品类", "月份", "销售额" }, 3);
            Assert.NotNull(intent);
            Assert.Equal("sumifs", intent.Template);
            Assert.Equal(1, intent.CriteriaColumn); // 「品类」在表头第 1 列
        }

        [Fact]
        public void ParseIntent_rank_and_cumulative_and_percent()
        {
            Assert.Equal("rank", FormulaBuilder.ParseIntent("按销售额排名", new[] { "月份", "销售额" }, 2).Template);
            Assert.Equal("cumulative_sum", FormulaBuilder.ParseIntent("算累计销售额", new[] { "月份", "销售额" }, 2).Template);
            Assert.Equal("percent_of_total", FormulaBuilder.ParseIntent("算每行占合计的占比", new[] { "月份", "销售额" }, 2).Template);
        }

        [Fact]
        public void ParseIntent_unknown_returns_null()
        {
            Assert.Null(FormulaBuilder.ParseIntent("帮我做个透视表", new[] { "月份", "金额" }, 2));
            Assert.Null(FormulaBuilder.ParseIntent("", new[] { "月份" }, 1));
        }

        [Fact]
        public void Column_out_of_source_width_fails_closed()
        {
            // 源宽 6 列（A:F），值列 10 越界 → 拒绝（实现者视角边界校验）
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(
                With(O("sum", "$A$1:$F$21"), o => { o.ValueColumn = 10; })));
        }

        [Fact]
        public void Header_rows_exceeding_total_rows_fails_closed()
        {
            // 区域只有 1 行而表头要 1 行 → 数据区为空 → 拒绝
            Assert.Throws<ArgumentException>(() => FormulaBuilder.Build(
                With(O("sum", "$A$1:$A$1"), o => { o.HeaderRows = 1; })));
        }

        [Fact]
        public void Criteria_cross_sheet_cellref_passes_through()
        {
            var o = With(O("sumifs", "$A$1:$C$21"), x => { x.ValueColumn = 3; x.CriteriaColumn = 1; x.Criteria = "参数表!$B$1"; });
            Assert.Equal("=SUMIFS($C$2:$C$21,$A$2:$A$21,参数表!$B$1)", FormulaBuilder.Build(o));
        }
    }
}
