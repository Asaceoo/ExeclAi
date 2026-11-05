using ExcelAiAssistant.Core.Tools;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>数据智能填充 BlankFillAnalyzer 纯逻辑测试（可脱离 Excel 运行）。</summary>
    public sealed class BlankFillAnalyzerTests
    {
        [Fact]
        public void FindBlanks_only_matches_empty_cells()
        {
            var block = new object[3, 2];
            block[0, 0] = "名称"; block[0, 1] = "数值";
            block[1, 0] = "华东"; block[1, 1] = 12;
            block[2, 0] = null;      // 空白格
            block[2, 1] = "待填";    // 占位文本 → 非空白，不应视为空缺
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            Assert.Single(blanks);
            Assert.Equal(3, blanks[0].Row);
            Assert.Equal(1, blanks[0].Col);
            Assert.Equal("A3", blanks[0].CellAddress);
            Assert.Equal(BlankCellKind.EmptyString, blanks[0].Kind);
        }

        [Fact]
        public void FindBlanks_treats_dbnull_and_whitespace_as_blank()
        {
            var block = new object[3, 1];
            block[0, 0] = "列A";
            block[1, 0] = "";
            block[2, 0] = "   ";
            block[2, 0] = System.DBNull.Value; // 覆盖为 DBNull
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            Assert.Equal(2, blanks.Count);
        }

        [Fact]
        public void FindBlanks_skips_header_row_when_hasHeader()
        {
            var block = new object[2, 1];
            block[0, 0] = null; // 表头空缺，hasHeader 时不计入
            block[1, 0] = "ok";
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            Assert.Empty(blanks);
        }

        [Fact]
        public void GroupByColumn_attaches_header_and_samples()
        {
            var block = new object[3, 2];
            block[0, 0] = "地区"; block[0, 1] = "销量";
            block[1, 0] = "华东"; block[1, 1] = 100;
            block[2, 0] = "华南"; block[2, 1] = null; // 空缺格
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            var groups = BlankFillAnalyzer.GroupByColumn(block, blanks, hasHeader: true, sampleLimit: 5);
            Assert.Single(groups);
            Assert.True(groups.ContainsKey(2));
            Assert.Equal("销量", groups[2].ColumnName);
            Assert.Equal(2, groups[2].Column);
            Assert.Single(groups[2].Blanks);
            Assert.Contains(100, groups[2].Samples);
        }

        [Fact]
        public void GroupByColumn_groups_multiple_columns()
        {
            var block = new object[3, 3];
            block[0, 0] = "h1"; block[0, 1] = "h2"; block[0, 2] = "h3";
            block[1, 0] = null; block[1, 1] = 5; block[1, 2] = null;
            block[2, 0] = "x"; block[2, 1] = null; block[2, 2] = "y";
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            var groups = BlankFillAnalyzer.GroupByColumn(block, blanks, hasHeader: true);
            Assert.Equal(3, groups.Count);
            Assert.True(groups.ContainsKey(1) && groups.ContainsKey(2) && groups.ContainsKey(3));
        }

        [Fact]
        public void CellAddress_converts_row_and_column()
        {
            Assert.Equal("A1", BlankFillAnalyzer.CellAddress(1, 1));
            Assert.Equal("AB27", BlankFillAnalyzer.CellAddress(27, 28));
            Assert.Equal("Z100", BlankFillAnalyzer.CellAddress(100, 26));
        }

        [Fact]
        public void IsLikelyBlank_recognizes_empty_variants()
        {
            Assert.True(BlankFillAnalyzer.IsLikelyBlank(null));
            Assert.True(BlankFillAnalyzer.IsLikelyBlank(""));
            Assert.True(BlankFillAnalyzer.IsLikelyBlank("   "));
            Assert.True(BlankFillAnalyzer.IsLikelyBlank(System.DBNull.Value));
            Assert.False(BlankFillAnalyzer.IsLikelyBlank("待填"));
            Assert.False(BlankFillAnalyzer.IsLikelyBlank("?"));
            Assert.False(BlankFillAnalyzer.IsLikelyBlank(0));
            Assert.False(BlankFillAnalyzer.IsLikelyBlank(12.5));
        }

        [Fact]
        public void Max_blanks_paging_is_applied_at_service_layer()
        {
            // 纯逻辑不负责分页常量，仅验证 FindBlanks 全量返回，分页在 Service 层按 max_blanks 截断
            var block = new object[6, 1];
            block[0, 0] = "h";
            for (var r = 1; r < 6; r++) block[r, 0] = null;
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            Assert.Equal(5, blanks.Count);
        }

        // ---------- D2 整行邻近上下文 ----------
        [Fact]
        public void BuildRowContext_marks_blank_and_keeps_same_row_values()
        {
            var block = new object[3, 3];
            block[0, 0] = "区域"; block[0, 1] = "产品"; block[0, 2] = "销量";
            block[1, 0] = "华东"; block[1, 1] = "苹果"; block[1, 2] = null; // 空白格
            block[2, 0] = "华南"; block[2, 1] = "香蕉"; block[2, 2] = 50;
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            var ctx = BlankFillAnalyzer.BuildRowContext(block, blanks, hasHeader: true);
            Assert.Single(ctx);
            Assert.Equal(2, ctx[0].Row);
            Assert.Equal(3, ctx[0].Cells.Count);
            Assert.Equal("区域", ctx[0].Cells[0].ColumnName);
            Assert.Equal("华东", ctx[0].Cells[0].Value);
            Assert.False(ctx[0].Cells[0].IsBlank);
            Assert.True(ctx[0].Cells[2].IsBlank);
            Assert.Null(ctx[0].Cells[2].Value);
        }

        [Fact]
        public void BuildRowContext_groups_distinct_rows_once()
        {
            var block = new object[4, 2];
            block[0, 0] = "h1"; block[0, 1] = "h2";
            block[1, 0] = null; block[1, 1] = "a";   // row2 blanks col1
            block[2, 0] = "b"; block[2, 1] = null;    // row3 blanks col2
            block[3, 0] = null; block[3, 1] = null;   // row4 blanks both
            var blanks = BlankFillAnalyzer.FindBlanks(block, hasHeader: true);
            var ctx = BlankFillAnalyzer.BuildRowContext(block, blanks, hasHeader: true);
            Assert.Equal(3, ctx.Count); // rows 2,3,4
            Assert.Equal(2, ctx[0].Row);
            Assert.Equal(3, ctx[1].Row);
            Assert.Equal(4, ctx[2].Row);
        }

        [Fact]
        public void BuildRowContext_empty_when_no_blanks()
        {
            var block = new object[2, 1];
            block[0, 0] = "h";
            block[1, 0] = 7;
            var ctx = BlankFillAnalyzer.BuildRowContext(block, new System.Collections.Generic.List<BlankCell>(), hasHeader: true);
            Assert.Empty(ctx);
        }
    }
}