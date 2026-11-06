using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 需求①「看板选区」：staging KPI 公式 → 明细数据源解析（ParseSourceFromFormula）单测。
    /// 覆盖 Excel 公式规范化四种形态（踩坑手册 §9）+ staging 自身族排除 + 非公式拒绝。
    /// </summary>
    public sealed class DashboardSourceParseTests
    {
        [Fact]
        public void Sumifs_with_quoted_underline_sheet_parses_source()
        {
            // 写入形态：明细表名下划线开头带引号（'_明细'!）
            const string f = "=SUMIFS('_明细'!$D$2:$D$19,'_明细'!$A$2:$A$19,看板!$B$2)";
            Assert.Equal("_明细!$D$2:$D$19", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Sumifs_bare_sheet_parses_source()
        {
            const string f = "=SUMIFS(明细!$D$2:$D$19,明细!$A$2:$A$19,B2)";
            Assert.Equal("明细!$D$2:$D$19", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Sumifs_normalized_stripped_underline_parses_source()
        {
            // Excel 规范化后：'_明细'! 读回为 明细!（剥引号剥首下划线）
            const string f = "=SUMIFS(明细!$D$2:$D$19,明细!$A$2:$A$19,看板!$B$2)";
            Assert.Equal("明细!$D$2:$D$19", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Space_in_sheet_name_keeps_quotes_strips_underline()
        {
            // 带空格表名（auto_rename 递增）：'名 2'! 保留引号
            const string f = "=SUMIFS('数 据'!$C$2:$C$9,'数 据'!$A$2:$A$9,B3)";
            Assert.Equal("数 据!$C$2:$C$9", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Self_staging_reference_is_skipped()
        {
            // staging 自身族（_X_calc / '名 2_calc'）引用应被跳过，继续找真正的明细引用
            const string f = "=IFERROR('_多维交叉看板_calc'!$B$3/明细!$D$2,0)";
            Assert.Equal("明细!$D$2", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Pure_self_reference_returns_null()
        {
            // 全部引用都是 staging 自身（pnl/target 完成率公式）→ null
            const string f = "=IFERROR(B2/C2,0)";
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Bare_staging_calc_reference_is_skipped()
        {
            // 第四形态（真机 T27b 教训）：裸名规范化 X_calc!B2（无下划线无空格）同属辅助表族，须跳过
            const string f = "=多维交叉看板_calc!B2";
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Staging_family_skipped_then_detail_found()
        {
            // 看板明细行公式：先引 staging（跳过），再无其他 → null（pnl/target 兜底语义）
            const string f = "=IFERROR(多维交叉看板_calc!B2,\"-\")";
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Board_self_reference_skipped_with_exclude_set()
        {
            // 真机 T27b 二轮教训：dim_cross staging A1 镜像看板 B3（裸名无 _calc 后缀），
            // 需排除集上下文才能识别为内部引用
            var exclude = new System.Collections.Generic.HashSet<string> { "多维交叉看板" };
            const string f = "=多维交叉看板!$B$3";
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(f, exclude));
            // 无排除集时仍返回（保持纯静态语义可用）
            Assert.Equal("多维交叉看板!$B$3", ExcelDashboardService.ParseSourceFromFormula(f));
        }

        [Fact]
        public void Non_formula_returns_null()
        {
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(null));
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula(string.Empty));
            Assert.Null(ExcelDashboardService.ParseSourceFromFormula("科目"));
        }
    }
}
