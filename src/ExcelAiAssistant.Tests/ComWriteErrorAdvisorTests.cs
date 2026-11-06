using System;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>写入路径 COM 错误（0x800A03EC 家族）可行动归因（ComWriteErrorAdvisor）。
    /// 背景：用户真机自检实证 write_excel_range 报裸 HRESULT，AI/用户无从下手（v3.7.1.59）。</summary>
    public class ComWriteErrorAdvisorTests
    {
        private static ComWriteErrorAdvisor.WriteContext Ctx(
            bool prot = false, bool ro = false, bool merged = false,
            bool array = false, int visible = -1)
        {
            return new ComWriteErrorAdvisor.WriteContext
            {
                Address = "$B$2:$D$9",
                SheetName = "Sheet1",
                SheetProtected = prot,
                WorkbookReadOnly = ro,
                TargetHasMergedCells = merged,
                TargetHasArrayFormula = array,
                SheetVisibleState = visible
            };
        }

        [Fact]
        public void Protected_sheet_wins_highest_priority()
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx(prot: true, merged: true));
            Assert.Contains("保护", advice);
            Assert.Contains("verify_protection", advice);
        }

        [Fact]
        public void Read_only_workbook_is_recognized()
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx(ro: true));
            Assert.Contains("只读", advice);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(2)]
        public void Hidden_sheet_is_recognized(int visibleState)
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx(visible: visibleState));
            Assert.Contains("隐藏", advice);
            Assert.Contains("取消隐藏", advice);
        }

        [Fact]
        public void Merged_cells_get_merge_advice()
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx(merged: true));
            Assert.Contains("合并", advice);
        }

        [Fact]
        public void Array_formula_block_gets_array_advice()
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx(array: true));
            Assert.Contains("数组公式", advice);
        }

        [Fact]
        public void No_signal_falls_back_to_ranked_checklist()
        {
            var advice = ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), Ctx());
            Assert.Contains("preview_excel_range", advice);
        }

        [Fact]
        public void Other_hresult_returns_null_for_caller_fallback()
        {
            Assert.Null(ComWriteErrorAdvisor.Advise(unchecked((int)0x8007139F), Ctx(prot: true)));
        }

        [Fact]
        public void Null_context_is_safe()
        {
            Assert.Null(ComWriteErrorAdvisor.Advise(unchecked((int)0x800A03EC), null));
        }
    }
}
