using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 写入/清洗链 0 行输出复现（工具分类审查第一批，工具闭环）：
    /// FilterData 0 匹配 / CleanData 全空行 → 0 行矩阵 → Resize 的钳行行为 → 组合后源区域首行被清空的数据破坏链。
    /// </summary>
    public class ZeroRowWriteChainTests
    {
        [Fact]
        public void ApplyFilter_zero_matches_returns_zero_row_matrix()
        {
            var source = new object[3, 2]
            {
                { "华东", 100 },
                { "华南", 200 },
                { "华北", 300 }
            };
            var filtered = DataTransformService.ApplyFilter(source, 1, "equals", "西南");
            Assert.Equal(0, filtered.MatchedRows);
            Assert.Equal(0, filtered.Values.GetLength(0)); // 0 行矩阵
        }

        [Fact]
        public void RemoveBlankRows_all_blank_returns_zero_row_matrix()
        {
            var source = new object[2, 2]
            {
                { "", null },
                { "  ", "" }
            };
            var removed = DataTransformService.RemoveBlankRows(source);
            Assert.Equal(2, removed.ChangedCount);
            Assert.Equal(0, removed.Values.GetLength(0)); // 0 行矩阵
        }

        [Fact]
        public void Resize_zero_rows_must_not_produce_single_row_address()
        {
            // 工具闭环复现：0 行 Resize 曾被 Math.Max(rows-1,0) 钳成 1 行地址（A1:C10+0行 → A1:C1），
            // 后续 WriteRangeCore 对该 1 行目标写 0 行数组（Refit 清空语义）→ 源区域首行被清空（数据破坏）。
            // 修复后 0 行/0 列必须返回 null（fail-closed），由 FilterData/CleanData 提前拦截输出人话提示。
            Assert.Null(ExcelRangeHelper.Resize("A1:C10", 0, 3));
            Assert.Null(ExcelRangeHelper.Resize("A1:C10", 3, 0));
            Assert.Null(ExcelRangeHelper.Resize("A1:C10", 0, 0));
            // 正常行为回归锁定
            Assert.Equal("A1:C1", ExcelRangeHelper.Resize("A1:C10", 1, 3));
            Assert.Equal("A1:C10", ExcelRangeHelper.Resize("A1:C10", 10, 3));
        }
    }
}
