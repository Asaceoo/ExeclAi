using System;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Core.Diagnostics
{
    /// <summary>
    /// 写入路径 COM 错误（0x800A03EC 家族）的可行动归因：纯逻辑可单测。
    /// 背景（v3.7.1.59）：用户真机自检实证 write_excel_range 报裸 "EXCEL_COM_ERROR HRESULT 0x800A03EC"，
    /// AI 与用户都不知道下一步该干什么。本归因器把 HRESULT + 写入上下文探测结果翻译成按优先级排列的人话建议。
    /// COM 对象探测由 Excel 层完成（探测纪律：零 FinalRelease、逐项 try-try-catch 失败即跳过），本类只做纯判定。
    /// </summary>
    public static class ComWriteErrorAdvisor
    {
        /// <summary>写入上下文探测结果（探测失败的项保持默认值，不阻断归因）。</summary>
        public sealed class WriteContext
        {
            public string Address { get; set; }
            public string SheetName { get; set; }
            public bool SheetProtected { get; set; }
            public bool WorkbookReadOnly { get; set; }
            public bool TargetHasMergedCells { get; set; }
            public bool TargetHasArrayFormula { get; set; }
            /// <summary>sheet.Visible 原始值：-1 可见 / 0 隐藏 / 2 深度隐藏。</summary>
            public int SheetVisibleState { get; set; } = -1;
        }

        /// <summary>返回可行动建议；无法归因（HRESULT 不在覆盖面且上下文无信号）返回 null，由调用方落通用话术。</summary>
        public static string Advise(int hResult, WriteContext ctx)
        {
            if (ctx == null)
            {
                return null;
            }

            const int hrObjectError = unchecked((int)0x800A03EC);
            if (hResult != hrObjectError)
            {
                return null;
            }

            // 按概率与可操作性排序：保护 > 只读 > 隐藏 > 合并 > 数组公式 > 兜底清单。
            if (ctx.SheetProtected)
            {
                return "目标工作表处于保护状态，写入被 Excel 拒绝：请先解除保护（审阅 → 撤销工作表保护），或先用 verify_protection 探测保护范围后改写非保护区域。";
            }

            if (ctx.WorkbookReadOnly)
            {
                return "工作簿当前为只读（可能以只读打开/被另一进程占用/受保护视图）：请先另存为可写副本或关闭占用它的 Excel 窗口后重试。";
            }

            if (ctx.SheetVisibleState == 0 || ctx.SheetVisibleState == 2)
            {
                return "目标工作表处于隐藏状态（Visible=" + ctx.SheetVisibleState + "）：隐藏表上的直接写入会被拒绝，请先取消隐藏（右键工作表标签 → 取消隐藏）再写入。";
            }

            if (ctx.TargetHasMergedCells)
            {
                return "目标区域含合并单元格，部分格不可独立写入：请取消合并（解除单元格合并），或把写入对齐到合并区左上角首格。";
            }

            if (ctx.TargetHasArrayFormula)
            {
                return "目标区域含数组公式（CSE/溢出区），不允许部分单元格改写：请整区删除数组公式后再写，或改写到其他空白区域。";
            }

            return "Excel 拒绝该写入（0x800A03EC），常见原因：①目标地址越界或含合并格；②工作表被保护；③目标被数组公式占用；④Excel 当前有未关闭的弹窗/编辑态。建议先用 preview_excel_range 确认目标区域可访问，再缩小范围重试。";
        }
    }
}
