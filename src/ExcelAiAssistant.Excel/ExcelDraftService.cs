using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// P1-1 草稿模式（对齐 Univer worktree collaboration）：
    /// AI 在隔离草稿副本上执行变更 → diff 预览 → 人工采纳/放弃。
    /// 会话状态内存持有（草稿生命周期 = 单会话内）。
    /// </summary>
    public sealed class ExcelDraftService
    {
        private readonly dynamic application;
        private readonly Dictionary<string, DraftSession> drafts = new Dictionary<string, DraftSession>(StringComparer.OrdinalIgnoreCase);

        private sealed class DraftSession
        {
            public string DraftSheet;
            public string SourceSheet;
        }

        public ExcelDraftService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        private dynamic GetWorkbook() => application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");

        /// <summary>start_draft：复制源 sheet 为隐藏草稿，记录源快照。</summary>
        public ToolResult StartDraft(Dictionary<string, object> args)
        {
            var source = Required(args, "sheet");
            dynamic workbook = null;
            try
            {
                workbook = GetWorkbook();
                dynamic src;
                try { src = workbook.Worksheets[source]; }
                catch { return new ToolResult { Ok = false, ErrorCode = "INVALID_SHEET", Data = "工作表不存在：" + source }; }

                // FIX：毫秒级命名防同秒撞名；源名过长时截断防 31 字符上限（审查询 D-1/C-4）
                var suffix = DateTime.Now.ToString("HHmmssfff");
                var baseName = source.Length > 18 ? source.Substring(0, 18) : source;
                var draftName = "draft_" + baseName + "_" + suffix;
                if (SheetExists(workbook, draftName))
                    return new ToolResult { Ok = false, ErrorCode = "SHEET_EXISTS", Data = "草稿已存在：" + draftName };

                var after = (dynamic)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                after.Name = draftName;
                // FIX（审查 C-1/C-3 + impl E2/E3）：隐藏草稿 + 按 UsedRange 原点对齐 + Formula 复制保公式
                try { after.Visible = 0; } catch { } // xlSheetHidden
                var used = src.UsedRange;
                int rows = Convert.ToInt32(used.Rows.Count);
                int cols = Convert.ToInt32(used.Columns.Count);
                int originRow = Convert.ToInt32(used.Row);
                int originCol = Convert.ToInt32(used.Column);
                if (rows > 0 && cols > 0)
                {
                    // 目标区域从同一原点开始（不再固定 A1）——避免复制后错位与回写重复
                    dynamic target = (dynamic)after.Range[after.Cells[originRow, originCol], after.Cells[originRow + rows - 1, originCol + cols - 1]];
                    try { target.Formula = rows * cols == 1 ? (object)used.Formula : used.Formula; }
                    catch { target.Value2 = rows * cols == 1 ? (object)used.Value2 : used.Value2; }
                    try { target.NumberFormat = rows * cols == 1 ? (object)used.NumberFormat : used.NumberFormat; } catch { }
                }

                drafts[draftName] = new DraftSession { DraftSheet = draftName, SourceSheet = source };
                return new ToolResult
                {
                    Ok = true,
                    Data = new { draft_sheet = draftName, source_sheet = source, hint = "在草稿 sheet 上执行后续变更，完成后调 diff_draft 预览，apply_draft 采纳或 discard_draft 放弃。" },
                    AffectedRanges = new List<string>()
                };
            }
            finally { }
        }

        /// <summary>diff_draft：对比草稿与源，产出单元格级差异列表。</summary>
        public ToolResult DiffDraft(Dictionary<string, object> args)
        {
            var draftName = Required(args, "draft_sheet");
            if (!drafts.TryGetValue(draftName, out var session))
                return new ToolResult { Ok = false, ErrorCode = "DRAFT_NOT_FOUND", Data = "草稿不存在或已结束：" + draftName };
            dynamic workbook = null;
            try
            {
                workbook = GetWorkbook();
                dynamic draft = workbook.Worksheets[draftName];
                dynamic src = workbook.Worksheets[session.SourceSheet];
                var diffs = new List<object>();
                // FIX（审查 C-2）：以源表 UsedRange 为比较域（草稿未覆盖的源数据也算变更），
                // 并同时比 Formula 与 Value（impl E2：公式改了值未变也要报差异）。
                var srcUsed = src.UsedRange;
                int rows = Convert.ToInt32(srcUsed.Rows.Count);
                int cols = Convert.ToInt32(srcUsed.Columns.Count);
                int originRow = Convert.ToInt32(srcUsed.Row);
                int originCol = Convert.ToInt32(srcUsed.Column);
                object srcVals = rows * cols == 1 ? (object)srcUsed.Value2 : srcUsed.Value2;
                object srcFormula = rows * cols == 1 ? (object)srcUsed.Formula : srcUsed.Formula;
                dynamic draftRange = draft.Range[draft.Cells[originRow, originCol], draft.Cells[originRow + rows - 1, originCol + cols - 1]];
                object draftVals = rows * cols == 1 ? (object)draftRange.Value2 : draftRange.Value2;
                object draftFormula = rows * cols == 1 ? (object)draftRange.Formula : draftRange.Formula;
                int changed = 0;
                for (int r = 1; r <= rows && changed < 500; r++)
                {
                    for (int c = 1; c <= cols && changed < 500; c++)
                    {
                        object dv = rows * cols == 1 ? draftVals : SafeGet(draftVals, r, c);
                        object sv = rows * cols == 1 ? srcVals : SafeGet(srcVals, r, c);
                        object df = rows * cols == 1 ? draftFormula : SafeGet(draftFormula, r, c);
                        object sf = rows * cols == 1 ? srcFormula : SafeGet(srcFormula, r, c);
                        // 公式优先：任一为公式且公式串不同 → 视为变更
                        var formulaChanged = !Equals(FormatVal(df), FormatVal(sf));
                        var valueChanged = !Equals(FormatVal(dv), FormatVal(sv));
                        if (formulaChanged || valueChanged)
                        {
                            diffs.Add(new { cell = ColName(originCol + c - 1) + (originRow + r - 1), before = FormatVal(sv), after = FormatVal(dv), before_formula = FormatVal(sf), after_formula = FormatVal(df) });
                            changed++;
                        }
                    }
                }
                return new ToolResult { Ok = true, Data = new { draft_sheet = draftName, source_sheet = session.SourceSheet, changed_cells = diffs.Count, diffs = diffs, note = diffs.Count == 0 ? "草稿与源一致（无变更）。" : "请向用户展示差异后调 apply_draft 或 discard_draft。" } };
            }
            finally { }
        }

        /// <summary>apply_draft：草稿值复制回源 sheet，删草稿。走覆盖审计（受 AffectedRanges）。</summary>
        public ToolResult ApplyDraft(Dictionary<string, object> args)
        {
            var draftName = Required(args, "draft_sheet");
            if (!drafts.TryGetValue(draftName, out var session))
                return new ToolResult { Ok = false, ErrorCode = "DRAFT_NOT_FOUND", Data = "草稿不存在或已结束：" + draftName };
            dynamic workbook = null;
            try
            {
                workbook = GetWorkbook();
                dynamic draft = workbook.Worksheets[draftName];
                dynamic src = workbook.Worksheets[session.SourceSheet];
                // FIX（impl E2/E3 + 审查 E4）：按源表原点回写；用 Formula 保公式；AffectedRanges 精确到列末
                var srcUsed = src.UsedRange;
                int rows = Convert.ToInt32(srcUsed.Rows.Count);
                int cols = Convert.ToInt32(srcUsed.Columns.Count);
                int originRow = Convert.ToInt32(srcUsed.Row);
                int originCol = Convert.ToInt32(srcUsed.Column);
                if (rows > 0 && cols > 0)
                {
                    dynamic draftRange = draft.Range[draft.Cells[originRow, originCol], draft.Cells[originRow + rows - 1, originCol + cols - 1]];
                    dynamic target = (dynamic)src.Range[src.Cells[originRow, originCol], src.Cells[originRow + rows - 1, originCol + cols - 1]];
                    object formulaData = rows * cols == 1 ? (object)draftRange.Formula : draftRange.Formula;
                    try { target.Formula = formulaData; }
                    catch { target.Value2 = rows * cols == 1 ? (object)draftRange.Value2 : draftRange.Value2; }
                }
                draft.Delete();
                drafts.Remove(draftName);
                return new ToolResult
                {
                    Ok = true,
                    Data = new { applied = draftName, source_sheet = session.SourceSheet, range = "'" + session.SourceSheet + "'!A1:" + ColName(cols) + rows },
                    AffectedRanges = new List<string> { "'" + session.SourceSheet + "'!A1:" + ColName(cols) + rows }
                };
            }
            finally { }
        }

        /// <summary>discard_draft：删除草稿不做任何修改。</summary>
        public ToolResult DiscardDraft(Dictionary<string, object> args)
        {
            var draftName = Required(args, "draft_sheet");
            if (!drafts.TryGetValue(draftName, out var session))
                return new ToolResult { Ok = false, ErrorCode = "DRAFT_NOT_FOUND", Data = "草稿不存在或已结束：" + draftName };
            dynamic workbook = null;
            try
            {
                workbook = GetWorkbook();
                workbook.Worksheets[draftName].Delete();
                drafts.Remove(draftName);
                return new ToolResult { Ok = true, Data = new { discarded = draftName } };
            }
            finally { }
        }

        // ===== helpers =====
        private static object SafeGet(object arr, int r, int c)
        {
            try { return ((object[,])arr).GetValue(r, c); } catch { return null; }
        }
        private static string FormatVal(object v) => v == null ? "" : Convert.ToString(v);
        private static string ColName(int c) { var s = ""; while (c > 0) { var m = (c - 1) % 26; s = (char)('A' + m) + s; c = (c - 1) / 26; } return s; }
        private bool SheetExists(dynamic workbook, string name)
        {
            foreach (var s in workbook.Worksheets)
                if (string.Equals(Convert.ToString(s.Name), name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
        private static string Required(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                var s = Convert.ToString(v);
                if (s.Length > 0) return s;
            }
            throw new ArgumentException("缺少必填参数：" + key);
        }
    }
}
