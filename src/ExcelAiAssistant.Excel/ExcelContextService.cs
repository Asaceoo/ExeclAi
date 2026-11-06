using System;
using ExcelAiAssistant.Core.Models;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    public sealed class ExcelContextService
    {
        private readonly dynamic application;

        public ExcelContextService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ExcelCurrentContext GetCurrentContext()
        {
            dynamic workbook = null;
            dynamic worksheet = null;
            dynamic selection = null;
            dynamic usedRange = null;

            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    throw new InvalidOperationException("当前没有打开的工作簿。");
                }

                worksheet = application.ActiveSheet;
                selection = application.Selection;
                usedRange = worksheet.UsedRange;

                return new ExcelCurrentContext
                {
                    ExcelVersion = application.Version,
                    Workbook = new ExcelWorkbookInfo
                    {
                        Name = workbook.Name,
                        FullPath = Convert.ToString(workbook.FullName),
                        IsReadOnly = Convert.ToBoolean(workbook.ReadOnly),
                        IsSaved = Convert.ToBoolean(workbook.Saved)
                    },
                    Worksheet = DescribeSheet(worksheet),
                    ActiveWorksheet = DescribeSheet(worksheet),
                    Selection = DescribeRange(selection),
                    UsedRange = DescribeRange(usedRange),
                    Protection = new ExcelProtectionInfo
                    {
                        WorkbookProtected = Convert.ToBoolean(workbook.ProtectStructure),
                        SheetProtected = Convert.ToBoolean(worksheet.ProtectContents)
                    },
                    CapturedAtUtc = DateTime.UtcNow
                };
            }
            finally
            {
                Release((object)usedRange);
                Release((object)selection);
                Release((object)worksheet);
                Release((object)workbook);
            }
        }

        private static ExcelSheetInfo DescribeSheet(dynamic worksheet)
        {
            dynamic usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                return new ExcelSheetInfo
                {
                    Name = worksheet.Name,
                    Index = Convert.ToInt32(worksheet.Index),
                    Visible = Convert.ToInt32(worksheet.Visible) == -1,
                    UsedRowCount = Convert.ToInt32(usedRange?.Rows.Count ?? 0),
                    UsedColumnCount = Convert.ToInt32(usedRange?.Columns.Count ?? 0)
                };
            }
            finally
            {
                Release((object)usedRange);
            }
        }

        private static ExcelRangeSummary DescribeRange(dynamic range)
        {
            if (range == null)
            {
                return null;
            }

            var rowCount = Convert.ToInt32(range.Rows.Count);
            var columnCount = Convert.ToInt32(range.Columns.Count);
            return new ExcelRangeSummary
            {
                Address = range.Address,
                RowCount = rowCount,
                ColumnCount = columnCount,
                CellCount = rowCount * columnCount
            };
        }

        internal static void Release(object comObject)
        {
            if (comObject == null || !System.Runtime.InteropServices.Marshal.IsComObject(comObject))
            {
                return;
            }

            // 真机实证（RibbonFixVerify P0）：Excel 对象模型里 Workbook/Worksheet 作为「获取器」返回并被
            // Excel 复用时，.NET 按 IUnknown 缓存为同一 RCW 包装——ActiveSheet、ActiveWorkbook、
            // Worksheets[x] 重复获取全部指向同一 RCW。对其中任意一个 FinalReleaseComObject 都会把共享 RCW
            // 与基础对象断开，累及其他持有者（不同工具层、嵌套 writer 各持同一 RCW），再访问即抛
            // 「COM 对象与其基础 RCW 分开后就不能再使用」。这正是「功能不能用」的根因。
            // 因此仅对实证为瞬态、每次调用返回独立新 RCW 的 Range 做强制释放；其余（Workbook/Worksheet/
            // Application/集合等共享缓存对象）一律交由 .NET 托管引用释放 + GC 终结器归还，杜绝跨层毒化。
            string typeName;
            try { typeName = comObject.GetType().FullName ?? string.Empty; }
            catch (Exception xaiEx) {
                AppLog.Warn("context", "ExcelContextService.cs:Release 静默兜底触发: " + xaiEx.Message); return; }
            if (typeName.IndexOf("Range", StringComparison.Ordinal) < 0)
            {
                return;
            }

            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject);
            }
            catch (System.Runtime.InteropServices.InvalidComObjectException)
            {
                AppLog.Warn("ExcelContextService", "[LogCover] ExcelContextService.cs:133  silent-catch covered (no exception variable)");
                // 幂等清理：同一瞬态 Range 被外层与嵌套 writer 各释放一次时，第二次 FinalRelease 已是空操作。
            }
        }
    }
}
