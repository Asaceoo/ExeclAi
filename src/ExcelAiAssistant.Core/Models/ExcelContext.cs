using System;

namespace ExcelAiAssistant.Core.Models
{
    public sealed class ExcelWorkbookInfo
    {
        public string Name { get; set; }

        public string FullPath { get; set; }

        public bool IsReadOnly { get; set; }

        public bool IsSaved { get; set; }
    }

    public sealed class ExcelSheetInfo
    {
        public string Name { get; set; }

        public int Index { get; set; }

        public bool Visible { get; set; }

        public int UsedRowCount { get; set; }

        public int UsedColumnCount { get; set; }
    }

    public sealed class ExcelRangeSummary
    {
        public string Address { get; set; }

        public int RowCount { get; set; }

        public int ColumnCount { get; set; }

        public int CellCount { get; set; }

        public object[,] SampleValues { get; set; }

        public string[] HeaderCandidates { get; set; } = new string[0];
    }

    public sealed class ExcelProtectionInfo
    {
        public bool WorkbookProtected { get; set; }

        public bool SheetProtected { get; set; }
    }

    public sealed class ExcelCurrentContext
    {
        public string ExcelVersion { get; set; }

        public ExcelWorkbookInfo Workbook { get; set; }

        public ExcelSheetInfo Worksheet { get; set; }

        public ExcelSheetInfo ActiveWorksheet { get; set; }

        public ExcelRangeSummary Selection { get; set; }

        public ExcelRangeSummary UsedRange { get; set; }

        public ExcelProtectionInfo Protection { get; set; }

        public DateTime CapturedAtUtc { get; set; } = DateTime.UtcNow;
    }
}
