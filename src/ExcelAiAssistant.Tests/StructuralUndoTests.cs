using System;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Operations;
using Newtonsoft.Json;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class StructuralUndoTests : IDisposable
    {
        private readonly string databasePath;
        private readonly OperationRepository repository;
        private readonly UndoService service;

        public StructuralUndoTests()
        {
            databasePath = Path.Combine(Path.GetTempPath(), $"struct-undo-{Guid.NewGuid():N}.db");
            repository = new OperationRepository(databasePath);
            service = new UndoService(repository, _ => true);
        }

        [Fact]
        public void Sheet_insert_payload_is_persisted_and_round_trips()
        {
            var operationId = service.CaptureStructural(
                "session",
                "workbook",
                @"C:\book.xlsx",
                "import_excel_file_sheet",
                new { file_path = @"C:\data.xlsx" },
                "sheet-insert",
                new { sheets = new[] { "数据A", "数据B" } },
                new SheetInsertUndoPayload { Sheets = new[] { "数据A", "数据B" } }).Id;

            var record = repository.Get(operationId);
            Assert.Equal("sheet-insert", record.UndoKind);
            Assert.True(record.CanUndo);
            Assert.Equal("import_excel_file_sheet", record.Tool);
            Assert.Equal("workbook", record.WorkbookHash);

            var payload = JsonConvert.DeserializeObject<SheetInsertUndoPayload>(record.PayloadJson);
            Assert.NotNull(payload);
            Assert.Equal(2, payload.Count);
            Assert.Equal("数据A", payload.Sheets[0]);
            Assert.Equal("数据B", payload.Sheets[1]);

            var affected = JsonConvert.DeserializeObject<dynamic>(record.AffectedJson);
            Assert.Equal(2, (int)affected.sheets.Count);
            Assert.Equal("数据B", (string)affected.sheets[1]);
        }

        [Fact]
        public void Dimension_insert_payload_is_persisted()
        {
            var operationId = service.CaptureStructural(
                "session",
                "workbook",
                @"C:\book.xlsx",
                "insert_excel_dimension",
                new { address = "3", dimension = "row", quantity = 2 },
                "dimension-insert",
                new { sheet = "Sheet1", dimension = "row", address = "3", quantity = 2 },
                new DimensionInsertUndoPayload
                {
                    Sheet = "Sheet1",
                    Dimension = "row",
                    Address = "3",
                    Count = 2
                }).Id;

            var record = repository.Get(operationId);
            Assert.Equal("dimension-insert", record.UndoKind);
            var payload = JsonConvert.DeserializeObject<DimensionInsertUndoPayload>(record.PayloadJson);
            Assert.NotNull(payload);
            Assert.Equal("row", payload.Dimension);
            Assert.Equal("3", payload.Address);
            Assert.Equal(2, payload.Count);
        }

        [Fact]
        public void Structural_record_is_listed_and_sorted_with_cells_records()
        {
            var cellsId = service.Capture(
                "session",
                "workbook",
                @"C:\book.xlsx",
                "write_excel_range",
                null,
                new CellSnapshot { Sheet = "Sheet1", Address = "A1", BeforeValues = new object[,] { { "old" } }, AfterValues = new object[,] { { "new" } } }).Id;

            var sheetId = service.CaptureStructural(
                "session",
                "workbook",
                @"C:\book.xlsx",
                "import_excel_file_sheet",
                null,
                "sheet-insert",
                new { sheets = new[] { "导入" } },
                new SheetInsertUndoPayload { Sheets = new[] { "导入" } }).Id;

            var records = repository.List(new OperationQuery { WorkbookHash = "workbook" });
            var ids = records.Select(r => r.Id).ToList();
            Assert.Contains(cellsId, ids);
            Assert.Contains(sheetId, ids);
            Assert.Equal(2, records.Count);
            Assert.Contains(records, r => r.UndoKind == "sheet-insert");
        }

        [Fact]
        public void Capture_structural_rejects_blank_undo_kind()
        {
            Assert.Throws<ArgumentException>(() => service.CaptureStructural(
                "session",
                "workbook",
                null,
                "tool",
                null,
                null,
                null,
                null));
        }

        public void Dispose()
        {
            if (File.Exists(databasePath))
            {
                File.Delete(databasePath);
            }
        }
    }
}