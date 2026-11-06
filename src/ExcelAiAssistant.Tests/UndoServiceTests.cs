using System;
using System.IO;
using ExcelAiAssistant.Core.Operations;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class UndoServiceTests : IDisposable
    {
        private readonly string databasePath;
        private readonly OperationRepository repository;
        public UndoServiceTests()
        {
            databasePath = Path.Combine(Path.GetTempPath(), $"undo-{Guid.NewGuid():N}.db");
            repository = new OperationRepository(databasePath);
        }

        [Fact]
        public void Captures_snapshot_and_restores_when_current_matches_after_values()
        {
            var restoredValues = default(object[,]);
            var service = new UndoService(repository, snapshot =>
            {
                restoredValues = snapshot.BeforeValues;
                return true;
            });
            var operationId = service.Capture(
                "session",
                "workbook",
                @"C:\book.xlsx",
                "write_excel_range",
                new { address = "A1:B2" },
                Snapshot("A1:B2", Before(), After())).Id;

            var result = service.Undo(operationId, After());

            Assert.True(result.Restored);
            Assert.False(result.ConflictDetected);
            Assert.Equal("old", restoredValues[0, 0]);
            Assert.Equal(123D, (double)restoredValues[1, 1]);
            Assert.False(repository.Get(operationId).CanUndo);
        }

        [Fact]
        public void Reports_conflict_without_restore_when_current_changed()
        {
            var called = false;
            var service = new UndoService(repository, _ =>
            {
                called = true;
                return true;
            });
            var operationId = service.Capture(
                "session",
                "workbook",
                null,
                "write_excel_range",
                null,
                Snapshot("A1:B2", Before(), After())).Id;

            var current = (object[,])After().Clone();
            current[0, 0] = "someone else changed this";
            var result = service.Undo(operationId, current);

            Assert.True(result.ConflictDetected);
            Assert.False(result.Restored);
            Assert.False(called);
        }

        private static CellSnapshot Snapshot(string address, object[,] before, object[,] after)
        {
            return new CellSnapshot
            {
                Sheet = "Sheet1",
                Address = address,
                BeforeValues = before,
                AfterValues = after
            };
        }

        private static object[,] Before()
        {
            return new object[,] { { "old", null }, { null, 123 } };
        }

        private static object[,] After()
        {
            return new object[,] { { "new", null }, { null, 456 } };
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


