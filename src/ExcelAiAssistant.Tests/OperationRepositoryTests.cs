using System;
using System.IO;
using ExcelAiAssistant.Core.Operations;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class OperationRepositoryTests : IDisposable
    {
        private readonly string databasePath;
        private readonly OperationRepository repository;

        public OperationRepositoryTests()
        {
            databasePath = Path.Combine(Path.GetTempPath(), $"operations-{Guid.NewGuid():N}.db");
            repository = new OperationRepository(databasePath);
        }

        [Fact]
        public void Adds_and_lists_records_newest_first()
        {
            repository.Add(new OperationRecord
            {
                Id = "older",
                WorkbookHash = "wb",
                Tool = "write_excel_range",
                CreatedAtUtc = DateTime.UtcNow.AddMinutes(-1)
            });
            repository.Add(new OperationRecord
            {
                Id = "newer",
                WorkbookHash = "wb",
                Tool = "format_excel_range"
            });

            var records = repository.List(new OperationQuery { WorkbookHash = "wb" });

            Assert.Equal(2, records.Count);
            Assert.Equal("newer", records[0].Id);
            Assert.Equal("older", records[1].Id);
        }

        [Fact]
        public void Marks_record_not_undoable()
        {
            var record = new OperationRecord();
            repository.Add(record);

            repository.MarkNotUndoable(record.Id);

            Assert.False(repository.Get(record.Id).CanUndo);
        }

        [Fact]
        public void Keeps_only_latest_records_for_workbook()
        {
            for (var index = 0; index < 3; index++)
            {
                repository.Add(new OperationRecord
                {
                    Id = $"record-{index}",
                    WorkbookHash = "wb",
                    CreatedAtUtc = DateTime.UtcNow.AddMinutes(-index)
                });
            }

            repository.KeepLatest("wb", 2);

            Assert.Equal(2, repository.List(new OperationQuery { WorkbookHash = "wb" }).Count);
            Assert.NotNull(repository.Get("record-0"));
            Assert.Null(repository.Get("record-2"));
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
