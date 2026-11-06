using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Data.SQLite;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Operations
{
    public sealed class OperationRepository
    {
        private readonly string connectionString;
        private readonly string databasePath;

        public OperationRepository(string databasePath)
        {
            this.databasePath = databasePath ?? throw new ArgumentNullException(nameof(databasePath));
            var directory = Path.GetDirectoryName(Path.GetFullPath(databasePath));
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            connectionString = $"Data Source={databasePath};Version=3;";
            Initialize();
        }

        public void Add(OperationRecord record)
        {
            if (record == null)
            {
                throw new ArgumentNullException(nameof(record));
            }

            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText = @"
INSERT INTO operations (
  id, session_id, workbook_hash, workbook_path, tool, turn_label, args_json, affected_json,
  undo_kind, payload_path, payload_json, before_digest, after_digest, can_undo, created_at_utc
) VALUES (
  @id, @session_id, @workbook_hash, @workbook_path, @tool, @turn_label, @args_json, @affected_json,
  @undo_kind, @payload_path, @payload_json, @before_digest, @after_digest, @can_undo, @created_at_utc
);";
            command.Parameters.AddWithValue("@id", record.Id);
            command.Parameters.AddWithValue("@session_id", record.SessionId ?? string.Empty);
            command.Parameters.AddWithValue("@workbook_hash", record.WorkbookHash ?? string.Empty);
            command.Parameters.AddWithValue("@workbook_path", record.WorkbookPath ?? string.Empty);
            command.Parameters.AddWithValue("@tool", record.Tool ?? string.Empty);
            command.Parameters.AddWithValue("@turn_label", record.TurnLabel ?? string.Empty);
            command.Parameters.AddWithValue("@args_json", record.ArgsJson ?? "{}");
            command.Parameters.AddWithValue("@affected_json", record.AffectedJson ?? "{}");
            command.Parameters.AddWithValue("@undo_kind", record.UndoKind ?? "cells");
            command.Parameters.AddWithValue("@payload_path", record.PayloadPath ?? string.Empty);
            command.Parameters.AddWithValue("@payload_json", record.PayloadJson ?? "{}");
            command.Parameters.AddWithValue("@before_digest", record.BeforeDigest ?? string.Empty);
            command.Parameters.AddWithValue("@after_digest", record.AfterDigest ?? string.Empty);
            command.Parameters.AddWithValue("@can_undo", record.CanUndo ? 1 : 0);
            command.Parameters.AddWithValue("@created_at_utc", record.CreatedAtUtc.ToString("O"));
            command.ExecuteNonQuery();
        }

        public IReadOnlyList<OperationRecord> List(OperationQuery query)
        {
            var effective = query ?? new OperationQuery();
            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText = @"
SELECT id, session_id, workbook_hash, workbook_path, tool, turn_label, args_json, affected_json,
       undo_kind, payload_path, payload_json, before_digest, after_digest, can_undo, created_at_utc
FROM operations
WHERE (@workbook_hash = '' OR workbook_hash = @workbook_hash)
ORDER BY created_at_utc DESC
LIMIT @limit;";
            command.Parameters.AddWithValue("@workbook_hash", effective.WorkbookHash ?? string.Empty);
            command.Parameters.AddWithValue("@limit", Math.Max(1, effective.Limit));

            var records = new List<OperationRecord>();
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                records.Add(new OperationRecord
                {
                    Id = reader.GetString(0),
                    SessionId = reader.IsDBNull(1) ? null : reader.GetString(1),
                    WorkbookHash = reader.IsDBNull(2) ? null : reader.GetString(2),
                    WorkbookPath = reader.IsDBNull(3) ? null : reader.GetString(3),
                    Tool = reader.GetString(4),
                    TurnLabel = reader.IsDBNull(5) ? null : reader.GetString(5),
                    ArgsJson = reader.IsDBNull(6) ? "{}" : reader.GetString(6),
                    AffectedJson = reader.IsDBNull(7) ? "{}" : reader.GetString(7),
                    UndoKind = reader.GetString(8),
                    PayloadPath = reader.IsDBNull(9) ? null : reader.GetString(9),
                    PayloadJson = reader.IsDBNull(10) ? "{}" : reader.GetString(10),
                    BeforeDigest = reader.IsDBNull(11) ? null : reader.GetString(11),
                    AfterDigest = reader.IsDBNull(12) ? null : reader.GetString(12),
                    CanUndo = reader.GetInt64(13) != 0,
                    CreatedAtUtc = DateTime.Parse(reader.GetString(14)).ToUniversalTime()
                });
            }

            return records.AsReadOnly();
        }

        public OperationRecord Get(string id)
        {
            return List(new OperationQuery { Limit = int.MaxValue })
                .FirstOrDefault(record => record.Id == id);
        }

        public void MarkNotUndoable(string id)
        {
            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE operations SET can_undo = 0 WHERE id = @id;";
            command.Parameters.AddWithValue("@id", id);
            command.ExecuteNonQuery();
        }

        public void KeepLatest(string workbookHash, int count)
        {
            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText =
                "DELETE FROM operations WHERE workbook_hash = @workbook_hash AND id IN (" +
                "SELECT id FROM operations WHERE workbook_hash = @workbook_hash " +
                "ORDER BY created_at_utc DESC LIMIT -1 OFFSET @count);";
            command.Parameters.AddWithValue("@workbook_hash", workbookHash ?? string.Empty);
            command.Parameters.AddWithValue("@count", Math.Max(0, count));
            command.ExecuteNonQuery();
        }

        public void DeleteOlderThan(DateTime cutoffUtc)
        {
            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText = "DELETE FROM operations WHERE created_at_utc < @cutoff;";
            command.Parameters.AddWithValue("@cutoff", cutoffUtc.ToString("O"));
            command.ExecuteNonQuery();
        }

        private SQLiteConnection Open()
        {
            var connection = new SQLiteConnection(connectionString);
            connection.Open();
            return connection;
        }

        private void Initialize()
        {
            using var connection = Open();
            using var command = connection.CreateCommand();
            command.CommandText = @"
CREATE TABLE IF NOT EXISTS operations (
  id TEXT PRIMARY KEY,
  session_id TEXT NOT NULL,
  workbook_hash TEXT NOT NULL,
  workbook_path TEXT,
  tool TEXT NOT NULL,
  turn_label TEXT,
  args_json TEXT NOT NULL,
  affected_json TEXT NOT NULL,
  undo_kind TEXT NOT NULL,
  payload_path TEXT,
  payload_json TEXT,
  before_digest TEXT,
  after_digest TEXT,
  can_undo INTEGER NOT NULL DEFAULT 1,
  created_at_utc TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_operations_workbook_created
  ON operations(workbook_hash, created_at_utc DESC);";
            command.ExecuteNonQuery();

            EnsureColumn(connection, "operations", "turn_label");
        }

        /// <summary>兼容旧库：若缺少新列则 ALTER TABLE 补上。</summary>
        private static void EnsureColumn(SQLiteConnection connection, string table, string column)
        {
            bool exists = false;
            using (var check = connection.CreateCommand())
            {
                check.CommandText = $"PRAGMA table_info({table});";
                using var reader = check.ExecuteReader();
                while (reader.Read())
                {
                    if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase))
                    {
                        exists = true;
                        break;
                    }
                }
            }

            if (!exists)
            {
                using var alter = connection.CreateCommand();
                alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} TEXT;";
                alter.ExecuteNonQuery();
            }
        }
    }
}
