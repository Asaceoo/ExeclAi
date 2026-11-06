using System;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Operations
{
    public sealed class UndoResult
    {
        public bool Restored { get; set; }

        public bool ConflictDetected { get; set; }

        public string Message { get; set; }
    }

    public sealed class UndoService
    {
        private readonly OperationRepository repository;
        private readonly Func<CellSnapshot, bool> restoreSnapshot;

        /// <summary>由宿主注入的对话轮次标签，记录操作时写入每条记录。</summary>
        public string TurnLabel { get; set; }

        public UndoService(
            OperationRepository repository,
            Func<CellSnapshot, bool> restoreSnapshot)
        {
            this.repository = repository ?? throw new ArgumentNullException(nameof(repository));
            this.restoreSnapshot = restoreSnapshot ?? throw new ArgumentNullException(nameof(restoreSnapshot));
        }

        public OperationRecord Capture(
            string sessionId,
            string workbookHash,
            string workbookPath,
            string tool,
            object args,
            CellSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            var record = new OperationRecord
            {
                SessionId = sessionId,
                WorkbookHash = workbookHash,
                WorkbookPath = workbookPath,
                Tool = tool,
                TurnLabel = TurnLabel,
                ArgsJson = JsonConvert.SerializeObject(args ?? new object()),
                AffectedJson = JsonConvert.SerializeObject(new
                {
                    sheet = snapshot.Sheet,
                    address = snapshot.Address
                }),
                UndoKind = "cells",
                PayloadJson = JsonConvert.SerializeObject(snapshot),
                BeforeDigest = Digest(snapshot.BeforeValues),
                AfterDigest = Digest(snapshot.AfterValues)
            };
            repository.Add(record);
            return record;
        }

        /// <summary>
        /// 记录非单元格类（结构化）撤销，例如工作表插入、行列插入。
        /// 恢复动作由 Excel 宿主按 UndoKind + PayloadJson 分派执行。
        /// </summary>
        public OperationRecord CaptureStructural(
            string sessionId,
            string workbookHash,
            string workbookPath,
            string tool,
            object args,
            string undoKind,
            object affected,
            object payload)
        {
            if (string.IsNullOrWhiteSpace(undoKind))
            {
                throw new ArgumentException("undoKind 不能为空。", nameof(undoKind));
            }

            var record = new OperationRecord
            {
                SessionId = sessionId,
                WorkbookHash = workbookHash,
                WorkbookPath = workbookPath,
                Tool = tool,
                TurnLabel = TurnLabel,
                ArgsJson = JsonConvert.SerializeObject(args ?? new object()),
                AffectedJson = JsonConvert.SerializeObject(affected ?? new object()),
                UndoKind = undoKind,
                PayloadJson = JsonConvert.SerializeObject(payload ?? new object())
            };
            repository.Add(record);
            return record;
        }

        public UndoResult Undo(string operationId, object[,] currentValues)
        {
            var record = repository.Get(operationId);
            if (record == null)
            {
                return new UndoResult { Message = "操作记录不存在。" };
            }

            if (!record.CanUndo)
            {
                return new UndoResult { Message = "该操作已不可回滚。" };
            }

            var snapshot = JsonConvert.DeserializeObject<CellSnapshot>(record.PayloadJson);
            if (snapshot != null)
            {
                snapshot.BeforeValues = NormalizeMatrix(snapshot.BeforeValues);
                snapshot.AfterValues = NormalizeMatrix(snapshot.AfterValues);
                snapshot.BeforeFormulas = NormalizeMatrix(snapshot.BeforeFormulas);
                snapshot.AfterFormulas = NormalizeMatrix(snapshot.AfterFormulas);
            }

            var expected = Digest(snapshot.AfterValues);
            var current = Digest(currentValues);
            if (!string.Equals(expected, current, StringComparison.Ordinal) &&
                !string.Equals(record.AfterDigest, current, StringComparison.Ordinal))
            {
                return new UndoResult
                {
                    ConflictDetected = true,
                    Message = "目标区域已被后续修改，回滚会覆盖当前值。"
                };
            }

            var restored = restoreSnapshot(snapshot);
            if (restored)
            {
                repository.MarkNotUndoable(record.Id);
            }

            return new UndoResult
            {
                Restored = restored,
                Message = restored ? "回滚完成。" : "恢复快照失败。"
            };
        }

        private static object[,] NormalizeMatrix(object[,] values)
        {
            if (values == null)
            {
                return null;
            }

            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    values[row, column] = NormalizeValue(values[row, column]);
                }
            }

            return values;
        }

        private static object NormalizeValue(object value)
        {
            return value switch
            {
                sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal =>
                    Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture),
                _ => value
            };
        }

        private static string Digest(object[,] values)
        {
            if (values == null)
            {
                using var nullHash = SHA256.Create();
                return Hex(nullHash.ComputeHash(Encoding.UTF8.GetBytes("null")));
            }

            using var sha256 = SHA256.Create();
            using var stream = new System.IO.MemoryStream();
            using var writer = new System.IO.BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
            writer.Write(values.GetLength(0));
            writer.Write(values.GetLength(1));
            for (var row = 0; row < values.GetLength(0); row++)
            {
                for (var column = 0; column < values.GetLength(1); column++)
                {
                    writer.Write(JsonConvert.SerializeObject(values[row, column]));
                }
            }

            writer.Flush();
            stream.Position = 0;
            return Hex(sha256.ComputeHash(stream));
        }

        private static string Hex(byte[] hash)
        {
            var builder = new StringBuilder(hash.Length * 2);
            foreach (var value in hash)
            {
                builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }
    }
}
