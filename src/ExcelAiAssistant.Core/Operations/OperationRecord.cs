using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Operations
{
    public sealed class OperationRecord
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string SessionId { get; set; }

        public string WorkbookHash { get; set; }

        public string WorkbookPath { get; set; }

        public string Tool { get; set; }

        /// <summary>所属对话轮次标签，例如"第 3 轮"。用于在历史中标识该操作由哪轮对话产生。</summary>
        public string TurnLabel { get; set; }

        public string ArgsJson { get; set; }

        public string AffectedJson { get; set; }

        public string UndoKind { get; set; }

        public string PayloadPath { get; set; }

        public string PayloadJson { get; set; }

        public string BeforeDigest { get; set; }

        public string AfterDigest { get; set; }

        public bool CanUndo { get; set; } = true;

        public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
    }

    public sealed class CellSnapshot
    {
        public string Sheet { get; set; }

        public string Address { get; set; }

        public object[,] BeforeValues { get; set; }

        public object[,] AfterValues { get; set; }

        public object[,] BeforeFormulas { get; set; }

        public object[,] AfterFormulas { get; set; }
    }

    public sealed class OperationQuery
    {
        public string WorkbookHash { get; set; }

        public int Limit { get; set; } = 50;
    }

    /// <summary>
    /// 记录结构化撤销（工作表插入/行列插入）的回调签名。
    /// 参数：tool, args, undoKind, affected, payload。
    /// </summary>
    public delegate void OperationRecorder(
        string tool,
        object args,
        string undoKind,
        object affected,
        object payload);

    /// <summary>
    /// 工作表插入类撤销载荷：撤销时删除这些新插入的工作表。
    /// </summary>
    public sealed class SheetInsertUndoPayload
    {
        public string[] Sheets { get; set; }

        public int Count => Sheets?.Length ?? 0;
    }

    /// <summary>
    /// 行列插入类撤销载荷：撤销时按 dimension/address/count 删除插入的行或列。
    /// </summary>
    public sealed class DimensionInsertUndoPayload
    {
        public string Sheet { get; set; }

        /// <summary>row 或 column。</summary>
        public string Dimension { get; set; }

        public string Address { get; set; }

        public int Count { get; set; }
    }

    /// <summary>
    /// 图表插入撤销载荷：撤销时删除对应名称的图表对象。
    /// </summary>
    public sealed class ChartInsertPayload
    {
        public string Sheet { get; set; }

        public string ChartName { get; set; }
    }

    /// <summary>
    /// 迷你图插入撤销载荷：撤销时清空目标区域的迷你图组。
    /// </summary>
    public sealed class SparklineInsertPayload
    {
        public string Sheet { get; set; }

        public string DestRange { get; set; }
    }
}
