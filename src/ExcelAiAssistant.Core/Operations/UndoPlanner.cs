using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Operations
{
    /// <summary>全量回滚计划的统计结果。</summary>
    public sealed class RollbackPlan
    {
        /// <summary>按最新→最早顺序、仅含当前可撤销记录的撤销顺序。</summary>
        public IReadOnlyList<OperationRecord> Order { get; set; }

        /// <summary>计划内可撤销的记录数。</summary>
        public int Eligible { get; set; }

        /// <summary>因不可撤销而跳过的记录数。</summary>
        public int NotEligible { get; set; }
    }

    /// <summary>
    /// 生成全量历史回滚的执行计划：保持输入顺序（仓库按 created_at_utc DESC 返回，即最新→最早），
    /// 只纳入当前可撤销的记录，保证以 LIFO 顺序依次撤销。
    /// </summary>
    public static class UndoPlanner
    {
        public static RollbackPlan CreatePlan(IEnumerable<OperationRecord> newestFirst)
        {
            var source = newestFirst ?? Array.Empty<OperationRecord>();
            var order = new List<OperationRecord>();
            var notEligible = 0;
            foreach (var record in source)
            {
                if (record != null && record.CanUndo)
                {
                    order.Add(record);
                }
                else
                {
                    notEligible++;
                }
            }

            return new RollbackPlan
            {
                Order = order,
                Eligible = order.Count,
                NotEligible = notEligible
            };
        }
    }
}