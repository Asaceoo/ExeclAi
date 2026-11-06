using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Operations;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    public sealed class UndoPlannerTests
    {
        private static OperationRecord Record(int createdSortKey, bool canUndo)
        {
            return new OperationRecord
            {
                Id = createdSortKey.ToString(),
                Tool = "write_excel_cell",
                CanUndo = canUndo,
                CreatedAtUtc = DateTime.UtcNow.AddSeconds(createdSortKey)
            };
        }

        [Fact]
        public void Empty_input_yields_empty_plan()
        {
            var plan = UndoPlanner.CreatePlan(Array.Empty<OperationRecord>());
            Assert.Empty(plan.Order);
            Assert.Equal(0, plan.Eligible);
            Assert.Equal(0, plan.NotEligible);
        }

        [Fact]
        public void Null_input_is_treated_as_empty()
        {
            var plan = UndoPlanner.CreatePlan(null);
            Assert.Empty(plan.Order);
        }

        [Fact]
        public void Plan_keeps_newest_first_order_and_preserves_can_undo()
        {
            // 仓库按 created_at DESC 返回，这里模拟最新在前。
            var records = new List<OperationRecord>
            {
                Record(3, true),
                Record(2, false),
                Record(1, true)
            };

            var plan = UndoPlanner.CreatePlan(records);

            Assert.Equal(2, plan.Eligible);
            Assert.Equal(1, plan.NotEligible);
            Assert.Equal(new[] { "3", "1" }, PlanIds(plan.Order));
        }

        [Fact]
        public void Null_records_are_ignored()
        {
            var plan = UndoPlanner.CreatePlan(new OperationRecord[] { Record(1, true), null });
            Assert.Single(plan.Order);
            Assert.Equal(1, plan.NotEligible);
        }

        private static IEnumerable<string> PlanIds(IEnumerable<OperationRecord> records)
        {
            foreach (var record in records)
            {
                yield return record.Id;
            }
        }
    }
}