using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 锁定覆盖写工具（排序/去重/查找替换/转置）的回滚字段透传：
    /// writer(WriteGenericRange) 产出的 operation_id 与 AffectedRanges 必须到达工具结果，
    /// 否则前端拿不到 operation_id、覆盖写操作不可 undo（回归锚点：ExcelDataToolsService 曾丢弃）。
    /// 全部走 fake COM + fake writer，不依赖真机 Excel。
    /// </summary>
    public sealed class ExcelDataToolsServiceTests
    {
        private const string FakeOpId = "op-fake-001";

        private Dictionary<string, object> _capturedArgs;
        private object[,] _capturedValues;
        private string _capturedTool;

        public sealed class FakeCell
        {
            public string Address { get; } = "$A$1";
        }

        public sealed class FakeCells
        {
            public FakeCell this[int row, int col] => new FakeCell();
        }

        public sealed class FakeRows
        {
            public int Count { get; set; }
        }

        public sealed class FakeColumns
        {
            public int Count { get; set; }
        }

        public sealed class FakeRange
        {
            public object[,] Values { get; set; }
            public string Address { get; } = "$A$1:$B$3";
            public FakeRows Rows { get; } = new FakeRows();
            public FakeColumns Columns { get; } = new FakeColumns();
            public FakeCells Cells { get; } = new FakeCells();
            public object Value2 => Values;
            public FakeRange this[string address] => this;
        }

        public sealed class FakeSheet
        {
            public string Name { get; } = "Sheet1";
            public FakeRange Range { get; set; }
        }

        public sealed class FakeWorksheets
        {
            public FakeSheet this[string name] => new FakeSheet();
        }

        public sealed class FakeWorkbook
        {
            public FakeWorksheets Worksheets { get; } = new FakeWorksheets();
        }

        public sealed class FakeApp
        {
            public FakeWorkbook ActiveWorkbook { get; } = new FakeWorkbook();
            public FakeSheet ActiveSheet { get; set; }
        }

        /// <summary>按行序列构造 1 基 object[,]（Excel COM Value2 形态）。</summary>
        private static object[,] Matrix1Based(params object[][] rows)
        {
            var result = new object[rows.Length + 1, rows[0].Length + 1];
            for (var r = 0; r < rows.Length; r++)
            {
                for (var c = 0; c < rows[r].Length; c++)
                {
                    result[r + 1, c + 1] = rows[r][c];
                }
            }

            return result;
        }

        private ToolResult FakeWriter(Dictionary<string, object> args, object[,] values, string valueMode, string tool)
        {
            // M16 起 transpose_range 缺省原地转置会追加一次"残留清空"写（同一契约操作），
            // 契约断言以首次写（转置本体）为准。
            if (_capturedArgs != null)
            {
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { "Sheet1!$A$1:$B$3" },
                    Data = new
                    {
                        operation_id = FakeOpId,
                        written_cells = values.GetLength(0) * values.GetLength(1),
                        value_mode = "value",
                        shape_refit = false
                    }
                };
            }

            _capturedArgs = args;
            _capturedValues = values;
            _capturedTool = tool;
            return new ToolResult
            {
                Ok = true,
                AffectedRanges = { "Sheet1!$A$1:$B$3" },
                Data = new
                {
                    operation_id = FakeOpId,
                    written_cells = values.GetLength(0) * values.GetLength(1),
                    value_mode = "value",
                    shape_refit = false
                }
            };
        }

        private ExcelDataToolsService MakeService(object[,] values, int rows, int cols)
        {
            var range = new FakeRange { Values = values };
            range.Rows.Count = rows;
            range.Columns.Count = cols;
            return new ExcelDataToolsService(new FakeApp { ActiveSheet = new FakeSheet { Range = range } }, FakeWriter);
        }

        [Fact]
        public void SortRange_result_carries_operation_id_and_affected_ranges()
        {
            var service = MakeService(Matrix1Based(
                new object[] { "丙", 3.0 },
                new object[] { "甲", 1.0 },
                new object[] { "乙", 2.0 }), 3, 2);

            var result = service.SortRange(new Dictionary<string, object>
            {
                ["address"] = "$A$1:$B$3",
                ["sort_column"] = 2,
                ["direction"] = "desc",
                ["header_rows"] = 0
            });

            Assert.True(result.Ok);
            Assert.Contains("Sheet1!$A$1:$B$3", result.AffectedRanges);
            var data = Assert.IsType<Dictionary<string, object>>(result.Data);
            Assert.Equal(FakeOpId, Convert.ToString(data["operation_id"]));
            Assert.Equal(true, data["sorted"]);
            Assert.Equal(true, data["row_conservation_ok"]);
            Assert.Equal(3, Convert.ToInt32(data["keys"]));
            Assert.Equal("sort_range", _capturedTool);
            Assert.Equal("丙", Convert.ToString(_capturedValues[0, 0]));
            Assert.Equal("乙", Convert.ToString(_capturedValues[1, 0]));
            Assert.Equal("甲", Convert.ToString(_capturedValues[2, 0]));
        }

        [Fact]
        public void DedupRange_result_carries_operation_id_and_affected_ranges()
        {
            var service = MakeService(Matrix1Based(
                new object[] { "品名", "数量" },
                new object[] { "甲", 1.0 },
                new object[] { "甲", 2.0 }), 3, 2);

            var result = service.DedupRange(new Dictionary<string, object>
            {
                ["address"] = "$A$1:$B$3",
                ["header_rows"] = 1,
                ["columns"] = new object[] { 1 }
            });

            Assert.True(result.Ok);
            var diagContains = result.AffectedRanges.Contains("Sheet1!$A$1:$B$3");
            Assert.True(diagContains, "AffectedRanges=[" + string.Join("|", result.AffectedRanges) + "] Data=" + (result.Data is Dictionary<string, object> dd ? string.Join("|", dd.Keys) : result.Data));
            Assert.Contains("Sheet1!$A$1:$B$3", result.AffectedRanges);
            var data = Assert.IsType<Dictionary<string, object>>(result.Data);
            Assert.Equal(FakeOpId, Convert.ToString(data["operation_id"]));
            Assert.Equal(1, Convert.ToInt32(data["deduped_rows_removed"]));
            Assert.Equal(1, Convert.ToInt32(data["survivors"]));
            Assert.Equal("dedup_range", _capturedTool);
        }

        [Fact]
        public void FindReplaceRange_result_carries_operation_id_and_affected_ranges()
        {
            var service = MakeService(Matrix1Based(
                new object[] { "apple", "APPLE" },
                new object[] { "banana", "apple pie" }), 2, 2);

            var result = service.FindReplaceRange(new Dictionary<string, object>
            {
                ["address"] = "$A$1:$B$2",
                ["find"] = "apple",
                ["replace"] = "pear"
            });

            Assert.True(result.Ok);
            Assert.Contains("Sheet1!$A$1:$B$3", result.AffectedRanges);
            var data = Assert.IsType<Dictionary<string, object>>(result.Data);
            Assert.Equal(FakeOpId, Convert.ToString(data["operation_id"]));
            Assert.Equal(3, Convert.ToInt32(data["replaced_cells"]));
            Assert.Equal("find_replace_range", _capturedTool);
        }

        [Fact]
        public void SortRange_writer_failure_preserves_error_without_crash()
        {
            // 覆盖 FakeWriter：失败时 Data 是错误文本（string），曾触发 MergeToolResult 对
            // string 反射索引器属性抛 TargetParameterCountException，掩盖真实错误。
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> failingWriter =
                (args, values, mode, tool) => new ToolResult
                {
                    Ok = false,
                    ErrorCode = "WRITE_FAILED",
                    Data = "写入失败：模拟错误"
                };
            var failingService = new ExcelDataToolsService(
                new FakeApp { ActiveSheet = new FakeSheet { Range = new FakeRange
                {
                    Values = Matrix1Based(
                        new object[] { "丙", 3.0 },
                        new object[] { "甲", 1.0 },
                        new object[] { "乙", 2.0 }),
                    Rows = { Count = 3 },
                    Columns = { Count = 2 }
                } } },
                failingWriter);

            var result = failingService.SortRange(new Dictionary<string, object>
            {
                ["address"] = "$A$1:$B$3",
                ["sort_column"] = 2,
                ["direction"] = "desc",
                ["header_rows"] = 0
            });

            Assert.False(result.Ok);
            Assert.Equal("WRITE_FAILED", result.ErrorCode);
            Assert.Equal("写入失败：模拟错误", Convert.ToString(result.Data));
        }

        [Fact]
        public void TransposeRange_result_carries_operation_id_and_affected_ranges()
        {
            var service = MakeService(Matrix1Based(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0, 5.0, 6.0 }), 2, 3);

            var result = service.TransposeRange(new Dictionary<string, object>
            {
                ["address"] = "$A$1:$C$2"
            });

            Assert.True(result.Ok);
            Assert.Contains("Sheet1!$A$1:$B$3", result.AffectedRanges);
            var data = Assert.IsType<Dictionary<string, object>>(result.Data);
            Assert.Equal(FakeOpId, Convert.ToString(data["operation_id"]));
            Assert.Equal(true, data["transposed"]);
            Assert.Equal("2x3", Convert.ToString(data["source"]));
            Assert.Equal("3x2", Convert.ToString(data["result"]));
            Assert.Equal("transpose_range", _capturedTool);
        }
    }
}
