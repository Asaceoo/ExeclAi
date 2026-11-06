using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>
    /// 参数默认值分歧审计（防回归锚点）：
    /// 同一工具走「参数框路径」（RibbonCommandPlanner 参数框默认值 → BuildArgs 实参）与
    /// 「AI 直接调用路径」（ExcelDataToolsService 缺省 fallback）必须行为一致。
    /// 任一默认值漂移（如参数框 has_header 默认改为 false 而工具 fallback 仍为 1）都会让
    /// 两条路径产出不同结果——本测试从参数框默认值、实参组装、工具缺省行为三端锁死。
    /// </summary>
    public sealed class ParamDefaultsAuditTests
    {
        private Dictionary<string, object> _capturedArgs;
        private object[,] _capturedValues;

        public sealed class FakeCell
        {
            public string Address { get; } = "$A$1";
        }

        public sealed class FakeCells
        {
            // CellTopLeft 走 COM 风格链式访问 range.Cells.Item[1,1]：Item 需返回可索引对象。
            // 索引器默认成员名即 Item，故承载索引器的是独立类。
            public FakeCellItems Item { get; } = new FakeCellItems();
        }

        public sealed class FakeCellItems
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
            // M16 起 transpose_range 缺省原地转置会追加一次"残留清空"写，契约断言以首次写为准。
            if (_capturedArgs != null)
            {
                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { "Sheet1!$A$1:$B$3" },
                    Data = new { operation_id = "op-audit", written_cells = values.GetLength(0) * values.GetLength(1) }
                };
            }

            _capturedArgs = args;
            _capturedValues = values;
            return new ToolResult
            {
                Ok = true,
                AffectedRanges = { "Sheet1!$A$1:$B$3" },
                Data = new { operation_id = "op-audit", written_cells = values.GetLength(0) * values.GetLength(1) }
            };
        }

        private ExcelDataToolsService MakeService(object[,] values, int rows, int cols)
        {
            var range = new FakeRange { Values = values };
            range.Rows.Count = rows;
            range.Columns.Count = cols;
            return new ExcelDataToolsService(new FakeApp { ActiveSheet = new FakeSheet { Range = range } }, FakeWriter);
        }

        private static Dictionary<string, object> DefaultArgs(string toolName, int cols)
        {
            var sel = new ParamSelection { SheetName = "Sheet1", Address = "$A$1", Rows = 3, Cols = cols };
            var built = RibbonCommandPlanner.BuildArgs(toolName, sel, new Dictionary<string, object>());
            return Assert.IsType<Dictionary<string, object>>(built);
        }

        private static string FieldDefault(List<ParamField> fields, string name)
        {
            return fields.First(f => f.Name == name).Default;
        }

        [Fact]
        public void SortRange_param_box_defaults_match_tool_fallback()
        {
            // 参数框默认值：排序列 1、升序、无次排序列、含表头
            var fields = RibbonCommandPlanner.BuildFields("sort_range", new ParamSelection { Cols = 2 });
            Assert.Equal("1", FieldDefault(fields, "sort_column"));
            Assert.Equal("asc", FieldDefault(fields, "direction"));
            Assert.Equal("0", FieldDefault(fields, "sort_column2"));
            Assert.Equal("asc", FieldDefault(fields, "direction2"));
            Assert.Equal("true", FieldDefault(fields, "has_header"));

            // 参数框默认 → 工具实参：header_rows=1、不注入次排序列
            var args = DefaultArgs("sort_range", 2);
            Assert.Equal(1, Convert.ToInt32(args["sort_column"]));
            Assert.Equal("asc", Convert.ToString(args["direction"]));
            Assert.Equal(1, Convert.ToInt32(args["header_rows"]));
            Assert.False(args.ContainsKey("sort_column2"), "参数框默认无次排序列时不得注入 sort_column2");

            // 工具缺省（AI 不传可选参数）→ 与参数框默认一致：表头保持、按第 1 列升序
            var service = MakeService(Matrix1Based(
                new object[] { "品名", "数量" },
                new object[] { "甲", 1.0 },
                new object[] { "丙", 3.0 }), 3, 2);
            var result = service.SortRange(new Dictionary<string, object> { ["address"] = "$A$1:$B$3" });
            Assert.True(result.Ok);
            Assert.Equal("品名", Convert.ToString(_capturedValues[0, 0]));
            Assert.Equal("丙", Convert.ToString(_capturedValues[1, 0]));
            Assert.Equal("甲", Convert.ToString(_capturedValues[2, 0]));
        }

        [Fact]
        public void DedupRange_param_box_defaults_match_tool_fallback()
        {
            // 参数框默认值：整行判重、保留首次、含表头、不区分大小写
            var fields = RibbonCommandPlanner.BuildFields("dedup_range", new ParamSelection { Cols = 2 });
            Assert.Equal("whole", FieldDefault(fields, "scope"));
            Assert.Equal("first", FieldDefault(fields, "keep"));
            Assert.Equal("true", FieldDefault(fields, "has_header"));
            Assert.Equal("false", FieldDefault(fields, "case_sensitive"));

            // 参数框默认 → 工具实参：不注入 columns（整行判重）、keep=first、header_rows=1
            var args = DefaultArgs("dedup_range", 2);
            Assert.Equal("first", Convert.ToString(args["keep"]));
            Assert.Equal(1, Convert.ToInt32(args["header_rows"]));
            Assert.False(Convert.ToBoolean(args["case_sensitive"]));
            Assert.False(args.ContainsKey("columns"), "参数框默认整行判重时不得注入 columns");

            // 工具缺省（AI 不传可选参数）→ 与参数框默认一致：整行判重、保留首次、表头保持
            var service = MakeService(Matrix1Based(
                new object[] { "品名", "数量" },
                new object[] { "甲", 1.0 },
                new object[] { "甲", 1.0 }), 3, 2);
            var result = service.DedupRange(new Dictionary<string, object> { ["address"] = "$A$1:$B$3" });
            Assert.True(result.Ok);
            Assert.Equal("品名", Convert.ToString(_capturedValues[0, 0]));
            Assert.Equal("甲", Convert.ToString(_capturedValues[1, 0]));
            Assert.True(_capturedValues[2, 0] is DBNull, "去重后尾部应补空（行守恒）");
        }

        [Fact]
        public void TransposeRange_param_box_defaults_match_tool_fallback()
        {
            // 参数框默认值：写入原区域左上角（top_left），不携带目标地址/工作表
            var fields = RibbonCommandPlanner.BuildFields("transpose_range", new ParamSelection { Cols = 2 });
            Assert.Equal("top_left", FieldDefault(fields, "target_mode"));

            // 参数框默认 → 工具实参：仅 address/sheet，无 target_address/target_sheet
            var args = DefaultArgs("transpose_range", 2);
            Assert.False(args.ContainsKey("target_address"), "参数框默认写回左上角时不得注入 target_address");
            Assert.False(args.ContainsKey("target_sheet"), "参数框默认写回左上角时不得注入 target_sheet");

            // 工具缺省（AI 不传可选参数）→ 与参数框默认一致：结果写回源区域左上角
            var service = MakeService(Matrix1Based(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 4.0, 5.0, 6.0 }), 2, 3);
            var result = service.TransposeRange(new Dictionary<string, object> { ["address"] = "$A$1:$C$2" });
            Assert.True(result.Ok);
            Assert.Equal("$A$1", Convert.ToString(_capturedArgs["address"]));
            Assert.Equal(3, _capturedValues.GetLength(0));
            Assert.Equal(2, _capturedValues.GetLength(1));
            Assert.Equal(1.0, Convert.ToDouble(_capturedValues[0, 0]));
            Assert.Equal(4.0, Convert.ToDouble(_capturedValues[0, 1]));
        }
    }
}
