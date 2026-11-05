using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 增强工具：环比/同比分析（rate_change_analysis）、批注管理（cell_comment）、
    /// 数据有效性（data_validation）。读采用无状态 COM 生命周期，写走一个 undo 友好的写回调。
    /// </summary>
    public sealed class ExcelEnhanceService
    {
        private const int MaxChangeRows = 20000;
        private const long MaxAnalyzeCells = 100000;

        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        private readonly OperationRecorder recorder;

        public ExcelEnhanceService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null,
            OperationRecorder recorder = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
            this.recorder = recorder;
        }

        /// <summary>环比/同比分析：读取 label+value 列，逐行计算绝对值与百分比变化。</summary>
        public ToolResult RateChangeAnalysis(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var outputAddress = Optional(args, "output_address");

            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                object[,] matrix;
                string sheetName;
                try
                {
                    matrix = ReadRange((object)sheet, address, out sheetName);
                }
                catch (GuardException gex)
                {
                    return Invalid(gex.Code, gex.Message);
                }
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                if (rows < 2)
                {
                    return Invalid("EMPTY", "至少需要一个表头与一行数据才能计算环比。");
                }

                var lfRaw = Math.Max(1, GetInt(args, "label_field", 1));
                var vfRaw = Math.Max(1, GetInt(args, "value_field", 2));
                if (lfRaw > cols || vfRaw > cols)
                {
                    var which = lfRaw > cols ? "label_field" : "value_field";
                    return Invalid("INVALID_COLUMN", $"{which}={Math.Max(lfRaw, vfRaw)} 超出区域列数 {cols}。");
                }

                var lf = lfRaw - 1;
                var vf = vfRaw - 1;
                var headerLabel = matrix[0, lf];
                var headerValue = matrix[0, vf];

                var series = new List<object>();
                var cap = Math.Max(0, Math.Min(rows - 1, MaxChangeRows));
                var result = new object[cap, 4];
                double? prev = null;
                var idx = 0;
                var numericRows = 0;
                for (var r = 1; r < rows; r++)
                {
                    if (idx >= MaxChangeRows)
                    {
                        break;
                    }

                    var label = matrix[r, lf] == null ? string.Empty : Convert.ToString(matrix[r, lf]);
                    var text = matrix[r, vf] == null ? null : Convert.ToString(matrix[r, vf]);
                    if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
                    {
                        continue;
                    }

                    double? abs = prev.HasValue ? value - prev.Value : (double?)null;
                    double? pct = prev.HasValue && Math.Abs(prev.Value) > 1e-12
                        ? (value - prev.Value) / prev.Value
                        : (double?)null;

                    result[idx, 0] = string.IsNullOrEmpty(label) ? (object)DBNull.Value : label;
                    result[idx, 1] = value;
                    result[idx, 2] = abs.HasValue ? (object)abs.Value : DBNull.Value;
                    result[idx, 3] = pct.HasValue ? (object)Math.Round(pct.Value, 4) : DBNull.Value;

                    series.Add(new { label = label, value = value, abs_change = abs, pct_change = pct });
                    prev = value;
                    idx++;
                    numericRows++;
                }

                var data = new
                {
                    label_field = headerLabel,
                    value_field = headerValue,
                    periods = numericRows,
                    method = "vs_previous_row",
                    series = series
                };

                if (!string.IsNullOrWhiteSpace(outputAddress) && numericRows > 0)
                {
                    var outMatrix = new object[idx, 4];
                    for (var i = 0; i < idx; i++)
                    {
                        outMatrix[i, 0] = result[i, 0];
                        outMatrix[i, 1] = result[i, 1];
                        outMatrix[i, 2] = result[i, 2];
                        outMatrix[i, 3] = result[i, 3];
                    }

                    var writeArgs = new Dictionary<string, object> { ["address"] = outputAddress, ["sheet"] = sheetName };
                    Release((object)sheet); sheet = null;
                    var res = writer != null
                        ? writer(writeArgs, outMatrix, "value", "rate_change_analysis")
                        : WriteDirect(sheetName, outputAddress, outMatrix);
                    return new ToolResult { Ok = res.Ok, ErrorCode = res.ErrorCode, Data = data };
                }

                return new ToolResult { Ok = true, Data = data };
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>批注管理：get / set / clear。set 时 text 必填。</summary>
        public ToolResult CellComment(Dictionary<string, object> args)
        {
            var action = (Optional(args, "action") ?? "get").ToLowerInvariant();
            if (action != "set" && action != "get" && action != "clear")
            {
                return Invalid("INVALID_VALUE", "action 仅支持 set/get/clear，当前为 " + action + "。");
            }

            var address = Required(args, "address");
            var text = Optional(args, "text");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic comment = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];
                if (action == "set")
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        return Invalid("INVALID_VALUE", "set 操作需要提供 text 参数。");
                    }

                    try
                    {
                        comment = range.Comment;
                        if (comment == null)
                        {
                            comment = range.AddComment(text);
                        }
                        else
                        {
                            comment.Text(text);
                        }

                        return new ToolResult { Ok = true, Data = new { action = "set", address, text } };
                    }
                    catch (Exception ex)
                    {
                        return Invalid("COMMENT_ERROR", "写入批注失败：" + ex.Message);
                    }
                }

                if (action == "clear")
                {
                    try
                    {
                        comment = range.Comment;
                        if (comment != null)
                        {
                            comment.Delete();
                        }

                        return new ToolResult { Ok = true, Data = new { action = "clear", address } };
                    }
                    catch (Exception ex)
                    {
                        return Invalid("COMMENT_ERROR", "清除批注失败：" + ex.Message);
                    }
                }

                // get
                comment = range.Comment;
                var current = comment != null ? Convert.ToString(comment.Text()) : null;
                return new ToolResult { Ok = true, Data = new { action = "get", address, comment = current ?? string.Empty } };
            }
            finally
            {
                Release((object)comment);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>数据有效性设置：下拉列表（list）/ 整数 interval 校验。默认 list。</summary>
        public ToolResult DataValidation(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var type = (Optional(args, "type") ?? "list").ToLowerInvariant();
            var source = Optional(args, "source");
            if (type != "list" && type != "interval")
            {
                return Invalid("INVALID_VALUE", "type 仅支持 list/interval，当前为 " + type + "。");
            }

            if (type == "list" && string.IsNullOrWhiteSpace(source))
            {
                return Invalid("INVALID_VALUE", "list 类型需要提供 source 参数（如 A1:A5，或用逗号分隔的列表值）。");
            }

            var operatorType = Optional(args, "operator") ?? "between";
            var value1 = Optional(args, "value1");
            var value2 = Optional(args, "value2");
            if (type == "interval")
            {
                if (string.IsNullOrWhiteSpace(value1) ||
                    !double.TryParse(value1.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    return Invalid("INVALID_VALUE", "interval 类型必须提供数值 value1（如 value1=10）。当前为 " + (value1 ?? "<空>") + "。");
                }

                if (operatorType.Equals("between", StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(value2) ||
                     !double.TryParse(value2.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out _)))
                {
                    return Invalid("INVALID_VALUE", "interval 的 between 需要数值 value1 与 value2（如 value1=0,value2=100）。");
                }
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                range = sheet.Range[address];

                // 先清除旧有效性规则
                try
                {
                    range.Validation.Delete();
                }
                catch (Exception)
                {
                }

                if (type == "list")
                {
                    var formula = source.StartsWith("=", StringComparison.Ordinal) || source.StartsWith("!", StringComparison.Ordinal)
                        ? source
                        : source.Contains(",")
                            ? "\"" + source + "\""
                            : "=" + source;
                    // xlValidateList=3, xlValidAlertStop=1, xlBetween=1。仅传 4 参(Type,AlertStyle,Operator,Formula1)：
                    // 实测本机 Excel 的 Validation.Add 经 net48 dynamic COM 最多接受 5 参，6 参(带 IgnoreBlank)会抛
                    // TargetParameterCountException「调用 Add 时出错」。省略第 6 参即采用 Excel 默认允许空白。
                    range.Validation.Add(3, 1, 1, formula);
                    return new ToolResult { Ok = true, Data = new { type = "list", address, source } };
                }

                // type == interval
                var op = ResolveOperator(operatorType);
                if (op < 0)
                {
                    return Invalid("INVALID_VALUE", "interval 类型 operator 仅支持 between/greater/less，当前为 " + operatorType);
                }

                var formula1 = BuildCompareFormula(value1);
                // xlValidateWholeNumber=1。between 用 5 参带上限，非 between 用 4 参。
                if (operatorType.Equals("between", StringComparison.OrdinalIgnoreCase))
                {
                    var formula2 = BuildCompareFormula(value2);
                    range.Validation.Add(1, 1, op, formula1, formula2);
                }
                else
                {
                    range.Validation.Add(1, 1, op, formula1);
                }

                return new ToolResult { Ok = true, Data = new Dictionary<string, object> { ["type"] = "interval", ["address"] = address, ["operator"] = operatorType, ["value1"] = value1, ["value2"] = value2 } };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private static string BuildCompareFormula(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "0";
            }

            var text = value.Trim();
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                return text;
            }

            if (text.StartsWith("=", StringComparison.Ordinal) || text.StartsWith(">", StringComparison.Ordinal) ||
                text.StartsWith("<", StringComparison.Ordinal))
            {
                return text;
            }

            // 视为单元格引用或公式
            return "=" + text;
        }

        private static int ResolveOperator(string op)
        {
            if (op.Equals("between", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (op.Equals("greater", StringComparison.OrdinalIgnoreCase))
            {
                return 5;
            }

            if (op.Equals("less", StringComparison.OrdinalIgnoreCase))
            {
                return 6;
            }

            return -1;
        }

        private ToolResult WriteDirect(string sheetName, string address, object[,] matrix)
        {
            // writer 为空时的兜底直写。
            dynamic wb = null;
            dynamic sh = null;
            try
            {
                Resolve(ref wb, ref sh, sheetName);
                var rng = sh.Range[address];
                rng.Value2 = matrix;
                Release((object)rng);
                return new ToolResult { Ok = true, Data = new { written = matrix.GetLength(0) } };
            }
            finally
            {
                Release((object)sh);
                Release((object)wb);
            }
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private object[,] ReadRange(dynamic sheetObject, string address, out string sheetName)
        {
            var range = sheetObject.Range[address];
            sheetName = Convert.ToString(sheetObject.Name);
            var rows = Convert.ToInt32(range.Rows.Count);
            var cols = Convert.ToInt32(range.Columns.Count);
            // 全量读入护栏：rows×cols 用 long 防 int 溢出，超限以受控信号拒绝（RateChange 捕获转 Invalid）。
            if (rows * (long)cols > MaxAnalyzeCells)
            {
                throw new GuardException("TOO_MANY_CELLS", $"数据区域过大（{rows} 行 × {cols} 列），超过分析上限 {MaxAnalyzeCells}，请缩小范围。");
            }

            var values = range.Value2;
            var matrix = Flatten(values, rows, cols);
            Release((object)range);
            return matrix;
        }

        private static object[,] Flatten(object values, int rows, int cols)
        {
            var matrix = new object[rows, cols];
            if (values == null)
            {
                return matrix;
            }

            var arr = values as object[,];
            if (arr == null)
            {
                matrix[0, 0] = values;
                return matrix;
            }

            var lo0 = arr.GetLowerBound(0);
            var lo1 = arr.GetLowerBound(1);
            var hi0 = arr.GetUpperBound(0);
            var hi1 = arr.GetUpperBound(1);
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    // Excel 的 Value2 返回 1-based 数组（GetLowerBound()==1），不能用
                    // lo+i<GetLength 判断（会整体丢最后一行/列），须按 [lower,upper] 区间判断。
                    if (lo0 + r <= hi0 && lo1 + c <= hi1)
                    {
                        matrix[r, c] = arr[lo0 + r, lo1 + c];
                    }
                }
            }

            return matrix;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            if (args.TryGetValue(key, out var value) && value != null && !string.IsNullOrWhiteSpace(Convert.ToString(value)))
            {
                return Convert.ToString(value);
            }

            throw new ArgumentException("缺少必填参数：" + key);
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (args.TryGetValue(key, out var value) && value != null)
            {
                if (int.TryParse(Convert.ToString(value), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
                {
                    return n;
                }
            }

            return defaultValue;
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args.TryGetValue(key, out var value) && value != null)
            {
                var text = Convert.ToString(value);
                if (bool.TryParse(text, out var b))
                {
                    return b;
                }

                if (text == "1")
                {
                    return true;
                }

                if (text == "0")
                {
                    return false;
                }
            }

            return defaultValue;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
        }
    }

    /// <summary>全量读入护栏的受控信号：只在区域行×列超上限时抛出，由 RateChange 捕获转 Invalid（TOO_MANY_CELLS）。</summary>
    internal sealed class GuardException : Exception
    {
        internal readonly string Code;

        internal GuardException(string code, string message) : base(message)
        {
            Code = code;
        }
    }
}