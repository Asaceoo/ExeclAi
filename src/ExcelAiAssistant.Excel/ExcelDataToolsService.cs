using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 数据加工类新工具：排序、去重、查找替换、转置、条件格式。
    /// 前四类通过 writer 回调统一写入（自动按结果尺寸 Resize + 记录撤销快照），
    /// 保证全程可回滚；条件格式为纯格式操作，提供 clear 模式便于还原。
    /// </summary>
    public sealed class ExcelDataToolsService
    {
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;

        public ExcelDataToolsService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
        }

        // ===================== 排序 =====================
        public ToolResult SortRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var sortColumn = PInt(args, "sort_column", 1);
            var direction = Optional(args, "direction") ?? "asc";
            if (!IsDirection(direction))
            {
                return Invalid("INVALID_DIRECTION", $"direction 仅支持 asc/desc，当前为 {direction}。");
            }

            var sortColumn2 = PInt(args, "sort_column2", 0);
            var direction2 = Optional(args, "direction2") ?? "asc";
            if (!IsDirection(direction2))
            {
                return Invalid("INVALID_DIRECTION", $"direction2 仅支持 asc/desc，当前为 {direction2}。");
            }

            var headerRows = Math.Max(0, PInt(args, "header_rows", 1));

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                var matrix = ReadMatrix(range, rows, cols);
                if (sortColumn < 1 || sortColumn > cols)
                {
                    return Invalid("INVALID_COLUMN", $"sort_column={sortColumn} 超出列范围 1..{cols}。");
                }

                if (sortColumn2 > 0 && (sortColumn2 < 1 || sortColumn2 > cols))
                {
                    return Invalid("INVALID_COLUMN", $"sort_column2={sortColumn2} 超出列范围 1..{cols}。");
                }

                // 已读完全部数据：释放本地 COM 句柄再交给 writer。
                // writer(WriteGenericRange) 会在 finally 中 FinalRelease 同一 COM Range/Worksheet，
                // 若此处不先释放，它会把本方法持有的同名 COM 对象一并销毁，令本方法 finally 的 Release 抛 InvalidComObjectException。
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;

                var rows2 = new List<RowRef>();
                for (var i = headerRows; i < rows; i++)
                {
                    rows2.Add(new RowRef(i));
                }

                // List.Sort 为不稳定（introsort）：主/次键全等时同键行相对顺序不保证。
                // 追加原始行下标作终局 tiebreaker，使同键行保持原始相对顺序（稳定），与 Excel 原生排序一致。
                rows2.Sort((x, y) =>
                {
                    var c = Most(x.Index, y.Index, sortColumn, direction, matrix);
                    if (c != 0)
                    {
                        return c;
                    }

                    if (sortColumn2 > 0)
                    {
                        var c2 = Most(x.Index, y.Index, sortColumn2, direction2, matrix);
                        if (c2 != 0)
                        {
                            return c2;
                        }
                    }

                    return x.Index.CompareTo(y.Index);
                });

                return WriteBack(args, matrix, rows2, headerRows, rows, cols, "sort_range");
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 去重 =====================
        public ToolResult DedupRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var keep = Optional(args, "keep") ?? "first";
            if (!IsKeepMode(keep))
            {
                return Invalid("INVALID_KEEP", $"keep 仅支持 first/last，当前为 {keep}。");
            }
            var headerRows = Math.Max(0, PInt(args, "header_rows", 1));
            var caseSensitive = PBool(args, "case_sensitive", false);
            var columnNums = new HashSet<int>();
            // columns 声明为数组，但 JSON 反序列化可能给出 JArray / IList / object[] / List<object> 等形态，
            // 只认 JArray 会静默退化为整行判重。按多形态防御读取。
            if (args.TryGetValue("columns", out var rawCols) && rawCols != null)
            {
                try
                {
                    if (rawCols is System.Collections.IEnumerable en && !(rawCols is string))
                    {
                        foreach (var item in en)
                        {
                            if (int.TryParse(Convert.ToString(item), out var cn))
                            {
                                columnNums.Add(cn);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                var matrix = ReadMatrix(range, rows, cols);

                // 判重列号越界/为 0 会使 DedupKey 跳过所有列 → 每行键都退化为空串 → 整表坍缩为单行。
                // 显式报错，勿静默降级为整行判重或空键坍缩。
                var badCol = FirstOutOfRangeColumn(columnNums, cols);
                if (badCol.HasValue)
                {
                    return Invalid("INVALID_COLUMN", $"判重列号 {badCol.Value} 超出列范围 1..{cols}。");
                }

                Func<int, bool> useCol = (c) => columnNums.Count == 0 || columnNums.Contains(c + 1);
                var survivors = new List<int>(); // 保留下来的原始行下标
                if (string.Equals(keep, "last", StringComparison.OrdinalIgnoreCase))
                {
                    var lastPos = new Dictionary<string, int>(StringComparer.Ordinal);
                    for (var i = headerRows; i < rows; i++)
                    {
                        lastPos[DedupKey(matrix[i], cols, useCol, caseSensitive)] = i;
                    }

                    var ordered = new List<int>(lastPos.Values);
                    ordered.Sort();
                    survivors.AddRange(ordered);
                }
                else
                {
                    var seen = new HashSet<string>(StringComparer.Ordinal);
                    for (var i = headerRows; i < rows; i++)
                    {
                        var key = DedupKey(matrix[i], cols, useCol, caseSensitive);
                        if (seen.Add(key))
                        {
                            survivors.Add(i);
                        }
                    }
                }

                // 结果仍写满原始行数：幸存行写到顶部、尾部补空，保证整块可回滚。
                // 数据已全部读取完毕：先释放本地 COM 句柄再交给 writer，与 SortRange 保持同型，
                // 避免 writer 的 FinalRelease 与本方法 finally 的 Release 竞态。
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var result = BuildWriteMatrix(matrix, rows, cols, headerRows, survivors);
                return WriteBackPadded(args, result, rows, cols, headerRows, survivors.Count, "dedup_range");
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 查找替换 =====================
        public ToolResult FindReplaceRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var find = Required(args, "find");
            var replace = Optional(args, "replace") ?? "";
            var matchCase = PBool(args, "match_case", false);
            var wholeCell = PBool(args, "whole_cell", false);
            if (string.IsNullOrEmpty(find))
            {
                return Invalid("INVALID_FIND", "查找文本（find）不能为空。");
            }

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                var matrix = ReadMatrix(range, rows, cols);
                var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var replaced = 0;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var v = matrix[r][c];
                        if (v == null)
                        {
                            continue;
                        }

                        var s = v as string;
                        if (s == null)
                        {
                            continue;
                        }

                        string nv;
                        if (wholeCell)
                        {
                            if (string.Equals(s, find, comparison))
                            {
                                nv = replace;
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            if (s.IndexOf(find, comparison) < 0)
                            {
                                continue;
                            }

                            nv = matchCase ? s.Replace(find, replace) : ReplaceIgnoreCase(s, find, replace);
                        }

                        if (nv != s)
                        {
                            matrix[r][c] = nv;
                            replaced++;
                        }
                    }
                }

                var writeArgs = new Dictionary<string, object>(args);
                var result = BuildMatrix(rows, cols, matrix);
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(writeArgs, result, "value", "find_replace_range");
                return MergeToolResult(res, new { replaced_cells = replaced });
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 转置 =====================
        public ToolResult TransposeRange(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var targetSheet = Optional(args, "target_sheet");
            var targetAddress = Optional(args, "target_address");

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                var matrix = ReadMatrix(range, rows, cols);

                var t = new object[cols, rows];
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        t[c, r] = NullToDBNull(matrix[r][c]);
                    }
                }

                // 缺省写出位 = 源区域左上角（schema 声明"缺省覆盖写回原区域"），而非 A1；
                // 缺省落到 A1 会与声明语义冲突且可能覆盖 A1 已有数据。
                var outAddr = string.IsNullOrWhiteSpace(targetAddress) ? CellTopLeft(range) : targetAddress;
                var isDefaultInPlace = string.IsNullOrWhiteSpace(targetAddress);
                var writeArgs = new Dictionary<string, object>
                {
                    ["address"] = outAddr,
                    ["sheet"] = string.IsNullOrWhiteSpace(targetSheet) ? sheetName : targetSheet
                };
                Release((object)range); range = null;
                Release((object)sheet); sheet = null;
                var res = writer(writeArgs, t, "value", "transpose_range");

                // M16 全量审查：非方阵缺省原地转置时，转置结果（cols×rows）与源区域（rows×cols）
                // 不重合——只写转置矩阵会让源区域尾部残留旧值（2×3 例：C1:C2）。补一次清空：
                // 右侧残留块（cols>rows 时为 rows×(cols-rows)）或底部残留块（rows>cols 时为
                // (rows-cols)×rows），与"区域替换语义"一致，且由 writer 留撤销快照。
                if (isDefaultInPlace && res.Ok && rows != cols)
                {
                    string ignoredSheet;
                    int topRow;
                    int topCol;
                    int ignoredR2;
                    int ignoredC2;
                    if (!ExcelRangeHelper.TryParseRange(outAddr, out ignoredSheet, out topRow, out topCol, out ignoredR2, out ignoredC2))
                    {
                        topRow = 1;
                        topCol = 1;
                    }

                    if (cols > rows)
                    {
                        // 右侧残留块：rows × (cols-rows)，从 (topRow, topCol+rows) 起
                        var rightBlock = new object[rows, cols - rows];
                        var rightArgs = new Dictionary<string, object>
                        {
                            ["address"] = ExcelRangeHelper.Resize(
                                ExcelRangeHelper.CellAddress(topRow, topCol + rows), rows, cols - rows),
                            ["sheet"] = writeArgs["sheet"]
                        };
                        res = writer(rightArgs, rightBlock, "value", "transpose_range_clear");
                    }
                    else
                    {
                        // 底部残留块：(rows-cols) × cols，从 (topRow+cols, topCol) 起
                        var bottomBlock = new object[rows - cols, cols];
                        var bottomArgs = new Dictionary<string, object>
                        {
                            ["address"] = ExcelRangeHelper.Resize(
                                ExcelRangeHelper.CellAddress(topRow + cols, topCol), rows - cols, cols),
                            ["sheet"] = writeArgs["sheet"]
                        };
                        res = writer(bottomArgs, bottomBlock, "value", "transpose_range_clear");
                    }
                }

                return MergeToolResult(res,
                    new { transposed = true, source = $"{rows}x{cols}", result = $"{cols}x{rows}" });
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 条件格式 =====================
        public ToolResult ConditionalFormat(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var rule = Required(args, "rule");
            var fill = ParseColor(Optional(args, "fill_color"));
            var text = ParseColor(Optional(args, "text_color"));
            var barColor = ParseColor(Optional(args, "bar_color"));

            dynamic sheet = null;
            dynamic range = null;
            try
            {
                sheet = Resolve(sheetName);
                range = sheet.Range[address];
                if (string.Equals(rule, "clear", StringComparison.OrdinalIgnoreCase))
                {
                    range.FormatConditions.Delete();
                    return new ToolResult { Ok = true, Data = new { rule = "clear", cleared = true } };
                }

                if (string.Equals(rule, "data_bar", StringComparison.OrdinalIgnoreCase))
                {
                    var fc = range.FormatConditions.AddDatabar();
                    try
                    {
                        fc.BarColor.Color = barColor.HasValue ? barColor.Value : ToRgb("#63be7b");
                    }
                    finally
                    {
                        Release((object)fc);
                    }

                    return new ToolResult { Ok = true, Data = new { rule = "data_bar" } };
                }

                var fill1 = fill.HasValue ? fill.Value : ToRgb("#ffd9d9");
                var text1 = text.HasValue ? text.Value : ToRgb("#9c0006");
                if (string.Equals(rule, "duplicates", StringComparison.OrdinalIgnoreCase))
                {
                    var fc = range.FormatConditions.AddUniqueValues();
                    try
                    {
                        fc.DupUnique = 2; // xlDuplicate
                        fc.Interior.Color = fill1;
                        fc.Font.Color = text1;
                    }
                    finally
                    {
                        Release((object)fc);
                    }

                    return new ToolResult { Ok = true, Data = new { rule = "duplicates" } };
                }

                // greater / less / between（数值）
                if (!string.Equals(rule, "greater", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(rule, "less", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(rule, "between", StringComparison.OrdinalIgnoreCase))
                {
                    return Invalid("INVALID_RULE", $"不支持的条件格式规则：{rule}。");
                }

                // 数值规则依赖 value（between 还需 value2）；缺失或非数值显式报错，避免静默按 0 高亮全部 / 上界为 0
                if (!TryPNumeric(args, "value", out var value1))
                {
                    return Invalid("INVALID_VALUE", $"条件格式规则 \"{rule}\" 需要数值型的 value 参数（边界值）。");
                }

                var op = string.Equals(rule, "greater", StringComparison.OrdinalIgnoreCase) ? 5L   // xlGreater
                    : string.Equals(rule, "less", StringComparison.OrdinalIgnoreCase) ? 6L        // xlLess
                    : 1L;                                                                            // xlBetween
                var value2 = 0.0;
                if (string.Equals(rule, "between", StringComparison.OrdinalIgnoreCase)
                    && !TryPNumeric(args, "value2", out value2))
                {
                    return Invalid("INVALID_VALUE", "条件格式 between 规则需要数值型的 value2 参数（上界）。");
                }

                var fc2 = string.Equals(rule, "between", StringComparison.OrdinalIgnoreCase)
                    ? range.FormatConditions.Add(1L, op, value1.ToString(CultureInfo.InvariantCulture), value2.ToString(CultureInfo.InvariantCulture))
                    : range.FormatConditions.Add(1L, op, value1.ToString(CultureInfo.InvariantCulture));
                try
                {
                    fc2.Interior.Color = fill1;
                    fc2.Font.Color = text1;
                }
                finally
                {
                    Release((object)fc2);
                }

                return new ToolResult { Ok = true, Data = new { rule = rule } };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
            }
        }

        // ===================== 内部工具 =====================
        private int Most(int x, int y, int col, string dir, object[][] matrix)
        {
            var a = matrix[x][col - 1];
            var b = matrix[y][col - 1];
            var c = CompareCell(a, b);
            return string.Equals(dir, "desc", StringComparison.OrdinalIgnoreCase) ? -c : c;
        }

        private ToolResult WriteBack(Dictionary<string, object> args, object[][] matrix, List<RowRef> order,
            int headerRows, int rows, int cols, string tool)
        {
            var result = new object[rows, cols];
            for (var i = 0; i < rows; i++)
            {
                var header = i < headerRows;
                var src = header ? i : order[i - headerRows].Index;
                for (var c = 0; c < cols; c++)
                {
                    result[i, c] = NullToDBNull(matrix[src][c]);
                }
            }

            // 行守恒硬护栏：排序必须是对身体行的纯重排——写入前比对排序前后身体行的多重集，
            // 不一致（丢行/错位/值被改）直接拒绝写入，从根上拦截「排序把数据排坏」这类系统性错误。
            var beforeKeys = new List<string>();
            for (var i = headerRows; i < rows; i++)
            {
                beforeKeys.Add(RowText(matrix[i], cols));
            }

            var afterKeys = new List<string>();
            for (var i = headerRows; i < rows; i++)
            {
                afterKeys.Add(RowText(result, i, cols));
            }

            beforeKeys.Sort(StringComparer.Ordinal);
            afterKeys.Sort(StringComparer.Ordinal);
            for (var i = 0; i < beforeKeys.Count; i++)
            {
                if (!string.Equals(beforeKeys[i], afterKeys[i], StringComparison.Ordinal))
                {
                    return Invalid("CONSERVATION_VIOLATED",
                        $"排序行守恒校验失败：排序前后身体行内容不一致（第 {i + 1} 行错位）。已拒绝写入，数据未被修改。");
                }
            }

            var res = writer(new Dictionary<string, object>(args), result, "value", tool);
            // writer(WriteGenericRange) 的 operation_id/AffectedRanges 必须透传，否则排序操作不可回滚。
            return MergeToolResult(res, new { sorted = true, keys = order.Count, row_conservation_ok = true });
        }

        /// <summary>
        /// 合并 writer 结果与工具自定义数据：保留 writer 的 operation_id、AffectedRanges 与
        /// written_cells/shape_refit 等字段，再并入工具自己的统计字段（如排序行数、去重行数）。
        /// 之前直接 new ToolResult 丢弃了 writer 的 operation_id，导致覆盖写操作无法 undo。
        /// </summary>
        private static ToolResult MergeToolResult(ToolResult writerResult, object data)
        {
            // writer 失败时其 Data 是错误文本（string）。对 string 反射 GetProperties 会命中
            // 索引器属性（Chars），无索引参数 GetValue 抛 TargetParameterCountException，
            // 把真实错误掩盖成崩溃——失败路径直接原样返回，不做任何合并。
            if (!writerResult.Ok)
            {
                return new ToolResult
                {
                    Ok = false,
                    ErrorCode = writerResult.ErrorCode,
                    Warning = writerResult.Warning,
                    Data = writerResult.Data,
                    AffectedRanges = writerResult.AffectedRanges
                };
            }

            var merged = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (writerResult.Data is Dictionary<string, object> writerData)
            {
                foreach (var kv in writerData)
                {
                    merged[kv.Key] = kv.Value;
                }
            }
            else if (writerResult.Data != null)
            {
                foreach (var property in writerResult.Data.GetType().GetProperties())
                {
                    if (property.GetIndexParameters().Length == 0)
                    {
                        merged[property.Name] = property.GetValue(writerResult.Data);
                    }
                }
            }

            if (data != null)
            {
                foreach (var property in data.GetType().GetProperties())
                {
                    if (property.GetIndexParameters().Length == 0)
                    {
                        merged[property.Name] = property.GetValue(data);
                    }
                }
            }

            return new ToolResult
            {
                Ok = writerResult.Ok,
                ErrorCode = writerResult.ErrorCode,
                Warning = writerResult.Warning,
                Truncated = writerResult.Truncated,
                AffectedRanges = writerResult.AffectedRanges,
                Data = merged
            };
        }

        /// <summary>行文本指纹（列间用控制符分隔），用于行多重集守恒比对。</summary>
        private static string RowText(object[] row, int cols)
        {
            var parts = new string[cols];
            for (var c = 0; c < cols; c++)
            {
                parts[c] = Convert.ToString(row[c]) ?? string.Empty;
            }

            return string.Join("\u0001", parts);
        }

        private static string RowText(object[,] matrix, int row, int cols)
        {
            var parts = new string[cols];
            for (var c = 0; c < cols; c++)
            {
                var value = matrix[row, c];
                parts[c] = value == null || value is DBNull ? string.Empty : Convert.ToString(value);
            }

            return string.Join("\u0001", parts);
        }

        private ToolResult WriteBackPadded(Dictionary<string, object> args, object[,] result,
            int rows, int cols, int headerRows, int survivorCount, string tool)
        {
            var res = writer(new Dictionary<string, object>(args), result, "value", tool);
            return MergeToolResult(res,
                new { deduped_rows_removed = rows - headerRows - survivorCount, survivors = survivorCount });
        }

        private object[,] BuildWriteMatrix(object[][] matrix, int rows, int cols, int headerRows, List<int> survivors)
        {
            var result = new object[rows, cols];
            for (var i = 0; i < headerRows && i < rows; i++)
            {
                WriteRow(result, matrix[i], i);
            }

            for (var k = 0; k < survivors.Count; k++)
            {
                WriteRow(result, matrix[survivors[k]], headerRows + k);
            }

            for (var r = headerRows + survivors.Count; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = DBNull.Value;
                }
            }

            return result;
        }

        private static void WriteRow(object[,] target, object[] row, int row0)
        {
            for (var c = 0; c < row.Length; c++)
            {
                target[row0, c] = NullToDBNull(row[c]);
            }
        }

        private string DedupKey(object[] row, int cols, Func<int, bool> useCol, bool caseSensitive)
        {
            var sb = new System.Text.StringBuilder();
            for (var c = 0; c < cols; c++)
            {
                if (!useCol(c))
                {
                    continue;
                }

                var v = row[c];
                var s = v == null ? "\u0000" : (v is double d ? d.ToString("R", CultureInfo.InvariantCulture) : Convert.ToString(v, CultureInfo.InvariantCulture));
                if (!caseSensitive)
                {
                    s = s.ToLowerInvariant();
                }

                sb.Append(s).Append('\u0001');
            }

            return sb.ToString();
        }

        private int CompareCell(object a, object b)
        {
            var da = ToDouble(a);
            var db = ToDouble(b);
            if (da.HasValue && db.HasValue)
            {
                return da.Value.CompareTo(db.Value);
            }

            // 数字 < 文本：与 verify_sort_filter 判据（ExcelVerifyJudgments.CompareCell）及 Excel 默认排序序一致，
            // 避免 AI 用 sort_range 排序后自检混合列误报。
            if (da.HasValue)
            {
                return -1;
            }

            if (db.HasValue)
            {
                return 1;
            }

            var sa = a == null ? string.Empty : Convert.ToString(a, CultureInfo.InvariantCulture);
            var sbb = b == null ? string.Empty : Convert.ToString(b, CultureInfo.InvariantCulture);
            return string.CompareOrdinal(sa, sbb);
        }

        private static double? ToDouble(object v)
        {
            if (v == null)
            {
                return null;
            }

            if (v is double d)
            {
                return d;
            }

            if (v is int i)
            {
                return i;
            }

            if (v is long l)
            {
                return l;
            }

            if (v is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var dr))
            {
                return dr;
            }

            return null;
        }

        private static string ReplaceIgnoreCase(string input, string find, string replace)
        {
            var sb = new System.Text.StringBuilder();
            var idx = 0;
            while (true)
            {
                var pos = input.IndexOf(find, idx, StringComparison.OrdinalIgnoreCase);
                if (pos < 0)
                {
                    sb.Append(input, idx, input.Length - idx);
                    break;
                }

                sb.Append(input, idx, pos - idx).Append(replace);
                idx = pos + find.Length;
            }

            return sb.ToString();
        }

        private object[][] ReadMatrix(dynamic range, int rows, int cols)
        {
            var obj = range.Value2;
            object[,] arr;
            if (obj is object[,] m)
            {
                arr = m;
            }
            else
            {
                arr = new object[1, 1] { { obj } };
            }

            var result = new object[rows][];
            for (var r = 0; r < rows; r++)
            {
                var rowVals = new object[cols];
                for (var c = 0; c < cols; c++)
                {
                    var v = (r + 1 <= arr.GetLength(0) && c + 1 <= arr.GetLength(1)) ? arr[r + 1, c + 1] : DBNull.Value;
                    // CVErr 错误单元格读回 Int32 统一归一为错误文本：否则排序/去重/转置/查找替换
                    // 会把 -2146826246 当数字参与比较并原样写回，错误值被静默改写成数字（M16 全量审查）。
                    rowVals[c] = (v is DBNull) ? null : ExcelErrorValues.Normalize(v);
                }

                result[r] = rowVals;
            }

            return result;
        }

        private static object[,] BuildMatrix(int rows, int cols, object[][] matrix)
        {
            var result = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = NullToDBNull(matrix[r][c]);
                }
            }

            return result;
        }

        private static object NullToDBNull(object v)
        {
            return v ?? DBNull.Value;
        }

        /// <summary>取区域左上角单元格的绝对地址（如 $B$2），用作转置缺省写出位。Cells.Item[1,1] 是链式无捕获行尾属性，
        /// 必须存临时变量 try-finally 释放，否则泄漏整棵 Range COM。</summary>
        private static string CellTopLeft(dynamic range)
        {
            dynamic cell = null;
            try
            {
                cell = range.Cells.Item[1, 1];
                return Convert.ToString(cell.Address);
            }
            catch (Exception)
            {
                return "A1";
            }
            finally
            {
                if (cell != null)
                {
                    try
                    {
                        ExcelContextService.Release((object)cell);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool IsDirection(string direction)
        {
            return string.Equals(direction, "asc", StringComparison.OrdinalIgnoreCase)
                || string.Equals(direction, "desc", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsKeepMode(string keep)
        {
            return string.Equals(keep, "first", StringComparison.OrdinalIgnoreCase)
                || string.Equals(keep, "last", StringComparison.OrdinalIgnoreCase);
        }

        private dynamic Resolve(string sheetName)
        {
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("当前没有打开的工作簿。");
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return application.ActiveSheet;
            }

            return workbook.Worksheets[sheetName];
        }

        private static int? ParseColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var hex = value.TrimStart('#');
            if (hex.Length == 6 && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var x))
            {
                return ToRgb("#" + hex);
            }

            return null;
        }

        private static int ToRgb(string hex)
        {
            var h = hex.TrimStart('#');
            var r = Convert.ToInt32(h.Substring(0, 2), 16);
            var g = Convert.ToInt32(h.Substring(2, 2), 16);
            var b = Convert.ToInt32(h.Substring(4, 2), 16);
            return r + (g << 8) + (b << 16);
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(dynamic com)
        {
            if (com != null)
            {
                try
                {
                    ExcelContextService.Release((object)com);
                }
                catch (Exception)
                {
                }
            }
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            // Y2-4：缺参抛 ArgumentException 走 INVALID_ARGS 通道（原返空串 → 下游 COM 误诊，对齐 ExcelToolHost.Required 范式）
            var text = Convert.ToString(args != null && args.TryGetValue(key, out var v) ? v : null);
            if (!string.IsNullOrEmpty(text)) { return text; }
            throw new ArgumentException($"缺少必填参数：{key}");
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null)
            {
                return Convert.ToString(v);
            }

            return null;
        }

        private static int PInt(Dictionary<string, object> args, string key, int fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null &&
                int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out var i))
            {
                return i;
            }

            return fallback;
        }

        private static double PNumeric(Dictionary<string, object> args, string key, double fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null &&
                double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                return d;
            }

            return fallback;
        }

        /// <summary>区别于 PNumeric：把「缺失」与「非法值」一并视为不可用，返回 false，供调用方显式报错而非按默认值静默处理。</summary>
        private static bool TryPNumeric(Dictionary<string, object> args, string key, out double value)
        {
            value = 0;
            return args != null && args.TryGetValue(key, out var v) && v != null &&
                double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static bool PBool(Dictionary<string, object> args, string key, bool fallback)
        {
            if (args != null && args.TryGetValue(key, out var v) && v != null && bool.TryParse(Convert.ToString(v), out var b))
            {
                return b;
            }

            return fallback;
        }

        private sealed class RowRef
        {
            public int Index { get; }

            public RowRef(int index)
            {
                Index = index;
            }
        }

        /// <summary>返回首个越界判重列号（&lt;1 或 &gt;cols），全部合法则返回 null。列为纯函数，便于回归测试锁死坍缩护栏。</summary>
        internal static int? FirstOutOfRangeColumn(IEnumerable<int> columnNums, int cols)
        {
            foreach (var cn in columnNums)
            {
                if (cn < 1 || cn > cols)
                {
                    return cn;
                }
            }

            return null;
        }
    }
}