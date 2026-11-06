using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 标记式清洗（P6-B）：识别异常值/空值/格式异常 → 源区视觉高亮（异色，易失）
    /// + 生成「异常标记诊断」sheet（持久可审计）。诊断 sheet 由宿主结构化撤销记录，
    /// 可整体删除并恢复。
    /// </summary>
    public sealed class ExcelFlagCellsService
    {
        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        private readonly Action<string, object, string, object, object> recordStructuralUndo;

        public ExcelFlagCellsService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null,
            Action<string, object, string, object, object> recordStructuralUndo = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
            this.recordStructuralUndo = recordStructuralUndo;
        }

        public ToolResult FlagCells(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var sheetName = Optional(args, "sheet");
            var flagTypes = ParseFlagTypes(args);
            foreach (var ft in flagTypes ?? Array.Empty<string>())
            {
                if (!IsKnownFlagType(ft))
                {
                    return Invalid("INVALID_FLAG_TYPE", "flag_types 含未知类型：" + ft + "，仅支持 outlier/format/empty。");
                }
            }

            var flagColors = BuildHighlightColors(args);
            if (flagColors == null)
            {
                return Invalid("INVALID_FILL_COLOR", "fill_colors 需为 type→\"#RRGGBB\" 映射，颜色须为 6 位十六进制。");
            }

            dynamic sourceSheet = null;
            dynamic range = null;
            try
            {
                var (sheet, workbook) = ResolveSource(sheetName);
                sourceSheet = sheet;
                range = sourceSheet.Range[address];
                var rows = Convert.ToInt32(range.Rows.Count);
                var cols = Convert.ToInt32(range.Columns.Count);
                // 选区绝对原点（1 基）：OutlierDetector 给出的 CellFlag.Row/Col 相对选区左上角（A1 语义），
                // 高亮/诊断要落回真实工作表坐标，必须叠加选区绝对行/列偏移；否则 address 非 A1 时全部错位。
                // 显式 int 收敛 dynamic 传染：Convert.ToInt32(range.Row) 带 dynamic 实参返回 dynamic，
                // 若用 var 声明会让 CreateDiagnosisSheet 变动态分派、返回元组退化为 dynamic，.Name 抛 RuntimeBinderException。
                int originRow = Convert.ToInt32(range.Row);
                int originCol = Convert.ToInt32(range.Column);
                // range 为 dynamic → ReadMatrix 是动态分派，返回 block 会退化为 dynamic；必须显式还原静态类型，
                // 否则 BuildPlan→plan→CreateDiagnosisSheet 这一串全部变成动态分派，diagInfo.Name 会抛
                // RuntimeBinderException（元组元素名非运行时成员）。
                object[,] block = ReadMatrix(range, rows, cols);

                var plan = OutlierDetector.BuildPlan(block, flagTypes);

                // 源区异色高亮（视觉引导，易失，不写数据）
                // sourceSheet 为 dynamic → 该调用是动态分派，返回值为 dynamic；元组元素名只是元数据非运行时成员。
                // 以运行时真实成员 Item1 取值（binder 找不到 TotalCells，会抛 RuntimeBinderException）。
                var highlightInfo = ApplyHighlights(sourceSheet, plan, originRow, originCol, flagColors);

                // 持久诊断 sheet（可审计，可整体删除+恢复）
                var diagInfo = CreateDiagnosisSheet(plan, sheetName, originRow, originCol);

                var counts = new Dictionary<string, int>();
                foreach (var kv in plan)
                {
                    counts[kv.Key] = kv.Value.Count;
                }

                var plannedTotal = counts.Values.Sum();
                var appliedTotal = (int)highlightInfo.Item1;
                var data = new Dictionary<string, object>
                {
                    ["source"] = $"{sheetName} {address}",
                    ["flagged"] = counts,
                    ["total_cells"] = appliedTotal,
                    ["planned_cells"] = plannedTotal,
                    ["diagnosis_sheet"] = diagInfo.Name,
                    ["diagnosis_rows"] = diagInfo.Rows,
                    ["record_undo"] = diagInfo.Recorded
                };
                if (appliedTotal < plannedTotal)
                {
                    // M16 全量审查（fail-closed）：受保护表等场景逐格高亮失败被吞，计划 N 格实际 0 格
                    // 曾被误读为"已标记"。如实上报缺口。
                    data["warning"] = $"有 {plannedTotal - appliedTotal} 个单元格标记失败（常见原因：工作表受保护），实际成功标记 {appliedTotal} 个。";
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = data
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sourceSheet);
            }
        }

        private (dynamic Sheet, dynamic Workbook) ResolveSource(string sheetName)
        {
            var workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
            var sheet = string.IsNullOrWhiteSpace(sheetName)
                ? (dynamic)application.ActiveSheet
                : (dynamic)workbook.Worksheets.Item[sheetName];
            return (sheet, workbook);
        }

        private static string[] ParseFlagTypes(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("flag_types", out var raw)) return null;
            // 兼容 JArray（侧栏 JSON 反序列化）与 List<object>/object[]（进程内直接调用）
            var list = new List<string>();
            if (raw is System.Collections.IEnumerable seq && !(raw is string))
            {
                foreach (var item in seq)
                {
                    string s = Convert.ToString(item);
                    if (!string.IsNullOrWhiteSpace(s)) list.Add(s.Trim().ToLowerInvariant());
                }
            }
            else
            {
                string s = Convert.ToString(raw);
                if (!string.IsNullOrWhiteSpace(s)) list.Add(s.Trim().ToLowerInvariant());
            }
            return list.Count > 0 ? list.ToArray() : null;
        }

        private static Dictionary<string, uint> HighlightColors = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase)
        {
            ["outlier"] = ToRgb("#bd0020"),
            ["format"]  = ToRgb("#ec7000"),
            ["empty"]   = ToRgb("#5b9bd5")
        };

        private static uint ToRgb(string hex)
        {
            // Excel BGR：把 "RRGGBB" 转成 BGR 长整型
            var h = hex.TrimStart('#');
            if (h.Length != 6) return 0x9C0006;
            var r = Convert.ToInt32(h.Substring(0, 2), 16);
            var g = Convert.ToInt32(h.Substring(2, 2), 16);
            var b = Convert.ToInt32(h.Substring(4, 2), 16);
            return (uint)((b << 16) | (g << 8) | r);
        }

        private (int TotalCells, Dictionary<string, int> ByType) ApplyHighlights(
            dynamic sheet, Dictionary<string, List<CellFlag>> plan, int originRow, int originCol,
            Dictionary<string, uint> colors)
        {
            var total = 0;
            var byType = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in plan)
            {
                if (!colors.TryGetValue(kv.Key, out var color)) continue;
                var cells = kv.Value;
                if (cells.Count == 0) continue;
                byType[kv.Key] = cells.Count;
                foreach (var cf in cells)
                {
                    dynamic cell = null;
                    try
                    {
                        cell = sheet.Range[ExcelRangeHelper.CellAddress(cf.Row + originRow - 1, cf.Col + originCol - 1)];
                        cell.Interior.Color = color;
                        total++;
                    }
                    catch (Exception)
                    {
                        // 单个标记失败不阻断整体
                    }
                    finally
                    {
                        Release((object)cell);
                    }
                }
            }
            return (total, byType);
        }

        private (string Name, int Rows, bool Recorded) CreateDiagnosisSheet(
            Dictionary<string, List<CellFlag>> plan, string sourceSheetName, int originRow, int originCol)
        {
            var totalRows = 0;
            foreach (var kv in plan) totalRows += kv.Value.Count;
            if (totalRows == 0)
            {
                return (null, 0, false);
            }

            var name = UniqueSheetName("异常标记诊断");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                workbook = application.ActiveWorkbook;
                sheet = workbook.Worksheets.Add();
                sheet.Name = name;
                var body = new object[totalRows + 1, 6];
                body[0, 0] = "工作表"; body[0, 1] = "行"; body[0, 2] = "列";
                body[0, 3] = "类型"; body[0, 4] = "值"; body[0, 5] = "原因/建议";
                int r = 1;
                foreach (var kv in plan)
                {
                    foreach (var cf in kv.Value)
                    {
                        body[r, 0] = sourceSheetName ?? "";
                        body[r, 1] = cf.Row + originRow - 1;
                        body[r, 2] = cf.Col + originCol - 1;
                        body[r, 3] = cf.Type;
                        body[r, 4] = cf.Value;
                        body[r, 5] = cf.Reason + "；" + cf.Suggestion;
                        r++;
                    }
                }

                var writeArgs = new Dictionary<string, object> { ["address"] = "A1", ["sheet"] = name };
                var writeRes = writer != null
                    ? writer(writeArgs, body, "value", "flag_cells")
                    : WriteDirect(sheet, body);
                var recorded = false;
                if (recordStructuralUndo != null && writeRes != null && writeRes.Ok)
                {
                    recordStructuralUndo("flag_cells", new Dictionary<string, object> { ["sheet"] = name }, "add_worksheet", new { sheet = name }, new { rows = totalRows + 1 });
                    recorded = true;
                }
                return (name, totalRows + 1, recorded);
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private ToolResult WriteDirect(dynamic sheet, object[,] body)
        {
            var rows = body.GetLength(0);
            var cols = body.GetLength(1);
            dynamic range = null;
            try
            {
                range = sheet.Range["A1:" + ExcelRangeHelper.CellAddress(rows, cols)];
                range.Value2 = body;
                return new ToolResult { Ok = true };
            }
            finally
            {
                Release((object)range);
            }
        }

        private string UniqueSheetName(string baseName)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var count = Convert.ToInt32(workbook.Worksheets.Count);
                for (var i = 1; i <= count; i++)
                {
                    names.Add(Convert.ToString(workbook.Worksheets.Item[i].Name));
                }
                if (!names.Contains(baseName)) return baseName;
                for (var n = 2; n < 100; n++)
                {
                    var candidate = baseName + n;
                    if (!names.Contains(candidate)) return candidate;
                }
                return baseName + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
            }
            finally
            {
                Release((object)workbook);
            }
        }

        private static object[,] ReadMatrix(dynamic range, int rows, int cols)
        {
            var result = new object[rows, cols];
            object value = null;
            try
            {
                value = range.Value2;
                if (value is object[,] array)
                {
                    for (var r = 0; r < rows; r++)
                    {
                        for (var c = 0; c < cols; c++)
                        {
                            result[r, c] = array[r + array.GetLowerBound(0), c + array.GetLowerBound(1)];
                        }
                    }
                    return result;
                }
                if (rows == 1 && cols == 1)
                {
                    result[0, 0] = value;
                }
                return result;
            }
            finally
            {
                Release((object)value);
            }
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static void Release(object value)
        {
            if (value == null) return;
            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(value))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool IsKnownFlagType(string type)
        {
            return string.Equals(type, "outlier", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "format", StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, "empty", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>默认色板并叠加用户自定义 fill_colors（type→"#RRGGBB"）。任一项非法返回 null，由调用方报 INVALID_FILL_COLOR。</summary>
        private static Dictionary<string, uint> BuildHighlightColors(Dictionary<string, object> args)
        {
            var colors = new Dictionary<string, uint>(HighlightColors, StringComparer.OrdinalIgnoreCase);
            if (args == null || !args.TryGetValue("fill_colors", out var raw) || raw == null)
            {
                return colors;
            }

            var map = ToDict(raw);
            if (map == null)
            {
                return null;
            }

            foreach (var kv in map)
            {
                var hex = kv.Value == null ? null : Convert.ToString(kv.Value);
                if (!TryParseColor(hex, out var color))
                {
                    return null;
                }

                colors[Convert.ToString(kv.Key)] = color;
            }

            return colors;
        }

        private static bool TryParseColor(string hex, out uint bgr)
        {
            bgr = 0;
            if (string.IsNullOrWhiteSpace(hex)) return false;
            var h = hex.Trim().TrimStart('#');
            if (h.Length != 6) return false;
            // 十六进制解析只能用整数 TryParse：double 不支持 AllowHexSpecifier 会抛 ArgumentException
            if (!int.TryParse(h.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r) ||
                !int.TryParse(h.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g) ||
                !int.TryParse(h.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
            {
                return false;
            }

            bgr = (uint)(((uint)b << 16) | ((uint)g << 8) | (uint)r);
            return true;
        }

        private static IDictionary<string, object> ToDict(object raw)
        {
            if (raw is IDictionary<string, object> map)
            {
                return map;
            }

            if (raw is Newtonsoft.Json.Linq.JObject obj)
            {
                return obj.ToObject<Dictionary<string, object>>();
            }

            return null;
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}