using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 公式解释（P6-A / 6.2）：真机读取公式并做结构化拆解（函数→参数→引用值）+ 关联公式追溯。
    /// 支持单格与多格选区（E1/E2/E3）、跨表/跨簿关联公式读取（E4/E5）、环优先深度兜底（E7）。
    /// 只读、不产生写操作；无法真机拆解/取值的项显式标记降级，不让模型对不可靠部分硬解释。
    /// </summary>
    public sealed class ExcelExplainService
    {
        private const int MaxRangeCellsForExpand = 200;
        private const int MaxSelectionCells = 50000;
        private const int MaxFormulaCells = 20;
        private const int DefaultMaxDepth = 3;
        private const int MaxDepthCap = 6;
        private readonly dynamic application;

        public ExcelExplainService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult ExplainFormula(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                return address.Trim().IndexOf(':') >= 0
                    ? ExplainRange(sheet, workbook, address.Trim(), PMaxDepth(args))
                    : ExplainSingleCell(sheet, address.Trim(), PMaxDepth(args));
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>多格选区解释（E1/E2/E3）：无公式格标注「这是无公式格」；>20 公式格只解释激活格；主格=首个有效公式格。</summary>
        private ToolResult ExplainRange(dynamic sheet, dynamic workbook, string address, int maxDepth)
        {
            dynamic range = null;
            try
            {
                range = sheet.Range[address];
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);
                if (rows * (long)cols > MaxSelectionCells)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            sheet = Convert.ToString(sheet.Name),
                            address = address,
                            has_formula = false,
                            note = "选区过大，请缩小到单元格级别再解释"
                        }
                    };
                }

                var formulaCells = new List<Dictionary<string, object>>();
                dynamic cell = null;
                for (var r = 1; r <= rows && formulaCells.Count <= MaxFormulaCells; r++)
                {
                    for (var c = 1; c <= cols && formulaCells.Count <= MaxFormulaCells; c++)
                    {
                        Release((object)cell);
                        cell = range.Cells.Item[r, c];
                        var f = ReadCellFormula(cell);
                        if (!string.IsNullOrWhiteSpace(f) && f.StartsWith("=", StringComparison.Ordinal))
                        {
                            formulaCells.Add(new Dictionary<string, object>
                            {
                                ["address"] = Convert.ToString(cell.Address),
                                ["formula"] = f,
                                ["row"] = r,
                                ["col"] = c
                            });
                        }
                    }
                }

                Release((object)cell);
                Release((object)range);
                range = null;

                if (formulaCells.Count == 0)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            sheet = Convert.ToString(sheet.Name),
                            address = address,
                            has_formula = false,
                            note = "这是无公式格"
                        }
                    };
                }

                var truncated = formulaCells.Count > MaxFormulaCells;
                string activationAddr = null;
                try { activationAddr = Convert.ToString(application.ActiveCell?.Address); }
                catch { activationAddr = null; }

                var targets = new List<Dictionary<string, object>>();
                if (truncated)
                {
                    // 只解释激活格（在选区内且有公式），否则取首个有效公式格
                    var activeCell = formulaCells.Find(f => ActivationWithinSelection(f, activationAddr))
                        ?? formulaCells[0];
                    targets.Add(activeCell);
                }
                else
                {
                    targets.AddRange(formulaCells);
                }

                var cellsView = new List<object>();
                foreach (var t in targets)
                {
                    dynamic targetRange = null;
                    try
                    {
                        targetRange = sheet.Range[Convert.ToString(t["address"])];
                        cellsView.Add(BuildSingleCellView(sheet, workbook, targetRange, maxDepth));
                    }
                    catch (Exception)
                    {
                        cellsView.Add(new { address = Convert.ToString(t["address"]), has_formula = false, note = "无法读取该公式格" });
                    }
                    finally
                    {
                        Release((object)targetRange);
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        address = address,
                        has_formula = true,
                        multi_cell = true,
                        cell_count = formulaCells.Count,
                        truncation = truncated ? $"{truncationMsg(formulaCells.Count, activationAddr)}" : null,
                        cells = cellsView
                    }
                };
            }
            finally
            {
                Release((object)range);
            }
        }

        private bool ActivationWithinSelection(Dictionary<string, object> f, string activationAddr)
        {
            if (string.IsNullOrWhiteSpace(activationAddr)) return false;
            var a = FormulaAnalyzer.CanonicalKey(activationAddr);
            var fAddr = FormulaAnalyzer.CanonicalKey(Convert.ToString(f["address"]));
            return a != null && string.Equals(a, fAddr, StringComparison.OrdinalIgnoreCase);
        }

        private string truncationMsg(int cellCount, string activationAddr)
        {
            return string.IsNullOrWhiteSpace(activationAddr)
                ? $"选区公式过多（共 {cellCount} 个公式格），仅解释首个公式格"
                : $"选区公式过多（共 {cellCount} 个公式格），仅解释激活格 {activationAddr}";
        }

        private ToolResult ExplainSingleCell(dynamic sheet, string address, int maxDepth)
        {
            dynamic range = null;
            try
            {
                range = sheet.Range[address];
                var formula = ReadCellFormula(range);
                if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("=", StringComparison.Ordinal))
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            sheet = Convert.ToString(sheet.Name),
                            address = Convert.ToString(range.Address),
                            has_formula = false,
                            note = "这是无公式格"
                        }
                    };
                }

                var view = BuildSingleCellView(sheet, application.ActiveWorkbook, range, maxDepth);
                return new ToolResult { Ok = true, Data = view };
            }
            finally
            {
                Release((object)range);
            }
        }

        private object BuildSingleCellView(dynamic sheet, dynamic workbook, dynamic range, int maxDepth)
        {
            var formula = ReadCellFormula(range);
            var node = FormulaAnalyzer.ParseTopLevel(formula);
            var argView = new List<object>();
            var degraded = new List<object>();

            // ring 键按工作表隔离，避免跨表同地址误判（E7）
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var self = SheetScopedKey(sheet, Convert.ToString(range.Address));
            if (self != null) visited.Add(self);

            var linkedCount = 0;
            var deepestReached = 0;
            for (var i = 0; i < node.Args.Count; i++)
            {
                var token = node.Args[i];
                var view = new Dictionary<string, object>
                {
                    ["index"] = i + 1,
                    ["kind"] = token.Kind.ToString().ToLowerInvariant(),
                    ["text"] = token.Text
                };
                string readNote;
                object refValue = TryReadReferenceValue(sheet, workbook, token, out readNote);
                if (readNote != null)
                {
                    view["note"] = readNote;
                }
                else
                {
                    view["value"] = refValue;
                }

                var linked = ReadLinkedFormula(sheet, workbook, token, maxDepth, visited, ref deepestReached);
                if (linked != null)
                {
                    view["linked_formula"] = linked;
                    linkedCount++;
                }

                argView.Add(view);

                if (token.Kind == ArgKind.Unparseable || token.Kind == ArgKind.ExternalRef || readNote != null)
                {
                    degraded.Add(new { index = i + 1, kind = token.Kind.ToString().ToLowerInvariant(), text = token.Text, note = readNote ?? "引用无法获取可核对的值" });
                }
            }

            return new
            {
                sheet = Convert.ToString(sheet.Name),
                address = Convert.ToString(range.Address),
                has_formula = true,
                formula = formula,
                func_name = node.Name,
                args = argView,
                degraded = degraded,
                linked = new { had_linked = linkedCount > 0, total_refs = node.Args.Count, deepest_reached = deepestReached }
            };
        }

        /// <summary>对参数尝试真机取值；不可取回 note 并标降级。</summary>
        private object TryReadReferenceValue(dynamic sheet, dynamic workbook, ArgToken token, out string note)
        {
            note = null;
            if (token.Kind == ArgKind.Constant || token.Kind == ArgKind.Expression || token.Kind == ArgKind.NestedFunc)
            {
                return null;
            }

            if (token.Kind == ArgKind.ExternalRef)
            {
                return TryReadExternalValue(token.Text, out note);
            }

            dynamic cell = null;
            try
            {
                cell = sheet.Range[token.Text];
                var value = cell.Value2;
                var matrix = FlattenValues(value, Convert.ToInt32(cell.Rows.Count), Convert.ToInt32(cell.Columns.Count));
                var rows = matrix.GetLength(0);
                var cols = matrix.GetLength(1);
                if (rows * cols <= MaxRangeCellsForExpand)
                {
                    return rows == 1 && cols == 1 ? matrix[0, 0] : matrix;
                }

                var sampleRows = 5;
                var sample = new object[Math.Min(sampleRows, rows), cols];
                for (var r = 0; r < sample.GetLength(0); r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        sample[r, c] = matrix[r, c];
                    }
                }

                note = $"引用区域较大（{rows}行x{cols}列），仅返回前 {sampleRows} 行示例";
                return sample;
            }
            catch (Exception)
            {
                note = "该引用无法真机取值";
                return null;
            }
            finally
            {
                Release((object)cell);
            }
        }

        /// <summary>
        /// 读取一个引用参数的「关联公式」（E4/E5）：CellRef / RangeRef（同表）与 ExternalRef（跨表/跨簿）均读取。
        /// 产出 { 位置(带表名), 公式, 摘要, deeper }；遇环优先置 degraded 不再下钻（E7）。</summary>
        private object ReadLinkedFormula(dynamic sheet, dynamic workbook, ArgToken token, int maxDepth, HashSet<string> visited, ref int deepestReached)
        {
            if (maxDepth < 1) return null;
            if (token.Kind != ArgKind.CellRef && token.Kind != ArgKind.RangeRef && token.Kind != ArgKind.ExternalRef)
            {
                return null;
            }

            dynamic refRange = null;
            dynamic refCell = null;
            try
            {
                if (token.Kind == ArgKind.ExternalRef)
                {
                    refRange = application.Range[token.Text];
                }
                else
                {
                    refRange = sheet.Range[token.Text];
                }

                refCell = token.Kind == ArgKind.RangeRef ? refRange.Cells.Item[1, 1] : refRange;
                return BuildLinked(sheet, refCell, maxDepth, visited, ref deepestReached);
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                Release((object)refCell);
                Release((object)refRange);
            }
        }

        private object BuildLinked(dynamic ownerSheet, dynamic refCell, int maxDepth, HashSet<string> visited, ref int deepestReached)
        {
            string refAddr;
            string refSheetName;
            try { refAddr = Convert.ToString(refCell.Address); refSheetName = Convert.ToString(refCell.Worksheet.Name); }
            catch { return null; }
            if (string.IsNullOrWhiteSpace(refAddr)) return null;

            string refKey = ScopedKey(refSheetName, FormulaAnalyzer.CanonicalKey(refAddr));
            var refFormula = ReadCellFormula(refCell);
            if (string.IsNullOrWhiteSpace(refFormula) || !refFormula.StartsWith("=", StringComparison.Ordinal))
            {
                return null;
            }

            // 环优先于深度：该被引用格已在访问链中 → 判环并给出触发位置（E7）
            if (refKey != null && visited.Contains(refKey))
            {
                return new
                {
                    pos = QualifyAddress(refSheetName, refAddr, null),
                    formula = refFormula,
                    summary = "循环/自引用，无法继续下钻",
                    degraded = "循环引用或自引用（触发位置：" + QualifyAddress(refSheetName, refAddr, null) + "）"
                };
            }

            // 间接环（2 跳+）：被引用公式所引用的任一格回落到已访问链 → 判环（E7）
            var indirect = FindVisitedTarget(refFormula, refSheetName, visited);
            if (indirect != null)
            {
                return new
                {
                    pos = QualifyAddress(refSheetName, refAddr, null),
                    formula = refFormula,
                    summary = "循环/自引用，无法继续下钻",
                    degraded = "循环引用或自引用（触发位置：" + indirect + "）"
                };
            }

            var child = refKey == null ? visited : new HashSet<string>(visited) { refKey };
            var summary = SummarizeLinked(refFormula, child, maxDepth - 1, refSheetName, ref deepestReached);
            deepestReached = Math.Max(deepestReached, 1);

            // pos 带工作表名（跨表可定位）；跨簿时由 range 工作表名体现
            return new
            {
                pos = QualifyAddress(refSheetName, refAddr, MaxBookScope(refCell)),
                formula = refFormula,
                summary,
                deeper = maxDepth > 1
            };
        }

        /// <summary>对被引用公式做轻量摘要；仅当还有深度余量时探测下一层关联公式并标注。</summary>
        private object SummarizeLinked(string formula, HashSet<string> visited, int depthLeft, string refSheetName, ref int deepestReached)
        {
            var node = FormulaAnalyzer.ParseTopLevel(formula);
            var text = node.Name == null ? "裸表达式" : $"{node.Name}({node.Args.Count} 个参数)";

            if (depthLeft > 0)
            {
                var deeperFound = false;
                foreach (var arg in node.Args)
                {
                    if (arg.Kind != ArgKind.CellRef && arg.Kind != ArgKind.RangeRef && arg.Kind != ArgKind.ExternalRef) continue;
                    dynamic range = null;
                    dynamic cell = null;
                    try
                    {
                        range = arg.Kind == ArgKind.ExternalRef ? application.Range[arg.Text] : application.Range[RefTextOnSheet(refSheetName, arg)];
                        cell = arg.Kind == ArgKind.RangeRef ? range.Cells.Item[1, 1] : range;
                        var f = ReadCellFormula(cell);
                        var key = ScopedKey(Convert.ToString(cell.Worksheet.Name), FormulaAnalyzer.CanonicalKey(CellText(cell.Address)));
                        if (!string.IsNullOrWhiteSpace(f) && f.StartsWith("=", StringComparison.Ordinal) && key != null && !visited.Contains(key))
                        {
                            deeperFound = true;
                            deepestReached = Math.Max(deepestReached, 2);
                            break;
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        Release((object)cell);
                        Release((object)range);
                    }
                }

                if (deeperFound) text += "（含下一层关联公式，可继续追问展开）";
            }

            return text;
        }

        /// <summary>把同表的相对引用文本当作跨表读取（供下钻在目标表上解析）。</summary>
        private string RefTextOnSheet(string sheetName, ArgToken arg)
        {
            if (arg.Kind == ArgKind.ExternalRef) return arg.Text;
            return QualifyAddress(sheetName, arg.Text, null);
        }

        /// <summary>解析被引用公式内的引用，若任一引用回落到已访问链（表!规范地址），返回其带表名的定位。</summary>
        private static string FindVisitedTarget(string formula, string refSheetName, HashSet<string> visited)
        {
            if (visited == null || visited.Count == 0 || string.IsNullOrWhiteSpace(formula)) return null;
            var refs = FormulaAnalyzer.ExtractCellRefs(formula);
            foreach (var r in refs)
            {
                if (string.IsNullOrEmpty(r)) continue;
                var (sheet, addr) = SplitSheetAddr(r, refSheetName);
                if (string.IsNullOrEmpty(addr)) continue;
                var canonical = FormulaAnalyzer.CanonicalKey(addr);
                if (canonical == null) continue;
                if (visited.Contains(sheet + "!" + canonical))
                {
                    return QualifyAddress(sheet, addr, null);
                }
            }

            return null;
        }

        /// <summary>把引用文本拆成 表名+地址：带 ! 的取最后一段为地址、前段去书签/引号当表名；否则默认落在当前表。</summary>
        private static (string Sheet, string Addr) SplitSheetAddr(string refText, string fallbackSheet)
        {
            var idx = refText.LastIndexOf('!');
            if (idx < 0) return (fallbackSheet ?? string.Empty, refText);
            var sheet = refText.Substring(0, idx);
            var addr = refText.Substring(idx + 1);
            var book = sheet.IndexOf('[');
            if (book >= 0)
            {
                var bookClose = sheet.IndexOf(']', book);
                if (bookClose > book) sheet = sheet.Substring(bookClose + 1);
            }

            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
            {
                sheet = sheet.Substring(1, sheet.Length - 2);
            }

            return (sheet, addr);
        }

        private static int PMaxDepth(Dictionary<string, object> args)
        {
            if (args != null && args.TryGetValue("max_depth", out var v) && v != null
                && int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out var d))
            {
                return Math.Max(1, Math.Min(MaxDepthCap, d));
            }

            return DefaultMaxDepth;
        }

        private object TryReadExternalValue(string text, out string note)
        {
            note = null;
            dynamic range = null;
            try
            {
                range = application.Range[text];
                var value = range.Value2;
                var matrix = FlattenValues(value, Convert.ToInt32(range.Rows.Count), Convert.ToInt32(range.Columns.Count));
                return matrix.GetLength(0) == 1 && matrix.GetLength(1) == 1
                    ? matrix[0, 0]
                    : matrix;
            }
            catch (Exception)
            {
                note = "跨表/跨簿引用无法真机取值";
                return null;
            }
            finally
            {
                Release((object)range);
            }
        }

        private static object[,] FlattenValues(object source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }

                return result;
            }

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ?? throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static string ReadCellFormula(dynamic cell)
        {
            try { return CellText(cell.Formula); }
            catch { return null; }
        }

        /// <summary>表作用域环键：表名!规范地址，跨表同地址不会误判成环。</summary>
        private static string SheetScopedKey(dynamic sheet, string addr)
        {
            string name;
            try { name = Convert.ToString(sheet.Name); }
            catch { name = null; }
            return ScopedKey(name, FormulaAnalyzer.CanonicalKey(addr));
        }

        private static string ScopedKey(string sheetName, string canonical)
        {
            if (canonical == null) return null;
            return (sheetName ?? string.Empty) + "!" + canonical;
        }

        /// <summary>地址带上工作表名（需要时加引号），供模型/用户跨表定位（E4）。</summary>
        private static string QualifyAddress(string sheetName, string addr, string bookPrefix)
        {
            var clean = (addr ?? string.Empty).Replace("$", "");
            var sheet = QuoteIfNeeded(sheetName);
            return (bookPrefix ?? string.Empty) + sheet + "!" + clean;
        }

        private static string QuoteIfNeeded(string name)
        {
            if (string.IsNullOrEmpty(name)) return "?";
            return name.IndexOfAny(new[] { ' ', '-', '(', ')', '!', '\'', '&', '#' }) >= 0
                ? "'" + name + "'"
                : name;
        }

        private static string MaxBookScope(dynamic cell)
        {
            try
            {
                var wb = cell.Worksheet.Parent;
                var bookName = Convert.ToString(wb.Name);
                return string.IsNullOrWhiteSpace(bookName) ? null : "[" + bookName + "]";
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string CellText(object value)
        {
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture);
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

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private static void Release(object value)
        {
            if (value == null)
            {
                return;
            }

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
    }
}