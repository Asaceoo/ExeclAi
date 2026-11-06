using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 写入/结构级工具：命名区域、行列插入删除、区域复制粘贴、生成目录页。
    /// </summary>
    public sealed class ExcelEditService
    {
        private readonly dynamic application;
        private static readonly Regex ColumnRange = new Regex(@"^[A-Za-z]+:[A-Za-z]+$", RegexOptions.Compiled);
        private static readonly Regex RowRange = new Regex(@"^\d+:\d+$", RegexOptions.Compiled);
        private OperationRecorder recorder;

        public ExcelEditService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void SetRecorder(OperationRecorder value)
        {
            recorder = value;
        }

        public ToolResult ManageNamedRange(Dictionary<string, object> args)
        {
            var action = Required(args, "action");
            if (!IsValidAction(action))
            {
                return Invalid("INVALID_ACTION", $"action 仅支持 query/create/update/delete，当前为 {action}。");
            }

            var name = Required(args, "name");
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                if (string.Equals(action, "query", StringComparison.OrdinalIgnoreCase))
                {
                    var found = FindName(workbook, name);
                    try
                    {
                        return new ToolResult
                        {
                            Ok = true,
                            Data = found == null
                                ? new { name, exists = false }
                                : new { name, exists = true, refers_to = found.RefersTo, scope = ScopeOf(Convert.ToString(found.Name), name) }
                        };
                    }
                    finally
                    {
                        Release((object)found); // M16：命中的 Name RCW 释放
                    }
                }

                if (string.Equals(action, "delete", StringComparison.OrdinalIgnoreCase))
                {
                    var found = FindName(workbook, name);
                    if (found == null)
                    {
                        return Invalid("NOT_FOUND", $"命名区域不存在：{name}");
                    }

                    try
                    {
                        found.Delete();
                        return new ToolResult { Ok = true, Data = new { name, deleted = true } };
                    }
                    finally
                    {
                        Release((object)found); // M16：命中 Name RCW 释放
                    }
                }

                // create / update
                var reference = Required(args, "reference");
                string sheetName;
                var inline = Optional(args, "sheet");
                var refSheet = ExtractSheet(reference);
                if (!string.IsNullOrEmpty(inline))
                {
                    sheetName = inline;
                    if (string.IsNullOrEmpty(refSheet))
                    {
                        Resolve(ref workbook, ref sheet, sheetName);
                    }
                }
                else
                {
                    sheetName = refSheet ?? Convert.ToString(application.ActiveSheet?.Name ?? "Sheet1");
                }

                Resolve(ref workbook, ref sheet, sheetName);
                var absolute = ExcelRangeHelper.ToAbsolute(reference);
                if (string.IsNullOrEmpty(absolute))
                {
                    return Invalid("INVALID_REFERENCE", $"引用区域地址无效：{reference}");
                }

                var refersTo = $"='{EscapeName(sheetName)}'!{absolute}";
                var existing = FindName(workbook, name);
                if (existing == null)
                {
                    if (string.Equals(action, "update", StringComparison.OrdinalIgnoreCase))
                    {
                        return Invalid("NOT_FOUND", $"命名区域不存在：{name}");
                    }

                    workbook.Names.Add(name, refersTo);
                    return new ToolResult { Ok = true, Data = new { name, created = true, refers_to = refersTo } };
                }

                if (string.Equals(action, "create", StringComparison.OrdinalIgnoreCase))
                {
                    Release((object)existing); // M16：命中 Name RCW 释放
                    return Invalid("ALREADY_EXISTS", $"命名区域已存在：{name}，请使用 update 进行更新。");
                }

                existing.RefersTo = refersTo;
                Release((object)existing); // M16：命中 Name RCW 释放
                return new ToolResult { Ok = true, Data = new { name, updated = true, refers_to = refersTo } };
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult InsertDimension(Dictionary<string, object> args)
        {
            return ModifyDimension(args, insert: true);
        }

        public ToolResult DeleteDimension(Dictionary<string, object> args)
        {
            return ModifyDimension(args, insert: false);
        }

        private ToolResult ModifyDimension(Dictionary<string, object> args, bool insert)
        {
            var address = Required(args, "address");
            var quantity = GetInt(args, "quantity", 1);
            if (quantity < 1 || quantity > 2000)
            {
                return Invalid("INVALID_QUANTITY", $"quantity 必须介于 1 到 2000，当前为 {quantity}。");
            }

            var dimension = Optional(args, "dimension") ?? InferDimension(address);
            if (dimension == null)
            {
                return Invalid("INVALID_DIMENSION", $"无法从地址推断行列方向：{address}，请传入 dimension 参数。");
            }

            if (!string.Equals(dimension, "row", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(dimension, "column", StringComparison.OrdinalIgnoreCase))
            {
                return Invalid("INVALID_DIMENSION", $"dimension 仅支持 row/column，当前为 {dimension}。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                for (var i = 0; i < quantity; i++)
                {
                    range = sheet.Range[address];
                    if (string.Equals(dimension, "row", StringComparison.OrdinalIgnoreCase))
                    {
                        var whole = range.EntireRow;
                        if (insert)
                        {
                            whole.Insert();
                        }
                        else
                        {
                            whole.Delete();
                        }

                        Release((object)whole);
                    }
                    else
                    {
                        var whole = range.EntireColumn;
                        if (insert)
                        {
                            whole.Insert();
                        }
                        else
                        {
                            whole.Delete();
                        }

                        Release((object)whole);
                    }

                    Release((object)range);
                    range = null;
                }

                if (insert && recorder != null)
                {
                    var sheetName = Convert.ToString(sheet.Name);
                    recorder(
                        "insert_excel_dimension",
                        args,
                        "dimension-insert",
                        new { sheet = sheetName, dimension, address, quantity },
                        new DimensionInsertUndoPayload
                        {
                            Sheet = sheetName,
                            Dimension = dimension,
                            Address = address,
                            Count = quantity
                        });
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { dimension, quantity, inserted = insert, deleted = !insert }
                };
            }
            finally
            {
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult CopyPasteRange(Dictionary<string, object> args)
        {
            var sourceAddress = Required(args, "source_address");
            var targetAddress = Required(args, "target_address");
            var pasteValues = GetBool(args, "paste_values", false);
            dynamic workbook = null;
            dynamic sourceSheet = null;
            dynamic targetSheet = null;
            dynamic source = null;
            dynamic target = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sourceSheet = GetSheet(workbook, Optional(args, "source_sheet"));
                source = sourceSheet.Range[sourceAddress];

                // 先捕获源地址为纯字符串，避免后续清剪贴板导致 RCW 失效后再次触碰 COM。
                var sourceAddressFinal = Convert.ToString(source.Address);

                source.Copy();
                // 拷贝后重新解析目标工作表与目标区域，降低 RCW 跨剪贴板操作失效的风险。
                targetSheet = GetSheet(workbook, Optional(args, "target_sheet"));
                target = targetSheet.Range[targetAddress];
                if (pasteValues)
                {
                    target.PasteSpecial(-4163, Type.Missing, Type.Missing, Type.Missing); // xlPasteValues
                }
                else
                {
                    target.PasteSpecial(-4104, Type.Missing, Type.Missing, Type.Missing); // xlPasteAll
                }

                // 立即捕获目标信息为纯字符串后再清理剪贴板。
                var targetAddressFinal = Convert.ToString(target.Address);
                var targetSheetNameFinal = Convert.ToString(targetSheet.Name);

                try
                {
                    application.CutCopyMode = false;
                }
                catch (Exception)
                {
                    // 忽略剪贴板异常。
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{targetSheetNameFinal}!{targetAddressFinal}" },
                    Data = new
                    {
                        target_address = targetAddressFinal,
                        paste_values = pasteValues,
                        source_address = sourceAddressFinal
                    }
                };
            }
            finally
            {
                Release((object)target);
                Release((object)source);
                Release((object)targetSheet);
                Release((object)sourceSheet);
                Release((object)workbook);
            }
        }

        public ToolResult CreateToc(Dictionary<string, object> args)
        {
            var tocName = Optional(args, "sheet") ?? "目录";
            var includeHidden = GetBool(args, "include_hidden", false);
            dynamic workbook = null;
            dynamic toc = null;
            var addedHyperlinks = new List<dynamic>();
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                toc = FindSheet(workbook, tocName) ?? workbook.Worksheets.Add();
                if (!Convert.ToString(toc.Name).Equals(tocName, StringComparison.Ordinal))
                {
                    toc.Name = tocName;
                }

                var sheetNames = new List<string>();
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    var name = Convert.ToString(worksheet.Name);
                    if (name.Equals(tocName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var hidden = Convert.ToInt32(worksheet.Visible) != -1; // -1=xlSheetVisible；0=Hidden、2=VeryHidden 均算隐藏
                    if (hidden && !includeHidden)
                    {
                        continue;
                    }

                    sheetNames.Add(name);
                }

                var headers = new[] { "工作表", "UsedRange", "行数", "列数" };
                for (var column = 0; column < headers.Length; column++)
                {
                    dynamic cell = toc.Cells[1, column + 1];
                    cell.Value2 = headers[column];
                    Release((object)cell);
                }

                var index = 2;
                foreach (var name in sheetNames)
                {
                    dynamic worksheet = null;
                    dynamic usedRange = null;
                    dynamic nameCell = null;
                    try
                    {
                        worksheet = workbook.Worksheets.Item[name];
                        usedRange = worksheet.UsedRange;
                        var usedAddress = Convert.ToString(usedRange?.Address ?? "None");
                        var usedRows = Convert.ToInt32(usedRange?.Rows.Count ?? 0);
                        var usedColumns = Convert.ToInt32(usedRange?.Columns.Count ?? 0);

                        nameCell = toc.Cells[index, 1];
                        var linkRange = toc.Range[nameCell, nameCell];
                        linkRange.Value2 = name;
                        var hyperlink = AddHyperlink(toc, linkRange, name);
                        if (hyperlink != null)
                        {
                            addedHyperlinks.Add(hyperlink);
                        }

                        Release((object)linkRange);
                        SetCell(toc, index, 2, usedAddress);
                        SetCell(toc, index, 3, usedRows);
                        SetCell(toc, index, 4, usedColumns);
                    }
                    finally
                    {
                        Release((object)nameCell);
                        Release((object)usedRange);
                        Release((object)worksheet);
                    }

                    index++;
                }

                var headerRange = toc.Range[$"A1:{ExcelRangeHelper.ColumnName(headers.Length)}1"];
                if (headerRange != null)
                {
                    headerRange.Font.Bold = true;
                    Release((object)headerRange);
                }

                dynamic tocColumns = null;
                try
                {
                    tocColumns = toc.Columns;
                    tocColumns.AutoFit();
                }
                finally
                {
                    Release((object)tocColumns);
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{Convert.ToString(toc.Name)}!A1:{ExcelRangeHelper.ColumnName(headers.Length)}{Math.Max(index - 1, 1)}" },
                    Data = new
                    {
                        sheet = Convert.ToString(toc.Name),
                        sheet_count = sheetNames.Count,
                        link_count = addedHyperlinks.Count
                    }
                };
            }
            finally
            {
                foreach (var hyperlink in addedHyperlinks)
                {
                    Release((object)hyperlink);
                }

                Release((object)toc);
                Release((object)workbook);
            }
        }

        private static void SetCell(dynamic sheet, int row, int column, object value)
        {
            dynamic cell = null;
            try
            {
                cell = sheet.Cells[row, column];
                cell.Value2 = value;
            }
            finally
            {
                Release((object)cell);
            }
        }

        private dynamic AddHyperlink(dynamic sheet, dynamic range, string targetSheet)
        {
            dynamic hyperlinks = null;
            try
            {
                hyperlinks = sheet.Hyperlinks;
                var subAddress = $"'{EscapeName(targetSheet)}'!A1";
                return hyperlinks.Add(range, string.Empty, subAddress);
            }
            catch (Exception)
            {
                // 链接添加失败时降级为纯文本。
                return null;
            }
            finally
            {
                Release((object)hyperlinks);
            }
        }

        private dynamic FindName(dynamic workbook, string name)
        {
            foreach (dynamic named in workbook.Names)
            {
                try
                {
                    var current = Convert.ToString(named.Name);
                    if (string.Equals(current, name, StringComparison.OrdinalIgnoreCase) ||
                        current.EndsWith("!" + name, StringComparison.OrdinalIgnoreCase))
                    {
                        return named; // 命中项所有权移交调用方（调用方负责释放）
                    }

                    // 未命中项 RCW 释放（M16 全量审查：原实现泄漏，每个名称一个）。
                    Release((object)named);
                }
                catch (Exception)
                {
                    Release((object)named);
                }
            }

            return null;
        }

        private static string ScopeOf(string fullName, string shortName)
        {
            return fullName.EndsWith("!" + shortName, StringComparison.OrdinalIgnoreCase)
                ? fullName.Substring(0, fullName.Length - shortName.Length - 1)
                : "workbook";
        }

        private dynamic FindSheet(dynamic workbook, string name)
        {
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                try
                {
                    if (string.Equals(Convert.ToString(worksheet.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet; // 命中项所有权移交调用方
                    }

                    Release((object)worksheet);
                }
                catch (Exception)
                {
                    Release((object)worksheet);
                }
            }

            return null;
        }

        private static bool IsValidAction(string action)
        {
            return string.Equals(action, "query", StringComparison.OrdinalIgnoreCase)
                || string.Equals(action, "create", StringComparison.OrdinalIgnoreCase)
                || string.Equals(action, "update", StringComparison.OrdinalIgnoreCase)
                || string.Equals(action, "delete", StringComparison.OrdinalIgnoreCase);
        }

        private static string InferDimension(string address)
        {
            var trimmed = address.Trim();
            if (ColumnRange.IsMatch(trimmed))
            {
                return "column";
            }

            if (RowRange.IsMatch(trimmed))
            {
                return "row";
            }

            return null;
        }

        private static string ExtractSheet(string reference)
        {
            var separator = reference.LastIndexOf('!');
            if (separator <= 0)
            {
                return null;
            }

            return reference.Substring(0, separator).Trim('\'', '"');
        }

        private static string EscapeName(string name)
        {
            return name.Replace("'", "''");
        }

        private dynamic GetSheet(dynamic workbook, string sheetName)
        {
            return string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = GetSheet(workbook, sheetName);
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = Optional(args, key);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            var raw = value is Newtonsoft.Json.Linq.JValue j ? j.Value : value;
            if (raw is int iv)
            {
                return iv;
            }

            if (raw is long lv)
            {
                return unchecked((int)lv);
            }

            // 非整数（如 "abc"、小数文本）不应抛异常打断工具，回退默认值让调用方走 INVALID_* 分类。
            return int.TryParse(Convert.ToString(raw), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : defaultValue;
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToBoolean(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
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
                // 释放阶段的 RCW 清理失败不应掩盖已经成功的业务操作。
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}