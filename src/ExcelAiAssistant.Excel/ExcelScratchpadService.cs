using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// AI 计算草稿：在专用工作表（默认 _AI_CALC）上追加输入/公式/期望输出并回填计算结果。
    /// </summary>
    public sealed class ExcelScratchpadService
    {
        public const string DefaultSheetName = "_AI_CALC";

        private static readonly string[] Headers =
        {
            "序号", "输入", "公式", "期望输出", "实际结果", "备注"
        };

        private readonly dynamic application;

        public ExcelScratchpadService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult Append(Dictionary<string, object> args)
        {
            var sheetName = Optional(args, "scratch_sheet") ?? DefaultSheetName;
            var input = Optional(args, "input");
            var formula = Optional(args, "formula");
            var expected = Optional(args, "expected_output");
            var note = Optional(args, "note");
            if (string.IsNullOrEmpty(input) && string.IsNullOrEmpty(formula) && string.IsNullOrEmpty(expected))
            {
                return Invalid("EMPTY_ENTRY", "请至少提供 input、formula 或 expected_output 之一。");
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic used = null;
            dynamic formulaCell = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                // 草稿表可能尚不存在：WPS/Excel 下 Worksheets.Item[不存在名] 直接抛
                // DISP_E_BADINDEX（M6 WPS 全量审查 F9 实测发现），故先查后建。
                sheet = EnsureSheet(workbook, sheetName);
                used = sheet.UsedRange;
                var lastRow = Math.Max(Convert.ToInt32(used.Rows.Count), 0);
                var empty = lastRow == 0 || string.IsNullOrWhiteSpace(Convert.ToString(sheet.Cells[1, 1].Value2));
                if (empty)
                {
                    WriteRow(sheet, 1, "序号", "输入", "公式", "期望输出", "实际结果", "备注");
                    lastRow = 1;
                }
                else
                {
                    var hasHeader = string.Equals(Convert.ToString(sheet.Cells[1, 2].Value2), "输入", StringComparison.OrdinalIgnoreCase);
                    if (hasHeader)
                    {
                        lastRow = used.Row + used.Rows.Count - 1;
                    }
                }

                var row = lastRow + 1;
                var seq = row - 1;
                WriteRow(sheet, row, Convert.ToString(seq), input, formula, expected, null, note);

                object result = null;
                if (!string.IsNullOrEmpty(formula))
                {
                    formulaCell = sheet.Cells[row, 5];
                    var candidate = formula.StartsWith("=", StringComparison.Ordinal) ? formula : "=" + formula;
                    formulaCell.Formula = candidate;
                    result = NormalizeValue(formulaCell.Value2);
                }
                else if (!string.IsNullOrEmpty(input))
                {
                    result = NormalizeValue(input);
                }

                return new ToolResult
                {
                    Ok = true,
                    AffectedRanges = { $"{sheetName}!A{row}:F{row}" },
                    Data = new
                    {
                        sheet = sheetName,
                        row,
                        seq,
                        input,
                        formula,
                        expected_output = expected,
                        computed = result
                    }
                };
            }
            finally
            {
                Release((object)formulaCell);
                Release((object)used);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        public ToolResult Clear(Dictionary<string, object> args)
        {
            var sheetName = Optional(args, "scratch_sheet") ?? DefaultSheetName;
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic used = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var existing = FindSheet(workbook, sheetName);
                if (existing == null)
                {
                    // 草稿表不存在即无内容可清，直接返回成功（此前抛 BADINDEX）。
                    return new ToolResult { Ok = true, Data = new { sheet = sheetName, cleared = false } };
                }

                sheet = existing;
                used = sheet.UsedRange;
                if (used != null && Convert.ToInt32(used.Rows.Count) > 0)
                {
                    used.ClearContents();
                }

                return new ToolResult { Ok = true, Data = new { sheet = sheetName, cleared = true } };
            }
            finally
            {
                Release((object)used);
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        private static object NormalizeValue(object value)
        {
            return value is Newtonsoft.Json.Linq.JValue j ? j.Value : value;
        }

        private static dynamic FindSheet(dynamic workbook, string sheetName)
        {
            foreach (dynamic s in workbook.Worksheets)
            {
                try
                {
                    if (string.Equals(Convert.ToString(s.Name), sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return s; // 命中项所有权移交调用方
                    }

                    ReleaseCom(s); // 未命中项 RCW 释放（M16 全量审查）
                }
                catch (Exception)
                {
                    ReleaseCom(s);
                }
            }

            return null;
        }

        /// <summary>自有 RCW 安全释放（失败不阻断收尾）。</summary>
        private static void ReleaseCom(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(comObject))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comObject);
                }
            }
            catch (Exception)
            {
            }
        }

        private dynamic EnsureSheet(dynamic workbook, string sheetName)
        {
            var existing = FindSheet(workbook, sheetName);
            if (existing != null)
            {
                return existing;
            }

            var created = workbook.Worksheets.Add();
            try
            {
                created.Name = sheetName;
            }
            catch (Exception)
            {
                // 重命名失败（同名边缘冲突）时返回新建表本身，写入仍可进行。
            }

            return created;
        }

        private static void WriteRow(dynamic sheet, int row, object seq, object input, object formula, object expected, object result, object note)
        {
            var col = 1;
            foreach (var value in new[] { seq, input, formula, expected, result, note })
            {
                dynamic cell = null;
                try
                {
                    cell = sheet.Cells[row, col];
                    if (col == 3 && value is string text && text.StartsWith("=", StringComparison.Ordinal))
                    {
                        cell.Value2 = "'" + text; // 以文本形式记录公式，不执行
                    }
                    else
                    {
                        cell.Value2 = value;
                    }
                }
                finally
                {
                    Release((object)cell);
                }

                col++;
            }
        }

        private static string Optional(Dictionary<string, object> args, string key)
        {
            return args.TryGetValue(key, out var value) && value != null
                ? Convert.ToString(value)
                : null;
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }
    }
}