using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// C2 报告生成：读取区域 → ReportTextBuilder 组装 → 按输出模式写入工作表或导出 MD/HTML 文件。
    /// </summary>
    public sealed class ExcelReportService
    {
        private const int MaxAnalyzedCells = 100000;

        private readonly dynamic application;

        public ExcelReportService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult GenerateReport(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var output = GetString(args, "output", "worksheet").ToLowerInvariant();
            if (output != "worksheet" && output != "markdown" && output != "html")
            {
                return new ToolResult { Ok = false, ErrorCode = "INVALID_OUTPUT", Data = $"output 仅支持 worksheet/markdown/html，当前为 {output}。" };
            }

            var sheetName = Optional(args, "sheet_name") ?? "分析报告";
            // 工具分类审查第七批修复（第四批同款缺陷家族残留）：sheet_name 直达 newSheet.Name 无预校验——
            // 超长/非法字符抛裸 COM 异常。复用第四批 ValidateSheetName 人话报错。
            var nameError = Excel.ExcelStructureService.ValidateSheetName(sheetName);
            if (nameError != null)
            {
                return new ToolResult { Ok = false, ErrorCode = "INVALID_SHEET_NAME", Data = nameError };
            }
            // 参数 Clamp 护栏：anomaly_std/correlation_threshold/max_columns 越界会让报告异常检测失效
            // （std=0 全标异常、threshold>1 相关性永不输出、max_columns=0 除零风险）。
            var anomalyStd = Math.Min(Math.Max(GetInt(args, "anomaly_std", 2), 1), 5);
            var outputPath = Optional(args, "output_path");
            var maxColumns = Math.Min(Math.Max(GetInt(args, "max_columns", 20), 1), 50);
            var correlationThreshold = Math.Min(Math.Max(GetDouble(args, "correlation_threshold", 0.7), 0.1), 1.0);

            // 章节白名单：逗号/顿号分隔的章节名（概览/趋势/异常/重点列摘要/相关性/建议）；缺省全部。
            List<string> sectionFilter = null;
            var sectionsRaw = GetString(args, "sections", null);
            if (!string.IsNullOrWhiteSpace(sectionsRaw))
            {
                sectionFilter = sectionsRaw
                    .Split(new[] { ',', '，', '、', ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();
            }

            var read = ReadMatrix(args, address);
            var report = ReportTextBuilder.Build(read.Matrix, read.Sheet, read.Address, anomalyStd, maxColumns, correlationThreshold, sectionFilter);
            var sections = report.Sections.ConvertAll(section => new
            {
                title = section.Title,
                lines = section.Lines.ToArray()
            });
            // P0-2（Q6 拍板）：生成后自检回显（数字对账/空段/占位符/覆盖匹配），对齐图表 auto_check 口径。
            var selfCheck = new
            {
                verdict = report.SelfCheck.Verdict,
                issues = report.SelfCheck.Issues.ToArray(),
                checked_sections = report.SelfCheck.CheckedCount
            };

            switch (output)
            {
                case "markdown":
                {
                    var filePath = WriteFileIfRequested(outputPath ?? DefaultReportPath(".md"), ".md", report.Markdown);
                    return new ToolResult
                    {
                        Ok = true,
                        Truncated = read.Truncated,
                        Data = new
                        {
                            mode = "markdown",
                            file_path = filePath,
                            self_check = selfCheck,
                            markdown = report.Markdown,
                            word_count = report.WordCount,
                            sections
                        }
                    };
                }

                case "html":
                {
                    var filePath = WriteHtml(outputPath, report.Html);
                    return new ToolResult
                    {
                        Ok = true,
                        Truncated = read.Truncated,
                        Data = new
                        {
                            mode = "html",
                            file_path = filePath,
                            self_check = selfCheck,
                            word_count = report.WordCount,
                            sections
                        }
                    };
                }

                default:
                {
                    var result = WriteWorksheet(read.Sheet, sheetName, report.ReportGrid);
                    return new ToolResult
                    {
                        Ok = true,
                        AffectedRanges = { result.SheetRef },
                        Truncated = read.Truncated,
                        Data = new
                        {
                            mode = "worksheet",
                            worksheet_name = result.SheetName,
                            self_check = selfCheck,
                            word_count = report.WordCount,
                            sections
                        }
                    };
                }
            }
        }

        private WorksheetWriteResult WriteWorksheet(string sourceSheet, string sheetName, object[,] grid)
        {
            dynamic workbook = null;
            dynamic newSheet = null;
            dynamic cell = null;
            dynamic range = null;
            dynamic columns = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                newSheet = workbook.Sheets.Add();
                var uniqueName = EnsureUniqueSheetName(workbook, sheetName);
                newSheet.Name = uniqueName;
                cell = newSheet.Cells[1, 1];
                range = newSheet.Range[cell, newSheet.Cells[grid.GetLength(0), grid.GetLength(1)]];
                range.Value2 = grid;
                // range.Columns 是链式行尾属性，若不捕获会被编译器丢弃中间 RCW 引用，泄漏整列 RCW。
                columns = range.Columns;
                columns.AutoFit();
                return new WorksheetWriteResult
                {
                    SheetName = uniqueName,
                    SheetRef = uniqueName + "!A1"
                };
            }
            finally
            {
                Release((object)columns);
                Release((object)range);
                Release((object)cell);
                Release((object)newSheet);
                Release((object)workbook);
            }
        }

        private static string EnsureUniqueSheetName(dynamic workbook, string baseName)
        {
            for (var attempt = 1; ; attempt++)
            {
                var candidate = attempt == 1 ? baseName : baseName + " (" + attempt + ")";
                if (HasSheet(workbook, candidate) == false)
                {
                    return candidate;
                }
            }
        }

        private static bool? HasSheet(dynamic workbook, string name)
        {
            try
            {
                var sheet = workbook.Sheets.Item[name];
                Release((object)sheet);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string WriteFileIfRequested(string requestedPath, string extension, string content)
        {
            if (string.IsNullOrWhiteSpace(requestedPath))
            {
                return null;
            }

            var path = ResolveFile(requestedPath, extension);
            WriteGuarded(path, content);
            return path;
        }

        private string WriteHtml(string requestedPath, string html)
        {
            var path = ResolveFile(requestedPath ?? DefaultReportPath(".html"), ".html");
            WriteGuarded(path, html);
            return path;
        }

        /// <summary>
        /// 报告落盘统一入口：目标已存在时先备份（可人工恢复），备份失败即中止——
        /// 旧实现直接 File.WriteAllText，模型被提示注入引导覆盖既有文档会静默截断且不可恢复。
        /// </summary>
        private static void WriteGuarded(string path, string content)
        {
            if (File.Exists(path))
            {
                try
                {
                    ExcelAiAssistant.Core.Security.PathGuard.CreateBackup(path);
                }
                catch (Exception backupException)
                {
                    throw new InvalidOperationException(
                        "目标文件已存在且覆盖前备份失败，已中止写入：" + backupException.Message, backupException);
                }
            }

            File.WriteAllText(path, content, new System.Text.UTF8Encoding(true));
        }

        private static string DefaultReportPath(string extension)
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var folder = Path.Combine(documents, "XAI报告");
            Directory.CreateDirectory(folder);
            return Path.Combine(folder,
                "分析报告_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension);
        }

        private static string ResolveFile(string rawPath, string extension)
        {
            // 防御性守门：拒绝对 NUL / 穿越逃逸的模型可控输出路径
            var path = ExcelAiAssistant.Core.Security.PathGuard.GuardAndComplete(rawPath, extension);
            var root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var full = Path.IsPathRooted(path) ? path : Path.Combine(root, path);
            var directory = Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(directory))
            {
                throw new ArgumentException("输出路径无效。");
            }

            Directory.CreateDirectory(directory);
            return full;
        }

        private ReadResult ReadMatrix(Dictionary<string, object> args, string address)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic value = null;
            dynamic sub = null;
            dynamic topLeft = null;
            dynamic bottomRight = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];
                range = sheet.Range[address];
                var totalRows = Convert.ToInt32(range.Rows.Count);
                var columns = Convert.ToInt32(range.Columns.Count);
                // long 乘法杜绝 int 溢出（否则整列/整表乘积回绕会绕过截断守卫直读全表）
                var truncated = totalRows * (long)columns > MaxAnalyzedCells;
                var rows = totalRows;
                if (truncated && columns > 0)
                {
                    // 真正截断：只读前 MaxAnalyzedCells 个单元对应的行，而非把整表 Value2 拉进内存
                    rows = Math.Max(1, (int)(MaxAnalyzedCells / columns));
                }

                if (rows >= totalRows)
                {
                    value = range.Value2;
                }
                else
                {
                    topLeft = sheet.Cells[range.Row, range.Column];
                    bottomRight = sheet.Cells[range.Row + rows - 1, range.Column + columns - 1];
                    sub = sheet.Range[topLeft, bottomRight];
                    value = sub.Value2;
                }

                var matrix = ExcelToolHostHelpers.FlattenValues(value, rows, columns);
                return new ReadResult
                {
                    Sheet = Convert.ToString(sheet.Name),
                    Address = Convert.ToString(range.Address),
                    Matrix = matrix,
                    Truncated = truncated
                };
            }
            finally
            {
                Release((object)bottomRight);
                Release((object)topLeft);
                Release((object)sub);
                Release((object)value);
                Release((object)range);
                Release((object)sheet);
                Release((object)workbook);
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

        private static string GetString(Dictionary<string, object> args, string key, string defaultValue)
        {
            return Optional(args, key) ?? defaultValue;
        }

        private static double GetDouble(Dictionary<string, object> args, string key, double defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToDouble(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value, CultureInfo.InvariantCulture);
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(value);
            }
        }

        private sealed class ReadResult
        {
            public string Sheet { get; set; }
            public string Address { get; set; }
            public object[,] Matrix { get; set; }
            public bool Truncated { get; set; }
        }

        private sealed class WorksheetWriteResult
        {
            public string SheetName { get; set; }
            public string SheetRef { get; set; }
        }
    }
}