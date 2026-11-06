using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 第三期-导入外部文件与批量处理：预览/导入分隔文本文件、导入 Excel 工作表、合并多文件。
    /// 纯分隔文件操作不依赖 COM，写入统一走 writer 回调以捕获 undo 快照。
    /// </summary>
    public sealed class ExcelImportService
    {
        private const int MaxCellsPerImport = 200000;

        private readonly dynamic application;
        private readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer;
        private readonly OperationRecorder recorder;

        public ExcelImportService(
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer = null,
            OperationRecorder recorder = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.writer = writer;
            this.recorder = recorder;
        }

        public ToolResult PreviewExternalFile(Dictionary<string, object> args)
        {
            var filePath = Required(args, "file_path");
            if (!File.Exists(filePath))
            {
                return Invalid("FILE_NOT_FOUND", "文件不存在：" + filePath);
            }

            var sampleRows = Math.Max(1, Math.Min(GetInt(args, "sample_rows", 10), 100));
            var delimiter = GetString(args, "delimiter", "auto");
            var requestedEncoding = GetString(args, "encoding", "auto");
            var encoding = ResolveEncoding(requestedEncoding);
            var info = new FileInfo(filePath);
            string[] lines;
            try
            {
                // 多读 1 行用于判定“文件还有更多行未预览”，而非以文件字节大小近似。
                // 工具分类审查第九批：auto 编码走 TextEncodingDetector（GBK 兜底），预览不再乱码。
                if (string.Equals(requestedEncoding, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    byte[] bytes;
                    using (var fs = File.OpenRead(filePath))
                    using (var ms = new MemoryStream())
                    {
                        fs.CopyTo(ms);
                        bytes = ms.ToArray();
                    }

                    lines = Core.Tools.TextEncodingDetector.AutoDecodeLines(bytes, Math.Max(sampleRows + 1, 5)).Item1;
                }
                else
                {
                    lines = ReadLines(filePath, encoding, Math.Max(sampleRows + 1, 5));
                }
            }
            catch (Exception ex)
            {
                return Invalid("READ_FAILED", "读取文件失败：" + ex.Message);
            }

            if (lines.Length == 0)
            {
                return Invalid("EMPTY_FILE", "文件为空。");
            }

            if (string.Equals(delimiter, "auto", StringComparison.OrdinalIgnoreCase) || delimiter.Length != 1)
            {
                delimiter = DetectDelimiter(lines).ToString();
            }

            var truncated = lines.Length > sampleRows;
            var parsed = lines.Take(sampleRows).Select(line => SplitLine(line, delimiter[0])).ToArray();
            var columnCount = parsed.Length == 0 ? 0 : parsed.Max(row => row.Length);
            var matrix = new object[parsed.Length, columnCount];
            for (var r = 0; r < parsed.Length; r++)
            {
                for (var c = 0; c < columnCount; c++)
                {
                    matrix[r, c] = c < parsed[r].Length ? parsed[r][c] : string.Empty;
                }
            }

            return new ToolResult
            {
                Ok = true,
                Truncated = truncated,
                Data = new
                {
                    file_path = filePath,
                    file_size_bytes = info.Length,
                    detected_delimiter = delimiter,
                    detected_encoding = EncodingReport(requestedEncoding, encoding, filePath),
                    sample_rows = parsed.Length,
                    column_count = columnCount,
                    sample = matrix
                }
            };
        }

        public ToolResult ImportDelimitedFile(Dictionary<string, object> args)
        {
            var filePath = Required(args, "file_path");
            if (!File.Exists(filePath))
            {
                return Invalid("FILE_NOT_FOUND", "文件不存在：" + filePath);
            }

            var address = Required(args, "target_address");
            var skipEmpty = GetBool(args, "skip_empty_rows", true);
            var delimiter = GetString(args, "delimiter", "auto");
            var encoding = ResolveEncoding(GetString(args, "encoding", "auto"));
            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length > 50 * 1024 * 1024)
            {
                return Invalid("FILE_TOO_LARGE", "文件超过 50MB，请先拆分或预处理。");
            }

            // 工具分类审查第九批修复（编码乱码真缺陷）：encoding=auto 时 ResolveEncoding 硬返 UTF-8，
            // 无 BOM 的 GBK/ANSI 中文文件（中文记事本 ANSI 另存常见）导入全乱码——改走
            // TextEncodingDetector.AutoDecodeLines（BOM→UTF-8 严格→GBK 兜底），显式编码仍走原路径。
            string[] lines;
            string usedEncoding;
            try
            {
                byte[] bytes;
                using (var fs = File.OpenRead(filePath))
                using (var ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    bytes = ms.ToArray();
                }

                if (string.Equals(GetString(args, "encoding", "auto"), "auto", StringComparison.OrdinalIgnoreCase))
                {
                    var decoded = Core.Tools.TextEncodingDetector.AutoDecodeLines(bytes, -1);
                    lines = decoded.Item1;
                    usedEncoding = decoded.Item2;
                }
                else
                {
                    lines = ReadLines(filePath, encoding, -1);
                    usedEncoding = encoding.WebName;
                }
            }
            catch (Exception ex)
            {
                return Invalid("READ_FAILED", "读取文件失败：" + ex.Message);
            }

            if (lines.Length == 0)
            {
                return Invalid("EMPTY_FILE", "文件为空。");
            }

            if (string.Equals(delimiter, "auto", StringComparison.OrdinalIgnoreCase) || delimiter.Length != 1)
            {
                delimiter = DetectDelimiter(lines).ToString();
            }

            var parsed = lines.Select(line => SplitLine(line, delimiter[0])).ToList();
            if (skipEmpty)
            {
                parsed = parsed.Where(row => row.Any(cell => !string.IsNullOrWhiteSpace(cell))).ToList();
            }

            var columnCount = parsed.Count == 0 ? 0 : parsed.Max(row => row.Length);
            if (columnCount == 0)
            {
                return Invalid("EMPTY_FILE", "文件无有效数据。");
            }

            if (parsed.Count * columnCount > MaxCellsPerImport)
            {
                return Invalid("TOO_MANY_CELLS", $"数据量 {parsed.Count}x{columnCount} 超过单次导入上限（{MaxCellsPerImport} cells）。");
            }

            var matrix = new object[parsed.Count, columnCount];
            for (var r = 0; r < parsed.Count; r++)
            {
                for (var c = 0; c < columnCount; c++)
                {
                    matrix[r, c] = c < parsed[r].Length ? parsed[r][c] : string.Empty;
                }
            }

            if (writer == null)
            {
                return Invalid("NO_WRITER", "缺少写入回调。");
            }

            var writeArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = GetString(args, "target_sheet", null),
                ["address"] = address
            };
            return WriteViaCallback(writeArgs, matrix, "value", "import_delimited_file_to_sheet",
                new { file_path = filePath, imported_rows = parsed.Count, imported_columns = columnCount });
        }

        public ToolResult ImportExcelFileSheet(Dictionary<string, object> args)
        {
            var filePath = Required(args, "file_path");
            if (!File.Exists(filePath))
            {
                return Invalid("FILE_NOT_FOUND", "文件不存在：" + filePath);
            }

            var copyAsNewSheet = GetBool(args, "copy_as_new_sheet", true);
            // schema（ThirdPhaseSchemas）声明 sheets 为 {name} 对象项数组，必须按 name 键解析；
            // 此前硬编码 path 键导致按 schema 传参解析出类名字符串、0 表命中（M6 复审修复）。
            var sheets = ArgsStringArray(args, "sheets", "name");
            dynamic workbook = null;
            dynamic targetWorkbook = null;
            dynamic source = null;
            var mates = new List<dynamic>();
            try
            {
                targetWorkbook = application.ActiveWorkbook ??
                                 throw new InvalidOperationException("当前没有打开的工作簿。");
                if (copyAsNewSheet)
                {
                    source = OpenSourceWorkbook(filePath, mates);
                    var requested = new List<string>();
                    foreach (dynamic worksheet in source.Worksheets)
                    {
                        if (sheets.Length == 0 || SheetSelected(sheets, Convert.ToString(worksheet.Name)))
                        {
                            requested.Add(Convert.ToString(worksheet.Name));
                        }
                    }

                    var imported = new List<string>();
                    var cellBudget = 0;
                    if (sheets.Length > 0 && requested.Count == 0)
                    {
                        // 指定 sheets 但 0 命中不得静默成功（M6 复审修复，与 NO_VALID_SOURCE 同防线）。
                        return Invalid("NO_MATCHED_SHEET",
                            "指定的 sheets 在源文件中均未命中：" + string.Join("、", sheets));
                    }
                    foreach (var name in requested)
                    {
                        dynamic sourceSheet = source.Worksheets.Item[name];
                        var beforeCount = Convert.ToInt32(targetWorkbook.Worksheets.Count);
                        dynamic used = null;
                        try
                        {
                            used = sourceSheet.UsedRange;
                            var usedRows = Convert.ToInt32(used?.Rows.Count ?? 0);
                            var usedCols = Convert.ToInt32(used?.Columns.Count ?? 0);
                            cellBudget += checked(usedRows * usedCols);
                            if (cellBudget > MaxCellsPerImport)
                            {
                                return Invalid("TOO_MANY_CELLS",
                                    $"{name} 累计数据量超过单次导入上限（{MaxCellsPerImport} cells）。");
                            }

                            sourceSheet.Copy(After: targetWorkbook.Worksheets.Item[beforeCount]);
                        }
                        finally
                        {
                            Release((object)used);
                            Release((object)sourceSheet);
                        }

                        dynamic copied = null;
                        try
                        {
                            copied = targetWorkbook.Worksheets.Item[beforeCount + 1];
                            imported.Add(Convert.ToString(copied.Name));
                        }
                        finally
                        {
                            Release((object)copied);
                        }
                    }

                    RecordSheetInsert(args, imported.ToArray());
                    return new ToolResult
                    {
                        Ok = true,
                        AffectedRanges = imported.Select(s => "INSERT_SHEET:" + s).ToList(),
                        Data = new { file_path = filePath, imported_sheets = imported.ToArray(), mode = "new_sheet" }
                    };
                }

                // 取值复制模式：把源表 UsedRange 复制到指定区域，可撤销。
                source = OpenSourceWorkbook(filePath, mates);
                if (writer == null)
                {
                    return Invalid("NO_WRITER", "缺少写入回调。");
                }

                var targetSheetName = GetString(args, "target_sheet", Convert.ToString(targetWorkbook.ActiveSheet.Name));
                var targetAddress = GetString(args, "target_address", "A1");
                var importedSheets = new List<string>();
                object[,] merged = new object[0, 0];
                foreach (dynamic worksheet in source.Worksheets)
                {
                    if (sheets.Length > 0 && !SheetSelected(sheets, Convert.ToString(worksheet.Name)))
                    {
                        continue;
                    }

                    dynamic used = null;
                    try
                    {
                        used = worksheet.UsedRange;
                        var rows = Convert.ToInt32(used?.Rows.Count ?? 0);
                        var columns = Convert.ToInt32(used?.Columns.Count ?? 0);
                        if (rows == 0 || columns == 0)
                        {
                            continue;
                        }

                        var matrix = FlattenValues(used.Value2, rows, columns);
                        merged = AppendMatrix(merged, matrix);
                        importedSheets.Add(Convert.ToString(worksheet.Name));
                    }
                    finally
                    {
                        Release((object)used);
                    }
                }

                var mergedRows = merged.GetLength(0);
                var mergedCols = merged.GetLength(1);
                if (sheets.Length > 0 && importedSheets.Count == 0)
                {
                    // 指定 sheets 但 0 命中不得静默成功（M6 复审修复）。
                    return Invalid("NO_MATCHED_SHEET",
                        "指定的 sheets 在源文件中均未命中：" + string.Join("、", sheets));
                }

                if (mergedRows * mergedCols > MaxCellsPerImport)
                {
                    return Invalid("TOO_MANY_CELLS",
                        $"取值复制数据量 {mergedRows}x{mergedCols} 超过单次导入上限（{MaxCellsPerImport} cells）。");
                }

                var writeArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["sheet"] = targetSheetName,
                    ["address"] = targetAddress
                };
                return WriteViaCallback(writeArgs, merged, "value", "import_excel_file_sheet",
                    new { file_path = filePath, imported_rows = merged.GetLength(0), imported_columns = merged.GetLength(1), imported_sheets = importedSheets.ToArray(), mode = "copy_values" });
            }
            finally
            {
                foreach (var mate in mates)
                {
                    try
                    {
                        mate.Close(false);
                    }
                    catch (Exception)
                    {
                    }

                    Release((object)mate);
                }

                Release((object)source);
                Release((object)targetWorkbook);
                Release((object)workbook);
            }
        }

        public ToolResult MergeDelimitedFiles(Dictionary<string, object> args)
        {
            var filePaths = ArgsStringArray(args, "file_paths");
            if (filePaths.Length == 0)
            {
                return Invalid("MISSING_FILES", "缺少 file_paths。");
            }

            var outputAddress = Required(args, "output_address");
            var hasHeader = GetBool(args, "has_header", true);
            var delimiter = GetString(args, "delimiter", "auto");
            var encoding = ResolveEncoding(GetString(args, "encoding", "auto"));
            var header = new List<string>();
            var allRows = new List<object[]>();
            var mergedFileCount = 0;
            var failedFiles = new List<string>();
            string firstDelimiter = delimiter;

            var autoDecode = string.Equals(GetString(args, "encoding", "auto"), "auto", StringComparison.OrdinalIgnoreCase);
            foreach (var path in filePaths)
            {
                if (!File.Exists(path))
                {
                    // 工具分类审查第九批修复（静默跳过）：不存在的路径此前静默 continue——
                    // source_count 少 1 无任何解释；现计入 failed_files 回显。
                    failedFiles.Add(path + "（文件不存在）");
                    continue;
                }

                string[] lines;
                try
                {
                    if (autoDecode)
                    {
                        byte[] bytes;
                        using (var fs = File.OpenRead(path))
                        using (var ms = new MemoryStream())
                        {
                            fs.CopyTo(ms);
                            bytes = ms.ToArray();
                        }

                        lines = Core.Tools.TextEncodingDetector.AutoDecodeLines(bytes, -1).Item1;
                    }
                    else
                    {
                        lines = ReadLines(path, encoding, -1);
                    }
                }
                catch (Exception)
                {
                    failedFiles.Add(path);
                    continue;
                }

                if (lines.Length == 0)
                {
                    continue;
                }

                var delim = firstDelimiter;
                if (string.Equals(delim, "auto", StringComparison.OrdinalIgnoreCase) || delim.Length != 1)
                {
                    delim = DetectDelimiter(lines).ToString();
                    if (firstDelimiter.Length != 1 || string.Equals(firstDelimiter, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        firstDelimiter = delim;
                    }
                }

                var parsed = lines.Select(line => SplitLine(line, delim[0])).ToList();
                var start = 0;
                if (hasHeader && parsed.Count > 0)
                {
                    if (header.Count == 0)
                    {
                        header.AddRange(parsed[0]);
                    }

                    start = 1;
                }

                var addedThisFile = false;
                for (var r = start; r < parsed.Count; r++)
                {
                    if (parsed[r].All(cell => string.IsNullOrWhiteSpace(cell)))
                    {
                        continue;
                    }

                    allRows.Add(parsed[r]);
                    addedThisFile = true;
                }

                if (addedThisFile)
                {
                    mergedFileCount++;
                }
            }

            if (allRows.Count == 0)
            {
                return Invalid("EMPTY_RESULT", "没有合并到任何有效行。");
            }

            var totalColumns = Math.Max(header.Count, allRows.Max(row => row.Length));
            var outputRows = hasHeader && header.Count > 0 ? allRows.Count + 1 : allRows.Count;
            if (outputRows * totalColumns > MaxCellsPerImport)
            {
                return Invalid("TOO_MANY_CELLS", $"合并数据量 {outputRows}x{totalColumns} 超过单次上限（{MaxCellsPerImport} cells）。");
            }

            var matrix = new object[outputRows, totalColumns];
            var outRow = 0;
            if (hasHeader && header.Count > 0)
            {
                for (var c = 0; c < totalColumns; c++)
                {
                    matrix[0, c] = c < header.Count ? header[c] : string.Empty;
                }

                outRow = 1;
            }

            foreach (var row in allRows)
            {
                for (var c = 0; c < totalColumns; c++)
                {
                    matrix[outRow, c] = c < row.Length ? row[c] : string.Empty;
                }

                outRow++;
            }

            var writeArgs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["sheet"] = GetString(args, "output_sheet", null),
                ["address"] = outputAddress
            };
            return WriteViaCallback(writeArgs, matrix, "value", "merge_delimited_files_to_sheet",
                new { combined_rows = outRow, combined_columns = totalColumns, source_count = mergedFileCount, failed_files = failedFiles.ToArray() });
        }

        private void RecordSheetInsert(Dictionary<string, object> args, string[] importedSheets)
        {
            if (recorder == null || importedSheets.Length == 0)
            {
                return;
            }

            recorder(
                "import_excel_file_sheet",
                args,
                "sheet-insert",
                new { sheets = importedSheets },
                new SheetInsertUndoPayload { Sheets = importedSheets });
        }

        private ToolResult WriteViaCallback(Dictionary<string, object> writeArgs, object[,] matrix, string valueMode, string tool, object extra)
        {
            var result = writer(writeArgs, matrix, valueMode, tool);
            if (!result.Ok)
            {
                return result;
            }

            var opId = TryGetOperationId(result);
            var data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (opId != null)
            {
                data["operation_id"] = opId;
            }

            data["imported_rows"] = matrix.GetLength(0);
            data["imported_columns"] = matrix.GetLength(1);
            // 把调用方传入的 extra(如 imported_sheets/source_count/failed_files) 并入输出，
            // 之前该参数被整体丢弃导致 source_count 等真实信息约不到 AI。
            if (extra != null)
            {
                foreach (var property in extra.GetType().GetProperties())
                {
                    data[property.Name] = property.GetValue(extra, null);
                }
            }

            return new ToolResult
            {
                Ok = true,
                AffectedRanges = result.AffectedRanges,
                Data = data
            };
        }

        private static string TryGetOperationId(ToolResult result)
        {
            if (result.Data == null || result.Data is string)
            {
                return null;
            }

            var property = result.Data.GetType().GetProperty("operation_id");
            if (property == null)
            {
                return null;
            }

            var value = property.GetValue(result.Data, null);
            return value == null ? null : Convert.ToString(value);
        }

        internal static string DetectDelimiter(string[] lines)
        {
            var candidates = new[] { '\t', ',', ';', '|', ' ' };
            var best = '\0';
            var bestScore = -1.0;
            foreach (var candidate in candidates)
            {
                var counts = new List<int>();
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    counts.Add(SplitLine(line, candidate).Length);
                }

                if (counts.Count < 2)
                {
                    continue;
                }

                var mean = counts.Average();
                var variance = counts.Select(c => (c - mean) * (c - mean)).Average();
                var averageCount = counts.Average();
                var score = variance == 0 ? averageCount : averageCount - variance;
                if (score > bestScore && counts.Max() > 1)
                {
                    bestScore = score;
                    best = candidate;
                }
            }

            return best == '\0' ? "," : best.ToString();
        }

        private static bool SheetSelected(string[] sheets, string name)
        {
            if (sheets.Length == 0)
            {
                return true;
            }

            foreach (var sheet in sheets)
            {
                if (string.Equals(sheet, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 打开源工作簿（M16 全量审查 P0 修复）：若该文件已被用户在宿主中打开，COM 的
        /// Workbooks.Open 会直接返回用户已打开的实例——此时绝不能在 finally 里 Close(false)
        /// （会把用户未保存的工作簿无提示关闭）。因此先按全路径比对 Workbooks 集合：命中
        /// 则复用既有实例且不登记 ownedMates（由调用方 finally 只关闭自己打开的簿）。
        /// </summary>
        private dynamic OpenSourceWorkbook(string filePath, List<dynamic> ownedMates)
        {
            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(filePath);
            }
            catch (Exception)
            {
                fullPath = filePath;
            }

            foreach (dynamic wb in application.Workbooks)
            {
                string openPath = null;
                try
                {
                    openPath = Convert.ToString(wb.FullName);
                }
                catch (Exception)
                {
                    // 未保存的新建簿 FullName 可能为空或抛错——跳过
                }

                if (!string.IsNullOrEmpty(openPath) &&
                    string.Equals(openPath, fullPath, StringComparison.OrdinalIgnoreCase))
                {
                    return wb; // 用户已打开：复用，不加入 ownedMates，调用方不负责关闭
                }

                // 工作簿级共享 RCW 不做强制释放（防"RCW 分离"——M16 真机 WPS 实证：
                // 对枚举到的 Workbook FinalRelease 会让宿主缓存 RCW 断开，后续访问报错）。
            }

            var opened = application.Workbooks.Open(filePath, ReadOnly: true);
            ownedMates.Add(opened);
            return opened;
        }

        private static string[] ReadLines(string filePath, Encoding encoding, int maxLines)
        {
            var lines = new List<string>();
            using (var stream = new StreamReader(filePath, encoding, detectEncodingFromByteOrderMarks: true))
            {
                string line;
                var count = 0;
                while ((line = stream.ReadLine()) != null)
                {
                    lines.Add(line);
                    count++;
                    if (maxLines > 0 && count >= maxLines)
                    {
                        break;
                    }
                }
            }

            return lines.ToArray();
        }

        private static string EncodingReport(string requested, Encoding fallback, string filePath)
        {
            if (!string.Equals(requested, "auto", StringComparison.OrdinalIgnoreCase))
            {
                return fallback.WebName;
            }

            return SniffBomEncoding(filePath);
        }

        internal static string SniffBomEncoding(string filePath)
        {
            try
            {
                using (var fs = File.OpenRead(filePath))
                {
                    var head = new byte[4];
                    var n = fs.Read(head, 0, 4);
                    if (n >= 4 && head[0] == 0xFF && head[1] == 0xFE && head[2] == 0x00 && head[3] == 0x00) return "utf-32-le";
                    if (n >= 4 && head[0] == 0x00 && head[1] == 0x00 && head[2] == 0xFE && head[3] == 0xFF) return "utf-32-be";
                    if (n >= 2 && head[0] == 0xFF && head[1] == 0xFE) return "utf-16-le";
                    if (n >= 2 && head[0] == 0xFE && head[1] == 0xFF) return "utf-16-be";
                    if (n >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF) return "utf-8-bom";
                }
            }
            catch (Exception)
            {
            }

            return "utf-8";
        }

        private static Encoding ResolveEncoding(string encoding)
        {
            if (string.Equals(encoding, "gbk", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(encoding, "gb2312", StringComparison.OrdinalIgnoreCase) ||
                encoding == "936")
            {
                try
                {
                    return Encoding.GetEncoding(936);
                }
                catch (Exception)
                {
                    return Encoding.UTF8;
                }
            }

            if (string.Equals(encoding, "ascii", StringComparison.OrdinalIgnoreCase))
            {
                return Encoding.ASCII;
            }

            return Encoding.UTF8;
        }

        private static string[] SplitLine(string line, char delimiter)
        {
            var result = new List<string>();
            if (line == null)
            {
                return result.ToArray();
            }

            var field = new StringBuilder();
            var inQuotes = false;
            for (var i = 0; i < line.Length; i++)
            {
                var ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            field.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        field.Append(ch);
                    }
                }
                else if (ch == '"' && field.Length == 0)
                {
                    inQuotes = true;
                }
                else if (ch == delimiter)
                {
                    result.Add(field.ToString());
                    field.Clear();
                }
                else
                {
                    field.Append(ch);
                }
            }

            result.Add(field.ToString());
            return result.ToArray();
        }

        private static object[,] AppendMatrix(object[,] current, object[,] next)
        {
            var width = Math.Max(current.GetLength(1), next.GetLength(1));
            var total = current.GetLength(0) + next.GetLength(0);
            var result = new object[total, width];
            for (var r = 0; r < current.GetLength(0); r++)
            {
                for (var c = 0; c < width; c++)
                {
                    result[r, c] = c < current.GetLength(1) ? current[r, c] : string.Empty;
                }
            }

            var offset = current.GetLength(0);
            for (var r = 0; r < next.GetLength(0); r++)
            {
                for (var c = 0; c < width; c++)
                {
                    result[offset + r, c] = c < next.GetLength(1) ? next[r, c] : string.Empty;
                }
            }

            return result;
        }

        private static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source == null || rows == 0 || columns == 0)
            {
                return result;
            }

            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }
            }
            else if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }

        private static string[] ArgsStringArray(Dictionary<string, object> args, string key, string itemKey = "path")
        {
            if (!args.TryGetValue(key, out var raw) || !(raw is System.Collections.IEnumerable enumerable) || raw is string)
            {
                return new string[0];
            }

            var result = new List<string>();
            foreach (var item in enumerable)
            {
                // 兼容 schema BoxArray 声明的对象项：此前 Convert.ToString(对象)
                // 会得到类名字符串，File.Exists 永远 false → merge_delimited_files 返回
                // EMPTY_RESULT（M6 WPS 全量审查 F9 实测发现）。
                // itemKey 由调用方按 schema 声明传入（如 sheets→name、file_paths→path）。
                string value;
                if (item is System.Collections.IDictionary map)
                {
                    if (!(map.Contains(itemKey) && map[itemKey] != null))
                    {
                        // 对象项缺声明键：跳过，禁止类名字符串混入结果（M6 复审修复）。
                        continue;
                    }

                    var mv = map[itemKey];
                    value = Convert.ToString(mv is Newtonsoft.Json.Linq.JValue jv ? jv.Value : mv);
                }
                else if (item is Newtonsoft.Json.Linq.JObject jo)
                {
                    if (!(jo.TryGetValue(itemKey, out var tv) && tv != null))
                    {
                        continue;
                    }

                    value = Convert.ToString(tv is Newtonsoft.Json.Linq.JValue jv2 ? jv2.Value : (object)tv);
                }
                else
                {
                    value = Convert.ToString(item is Newtonsoft.Json.Linq.JValue j3 ? j3.Value : item);
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    result.Add(value);
                }
            }

            return result.ToArray();
        }

        private static string Required(Dictionary<string, object> args, string key)
        {
            var value = GetString(args, key, null);
            return string.IsNullOrWhiteSpace(value)
                ? throw new ArgumentException($"缺少必填参数：{key}")
                : value;
        }

        private static string GetString(Dictionary<string, object> args, string key, string defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToString(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static int GetInt(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
        }

        private static bool GetBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToBoolean(value is Newtonsoft.Json.Linq.JValue j ? j.Value : value);
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
}