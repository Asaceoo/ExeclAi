using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 工作簿级工具：页面布局、导出 PDF、合并/拆分工作簿、外部链接检查与替换。
    /// </summary>
    public sealed class ExcelWorkbookService
    {
        private readonly dynamic application;

        public ExcelWorkbookService(dynamic application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public ToolResult ApplyPageSetup(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = ResolveSheet(workbook, Optional(args, "sheet"));
                var page = sheet.PageSetup;

                var orientation = Optional(args, "orientation");
                if (!string.IsNullOrWhiteSpace(orientation) && orientation != "landscape" && orientation != "portrait")
                {
                    return Invalid("INVALID_ORIENTATION", $"orientation 仅支持 portrait/landscape，当前为 {orientation}。");
                }

                if (!string.IsNullOrWhiteSpace(orientation))
                {
                    page.Orientation = orientation == "landscape" ? 2 : 1; // xlLandscape / xlPortrait
                }

                if (args.TryGetValue("paper_size", out var paperRaw) && paperRaw != null)
                {
                    var paperName = Convert.ToString(paperRaw);
                    if (!IsKnownPaperSize(paperName))
                    {
                        return Invalid("INVALID_PAPER_SIZE", $"paper_size 仅支持 a4/a3/a5/letter/legal，当前为 {paperName}。");
                    }

                    page.PaperSize = PaperSize(paperName);
                }

                if (GetInt(args, "fit_to_pages_wide", 0) > 0)
                {
                    page.Zoom = false;
                    page.FitToPagesWide = GetInt(args, "fit_to_pages_wide", 1);
                }

                if (GetInt(args, "fit_to_pages_tall", 0) > 0)
                {
                    page.Zoom = false;
                    page.FitToPagesTall = GetInt(args, "fit_to_pages_tall", 1);
                }

                var scale = GetInt(args, "scale", 0);
                if (scale >= 10 && scale <= 400)
                {
                    page.Zoom = true;
                    page.Zoom = scale;
                }

                // schema 把 margins 声明为数组（[{top,bottom,left,right}]），但历史调用方可能传单对象，
                // 两种形态都归一为单字典后应用；数值兼容 int/long/float/数值串，避免整数边距被静默丢弃。
                var margins = ResolveMargins(args);
                if (margins != null)
                {
                    if (TryMarginsValue(margins, "top", out var top)) page.TopMargin = PointsFromCm(top);
                    if (TryMarginsValue(margins, "bottom", out var bottom)) page.BottomMargin = PointsFromCm(bottom);
                    if (TryMarginsValue(margins, "left", out var left)) page.LeftMargin = PointsFromCm(left);
                    if (TryMarginsValue(margins, "right", out var right)) page.RightMargin = PointsFromCm(right);
                }

                var printArea = Optional(args, "print_area");
                if (printArea != null)
                {
                    page.PrintArea = printArea;
                }

                if (GetBool(args, "center_horizontally", false)) page.CenterHorizontally = true;
                if (GetBool(args, "center_vertically", false)) page.CenterVertically = true;

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        sheet = Convert.ToString(sheet.Name),
                        orientation = orientation,
                        print_area = printArea,
                        applied = true
                    }
                };
            }
            finally
            {
                Release((object)sheet);
                Release((object)workbook);
            }
        }

        /// <summary>把 schema 的 margins（单个对象或数组）归一为页边距字典；无法识别返回 null。
        /// 容忍 JSON 反序列化出的 Dictionary / JObject / JArray / IList / JSON 字符串等形态。
        /// 注意 Newtonsoft 的 JObject 会匹配通用 IList 分支（被当元素列表迭代），故须把 JObject/JArray 特判提前。</summary>
        private static Dictionary<string, object> ResolveMargins(Dictionary<string, object> args)
        {
            if (!args.TryGetValue("margins", out var raw) || raw == null)
            {
                return null;
            }

            var merged = new Dictionary<string, object>();
            var found = false;
            switch (raw)
            {
                case Newtonsoft.Json.Linq.JObject jo:
                    found = MergeMarginsItem(merged, ToDict(jo)) || found;
                    break;
                case Newtonsoft.Json.Linq.JArray ja:
                    foreach (var item in ja)
                    {
                        found = MergeMarginsItem(merged, AsMarginsDict(item)) || found;
                    }

                    break;
                case IDictionary<string, object> dict:
                    foreach (var kv in dict)
                    {
                        merged[kv.Key] = kv.Value;
                        found = true;
                    }

                    break;
                case IDictionary<string, Newtonsoft.Json.Linq.JToken> jdict:
                    foreach (var kv in jdict)
                    {
                        merged[kv.Key] = kv.Value.Type == Newtonsoft.Json.Linq.JTokenType.Null ? null : kv.Value;
                        found = true;
                    }

                    break;
                case System.Collections.IList list:
                    foreach (var item in list)
                    {
                        found = MergeMarginsItem(merged, AsMarginsDict(item)) || found;
                    }

                    break;
                case string s when !string.IsNullOrWhiteSpace(s):
                    try
                    {
                        found = MergeMarginsItem(merged, ToDict(Newtonsoft.Json.Linq.JObject.Parse(s))) || found;
                    }
                    catch
                    {
                    }
                    break;
            }

            return found ? merged : null;
        }

        private static Dictionary<string, object> AsMarginsDict(object value)
        {
            switch (value)
            {
                case IDictionary<string, object> d:
                    return new Dictionary<string, object>(d);
                case Newtonsoft.Json.Linq.JObject jo:
                    return ToDict(jo);
                default:
                    return null;
            }
        }

        private static Dictionary<string, object> ToDict(Newtonsoft.Json.Linq.JObject jo)
        {
            try
            {
                return jo.ToObject<Dictionary<string, object>>();
            }
            catch
            {
                return null;
            }
        }

        private static bool MergeMarginsItem(Dictionary<string, object> merged, Dictionary<string, object> item)
        {
            if (item == null)
            {
                return false;
            }

            foreach (var kv in item)
            {
                merged[kv.Key] = kv.Value;
            }

            return true;
        }

        /// <summary>兼容 int/long/float/decimal/数值串的页边距读取；缺失或不可解析返回 false（不做任何改动）。</summary>
        private static bool TryMarginsValue(Dictionary<string, object> margins, string key, out double value)
        {
            value = 0;
            if (!margins.TryGetValue(key, out var raw) || raw == null)
            {
                return false;
            }

            var d = ToCmNumber(raw);
            if (d == null)
            {
                return false;
            }

            value = d.Value;
            return true;
        }

        private static double? ToCmNumber(object raw)
        {
            try
            {
                switch (raw)
                {
                    case double dd:
                        return dd;
                    case int ni:
                        return ni;
                    case long nl:
                        return nl;
                    case float nf:
                        return nf;
                    case decimal md:
                        return (double)md;
                    case Newtonsoft.Json.Linq.JValue jv:
                        var num = jv.Type is Newtonsoft.Json.Linq.JTokenType.Integer
                                       or Newtonsoft.Json.Linq.JTokenType.Float
                            ? Convert.ToDouble(jv.Value, CultureInfo.InvariantCulture)
                            : (double?)null;
                        return num ?? (double.TryParse(Convert.ToString(jv.Value), NumberStyles.Float, CultureInfo.InvariantCulture, out var jp) ? jp : (double?)null);
                    default:
                        return double.TryParse(Convert.ToString(raw), NumberStyles.Float, CultureInfo.InvariantCulture, out var r) ? r : (double?)null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static int PaperSize(string paper)
        {
            switch (paper.ToLowerInvariant())
            {
                case "letter": return 1;
                case "a3": return 8;
                case "a4": return 9;
                case "a5": return 11;
                case "legal": return 5;
                default: return 9;
            }
        }

        private static bool IsKnownPaperSize(string paper)
        {
            switch (paper?.Trim().ToLowerInvariant())
            {
                case "letter":
                case "legal":
                case "a3":
                case "a4":
                case "a5":
                    return true;
                default:
                    return false;
            }
        }

        private static double PointsFromCm(double cm)
        {
            return cm * 28.3464567;
        }

        public ToolResult ExportSheetsToPdf(Dictionary<string, object> args)
        {
            var filePath = ExcelAiAssistant.Core.Security.PathGuard.GuardAndComplete(
                Required(args, "file_path"), ".pdf");
            // M28 审查：导出不得静默覆盖既有文件（NormalWrite 的 overwrite 仅作门槛，仍需显式）
            if (!TryReadBoolArg(args, "overwrite", out var overwriteExport, out var overwriteError))
            {
                return Invalid("INVALID_ARGS", overwriteError);
            }

            if (System.IO.File.Exists(filePath) && !overwriteExport)
            {
                return Invalid("OVERWRITE_DENIED",
                    "目标文件已存在：" + filePath + "。为防止覆盖既有文件，请显式传 overwrite=true（或换一个 file_path）。");
            }

            // M28（第三轮续）：覆盖既有 PDF 前先备份（可恢复），备份失败则中止
            string backupPath = null;
            if (System.IO.File.Exists(filePath))
            {
                try { backupPath = ExcelAiAssistant.Core.Security.PathGuard.CreateBackup(filePath); }
                catch (Exception backupException)
                {
                    return Invalid("BACKUP_FAILED", "覆盖前备份既有 PDF 失败，已中止导出：" + backupException.Message);
                }
            }

            var sheets = args.TryGetValue("sheets", out var rawSheets) && rawSheets is IEnumerable<object> list
                ? list.Where(item => item != null).Select(item => Convert.ToString(item)).ToArray()
                : new string[0];
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                // 仅导出部分工作表时，新建临时工作簿承载选定副本；否则整簿导出。
                dynamic exportWb = workbook;
                List<string> missingSheets = null;
                if (sheets.Length > 0)
                {
                    List<string> missing = null;
                    exportWb = BuildTemporaryWorkbook(workbook, sheets, out missing);
                    missingSheets = missing;
                    // 全部表都不存在：不要导出「空白占位簿」误导用户，关掉临时簿并明确报错。
                    if (missing.Count == sheets.Length)
                    {
                        try
                        {
                            if (!ReferenceEquals(exportWb, workbook))
                            {
                                exportWb.Close(false);
                            }
                        }
                        catch
                        {
                        }

                        Release((object)exportWb);
                        return Invalid("SHEET_NOT_FOUND",
                            "请求导出的工作表均不存在：" + string.Join(", ", missing));
                    }
                }

                try
                {
                    // xlTypePDF = 0。ExportAsFixedFormat 在脚本环境可能因打印驱动受限，捕获并提示。
                    exportWb.ExportAsFixedFormat(0, filePath);
                }
                catch (Exception ex)
                {
                    return Invalid("PDF_EXPORT_FAILED", "导出 PDF 失败：" + ex.Message + "（请确认系统已安装可用打印机/PDF 驱动）。");
                }
                finally
                {
                    if (!ReferenceEquals(exportWb, workbook))
                    {
                        try
                        {
                            exportWb.Close(false);
                        }
                        catch
                        {
                        }

                        Release((object)exportWb);
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        file_path = filePath,
                        sheets = sheets.Length == 0 ? new[] { "ALL" } : sheets,
                        exported = true,
                        backup_path = backupPath,
                        skipped = missingSheets != null && missingSheets.Count > 0 ? missingSheets.ToArray() : null
                    }
                };
            }
            finally
            {
                Release((object)workbook);
            }
        }

        private dynamic BuildTemporaryWorkbook(dynamic sourceWorkbook, string[] sheetNames, out List<string> missing)
        {
            missing = new List<string>();
            dynamic temp = null;
            try
            {
                temp = application.Workbooks.Add();
                dynamic firstSheet = temp.Worksheets.Item[1];
                var inserted = false;
                foreach (var name in sheetNames)
                {
                    dynamic sourceSheet = null;
                    try
                    {
                        sourceSheet = sourceWorkbook.Worksheets.Item[name];
                        dynamic afterAnchor = temp.Worksheets.Item[temp.Worksheets.Count];
                        try
                        {
                            sourceSheet.Copy(After: afterAnchor);
                            inserted = true;
                        }
                        finally
                        {
                            Release((object)afterAnchor);
                        }
                    }
                    catch
                    {
                        missing.Add(name);
                    }
                    finally
                    {
                        Release((object)sourceSheet);
                    }
                }

                if (inserted)
                {
                    firstSheet.Delete();
                }

                // M16 全量审查：firstSheet RCW 释放（原实现从不释放，每次导出 PDF 泄漏一个）。
                Release((object)firstSheet);
                return temp;
            }
            catch
            {
                Release((object)temp);
                throw;
            }
        }

        public ToolResult MergeWorkbooks(Dictionary<string, object> args)
        {
            var sources = ArgsNameArray(args, "source_workbooks", "path");
            if (sources.Count == 0)
            {
                return Invalid("MISSING_SOURCES", "缺少必填参数：source_workbooks。");
            }

            dynamic workbook = null;
            var mates = new List<dynamic>();
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var addedSheets = new List<string>();
                var mergedCount = 0;
                foreach (var path in sources)
                {
                    // M16 全量审查：源路径与 export_sheets_to_pdf/split_workbook 同口径过 PathGuard
                    //（防 '..' 穿越）；随后保留存在性检查（跳过不存在的源，全部无效时有显式报错）。
                    var guarded = ExcelAiAssistant.Core.Security.PathGuard.GuardOutputPath(path);
                    if (!System.IO.File.Exists(guarded))
                    {
                        continue;
                    }

                    dynamic mate = null;
                    try
                    {
                        mate = application.Workbooks.Open(guarded, ReadOnly: true);
                        mates.Add(mate);
                        foreach (dynamic worksheet in mate.Worksheets)
                        {
                            var copied = worksheet.Copy(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
                            addedSheets.Add(Convert.ToString(worksheet.Name));
                        }

                        mergedCount++;
                    }
                    catch (Exception ex)
                    {
                        return Invalid("MERGE_FAILED", "合并失败：" + ex.Message);
                    }
                }

                // 0 个源文件成功合并不得静默返回 OK（与 M6 缺陷 #1 同类的静默成功残余，
                // M6 复审修复：路径全部无效时显式报错并列出提供值）。
                if (mergedCount == 0)
                {
                    return Invalid("NO_VALID_SOURCE",
                        "提供的源文件均未成功合并（路径不存在或无法打开）：" + string.Join("、", sources));
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { source_count = mergedCount, added_sheets = addedSheets }
                };
            }
            finally
            {
                foreach (var mate in mates)
                {
                    try
                    {
                        mate.Close(false);
                    }
                    catch
                    {
                    }

                    Release((object)mate);
                }

                Release((object)workbook);
            }
        }

        public ToolResult SplitWorkbook(Dictionary<string, object> args)
        {
            // M16 全量审查：output_folder 与 export_sheets_to_pdf 同口径过 PathGuard，防 '..' 穿越。
            var outputFolder = ExcelAiAssistant.Core.Security.PathGuard.GuardOutputPath(Required(args, "output_folder"));
            if (!System.IO.Directory.Exists(outputFolder))
            {
                System.IO.Directory.CreateDirectory(outputFolder);
            }

            var sheets = ArgsNameArray(args, "sheets", "name");
            if (!TryReadBoolArg(args, "overwrite", out var overwrite, out var overwriteError))
            {
                return Invalid("INVALID_ARGS", overwriteError);
            }

            dynamic workbook = null;
            bool? prevAlerts = null; // M28 审查：读取前异常时不得把宿主 DisplayAlerts 误写回 true
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var written = new List<string>();
                var failed = new List<string>();
                var skipped = new List<string>();
                var backups = new List<string>(); // M28（第三轮续）：覆盖前备份的原文件
                var failedDetails = new List<object>();   // M28 审查（AI 使用方）：结构化失败明细
                var skippedDetails = new List<object>();
                var targetNames = sheets.Count == 0
                    ? AllSheetNames(workbook)
                    : sheets;
                // M28/R3：SaveAs 覆盖既有文件、删除临时簿默认表都会弹确认框——DisplayAlerts=true 的
                // 宿主 Excel 上 out-of-proc 调用会模态卡死，全循环显式压制、结束恢复。
                prevAlerts = Convert.ToBoolean(application.DisplayAlerts);
                application.DisplayAlerts = false;
                foreach (var name in targetNames)
                {
                    dynamic sourceSheet = null;
                    dynamic temp = null;
                    dynamic defaultSheet = null;
                    try
                    {
                        sourceSheet = workbook.Worksheets.Item[name];
                        temp = application.Workbooks.Add();
                        // M28/R3：Copy 前捕获临时簿自带的空白默认表，Copy 后删除——
                        // 原实现把它留在拆出文件里（每个拆出文件都多一张空白 Sheet1）。
                        defaultSheet = temp.Worksheets.Item[1];
                        // M28 审查（AI 使用方）：先给默认表改名——源表与默认表同名（Sheet1）时
                        // Copy 会让副本被 Excel 加 " (2)" 后缀，拆出文件里的表名与文件名不一致。
                        try { defaultSheet.Name = "__jyyj_tmp_" + Guid.NewGuid().ToString("N").Substring(0, 8); } catch { }
                        sourceSheet.Copy(After: defaultSheet);
                        defaultSheet.Delete();
                        var filePath = System.IO.Path.Combine(outputFolder, SanitizeFileName(Convert.ToString(sourceSheet.Name)) + ".xlsx");
                        if (System.IO.File.Exists(filePath))
                        {
                            if (!overwrite)
                            {
                                // M28 审查：拆出文件覆盖既有文件属破坏性写入，默认跳过并如实报告；
                                // 跳过≠失败（用户已有文件是正常状态），单独进 skipped_sheets，不污染 ok 判定。
                                skipped.Add(name + "：目标已存在（overwrite=true 可覆盖）：" + filePath);
                                skippedDetails.Add(new { sheet = name, code = "TARGET_EXISTS", path = filePath, message = "目标已存在（overwrite=true 可覆盖）" });
                                continue;
                            }

                            // M28（第三轮续）：覆盖前备份，备份失败则该表按失败处理（不静默覆盖）
                            try
                            {
                                var backup = ExcelAiAssistant.Core.Security.PathGuard.CreateBackup(filePath);
                                if (backup != null) { backups.Add(backup); }
                            }
                            catch (Exception backupException)
                            {
                                failed.Add(name + "：覆盖前备份失败：" + backupException.Message);
                                failedDetails.Add(new { sheet = name, code = "BACKUP_FAILED", path = filePath, message = backupException.Message });
                                continue;
                            }
                        }

                        temp.SaveAs(filePath, 51); // xlOpenXMLWorkbook
                        temp.Close(false);
                        written.Add(filePath);
                    }
                    catch (Exception ex)
                    {
                        // M28/R3（假失败修复）：原实现任何一张表失败都整体 return ok:false——
                        // 即使前面的表已经 SaveAs 落盘（「文件已生成却报拆分失败」）。
                        // 改为按表收集：部分成功返回 ok:true + failed_sheets 明细，全部失败才报 SPLIT_FAILED。
                        failed.Add(name + "：" + ex.Message);
                        failedDetails.Add(new { sheet = name, code = "SPLIT_FAILED", message = ex.Message });
                    }
                    finally
                    {
                        // M16 全量审查：失败路径上 temp 未 SaveAs/Close，会残留空白工作簿——
                        // finally 中兜底关闭（written 路径已 Close，重复 Close 幂等吞异常）。
                        // FIX（M28/R3 假失败根因）：temp 为 dynamic，成功路径 temp.Close 后 RCW 已断连，
                        // 动态绑定器对断连 RCW 做 != null 会抛 COMException 0x80010114（对象不存在）——
                        // 这正是"文件已生成却报拆分失败"的触发点。静态转 object 判空不经过 DLR 绑定。
                        if ((object)temp != null)
                        {
                            try { temp.Close(false); } catch { }
                        }

                        Release((object)defaultSheet);
                        Release((object)temp);
                        Release((object)sourceSheet);
                    }
                }

                if (written.Count == 0 && failed.Count == 0 && skipped.Count > 0)
                {
                    // 全部因目标已存在而跳过：这是正常结果（幂等重跑），不是失败
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new
                        {
                            output_folder = outputFolder,
                            written_files = written,
                            skipped_sheets = skipped,
                            skipped_count = skipped.Count,
                            skipped_details = skippedDetails,
                            backups = backups,
                            failed_sheets = failed,
                            failed_details = failedDetails,
                            failed_count = 0,
                            note = "目标文件均已存在，本次未覆盖（overwrite=true 可覆盖）。"
                        }
                    };
                }

                if (written.Count == 0)
                {
                    // M28 审查（第三轮）：全失败也保留结构化明细（不再只回拼接文本）
                    return new ToolResult
                    {
                        Ok = false,
                        ErrorCode = "SPLIT_FAILED",
                        Data = new
                        {
                            message = "拆分失败（0 个文件生成）：" + string.Join("；", failed.Concat(skipped)),
                            output_folder = outputFolder,
                            failed_sheets = failed,
                            failed_details = failedDetails,
                            skipped_sheets = skipped,
                            skipped_details = skippedDetails,
                            backups = backups,
                            failed_count = failed.Count,
                            skipped_count = skipped.Count
                        }
                    };
                }

                return new ToolResult
                {
                    Ok = true,
                    Warning = failed.Count > 0
                        ? ("部分拆分失败：" + failed.Count + " 张表未拆分，详见 failed_sheets/failed_details。")
                        : null,
                    Data = new
                    {
                        output_folder = outputFolder,
                        written_files = written,
                        skipped_sheets = skipped,
                        skipped_count = skipped.Count,
                        skipped_details = skippedDetails,
                        backups = backups,
                        failed_sheets = failed,
                        failed_details = failedDetails,
                        failed_count = failed.Count
                    }
                };
            }
            finally
            {
                if (prevAlerts.HasValue)
                {
                    try { application.DisplayAlerts = prevAlerts.Value; } catch { }
                }

                Release((object)workbook);
            }
        }

        /// <summary>
        /// 保存当前工作簿（M28/R3 审查 B：MCP/CLI 等外部入口写完必须能落盘——此前 server 进程退出即全丢，改动无存）。
        /// 无 file_path：保存回当前文件（工作簿从未存盘时显式报错，防无效调用）；
        /// 有 file_path：另存为该路径（PathGuard 防穿越且返回绝对路径；扩展名限 .xlsx/.xlsm/.xls，FileFormat 按扩展名推导）。
        /// </summary>
        public ToolResult SaveWorkbook(Dictionary<string, object> args)
        {
            var filePath = Optional(args, "file_path");
            dynamic workbook = null;
            bool? prevAlerts = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    // 未存盘的工作簿没有磁盘路径，Save() 会弹「另存为」对话框（DisplayAlerts 关不掉）——
                    // out-of-proc 场景直接卡死，必须显式要求 file_path。
                    var currentPath = Convert.ToString(workbook.Path);
                    if (string.IsNullOrEmpty(currentPath))
                    {
                        return Invalid("NEED_FILE_PATH",
                            "当前工作簿从未保存过，没有可保存的磁盘路径。请提供 file_path 指定另存位置。");
                    }

                    workbook.Save();
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { saved = true, file_path = System.IO.Path.Combine(currentPath, Convert.ToString(workbook.Name)) }
                    };
                }

                var guarded = ExcelAiAssistant.Core.Security.PathGuard.GuardOutputPath(filePath);
                var ext = System.IO.Path.GetExtension(guarded);
                if (string.IsNullOrWhiteSpace(ext))
                {
                    return Invalid("INVALID_EXTENSION", "file_path 缺少扩展名（支持 .xlsx/.xlsm/.xls）。");
                }

                // SaveAs FileFormat 按扩展名推导（与 split_workbook 的 51 同源口径）
                // M28 审查：不支持 .csv——SaveAs(6) 只保存活动表，多表工作簿会被静默丢表（数据丢失级），
                // 需要 CSV 请显式用导出类工具并自行确认单表语义。
                var format = string.Equals(ext, ".xlsm", StringComparison.OrdinalIgnoreCase) ? 52
                    : string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase) ? 56
                    : string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) ? 51
                    : 0;
                if (format == 0)
                {
                    return Invalid("INVALID_EXTENSION",
                        "file_path 扩展名仅支持 .xlsx/.xlsm/.xls，当前：" + ext);
                }

                // M28 审查：SaveAs 到已存在的**其他文件**是破坏性覆盖，必须显式 overwrite=true；
                // 覆盖当前工作簿自身文件（保存语义）不受限。
                if (!TryReadBoolArg(args, "overwrite", out var overwrite, out var overwriteError))
                {
                    return Invalid("INVALID_ARGS", overwriteError);
                }

                var currentFullName = string.Empty;
                try { currentFullName = Convert.ToString(workbook.FullName) ?? string.Empty; } catch { }
                if (System.IO.File.Exists(guarded) &&
                    !string.Equals(guarded, currentFullName, StringComparison.OrdinalIgnoreCase) &&
                    !overwrite)
                {
                    return Invalid("OVERWRITE_DENIED",
                        "目标文件已存在：" + guarded + "。为防止覆盖其他文件，请显式传 overwrite=true 后重试（或换一个 file_path）。");
                }

                // M28（第三轮续）：覆盖其他既有文件前先备份（可恢复），备份失败则中止
                string backupPath = null;
                if (overwrite && System.IO.File.Exists(guarded) &&
                    !string.Equals(guarded, currentFullName, StringComparison.OrdinalIgnoreCase))
                {
                    try { backupPath = ExcelAiAssistant.Core.Security.PathGuard.CreateBackup(guarded); }
                    catch (Exception backupException)
                    {
                        return Invalid("BACKUP_FAILED", "覆盖前备份既有文件失败，已中止写入：" + backupException.Message);
                    }
                }

                // SaveAs 覆盖既有文件会弹确认框——DisplayAlerts=true 的宿主 Excel 上 out-of-proc 调用会模态卡死，显式压制。
                prevAlerts = Convert.ToBoolean(application.DisplayAlerts);
                application.DisplayAlerts = false;
                workbook.SaveAs(guarded, format);
                return new ToolResult
                {
                    Ok = true,
                    Data = new { saved = true, file_path = guarded, format = format, backup_path = backupPath }
                };
            }
            finally
            {
                if (prevAlerts.HasValue)
                {
                    try { application.DisplayAlerts = prevAlerts.Value; } catch { }
                }

                Release((object)workbook);
            }
        }

        public ToolResult ListExternalLinks(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var links = new List<DictLink>();
                EnumerateLinks(workbook, links);
                return new ToolResult
                {
                    Ok = true,
                    Data = new { link_count = links.Count, links = links.Select(l => new { source = l.Source, kind = l.Kind }) }
                };
            }
            finally
            {
                Release((object)workbook);
            }
        }

        private void EnumerateLinks(dynamic workbook, List<DictLink> links)
        {
            try
            {
                var sources = workbook.LinkSources(1); // xlLinkTypeExcelLinks
                if (sources is object[,] array)
                {
                    for (var i = 0; i < array.GetLength(0); i++)
                    {
                        var source = Convert.ToString(array[i, 1] ?? array[i, 0]);
                        links.Add(new DictLink { Source = source, Kind = "excel" });
                    }
                }
            }
            catch
            {
            }

            try
            {
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    try
                    {
                        foreach (dynamic connected in worksheet.QueryTables)
                        {
                            var connection = Convert.ToString(connected.Connection);
                            links.Add(new DictLink { Source = connection, Kind = "query" });
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        public ToolResult ReplaceExternalLinks(Dictionary<string, object> args)
        {
            var oldLink = Required(args, "old_link");
            var newLink = Required(args, "new_link");
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                workbook.ChangeLink(oldLink, newLink, 1); // xlLinkTypeExcelLinks
                return new ToolResult
                {
                    Ok = true,
                    Data = new { old_link = oldLink, new_link = newLink, replaced = true }
                };
            }
            finally
            {
                Release((object)workbook);
            }
        }

        /// <summary>
        /// 从数组参数提取字符串列表：兼容纯字符串项与对象项（schema BoxArray 声明的
        /// {path}/{name} 结构）。此前直接 Convert.ToString(item) 会把对象项序列化成
        /// 类名（"System.Collections.Generic.Dictionary`2…"），导致按 schema 传参必然
        /// 失效：merge_workbooks 静默合入 0 个文件、split_workbook 抛 DISP_E_BADINDEX。
        /// （M6 WPS 全量审查 F9 实测发现，Excel 版同源同病。）
        /// </summary>
        private static List<string> ArgsNameArray(Dictionary<string, object> args, string key, string itemKey)
        {
            var result = new List<string>();
            if (!args.TryGetValue(key, out var raw) || !(raw is System.Collections.IEnumerable enumerable) || raw is string)
            {
                return result;
            }

            foreach (var item in enumerable)
            {
                var value = ExtractItemString(item, itemKey);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    result.Add(value);
                }
            }

            return result;
        }

        private static string ExtractItemString(object item, string itemKey)
        {
            if (item == null)
            {
                return null;
            }

            if (item is System.Collections.IDictionary map)
            {
                // 对象项缺声明键：返回 null 让调用方跳过，禁止类名字符串混入结果（M6 复审修复）。
                if (!(map.Contains(itemKey) && map[itemKey] != null))
                {
                    return null;
                }

                var v = map[itemKey];
                return Convert.ToString(v is Newtonsoft.Json.Linq.JValue jv ? jv.Value : v);
            }

            if (item is Newtonsoft.Json.Linq.JObject jo)
            {
                if (!(jo.TryGetValue(itemKey, out var tv) && tv != null))
                {
                    return null;
                }

                return Convert.ToString(tv is Newtonsoft.Json.Linq.JValue jv2 ? jv2.Value : (object)tv);
            }

            return Convert.ToString(item is Newtonsoft.Json.Linq.JValue j3 ? j3.Value : item);
        }

        private static List<string> AllSheetNames(dynamic workbook)
        {
            var names = new List<string>();
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                names.Add(Convert.ToString(worksheet.Name));
            }

            return names;
        }

        private dynamic ResolveSheet(dynamic workbook, string sheetName)
        {
            return string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static string SanitizeFileName(string name)
        {
            var invalid = System.IO.Path.GetInvalidFileNameChars();
            var chars = name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
            return new string(chars);
        }

        /// <summary>M28 审查（第三轮）：布尔参数严格解析——只接受 true/false/1/0，非法值返回参数错误（不是 COM 错误）。</summary>
        private static bool TryReadBoolArg(Dictionary<string, object> args, string key, out bool value, out string error)
        {
            value = false;
            error = null;
            if (!args.TryGetValue(key, out var raw) || raw == null)
            {
                return true;
            }

            if (raw is Newtonsoft.Json.Linq.JValue jv) { raw = jv.Value; }
            if (raw is bool boolean) { value = boolean; return true; }
            var text = Convert.ToString(raw)?.Trim();
            if (bool.TryParse(text, out value)) { return true; }
            if (text == "1") { value = true; return true; }
            if (text == "0") { value = false; return true; }
            error = "参数 " + key + " 需布尔值（true/false，或 1/0），当前：" + Convert.ToString(raw);
            return false;
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
            catch
            {
            }
        }

        private static ToolResult Invalid(string code, string message)
        {
            return new ToolResult { Ok = false, ErrorCode = code, Data = message };
        }

        private sealed class DictLink
        {
            public string Source { get; set; }

            public string Kind { get; set; }
        }
    }
}