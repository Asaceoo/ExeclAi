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
            dynamic workbook = null;
            var temps = new List<dynamic>();
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                var written = new List<string>();
                var targetNames = sheets.Count == 0
                    ? AllSheetNames(workbook)
                    : sheets;
                foreach (var name in targetNames)
                {
                    dynamic sourceSheet = null;
                    dynamic temp = null;
                    try
                    {
                        sourceSheet = workbook.Worksheets.Item[name];
                        temp = application.Workbooks.Add();
                        temps.Add(temp);
                        sourceSheet.Copy(After: temp.Worksheets.Item[1]);
                        var filePath = System.IO.Path.Combine(outputFolder, SanitizeFileName(Convert.ToString(sourceSheet.Name)) + ".xlsx");
                        temp.SaveAs(filePath, 51); // xlOpenXMLWorkbook
                        temp.Close(false);
                        written.Add(filePath);
                    }
                    catch (Exception ex)
                    {
                        return Invalid("SPLIT_FAILED", "拆分失败（" + name + "）：" + ex.Message);
                    }
                    finally
                    {
                        // M16 全量审查：失败路径上 temp 未 SaveAs/Close，会残留空白工作簿——
                        // finally 中兜底关闭（written 路径已 Close，重复 Close 幂等吞异常）。
                        if (temp != null)
                        {
                            try { temp.Close(false); } catch { }
                        }

                        Release((object)temp);
                        Release((object)sourceSheet);
                    }
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { output_folder = outputFolder, written_files = written }
                };
            }
            finally
            {
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