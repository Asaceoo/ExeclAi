using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    public sealed class ExcelStructureService
    {
        private readonly dynamic application;
        private readonly Action<string, object, string, object, object> recordStructuralUndo;

        public ExcelStructureService(
            dynamic application,
            Action<string, object, string, object, object> recordStructuralUndo = null)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.recordStructuralUndo = recordStructuralUndo;
        }

        public ToolResult AddWorksheet(Dictionary<string, object> args)
        {
            var name = Required(args, "name");
            var nameError = ValidateSheetName(name);
            if (nameError != null)
            {
                return Invalid("INVALID_SHEET_NAME", nameError);
            }
            dynamic workbook = null;
            dynamic before = null;
            dynamic after = null;
            dynamic added = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");

                if (SheetExists(workbook, name))
                {
                    return Invalid("SHEET_EXISTS", $"工作表已存在：{name}");
                }

                if (TryGetOptional(args, "before_sheet", out var beforeName))
                {
                    before = workbook.Worksheets.Item[beforeName];
                    added = workbook.Worksheets.Add(Before: before);
                }
                else if (TryGetOptional(args, "after_sheet", out var afterName))
                {
                    after = workbook.Worksheets.Item[afterName];
                    added = workbook.Worksheets.Add(After: after);
                }
                else
                {
                    added = workbook.Worksheets.Add();
                }

                added.Name = name;
                // 新建工作表必须留结构撤销（undoKind=sheet-insert 由宿主 Undo 分派删除该表），
                // 否则 add_excel_worksheet 返回无 operation_id、用户无法回滚本批次新增的表。
                try
                {
                    recordStructuralUndo?.Invoke(
                        "add_excel_worksheet", args, "sheet-insert",
                        new { sheet = name },
                        new { sheets = new[] { name } });
                }
                catch (Exception undoException)
                {
                    AppLog.Warn("structure", "add_excel_worksheet 结构撤销录制失败（建表本身已成功）：" + undoException.Message);
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { sheet = name, index = Convert.ToInt32(added.Index) }
                };
            }
            finally
            {
                Release((object)added);
                Release((object)after);
                Release((object)before);
                workbook = null;
            }
        }

        public ToolResult RenameWorksheet(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));
                var oldName = Convert.ToString(sheet.Name);
                var newName = Required(args, "name");
                if (string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
                {
                    return new ToolResult { Ok = true, Data = new { old_name = oldName, new_name = newName } };
                }

                if (SheetExists(workbook, newName))
                {
                    return Invalid("SHEET_EXISTS", $"工作表已存在：{newName}");
                }

                var renameError = ValidateSheetName(newName);
                if (renameError != null)
                {
                    return Invalid("INVALID_SHEET_NAME", renameError);
                }

                sheet.Name = newName;
                // 重命名可回滚：记录旧名/新名，宿主 Undo 按 sheet-rename 分派还原。
                try
                {
                    recordStructuralUndo?.Invoke(
                        "rename_excel_worksheet", args, "sheet-rename",
                        new { sheet = oldName, renamed_to = newName },
                        new { old_name = oldName, new_name = newName });
                }
                catch (Exception undoException)
                {
                    AppLog.Warn("structure", "rename_excel_worksheet 结构撤销录制失败（改名本身已成功）：" + undoException.Message);
                }

                return new ToolResult { Ok = true, Data = new { old_name = oldName, new_name = newName, undoable = true } };
            }
            finally
            {
                sheet = null;
                workbook = null;
            }
        }

        public ToolResult DeleteWorksheet(Dictionary<string, object> args)
        {
            dynamic workbook = null;
            dynamic sheet = null;
            try
            {
                workbook = application.ActiveWorkbook ??
                           throw new InvalidOperationException("当前没有打开的工作簿。");
                sheet = string.IsNullOrWhiteSpace(Optional(args, "sheet"))
                    ? application.ActiveSheet
                    : workbook.Worksheets.Item[Optional(args, "sheet")];

                if (Convert.ToInt32(workbook.Worksheets.Count) == 1)
                {
                    return Invalid("LAST_WORKSHEET", "不能删除工作簿中唯一的工作表。");
                }

                var name = Convert.ToString(sheet.Name);
                sheet.Delete();
                // 删除工作表不复原：不保存整表内容/格式/图表快照，回滚无法恢复——必须如实告知，不给假承诺。
                return new ToolResult
                {
                    Ok = true,
                    Warning = "已删除工作表「" + name + "」。该操作不可回滚（未保留表内容快照）；如需保留，请先复制该表再删除。",
                    Data = new { deleted = name, undoable = false }
                };
            }
            finally
            {
                sheet = null;
                workbook = null;
            }
        }

        /// <summary>
        /// 把区域转为智能表（ListObject）。WPS ET 同样支持 ListObjects.Add，宿主无关。
        /// </summary>
        public ToolResult CreateTable(Dictionary<string, object> args)
        {
            var address = Required(args, "address");
            var requestedName = Optional(args, "name");
            var hasHeaders = !TryGetOptional(args, "has_headers", out var hasHeadersRaw) || ToBool(hasHeadersRaw);
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic tables = null;
            dynamic added = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));

                if (Convert.ToBoolean(sheet.ProtectContents))
                {
                    return Invalid("SHEET_PROTECTED", "工作表处于保护状态，无法创建智能表；请先解除保护（可用 verify_protection 探测保护状态）。");
                }

                range = sheet.Range[address];
                tables = sheet.ListObjects;

                if (!string.IsNullOrWhiteSpace(requestedName))
                {
                    // 工具分类审查第八批修复（第四批命名预校验家族残留）：ListObject 表名有独立规则
                    // （禁空格/禁数字开头/禁与单元格引用冲突/仅字母数字下划线）——违规此前抛裸 COM 异常
                    // 无法自纠，现人话 INVALID_TABLE_NAME。
                    var tableError = ValidateTableName(requestedName);
                    if (tableError != null)
                    {
                        return Invalid("INVALID_TABLE_NAME", tableError);
                    }

                    if (TableExists(tables, requestedName))
                    {
                        return Invalid("TABLE_NAME_EXISTS", $"同表已存在同名智能表：{requestedName}");
                    }
                }

                // xlSrcRange=1；XlListObjectHasHeaders：xlYes=1 / xlNo=2（COM 枚举值，避开 dynamic 绑定对枚举的解析）。
                added = tables.Add(1, range, hasHeaders ? 1 : 2);

                if (!string.IsNullOrWhiteSpace(requestedName))
                {
                    added.Name = requestedName;
                }

                var style = Optional(args, "table_style");
                if (!string.IsNullOrWhiteSpace(style))
                {
                    added.TableStyle = style;
                }

                var finalName = Convert.ToString(added.Name);
                var finalAddress = Convert.ToString(added.Range.Address);
                var rowCount = Convert.ToInt32(added.Range.Rows.Count);
                var colCount = Convert.ToInt32(added.Range.Columns.Count);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        name = finalName,
                        address = finalAddress,
                        row_count = rowCount,
                        column_count = colCount,
                        has_headers = hasHeaders,
                        hint = "建议随后调用 verify_table 核对表名/表头/行列数。"
                    }
                };
            }
            finally
            {
                Release((object)added);
                Release((object)tables);
                Release((object)range);
                sheet = null;
                workbook = null;
            }
        }

        /// <summary>
        /// 在指定锚点单元格左上角插入本地图片（Shapes.AddPicture，宿主无关，WPS ET 同样支持）。
        /// LinkToFile=msoFalse(0) / SaveWithDocument=msoTrue(1)：图片随工作簿保存，不依赖源文件。
        /// 宽高单位为磅；缺省传 -1 -1 由宿主按图片原始尺寸插入。
        /// </summary>
        public ToolResult InsertImage(Dictionary<string, object> args)
        {
            var imagePath = Required(args, "image_path");
            var anchor = Required(args, "anchor");
            // 宽高（磅）：缺省 -1 -1 由宿主按图片原始尺寸插入；显式传值必须为正数
            // （审查 F1：0/负数会插入不可见图片或触发宿主异常，非数字会抛 FormatException 被误报为 COM 错误）。
            double width = -1d;
            double height = -1d;
            if (TryGetOptional(args, "width", out var widthText))
            {
                if (!double.TryParse(widthText, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var parsedWidth) || parsedWidth <= 0)
                {
                    return Invalid("INVALID_ARGS", $"width 必须为正数（磅），实际收到 '{widthText}'；缺省不传则按图片原始尺寸插入。");
                }

                width = parsedWidth;
            }

            if (TryGetOptional(args, "height", out var heightText))
            {
                if (!double.TryParse(heightText, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var parsedHeight) || parsedHeight <= 0)
                {
                    return Invalid("INVALID_ARGS", $"height 必须为正数（磅），实际收到 '{heightText}'；缺省不传则按图片原始尺寸插入。");
                }

                height = parsedHeight;
            }
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic shapes = null;
            dynamic shape = null;
            try
            {
                if (!System.IO.File.Exists(imagePath))
                {
                    return Invalid("IMAGE_NOT_FOUND", $"图片文件不存在：{imagePath}");
                }

                // 审查 F4 补强（Excel 真机实证）：Excel 的 AddPicture 对垃圾字节文件可能宽容接受（静默插入
                // 无效图片），WPS 则抛 E_FAIL——宿主过滤器行为不可控。前置魔数白名单嗅探（PNG/JPEG/GIF/
                // BMP/TIFF/EMF/WMF），未识别格式在触碰宿主前显式拒绝，双宿主行为确定一致。
                if (!HasRecognizedImageHeader(imagePath))
                {
                    return Invalid("IMAGE_INVALID", $"文件不是可识别的图片格式（支持 PNG/JPEG/GIF/BMP/TIFF/EMF/WMF）：{imagePath}");
                }

                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));

                if (Convert.ToBoolean(sheet.ProtectContents))
                {
                    return Invalid("SHEET_PROTECTED", "工作表处于保护状态，无法插入图片；请先解除保护（可用 verify_protection 探测保护状态）。");
                }

                range = sheet.Range[anchor];
                shapes = sheet.Shapes;
                // msoFalse=0（不链接到源文件）、msoTrue=1（随文档保存）；COM 枚举用整数值避开 dynamic 绑定歧义。
                // 审查 F4：损坏/非图片文件 AddPicture 抛裸 COM 异常（探针实测 E_FAIL），包装为显式错误码。
                try
                {
                    shape = shapes.AddPicture(imagePath, 0, 1, range.Left, range.Top, width, height);
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return Invalid("IMAGE_INVALID", $"图片无法插入（宿主拒绝，可能不是有效图片文件或格式不受支持）：{imagePath}；宿主报错：{ex.Message}");
                }

                var shapeName = Convert.ToString(shape.Name);
                var left = Convert.ToDouble(shape.Left);
                var top = Convert.ToDouble(shape.Top);
                var finalWidth = Convert.ToDouble(shape.Width);
                var finalHeight = Convert.ToDouble(shape.Height);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        shape_name = shapeName,
                        anchor,
                        left,
                        top,
                        width = finalWidth,
                        height = finalHeight,
                        hint = "建议随后调用 verify_image 核对图片锚点/尺寸。"
                    }
                };
            }
            finally
            {
                Release((object)shape);
                Release((object)shapes);
                Release((object)range);
                sheet = null;
                workbook = null;
            }
        }

        /// <summary>检查文件头魔数是否为受支持的图片格式（前 8 字节足够区分常见位图/矢量图容器）。</summary>
        private static bool HasRecognizedImageHeader(string path)
        {
            try
            {
                var header = new byte[8];
                var offset = 0;
                using (var stream = System.IO.File.OpenRead(path))
                {
                    int read;
                    while (offset < header.Length && (read = stream.Read(header, offset, header.Length - offset)) > 0)
                    {
                        offset += read;
                    }
                }

                if (offset < 4)
                {
                    return false;
                }

                // PNG: 89 50 4E 47；JPEG: FF D8 FF；GIF: "GIF8"；BMP: "BM"。
                if (header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47) return true;
                if (header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF) return true;
                if (header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x38) return true;
                if (header[0] == 0x42 && header[1] == 0x4D) return true;
                // TIFF: "II*\0" 或 "MM\0*"。
                if (header[0] == 0x49 && header[1] == 0x49 && header[2] == 0x2A && header[3] == 0x00) return true;
                if (header[0] == 0x4D && header[1] == 0x4D && header[2] == 0x00 && header[3] == 0x2A) return true;
                // EMF: 01 00 00 00（头部后续含 " EMF" 签名）；WMF placeable: D7 CD C6 9A，标准: 01 00 09 00。
                if (header[0] == 0xD7 && header[1] == 0xCD && header[2] == 0xC6 && header[3] == 0x9A) return true;
                if (header[0] == 0x01 && header[1] == 0x00 && header[2] == 0x09 && header[3] == 0x00) return true;
                if (header[0] == 0x01 && header[1] == 0x00 && header[2] == 0x00 && header[3] == 0x00
                    && offset >= 8 && header[4] == 0x20 && header[5] == 0x00 && header[6] == 0x00
                    && header[7] == 0x00)
                {
                    return true;
                }

                return false;
            }
            catch (Exception xaiEx) {
                AppLog.Warn("structure", "ExcelStructureService.cs:HasRecognizedImageHeader 静默兜底触发: " + xaiEx.Message);
                // 文件读取失败（被占用/权限不足）交由宿主报错，不在嗅探层误杀。
                return true;
            }
        }

        /// <summary>
        /// 冻结/解冻窗口窗格。冻结是窗口级操作（绑定到活动窗口下目标工作表的视图），
        /// 因此先激活目标工作表再设置 ActiveWindow.FreezePanes；cell 为冻结锚点，
        /// 其上方所有行与左侧所有列被冻结（如 B2 → 冻结第 1 行和 A 列）。A1 无法作为锚点。
        /// </summary>
        public ToolResult FreezePanes(Dictionary<string, object> args)
        {
            var unfreeze = TryGetOptional(args, "unfreeze", out var unfreezeRaw) && ToBool(unfreezeRaw);
            var cell = Optional(args, "cell");
            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic window = null;
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));

                if (!unfreeze && string.IsNullOrWhiteSpace(cell))
                {
                    return Invalid("MISSING_CELL", "冻结窗格需要提供锚点单元格（cell，如 B2）；仅解除冻结时可省略。");
                }

                sheet.Activate();
                window = application.ActiveWindow;

                // 先解除已有冻结并把分割量归零，避免旧冻结/分割残留污染本次设置。
                window.FreezePanes = false;
                window.SplitRow = 0;
                window.SplitColumn = 0;
                if (unfreeze)
                {
                    return new ToolResult
                    {
                        Ok = true,
                        Data = new { frozen = false, sheet = Convert.ToString(sheet.Name) }
                    };
                }

                range = sheet.Range[cell];
                var anchorRow = Convert.ToInt32(range.Row);
                var anchorColumn = Convert.ToInt32(range.Column);
                if (anchorRow == 1 && anchorColumn == 1)
                {
                    return Invalid("INVALID_ANCHOR", "A1 不能作为冻结锚点（其上方无行、左侧无列）；请改用如 B2（冻结第 1 行和 A 列）、A2（仅冻结第 1 行）、B1（仅冻结 A 列）。");
                }

                range.Select();
                // WPS 实测（M13 探针）：Excel 惯用的「Select 锚点 + FreezePanes=true」在 WPS ET 无效
                // （赋值被静默忽略）；双宿主通用写法是先设 SplitRow/SplitColumn 再 FreezePanes=true。
                window.SplitRow = anchorRow - 1;
                window.SplitColumn = anchorColumn - 1;
                window.FreezePanes = true;

                var frozen = Convert.ToBoolean(window.FreezePanes);
                if (!frozen)
                {
                    return Invalid("FREEZE_FAILED", "宿主未接受冻结设置（FreezePanes 读回仍为 false）；请确认锚点单元格不在合并区域或拆分视图内。");
                }

                var splitRow = Convert.ToInt32(window.SplitRow);
                var splitColumn = Convert.ToInt32(window.SplitColumn);
                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        frozen = true,
                        anchor = cell,
                        split_row = splitRow,
                        split_column = splitColumn,
                        hint = "建议随后调用 verify_freeze_panes 核对冻结状态与位置。"
                    }
                };
            }
            finally
            {
                Release((object)window);
                Release((object)range);
                sheet = null;
                workbook = null;
            }
        }

        /// <summary>
        /// 超链接管理（M15）：action=list 枚举整表超链接（可按 address 过滤）；add 在目标单元格新建
        /// （link_url 外部网址 / link_subaddress 工作簿内部位置，二选一）；update 修改目标单元格超链接；
        /// remove 删除区域内的超链接；remove_all 清空整表。写动作前置工作表保护检查。
        /// </summary>
        public ToolResult ManageHyperlinks(Dictionary<string, object> args)
        {
            var action = (Optional(args, "action") ?? "list").Trim().ToLowerInvariant();
            if (action != "list" && action != "add" && action != "update" && action != "remove" && action != "remove_all")
            {
                return Invalid("INVALID_ARGS", $"未知 action：'{action}'；支持 list/add/update/remove/remove_all。");
            }

            var address = Optional(args, "address");
            var mutating = action != "list";
            if (mutating && action != "remove_all" && string.IsNullOrWhiteSpace(address))
            {
                return Invalid("INVALID_ARGS", $"action={action} 需要提供 address（目标单元格或区域，如 B2 / A1:A5）。");
            }

            var linkUrl = Optional(args, "link_url");
            var linkSubAddress = Optional(args, "link_subaddress");
            if (action == "add")
            {
                if (string.IsNullOrWhiteSpace(linkUrl) && string.IsNullOrWhiteSpace(linkSubAddress))
                {
                    return Invalid("INVALID_ARGS", "action=add 需要提供 link_url（外部网址）或 link_subaddress（内部位置，如 目录!A1）之一。");
                }

                if (!string.IsNullOrWhiteSpace(linkUrl) && !string.IsNullOrWhiteSpace(linkSubAddress))
                {
                    return Invalid("INVALID_ARGS", "link_url 与 link_subaddress 互斥：外部网址走 link_url，工作簿内部位置走 link_subaddress。");
                }
            }

            dynamic workbook = null;
            dynamic sheet = null;
            dynamic range = null;
            dynamic hyperlinks = null;
            var handles = new List<HyperlinkHandle>();
            try
            {
                Resolve(ref workbook, ref sheet, Optional(args, "sheet"));

                if (mutating && Convert.ToBoolean(sheet.ProtectContents))
                {
                    return Invalid("SHEET_PROTECTED", "工作表处于保护状态，无法修改超链接；请先解除保护（可用 verify_protection 探测保护状态）。");
                }

                if (!string.IsNullOrWhiteSpace(address))
                {
                    // 地址非法时显式报错，而非让宿主抛报错文案不可控的裸 COM 异常。
                    try
                    {
                        range = sheet.Range[address];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return Invalid("INVALID_ADDRESS", $"目标地址 '{address}' 无法解析为有效单元格/区域。");
                    }
                }

                hyperlinks = sheet.Hyperlinks;
                switch (action)
                {
                    case "list":
                        return ListHyperlinks(hyperlinks, range);
                    case "add":
                        return AddHyperlinkAt(hyperlinks, range, address, linkUrl, linkSubAddress,
                            Optional(args, "display_text"), Optional(args, "tooltip"));
                    case "update":
                        return UpdateHyperlinkAt(hyperlinks, range, linkUrl, linkSubAddress, Optional(args, "display_text"));
                    case "remove":
                        return RemoveHyperlinksIn(hyperlinks, range, address);
                    case "remove_all":
                        return RemoveAllHyperlinks(hyperlinks);
                    default:
                        return Invalid("INVALID_ARGS", $"未知 action：'{action}'。");
                }
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }

                Release((object)hyperlinks);
                Release((object)range);
                sheet = null;
                workbook = null;
            }
        }

        private sealed class HyperlinkHandle : IDisposable
        {
            public int Row { get; set; }

            public int Column { get; set; }

            /// <summary>链接锚定范围的行跨度（单格链接为 1；用户手工建的多格链接可能 &gt; 1）。</summary>
            public int RowSpan { get; set; }

            /// <summary>链接锚定范围的列跨度（单格链接为 1）。</summary>
            public int ColSpan { get; set; }

            public string Cell { get; set; }

            public dynamic Link { get; set; }

            /// <summary>链接范围是否覆盖目标格（审查 F1：多格链接不能只按首格判重）。</summary>
            public bool Covers(int row, int column)
            {
                return row >= Row && row < Row + RowSpan
                       && column >= Column && column < Column + ColSpan;
            }

            /// <summary>链接范围是否与目标区域相交（审查 F1：删除按相交判定，首格在区外也要能删）。</summary>
            public bool Intersects(int top, int left, int rows, int cols)
            {
                return Row < top + rows && Row + RowSpan > top
                       && Column < left + cols && Column + ColSpan > left;
            }

            public void Dispose()
            {
                if (Link != null && System.Runtime.InteropServices.Marshal.IsComObject((object)Link))
                {
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)Link); }
                    catch (Exception)
                    {
                        AppLog.Warn("ExcelStructureService", "[LogCover] ExcelStructureService.cs:559  silent-catch covered (no exception variable)");
                        // 链接被 Delete 后宿主 RCW 状态不可控，释放失败不阻断收尾。
                    }
                }

                Link = null;
            }
        }

        /// <summary>枚举整表超链接为 (行, 列, 跨度, COM 引用) 快照；调用方负责 Dispose 每个 handle。</summary>
        private List<HyperlinkHandle> CollectHyperlinks(dynamic hyperlinks)
        {
            var result = new List<HyperlinkHandle>();
            var total = Convert.ToInt32(hyperlinks.Count);
            for (var index = 1; index <= total; index++)
            {
                dynamic item = null;
                dynamic anchor = null;
                try
                {
                    item = hyperlinks.Item[index];
                    anchor = item.Range;
                    result.Add(new HyperlinkHandle
                    {
                        Row = Convert.ToInt32(anchor.Row),
                        Column = Convert.ToInt32(anchor.Column),
                        RowSpan = Convert.ToInt32(anchor.Rows.Count),
                        ColSpan = Convert.ToInt32(anchor.Columns.Count),
                        Cell = ExcelRangeHelper.ColumnName(Convert.ToInt32(anchor.Column)) + Convert.ToInt32(anchor.Row),
                        Link = item
                    });
                    item = null; // 所有权移交 handle，finally 不再释放
                }
                finally
                {
                    Release((object)anchor);
                    Release((object)item);
                }
            }

            return result;
        }

        private static bool IsWithinBounds(HyperlinkHandle handle, dynamic range)
        {
            return handle.Row >= Convert.ToInt32(range.Row)
                   && handle.Row < Convert.ToInt32(range.Row) + Convert.ToInt32(range.Rows.Count)
                   && handle.Column >= Convert.ToInt32(range.Column)
                   && handle.Column < Convert.ToInt32(range.Column) + Convert.ToInt32(range.Columns.Count);
        }

        private ToolResult ListHyperlinks(dynamic hyperlinks, dynamic filterRange)
        {
            var links = new List<object>();
            var handles = CollectHyperlinks(hyperlinks);
            try
            {
                foreach (var handle in handles)
                {
                    if (filterRange != null && !IsWithinBounds(handle, filterRange))
                    {
                        continue;
                    }

                    var text = SafeProperty(() => handle.Link.TextToDisplay);
                    var url = SafeProperty(() => handle.Link.Address);
                    var subAddress = SafeProperty(() => handle.Link.SubAddress);
                    var tooltip = SafeProperty(() => handle.Link.ScreenTip);
                    links.Add(new
                    {
                        cell = handle.Cell,
                        text = text ?? string.Empty,
                        url = url ?? string.Empty,
                        subaddress = subAddress ?? string.Empty,
                        tooltip = tooltip ?? string.Empty
                    });
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        action = "list",
                        hyperlink_count = links.Count,
                        total_in_sheet = handles.Count,
                        links
                    }
                };
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }
            }
        }

        private ToolResult AddHyperlinkAt(
            dynamic hyperlinks, dynamic range, string address,
            string linkUrl, string linkSubAddress, string displayText, string tooltip)
        {
            if (Convert.ToInt32(range.Rows.Count) * Convert.ToInt32(range.Columns.Count) > 1)
            {
                return Invalid("INVALID_ARGS", $"action=add 的 address 必须是单个单元格，实际收到区域 '{address}'。");
            }

            var anchorRow = Convert.ToInt32(range.Row);
            var anchorColumn = Convert.ToInt32(range.Column);
            var handles = CollectHyperlinks(hyperlinks);
            try
            {
                foreach (var handle in handles)
                {
                    // 审查 F1：判重按「目标格被任一现有链接范围覆盖」，多格链接覆盖到的格同样拦截。
                    if (handle.Covers(anchorRow, anchorColumn))
                    {
                        return Invalid("HYPERLINK_EXISTS", $"单元格 '{handle.Cell}' 已存在覆盖该格的超链接（范围 {handle.RowSpan} 行 {handle.ColSpan} 列）；如需更换请用 action=update，如需删除请用 action=remove。");
                    }
                }

                dynamic added = null;
                try
                {
                    // 与 create_toc_sheet 相同的宿主无关写法：3 参固定（审查 F4），外部网址走 Address、
                    // 内部位置走 SubAddress，另一项传空串，避免依赖宿主对省略参数的默认绑定。
                    added = string.IsNullOrWhiteSpace(linkSubAddress)
                        ? hyperlinks.Add(range, linkUrl, string.Empty)
                        : hyperlinks.Add(range, string.Empty, linkSubAddress);

                    var warnings = new List<string>();
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(tooltip))
                        {
                            added.ScreenTip = tooltip;
                        }

                        if (!string.IsNullOrWhiteSpace(displayText))
                        {
                            added.TextToDisplay = displayText;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        // 链接本体已创建成功，附加属性设置失败不推翻本次操作，如实附警告。
                        warnings.Add("链接已创建，但 ScreenTip/TextToDisplay 设置失败：" + ex.Message);
                    }

                    var cell = ExcelRangeHelper.ColumnName(anchorColumn) + anchorRow;
                    var data = new Dictionary<string, object>
                    {
                        ["action"] = "add",
                        ["cell"] = cell,
                        ["url"] = SafeProperty(() => added.Address) ?? string.Empty,
                        ["subaddress"] = SafeProperty(() => added.SubAddress) ?? string.Empty,
                        ["text"] = SafeProperty(() => added.TextToDisplay) ?? string.Empty,
                        ["hint"] = "建议随后调用 verify_hyperlink 核对链接目标。"
                    };
                    if (warnings.Count > 0)
                    {
                        data["warnings"] = warnings;
                    }

                    return new ToolResult { Ok = true, Data = data };
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return Invalid("HYPERLINK_ADD_FAILED", $"宿主拒绝创建超链接（地址格式可能非法）：{ex.Message}");
                }
                finally
                {
                    // 审查 F2：Add 返回的 Hyperlink COM 引用必须释放，长会话不累积 RCW。
                    Release((object)added);
                }
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }
            }
        }

        private ToolResult UpdateHyperlinkAt(
            dynamic hyperlinks, dynamic range, string linkUrl, string linkSubAddress, string displayText)
        {
            if (linkUrl == null && linkSubAddress == null && displayText == null)
            {
                return Invalid("INVALID_ARGS", "action=update 需要至少提供 link_url / link_subaddress / display_text 之一。");
            }

            if (Convert.ToInt32(range.Rows.Count) * Convert.ToInt32(range.Columns.Count) > 1)
            {
                return Invalid("INVALID_ARGS", "action=update 的 address 必须是单个单元格。");
            }

            var anchorRow = Convert.ToInt32(range.Row);
            var anchorColumn = Convert.ToInt32(range.Column);
            var handles = CollectHyperlinks(hyperlinks);
            try
            {
                HyperlinkHandle target = null;
                foreach (var handle in handles)
                {
                    if (handle.Row == anchorRow && handle.Column == anchorColumn)
                    {
                        target = handle;
                        break;
                    }
                }

                if (target == null)
                {
                    return Invalid("HYPERLINK_NOT_FOUND", $"单元格 '{ExcelRangeHelper.ColumnName(anchorColumn)}{anchorRow}' 上没有超链接可更新。");
                }

                try
                {
                    // M15 探针实证（Excel 真机）：外→外改 Address、任意→内改 SubAddress 均可靠生效；
                    // 但对链接赋 SubAddress="" 会把外链打回「内链空目标」（Address 被清空）——禁止该赋值。
                    // 内→外切换属性赋值语义不可控，统一走 Delete+Add 重建（探针 C 实证可靠）。
                    var currentSub = SafeProperty(() => target.Link.SubAddress) ?? string.Empty;
                    var wasInternal = currentSub.Trim().Length > 0;
                    var needsRebuild = linkUrl != null && wasInternal;

                    if (needsRebuild)
                    {
                        var preservedText = displayText;
                        if (preservedText == null)
                        {
                            var existingText = SafeProperty(() => target.Link.TextToDisplay);
                            if (!string.IsNullOrEmpty(existingText))
                            {
                                preservedText = existingText;
                            }
                        }

                        dynamic anchorRange = null;
                        dynamic rebuilt = null;
                        try
                        {
                            anchorRange = target.Link.Range;
                            target.Link.Delete();
                            rebuilt = hyperlinks.Add(anchorRange, linkUrl, string.Empty);
                            if (preservedText != null)
                            {
                                rebuilt.TextToDisplay = preservedText;
                            }
                        }
                        finally
                        {
                            Release((object)rebuilt);
                            Release((object)anchorRange);
                        }
                    }
                    else
                    {
                        if (linkUrl != null)
                        {
                            target.Link.Address = linkUrl;
                        }
                        else if (linkSubAddress != null)
                        {
                            target.Link.SubAddress = linkSubAddress;
                        }

                        if (displayText != null)
                        {
                            target.Link.TextToDisplay = displayText;
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    return Invalid("HYPERLINK_UPDATE_FAILED", "宿主拒绝更新超链接（地址格式可能非法）：" + ex.Message);
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new
                    {
                        action = "update",
                        cell = target.Cell,
                        url = SafeProperty(() => target.Link.Address) ?? string.Empty,
                        subaddress = SafeProperty(() => target.Link.SubAddress) ?? string.Empty,
                        text = SafeProperty(() => target.Link.TextToDisplay) ?? string.Empty,
                        hint = "建议随后调用 verify_hyperlink 核对链接目标。"
                    }
                };
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }
            }
        }

        private ToolResult RemoveHyperlinksIn(dynamic hyperlinks, dynamic range, string address)
        {
            var top = Convert.ToInt32(range.Row);
            var left = Convert.ToInt32(range.Column);
            var rows = Convert.ToInt32(range.Rows.Count);
            var cols = Convert.ToInt32(range.Columns.Count);
            var handles = CollectHyperlinks(hyperlinks);
            try
            {
                var removed = new List<string>();
                foreach (var handle in handles)
                {
                    // 审查 F1：按范围相交判定，锚在区外的多格链接（如 A1:A3 删 A2）也能删干净。
                    if (handle.Intersects(top, left, rows, cols))
                    {
                        handle.Link.Delete();
                        removed.Add(handle.Cell);
                    }
                }

                if (removed.Count == 0)
                {
                    return Invalid("HYPERLINK_NOT_FOUND", $"区域 '{address}' 内没有超链接可删除。");
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { action = "remove", removed = removed.Count, cells = removed }
                };
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }
            }
        }

        private ToolResult RemoveAllHyperlinks(dynamic hyperlinks)
        {
            var handles = CollectHyperlinks(hyperlinks);
            try
            {
                if (handles.Count == 0)
                {
                    return Invalid("HYPERLINK_NOT_FOUND", "工作表上没有任何超链接。");
                }

                foreach (var handle in handles)
                {
                    handle.Link.Delete();
                }

                return new ToolResult
                {
                    Ok = true,
                    Data = new { action = "remove_all", removed = handles.Count }
                };
            }
            finally
            {
                foreach (var handle in handles)
                {
                    handle.Dispose();
                }
            }
        }

        /// <summary>读取 COM 属性并吞掉宿主抛出的异常（如部分超链接类型的 TextToDisplay 不可读）。</summary>
        private static string SafeProperty(Func<object> getter)
        {
            try
            {
                var value = getter();
                return value == null ? null : Convert.ToString(value);
            }
            catch (Exception xaiEx) {
                AppLog.Warn("structure", "ExcelStructureService.cs:SafeProperty 静默兜底触发: " + xaiEx.Message);
                return null;
            }
        }

        private static bool TableExists(dynamic tables, string name)
        {
            var count = Convert.ToInt32(tables.Count);
            for (var index = 1; index <= count; index++)
            {
                dynamic item = null;
                try
                {
                    item = tables.Item[index];
                    if (string.Equals(Convert.ToString(item.Name), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    Release((object)item);
                }
            }

            return false;
        }

        private static bool ToBool(object value)
        {
            if (value is bool b)
            {
                return b;
            }

            var text = Convert.ToString(value);
            return text == "1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase) || text == "是";
        }

        /// <summary>
        /// 工作表名预校验（工具分类审查第四批修复）：Excel 限制 ≤31 字符且不含 : \ / ? * [ ]——
        /// AI 起名常见违规（如带斜杠的「2026/09」、超长描述名）此前直接抛裸 COM 异常无法自纠，
        /// 现返回人话错误（INVALID_SHEET_NAME），AI 可立即改参数重试。返回 null 表示合法。
        /// </summary>
        public static string ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "工作表名不能为空。";
            }

            var trimmed = name.Trim();
            if (trimmed.Length == 0)
            {
                return "工作表名不能为空白。";
            }

            if (trimmed.Length > 31)
            {
                return $"工作表名超长（{trimmed.Length} 字符），Excel 上限 31 字符。当前：「{trimmed}」";
            }

            var illegal = ":\\/?*[]";
            foreach (var ch in trimmed)
            {
                if (illegal.IndexOf(ch) >= 0)
                {
                    return $"工作表名不能包含 : \\ / ? * [ ] 字符。当前：「{trimmed}」（非法字符 {ch}）。";
                }
            }

            return null;
        }

        /// <summary>
        /// 智能表（ListObject）名预校验（工具分类审查第八批，第四批命名预校验家族）：
        /// 规则——非空、长度 ≤255、仅字母/数字/下划线、不得以数字开头、不得与 A1/R1C1 单元格引用形态冲突。
        /// 返回 null 表示合法，否则人话错误。
        /// </summary>
        public static string ValidateTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "智能表名不能为空。";
            }

            var trimmed = name.Trim();
            if (trimmed.Length > 255)
            {
                return $"智能表名超长（{trimmed.Length} 字符），上限 255 字符。";
            }

            foreach (var ch in trimmed)
            {
                if (!(char.IsLetterOrDigit(ch) || ch == '_'))
                {
                    return $"智能表名只能包含字母/数字/下划线（不允许空格与符号）。当前：「{trimmed}」";
                }
            }

            if (char.IsDigit(trimmed[0]))
            {
                return $"智能表名不能以数字开头。当前：「{trimmed}」";
            }

            // 与 A1 引用形态冲突（如 "B2"、"XFD1048576"）——Excel 将其视为保留名拒绝。
            if (System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^[A-Za-z]{1,3}[0-9]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                return $"智能表名不能与单元格引用形态相同（如 B2）。当前：「{trimmed}」";
            }

            return null;
        }

        private void Resolve(ref dynamic workbook, ref dynamic sheet, string sheetName)
        {
            workbook = application.ActiveWorkbook ??
                       throw new InvalidOperationException("当前没有打开的工作簿。");
            sheet = string.IsNullOrWhiteSpace(sheetName)
                ? application.ActiveSheet
                : workbook.Worksheets.Item[sheetName];
        }

        private static bool SheetExists(dynamic workbook, string name)
        {
            foreach (dynamic worksheet in workbook.Worksheets)
            {
                if (string.Equals(Convert.ToString(worksheet.Name), name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetOptional(
            Dictionary<string, object> args,
            string key,
            out string value)
        {
            value = Optional(args, key);
            return !string.IsNullOrWhiteSpace(value);
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
