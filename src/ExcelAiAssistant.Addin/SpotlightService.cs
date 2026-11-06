using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Addin
{
    /// <summary>功能区「视图组」视图模式：聚光灯 / 指针。</summary>
    public enum SpotlightMode
    {
        Off = 0,
        Spotlight = 1,
        Pointer = 2
    }

    /// <summary>
    /// 聚光灯 / 指针：Excel 原生形状（Shape）方案。
    ///
    /// 历经三代的最终形态。第一代「条件格式 + 定义名」延迟在 Excel 重算链路（数百毫秒）；
    /// 第二代「Win32 分层窗叠加」在此环境的远程会话里被 DWM 丢弃半透明分层窗口（alpha&lt;255 不渲染，
    /// 且叠加窗还叠了 DPI 虚拟化/z 序/坐标换算三重坑）。Shape 方案由 Excel 自己渲染，与条件格式同管线：
    /// 半透明真实合成、必然可见；Shape 用工作表点坐标定位（与 Range.Left/Top 同系），
    /// 天然跟随滚动/缩放/DPI，无任何屏幕坐标换算；移动只写 4 个属性，零重算，选区变化即时跟随。
    ///
    /// 结构：每张工作表最多 2 个聚光灯形状（列带 + 行带）或 1 个指针包围盒形状，
    /// 按固定命名（JYYJ_VIEW_*）创建/复用/删除，不碰用户形状。
    /// 15ms 轮询只读探测（场景键 = 表|活动格|可见区），无变化零写入。
    /// 残留治理：启动/开启/关闭时清理旧版 CF 方案遗留（JYYJ_VIEW_MODE 哨兵规则 + JYYJ_* 定义名）
    /// 与本方案的 JYYJ_VIEW_* 形状。
    /// </summary>
    public sealed class SpotlightService : IDisposable
    {
        private const int XlExpression = 2;                 // xlExpression（表达式规则才读 Formula1）
        private const int MsoShapeRectangle = 1;
        private const int MsoFalse = 0;
        private const int MsoTrue = -1;
        private const int XlFreeFloating = 3;               // 形状不随单元格移动/缩放
        private const int YellowRgb = 65535;                // Excel RGB: R + G*256 + B*65536 → 黄
        private const int RedRgb = 255;                     // 红
        private const double FillTransparency = 0.55;       // 半透明填充，格内文字仍可读
        private const int TickIntervalMs = 15;              // 选区轮询：场景键未变零写入，变化 ≤15ms 内跟随
        private const int BusyIntervalMs = 200;             // 忙态退避
        private const string SentinelName = "JYYJ_VIEW_MODE";
        private const string SentinelRefersTo = "=1";
        private const string SentinelToken = "JYYJ_VIEW_MODE";
        private const string SelName = "JYYJ_SEL";          // 旧版选区判定表达式名（清理用）
        private const string InitialSelRefersTo = "=0";     // 旧版初始值（所有权判定用）
        private const string CellRowName = "JYYJ_R";        // 旧版活动格易失名（清理用）
        private const string CellColName = "JYYJ_C";
        private const string CellRowRefersTo = "=CELL(\"row\")";
        private const string CellColRefersTo = "=CELL(\"col\")";
        private const string ColShapeName = "JYYJ_VIEW_COL";   // 池化后缀 1..MaxAreas
        private const string RowShapeName = "JYYJ_VIEW_ROW";
        private const string BoxShapeName = "JYYJ_VIEW_BOX";
        private const int MaxAreas = 8;                        // 多选区域形状池上限（超出部分不高亮）
        // 旧版（边界定义名方案）遗留的 4 个数值常量名，清理时兼容移除
        private static readonly string[] LegacyBoundNames =
            { "JYYJ_SEL_R1", "JYYJ_SEL_R2", "JYYJ_SEL_C1", "JYYJ_SEL_C2" };

        // 忙态 HRESULT：VBA_E_IGNORE（重入）、RPC 调用被拒/需重试（UIA 自动化期间常见）
        private const int VbaIgnore = unchecked((int)0x800AC472);
        private const int RpcCallRejected = unchecked((int)0x80010001);
        private const int RpcServerCallRetryLater = unchecked((int)0x8001010A);

        private readonly dynamic _app;
        private readonly Timer _timer;
        private SpotlightMode _mode = SpotlightMode.Off;
        private string _sceneKey;                           // 上次场景键，未变则整轮零写入
        private string _lastSheetName = "";                 // 上次绘制形状所在表（切表时清理旧表形状）
        private string _lastError;
        private bool _loggedFirstScene;

        public SpotlightService(dynamic excelApplication)
        {
            _app = excelApplication;
            _timer = new Timer { Interval = TickIntervalMs };
            _timer.Tick += OnTick;
        }

        public SpotlightMode Mode => _mode;

        /// <summary>启动/关闭自愈入口：清理旧版 CF 遗留与本方案形状残留。</summary>
        public static void CleanupAllWorkbooks(dynamic excelApplication)
        {
            if (excelApplication == null) return;
            CleanupAllOpenWorkbooks(excelApplication);
        }

        public void SetMode(SpotlightMode mode)
        {
            DiagnosticsLog.Write($"SpotlightService SetMode mode={mode} (was {_mode})");
            if (_mode == mode) return;
            _timer.Stop();
            _mode = mode;
            SaveViewMode(mode);
            if (mode == SpotlightMode.Off)
            {
                // 关闭：清掉本方案形状 + 旧版 CF 遗留
                CleanupAllOpenWorkbooks(_app);
                _sceneKey = null;
            }
            else
            {
                CleanupLegacyOnly(_app);
                _sceneKey = null;
                _timer.Start();
                RenderTick();
            }
        }

        public void Dispose()
        {
            _timer.Stop();
            _timer.Tick -= OnTick;
            try { CleanupAllOpenWorkbooks(_app); }
            catch (Exception exception) { DiagnosticsLog.Write($"SpotlightService dispose cleanup: {exception.GetType().Name}"); }
            _mode = SpotlightMode.Off;
        }

        private void OnTick(object sender, EventArgs e)
        {
            if (_mode == SpotlightMode.Off) return;
            // Excel 忙态（外部自动化/重入）时短退避重试，空闲恢复 15ms
            _timer.Interval = RenderTick() ? TickIntervalMs : BusyIntervalMs;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);

        /// <summary>只读轮询一轮：场景键（表|选区|可见区|窗口矩形）未变直接跳过；
        /// 变化则把形状移动到选区各区域的行列十字/包围盒。写入仅为形状几何，无重算。</summary>
        private bool RenderTick()
        {
            dynamic win = null;
            dynamic sel = null;
            dynamic activeCell = null;
            dynamic visibleRange = null;
            dynamic sheet = null;
            dynamic shapes = null;
            try
            {
                win = _app.ActiveWindow;
                if (win == null) return HideOverlay("no active window");

                IntPtr hwnd = (IntPtr)(long)Convert.ToInt64(win.Hwnd);
                if (!GetWindowRect(hwnd, out NativeRect wr) || wr.Right <= wr.Left || wr.Bottom <= wr.Top
                    || wr.Left < -20000 || wr.Top < -20000) // 最小化窗口坐标 ~-32000
                {
                    return HideOverlay("window rect unavailable");
                }

                sheet = _app.ActiveSheet;
                if (sheet == null) return HideOverlay("no active sheet");
                visibleRange = win.VisibleRange;
                activeCell = _app.ActiveCell;
                if (activeCell == null) return HideOverlay("no active cell");

                // 防拖动：用户点到高亮形状时 Excel 会选中它（可被拖走），立即退选回活动格；
                // 并复位场景键强制重绘 —— 拖放后形状弹回原位（拖放瞬间选区=形状，场景键未变会漏重绘）
                sel = _app.Selection;
                string selAddress = "";
                if (sel != null)
                {
                    // 先探测 Address（Range 常态，零异常）；失败才检查是否为本方案形状被选中
                    bool selIsRange = true;
                    try { selAddress = Convert.ToString(sel.Address); }
                    catch { selIsRange = false; }

                    if (!selIsRange && IsOwnShapeSelected(sel))
                    {
                        activeCell.Select();
                        Marshal.ReleaseComObject(sel);
                        sel = _app.Selection; // 退选后重读（此时应为 Range）
                        _sceneKey = null;     // 强制重绘：拖放后形状弹回原位
                        try { selAddress = Convert.ToString(sel.Address); } catch { selAddress = ""; }
                    }
                }

                string sheetName = Convert.ToString(sheet.Name);
                string visibleAddress = Convert.ToString(visibleRange.Address);
                string windowRect = wr.Left + "," + wr.Top + "," + wr.Right + "," + wr.Bottom;
                string key = sheetName + "|" + _mode + "|" + selAddress + "|" + visibleAddress + "|" + windowRect;
                if (key == _sceneKey) return true; // 快速路径：场景未变零 COM 写
                bool sheetChanged = _sceneKey == null || !key.StartsWith(sheetName + "|", StringComparison.Ordinal);
                _sceneKey = key;

                // 指针模式跟随整个选区包围盒；聚光灯用选区各区域的行列十字
                shapes = sheet.Shapes;
                if (sheetChanged && !string.IsNullOrEmpty(_lastSheetName))
                {
                    // 切表：清掉旧表上残留的高亮形状（避免切回时看到过期位置）。
                    // 首场景 _lastSheetName 为空，跳过（Sheets.Item("") 会抛 BADINDEX）。
                    dynamic oldSheet = null;
                    try
                    {
                        oldSheet = _app.Sheets.Item(_lastSheetName);
                        if (oldSheet != null) DeleteViewShapes(oldSheet);
                    }
                    catch
                    {
                        // 旧表已删/改名：忽略（预期清理路径，不打日志）。
                    }
                    finally
                    {
                        if (oldSheet != null)
                        {
                            try { Marshal.ReleaseComObject(oldSheet); } catch { }
                        }
                    }

                    _lastSheetName = sheetName;
                }

                if (sel == null) return HideOverlay("no selection");

                dynamic box = null;
                dynamic areas = null;
                try
                {
                    if (_mode == SpotlightMode.Pointer)
                    {
                        box = EnsureShape(shapes, BoxShapeName, red: true);
                        box.Left = sel.Left;
                        box.Top = sel.Top;
                        box.Width = sel.Width;
                        box.Height = sel.Height;
                        DeleteAllBandShapes(shapes);
                        return true;
                    }

                    areas = sel.Areas;
                    int areaCount = Convert.ToInt32(areas.Count);
                    if (areaCount <= 0) return HideOverlay("no areas");
                    bool capped = areaCount > MaxAreas;
                    if (capped) areaCount = MaxAreas; // 超出上限的多选区域降级（只画前 N 块）

                    double visLeft = visibleRange.Left;
                    double visTop = visibleRange.Top;
                    double visRight = visLeft + visibleRange.Width;
                    double visBottom = visTop + visibleRange.Height;

                    // 每块区域各出「行∪列」十字带（与旧版条件格式语义一致）
                    for (int i = 1; i <= areaCount; i++)
                    {
                        dynamic area = null;
                        dynamic col = null;
                        dynamic row = null;
                        try
                        {
                            area = areas[i];
                            col = EnsureShape(shapes, ColShapeName + i, red: false);
                            col.Left = area.Left;
                            col.Top = 0;
                            col.Width = area.Width;
                            col.Height = visBottom;

                            row = EnsureShape(shapes, RowShapeName + i, red: false);
                            row.Left = 0;
                            row.Top = area.Top;
                            row.Width = visRight;
                            row.Height = area.Height;
                        }
                        finally
                        {
                            foreach (dynamic rcw in new[] { area, col, row })
                            {
                                if (rcw != null)
                                {
                                    try { Marshal.ReleaseComObject(rcw); } catch { }
                                }
                            }
                        }
                    }

                    // 回收未使用的池形状
                    for (int i = areaCount + 1; i <= MaxAreas; i++)
                    {
                        DeleteShape(shapes, ColShapeName + i);
                        DeleteShape(shapes, RowShapeName + i);
                    }

                    DeleteShape(shapes, BoxShapeName);
                    return true;
                }
                finally
                {
                    foreach (dynamic comObject in new[] { box, areas })
                    {
                        if (comObject != null)
                        {
                            try { Marshal.ReleaseComObject(comObject); } catch { }
                        }
                    }
                }
            }
            catch (COMException exception) when (IsBusy(exception))
            {
                _sceneKey = null; // 忙态写入未提交：复位以便下 tick 重试
                return false;
            }
            catch (Exception exception)
            {
                // 单元格编辑模式/保护表/图表页等写入被拒：复位场景键以便恢复后重试；
                // 返回 false 走 200ms 退避（保护表等永久失败场景避免 15ms 异常空转）；同类错误只记一次
                _sceneKey = null;
                string message = $"{exception.GetType().Name} {exception.Message}";
                if (!_loggedFirstScene || message != _lastError)
                {
                    _lastError = message;
                    _loggedFirstScene = true;
                    DiagnosticsLog.Write($"SpotlightService render skipped: {message}");
                }

                return false;
            }
            finally
            {
                foreach (dynamic comObject in new[] { sel, activeCell, visibleRange, sheet, shapes, win })
                {
                    if (comObject != null)
                    {
                        try { Marshal.ReleaseComObject(comObject); } catch { }
                    }
                }
            }
        }

        private bool HideOverlay(string reason)
        {
            if (_sceneKey != null)
            {
                DiagnosticsLog.Write($"SpotlightService hide: {reason}");
                _sceneKey = null;
            }

            return true;
        }

        /// <summary>按名取本方案形状；不存在则创建（半透明填充无边框、不打印、不随单元格变动）。</summary>
        private static dynamic EnsureShape(dynamic shapes, string name, bool red)
        {
            dynamic existing = null;
            try
            {
                existing = shapes.Item(name);
                return existing;
            }
            catch
            {
                // 名字不存在：走创建（首次使用的主路径，非异常，不打日志）。
            }

            dynamic shape = shapes.AddShape(MsoShapeRectangle, 0, 0, 10, 10);
            shape.Name = name;
            if (red)
            {
                // 指针：只画红框不填充（旧版语义），不遮挡单元格内容
                shape.Fill.Visible = MsoFalse;
                shape.Line.Visible = MsoTrue;
                shape.Line.ForeColor.RGB = RedRgb;
                shape.Line.Weight = 1.25;
            }
            else
            {
                // 聚光灯：半透明黄色填充带
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = YellowRgb;
                shape.Fill.Transparency = FillTransparency;
                shape.Line.Visible = MsoFalse;
            }

            shape.Placement = XlFreeFloating;
            // Locked/PrintObject 在部分版本 dynamic 绑定失败，反射兜底；失败仅影响打印/受保护表场景，不阻塞
            TrySetComProperty(shape, "Locked", true);
            TrySetComProperty(shape, "PrintObject", false);
            DiagnosticsLog.Write($"SpotlightService shape created name={name} red={red}");
            return shape;
        }

        /// <summary>反射设置 COM 属性（PrintObject/Locked 等部分版本 dynamic 绑定不可用）。</summary>
        private static void TrySetComProperty(object target, string propertyName, object value)
        {
            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.SetProperty,
                    null,
                    target,
                    new[] { value });
            }
            catch { }
        }

        /// <summary>当前 Selection 是否为本方案的高亮形状（用户点到形状会被 Excel 选中并可拖走，需立即退选）。</summary>
        private static bool IsOwnShapeSelected(dynamic sel)
        {
            try
            {
                string name = Convert.ToString(sel.Name);
                return name != null && name.StartsWith("JYYJ_VIEW", StringComparison.Ordinal);
            }
            catch (Exception)
            {
                return false; // Range 等：Range.Name 未命名时抛异常（常态探测，不打日志）。
            }
        }

        /// <summary>删除全部行列十字带形状（指针模式切换时）。</summary>
        private static void DeleteAllBandShapes(dynamic shapes)
        {
            for (int i = 1; i <= MaxAreas; i++)
            {
                DeleteShape(shapes, ColShapeName + i);
                DeleteShape(shapes, RowShapeName + i);
            }
        }

        private static void DeleteShape(dynamic shapes, string name)
        {
            dynamic shape = null;
            try
            {
                shape = shapes.Item(name);
                shape.Delete();
            }
            catch (Exception xaiEx3)
            {
                // 「不存在即已删」为契约内幂等清理（FR-K 审查降噪：预期未命中不打日志）；
                // 仅非「未找到」类失败（如保护态删不掉）才告警。
                if (xaiEx3.Message.IndexOf("未找到", StringComparison.Ordinal) < 0)
                {
                    AppLog.Warn("spotlight", "SpotlightService.cs:DeleteShape 删除高亮形状失败（非预期）: " + xaiEx3.Message);
                }
            }
            finally
            {
                if (shape != null)
                {
                    try { Marshal.ReleaseComObject(shape); } catch { }
                }
            }
        }

        /// <summary>只清旧版 CF 遗留，不动形状（开启模式时形状归本轮管理）。</summary>
        private static void CleanupLegacyOnly(dynamic app)
        {
            CleanupAllOpenWorkbooks(app, includeShapes: false);
        }

        /// <summary>扫全部打开工作簿：删除带哨兵标记的旧版 CF 规则、JYYJ_* 定义名，
        /// 以及本方案的 JYYJ_VIEW_* 形状。</summary>
        private static void CleanupAllOpenWorkbooks(dynamic app, bool includeShapes = true)
        {
            try
            {
                dynamic workbooks = app.Workbooks;
                int count = Convert.ToInt32(workbooks.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = null;
                    try
                    {
                        wb = workbooks[i];
                        CleanupWorkbook(wb, includeShapes);
                    }
                    catch (Exception exception)
                    {
                        DiagnosticsLog.Write($"SpotlightService cleanup workbook {i}: {exception.GetType().Name}");
                    }
                    finally
                    {
                        if (wb != null) Marshal.ReleaseComObject(wb);
                    }
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"SpotlightService CleanupAllOpenWorkbooks: {exception.GetType().Name}");
            }
        }

        private static void CleanupWorkbook(dynamic wb, bool includeShapes)
        {
            // 旧版 CF 遗留以哨兵定义名为准；本方案形状以固定命名清理
            bool hadSentinel = TryRemoveSentinelName(wb);
            if (hadSentinel) RemoveBoundNames(wb);
            if (!hadSentinel && !includeShapes) return;

            dynamic worksheets = wb.Worksheets;
            int sheetCount = Convert.ToInt32(worksheets.Count);
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic ws = null;
                try
                {
                    ws = worksheets[i];
                    if (hadSentinel) RemoveMarkedRules(ws);
                    if (includeShapes) DeleteViewShapes(ws);
                }
                catch (Exception exception)
                {
                    DiagnosticsLog.Write($"SpotlightService cleanup sheet {i}: {exception.GetType().Name}");
                }
                finally
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                }
            }
        }

        /// <summary>删除本方案的视图形状（按固定命名池；保护表等不可写场景静默跳过）。</summary>
        private static void DeleteViewShapes(dynamic ws)
        {
            dynamic shapes = null;
            try
            {
                shapes = ws.Shapes;
                DeleteAllBandShapes(shapes);
                // 兼容：清理旧版无池化后缀的单形状命名
                DeleteShape(shapes, ColShapeName);
                DeleteShape(shapes, RowShapeName);
                DeleteShape(shapes, BoxShapeName);
            }
            catch { }
            finally
            {
                if (shapes != null)
                {
                    try { Marshal.ReleaseComObject(shapes); } catch { }
                }
            }
        }

        private const string ViewModeRegPath = @"Software\JYYJ\ExcelAiAssistant";
        private const string ViewModeRegValue = "ViewMode";

        /// <summary>持久化视图模式：重启 Excel 后由 OnStartupComplete 恢复，无需再次点按钮。</summary>
        private static void SaveViewMode(SpotlightMode mode)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(ViewModeRegPath))
                {
                    key.SetValue(ViewModeRegValue, (int)mode, Microsoft.Win32.RegistryValueKind.DWord);
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"SpotlightService SaveViewMode: {exception.GetType().Name}");
            }
        }

        /// <summary>读取上次保存的视图模式；未保存过或读失败返回 Off（不自动开启）。</summary>
        public static SpotlightMode GetSavedViewMode()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(ViewModeRegPath))
                {
                    if (key == null) return SpotlightMode.Off;
                    if (key.GetValue(ViewModeRegValue) is int saved
                        && saved >= (int)SpotlightMode.Off
                        && saved <= (int)SpotlightMode.Pointer)
                    {
                        return (SpotlightMode)saved;
                    }
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"SpotlightService GetSavedViewMode: {exception.GetType().Name}");
            }

            return SpotlightMode.Off;
        }

        /// <summary>
        /// 删除带哨兵标记的表达式 CF 规则（逆序遍历，边删边取安全）。
        /// 非表达式规则（色阶/数据条/图标集）无 Formula1 属性，先按 Type 过滤避免 dynamic 绑定异常。
        /// 忙态退避重试；非忙态异常直接放弃，避免死循环。
        /// </summary>
        private static void RemoveMarkedRules(dynamic sheet)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                dynamic fcs = null;
                try
                {
                    fcs = sheet.Cells.FormatConditions;
                    int count = Convert.ToInt32(fcs.Count);
                    for (int i = count; i >= 1; i--)
                    {
                        dynamic fc = null;
                        try
                        {
                            fc = fcs[i];
                            int type = Convert.ToInt32(fc.Type);
                            if (type != XlExpression) continue;
                            string formula1 = "";
                            try { formula1 = Convert.ToString(fc.Formula1); } catch { continue; }
                            if (formula1.IndexOf(SentinelToken, StringComparison.Ordinal) >= 0)
                            {
                                try
                                {
                                    fc.Delete();
                                    DiagnosticsLog.Write("SpotlightService legacy rule deleted");
                                }
                                catch (Exception de) { DiagnosticsLog.Write($"SpotlightService rule delete failed: {de.GetType().Name} {de.Message}"); }
                            }
                        }
                        finally
                        {
                            if (fc != null)
                            {
                                try { Marshal.ReleaseComObject(fc); } catch { }
                            }
                        }
                    }

                    return;
                }
                catch (COMException exception) when (IsBusy(exception))
                {
                    System.Threading.Thread.Sleep(150);
                }
                catch (Exception)
                {
                    // 清理期预期失败（名称不存在等）：保持原语义停止继续清理，不打日志。
                    return;
                }
                finally
                {
                    if (fcs != null)
                    {
                        try { Marshal.ReleaseComObject(fcs); } catch { }
                    }
                }
            }
        }

        /// <summary>移除旧版选区/易失名/边界定义名（仅本加载项创建的，同名用户定义保留）。</summary>
        private static void RemoveBoundNames(dynamic wb)
        {
            foreach (string name in LegacyBoundNames)
            {
                string refersTo = GetNameRefersTo(wb, name);
                if (refersTo == null) continue;
                if (!IsNumericConstant(refersTo)) continue;
                DeleteNameByName(wb, name);
            }

            string selRefersTo = GetNameRefersTo(wb, SelName);
            if (selRefersTo != null && IsOwnSelectionName(selRefersTo)) DeleteNameByName(wb, SelName);

            string rowRefersTo = GetNameRefersTo(wb, CellRowName);
            if (rowRefersTo != null && IsOwnCellName(rowRefersTo, CellRowRefersTo)) DeleteNameByName(wb, CellRowName);
            string colRefersTo = GetNameRefersTo(wb, CellColName);
            if (colRefersTo != null && IsOwnCellName(colRefersTo, CellColRefersTo)) DeleteNameByName(wb, CellColName);
        }

        private static bool IsNumericConstant(string refersTo)
        {
            if (string.IsNullOrWhiteSpace(refersTo)) return false;
            string t = refersTo.Trim();
            if (!t.StartsWith("=", StringComparison.Ordinal)) return false;
            return int.TryParse(t.Substring(1).Trim(), out _);
        }

        /// <summary>判定 RefersTo 是否为本加载项创建的 JYYJ_SEL（数值 0、旧版 =FALSE 或判定表达式 =OR(...)）。</summary>
        private static bool IsOwnSelectionName(string refersTo)
        {
            if (string.IsNullOrWhiteSpace(refersTo)) return false;
            string t = refersTo.Trim();
            if (string.Equals(t, "=FALSE", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(t, InitialSelRefersTo, StringComparison.OrdinalIgnoreCase)) return true;
            return t.StartsWith("=OR(", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判定 RefersTo 是否为本加载项创建的易失名（与期望值完全一致）。</summary>
        private static bool IsOwnCellName(string refersTo, string ownRefersTo)
        {
            if (string.IsNullOrWhiteSpace(refersTo)) return false;
            return string.Equals(refersTo.Trim(), ownRefersTo, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>按名直取并删除定义名（O(1)）；不存在则视为已删。Names.Item 名称匹配不区分大小写。</summary>
        private static void DeleteNameByName(dynamic wb, string nameToDelete)
        {
            dynamic name = null;
            try
            {
                name = wb.Names.Item(nameToDelete);
                name.Delete();
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"SpotlightService DeleteNameByName {nameToDelete}: {exception.GetType().Name}");
            }
            finally
            {
                if (name != null)
                {
                    try { Marshal.ReleaseComObject(name); } catch { }
                }
            }
        }

        private static bool IsBusy(Exception exception)
        {
            if (!(exception is COMException com)) return false;
            int code = com.ErrorCode;
            return code == VbaIgnore || code == RpcCallRejected || code == RpcServerCallRetryLater;
        }

        /// <summary>尝试移除哨兵定义名。返回 false 表示不存在（无需清理）或为同名用户定义（保留）。</summary>
        private static bool TryRemoveSentinelName(dynamic wb)
        {
            string refersTo = GetNameRefersTo(wb, SentinelName);
            if (refersTo == null) return false;
            if (!string.Equals(refersTo.Trim(), SentinelRefersTo, StringComparison.OrdinalIgnoreCase))
            {
                DiagnosticsLog.Write($"SpotlightService TryRemoveSentinelName: 保留同名用户定义 {SentinelName}={refersTo}");
                return false;
            }

            DeleteNameByName(wb, SentinelName);
            return true;
        }

        /// <summary>返回指定定义名的 RefersTo；不存在返回 null。按名直查 O(1)（Names.Item 匹配不区分大小写，
        /// 名称不存在时抛 COM 异常视为缺失）。</summary>
        private static string GetNameRefersTo(dynamic wb, string nameToFind)
        {
            dynamic name = null;
            try
            {
                name = wb.Names.Item(nameToFind);
                return Convert.ToString(name.RefersTo);
            }
            catch (Exception)
            {
                // 名称不存在/引用失效 = 聚光灯未在此表初始化的预期探测结果，静默返回 null（FR-K 审查降噪）。
                return null;
            }
            finally
            {
                if (name != null)
                {
                    try { Marshal.ReleaseComObject(name); } catch { }
                }
            }
        }
    }
}
