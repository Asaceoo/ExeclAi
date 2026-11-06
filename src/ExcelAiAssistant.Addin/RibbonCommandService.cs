using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Excel;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 「JYYJ快捷」功能区统一入口：选区解析 → 参数框 → 安全分级 → 二次确认 → 调既有工具 → 历史/日志 → 摘要。
    /// 持有一个进程内 ExcelToolHost（复用侧边栏同款会话历史，侧边栏未开则自建同路径仓库兜底），
    /// 全部在 Ribbon onAction 的 UI 线程同步执行（WebView2 参数框依赖 UI 线程，绝不 Task.Run）。仅调既有已注册工具，不改动其执行体。
    /// </summary>
    public sealed class RibbonCommandService
    {
        private readonly dynamic application;
        private readonly string sessionId = Guid.NewGuid().ToString("N");
        private ExcelToolHost cachedHost;
        private bool executing;

        private enum SelectionStatus { Ok, Empty, NonRect, ReadError }

        public RibbonCommandService(dynamic excelApplication)
        {
            application = excelApplication ?? throw new ArgumentNullException(nameof(excelApplication));
        }

        /// <summary>
        /// 超级看板专用入口（FR-SB-10）：当前选区预填 → 专用配置表单（维度映射/自定义指标/透视字段/切片器）
        /// → 构建 render_super_dashboard 实参 → 执行 → 结果人话提示。区别于通用 RunQuickCommand 的字段化参数框。
        /// </summary>
        public void ExecuteSuperBoard()
        {
            if (executing)
            {
                DiagnosticsLog.Write("JYYJ快捷 busy ignore tool=render_super_dashboard");
                return;
            }
            executing = true;
            try
            {
                TryResolveSelection(out var sel, out _, out _, out _, out var status);
                if (status == SelectionStatus.NonRect)
                {
                    MessageBox.Show("请选择矩形数据区域（含表头行）。", "超级看板", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var sheetName = sel?.SheetName ?? "";
                var address = sel?.Address ?? "";
                var columns = sel?.ColumnNames ?? new List<string>();

                // v3.7.1.49 宿主拦截（Q7 拍板：超级看板仅 Excel 宿主）：加载项运行在宿主进程内，
                // 进程名判定在此可靠（ExcelToolHost.CreateRegistry 的同名判定在 out-of-proc 巡检 harness 内失效）
                var hostProcess = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
                if (hostProcess.Equals("et", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("超级看板仅 Microsoft Excel 宿主支持（WPS 下不可用）。", "超级看板", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 表单专用 registry：dashboard 实例仅在此路径注册（不进宿主 CreateRegistry，守 D11/WPS 基线）
                var dashboard = new ExcelDashboardService((object)application);
                var pivotChartSb = new ExcelPivotChartService((object)application, null, null);
                var registry = new ToolRegistry();
                Func<Dictionary<string, object>, ToolResult> stubSb = _ =>
                    new ToolResult { Ok = false, ErrorCode = "NOT_UNDER_TEST", Data = "skip" };
                ExcelToolHostRegistration.RegisterM3Tools(registry, stubSb, stubSb, null, pivotChartSb, dashboard);

                Func<string, Dictionary<string, object>, ToolResult> executeTool = (tool, toolArgs) =>
                {
                    if (!registry.TryGet(tool, out var descriptor))
                    {
                        return new ToolResult { Ok = false, ErrorCode = "TOOL_NOT_REGISTERED", Data = "底层工具未注册：" + tool + "。请修复安装或联系支持。" };
                    }
                    Cursor.Current = Cursors.WaitCursor;   // 审查 T4：大表渲染期间给出忙碌反馈
                    var r = descriptor.Execute(toolArgs);
                    Cursor.Current = Cursors.Default;
                    return r ?? new ToolResult { Ok = false, ErrorCode = "NULL_RESULT", Data = "工具无返回结果。" };
                };

                using (var wizard = new SuperBoardWizardForm(sheetName, address, columns, executeTool))
                {
                    wizard.ShowDialog();
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write("ExecuteSuperBoard failed: " + exception);
                MessageBox.Show("执行失败：" + exception.Message, "超级看板", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                executing = false;
            }
        }

        public void Execute(string toolName)
        {            // 防重入：执行期（含参数框/COM 写）忽略新请求，避免连点导致同一工具串行重复执行。
            if (executing)
            {
                DiagnosticsLog.Write($"JYYJ快捷 busy ignore tool={toolName}");
                return;
            }

            executing = true;
            try
            {
                DiagnosticsLog.Write($"JYYJ快捷 execute tool={toolName}");
                if (!TryResolveSelection(out var sel, out var matrix, out var topRow, out var startCol, out var status))
                {
                    switch (status)
                    {
                        case SelectionStatus.Empty:
                            MessageBox.Show("请先选中数据区域。", toolName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        case SelectionStatus.NonRect:
                            MessageBox.Show("请选择矩形区域。", toolName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                    }
                    return;
                }

                var fields = RibbonCommandPlanner.BuildFields(toolName, sel);
                var safety = RibbonCommandPlanner.Classify(toolName);
                var overlayHint = RibbonCommandPlanner.OverlayHint(safety, sel);
                var dangerHint = safety == CommandSafety.DeleteRows
                    ? RibbonCommandPlanner.DeleteWarning(toolName == "delete_blank_rows_range"
                        ? RibbonCommandPlanner.CountRemovedBlankRows(matrix, false)
                        : RibbonCommandPlanner.CountRemovedRows(matrix, 1, new int[0], false, false))
                    : null;
                var addressTitle = $"{QuoteIfNeeded(sel.SheetName)}!{sel.Address}（{sel.Rows} 行 × {sel.Cols} 列）";

                using (var dialog = new RibbonParameterDialog(toolName, addressTitle, fields, overlayHint, dangerHint))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        DiagnosticsLog.Write($"JYYJ快捷 cancelled tool={toolName}");
                        return; // 取消：零执行零变更
                    }

                    var buildResult = RibbonCommandPlanner.BuildArgs(toolName, sel, dialog.Values);
                    if (buildResult is string error)
                    {
                        MessageBox.Show(error, toolName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    ExecuteTool(toolName, sel, matrix, topRow, startCol, safety, (Dictionary<string, object>)buildResult);
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"JYYJ快捷 execute failed tool={toolName}: {exception}");
                MessageBox.Show("执行失败：" + exception.Message, toolName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                executing = false;
            }
        }

        private void ExecuteTool(
            string toolName,
            ParamSelection sel,
            object[,] matrix,
            int topRow,
            int startCol,
            CommandSafety safety,
            Dictionary<string, object> args)
        {
            // 目标带数据守卫（统一标准）：text_split 拆分带区 / text_merge 结果列 / transpose 指定位置，
            // 目标可静态预算即写前检查占用，避免静默覆盖用户已有数据。
            if (IsTargetOccupied(toolName, sel, matrix, topRow, startCol, args, out var guardHint))
            {
                MessageBox.Show(guardHint, toolName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 删行类结构操作：参数已确定，二次红色警示 + 明确再确认；取消则零执行。预算提前算出，供确认与执行后自检共用。
            var removedBudget = -1;
            if (RibbonCommandPlanner.RequiresSecondaryConfirm(safety))
            {
                removedBudget = toolName == "delete_blank_rows_range"
                    ? RibbonCommandPlanner.CountRemovedBlankRows(matrix, IsTrueArg(args, "only_fully_blank"))
                    : CountRemovedRows(matrix, args);
                var msg = RibbonCommandPlanner.DeleteWarning(Math.Max(0, removedBudget));
                if (MessageBox.Show(msg, "确认删除", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK)
                {
                    DiagnosticsLog.Write($"JYYJ快捷 {toolName} 二次确认取消 removed={removedBudget}");
                    return;
                }
            }

            var host = GetHost();
            var registry = host.CreateRegistry();
            if (!registry.TryGet(toolName, out var descriptor))
            {
                MessageBox.Show($"底层工具未注册：{toolName}", toolName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = descriptor.Execute(args);
            if (result == null)
            {
                MessageBox.Show("工具无返回结果。", toolName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DiagnosticsLog.Write($"JYYJ快捷 {toolName} ok={result.Ok} code={result.ErrorCode}");
            // 参数级审计日志：选区 + 完整实参落盘，兑现「运行日志全量留痕」承诺（设置页可查 AI/功能区操作）。
            DiagnosticsLog.Write($"JYYJ快捷 {toolName} sel={QuoteIfNeeded(sel.SheetName)}!{sel.Address} args={JsonConvert.SerializeObject(args)}");
            var summary = RibbonCommandPlanner.BuildSummary(toolName, result, sel);
            var verifyFail = PostVerify(toolName, sel, matrix, topRow, startCol, args, removedBudget, result, out var sampleText);
            if (verifyFail != null)
            {
                MessageBox.Show($"执行已完成，但自检未通过：\n{verifyFail}\n\n摘要：\n{summary}",
                    toolName + " · 自检告警", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!string.IsNullOrEmpty(sampleText))
            {
                summary += "\n" + sampleText;
            }
            MessageBox.Show(summary, toolName + " · 完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ================= 执行后自检（fail-closed：回读关键锚点，不符则显著告警） =================

        private string PostVerify(string toolName, ParamSelection sel, object[,] matrix, int topRow, int startCol,
            Dictionary<string, object> args, int removedBudget, ToolResult result, out string sampleText)
        {
            sampleText = null;
            if (!RibbonCommandPlanner.NeedsPostVerify(toolName))
            {
                return null;
            }
            try
            {
                switch (toolName)
                {
                    case "transpose_range":
                        return VerifyTranspose(sel, matrix, topRow, startCol, args);
                    case "dedup_range":
                    case "delete_blank_rows_range":
                        return VerifyDeletion(removedBudget, result);
                    case "fill_random_numbers":
                        return VerifyRandom(sel, topRow, startCol, args, out sampleText);
                    default:
                        return null;
                }
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"JYYJ快捷 {toolName} 自检异常（不阻断）: {exception.GetType().Name}: {exception.Message}");
                return null; // 自检是保护性检查：读取异常不阻断正常流程
            }
        }

        private string VerifyTranspose(ParamSelection sel, object[,] matrix, int topRow, int startCol, Dictionary<string, object> args)
        {
            // 目标区：指定位置→解析地址（支持跨表）；原位→源区位置。两者都可回读验证转置维度与锚点。
            var rows = sel.Cols;
            var cols = sel.Rows;
            string sheetName;
            int top, left;
            if (args.TryGetValue("target_address", out var addr)
                && RibbonCommandPlanner.TryParseAddress(Convert.ToString(addr), out var targetRow, out var targetCol))
            {
                sheetName = args.TryGetValue("target_sheet", out var ts) ? Convert.ToString(ts) : sel.SheetName;
                top = targetRow;
                left = targetCol;
            }
            else
            {
                sheetName = sel.SheetName;
                top = topRow;
                left = startCol;
            }

            var target = ReadBlockValues(sheetName, top, left, rows, cols);
            return RibbonCommandPlanner.VerifyTransposeResult(matrix, target);
        }

        private static string VerifyDeletion(int removedBudget, ToolResult result)
        {
            if (removedBudget < 0)
            {
                return null;
            }
            if (!(result.Data is IDictionary<string, object> dict) || dict.Count == 0)
            {
                return null;
            }
            var key = dict.ContainsKey("deduped_rows_removed") ? "deduped_rows_removed" : "deleted_rows";
            if (dict.TryGetValue(key, out var raw))
            {
                var removed = Convert.ToInt32(raw);
                if (removed != removedBudget)
                {
                    return $"删除行数 {removed} 与预估 {removedBudget} 不一致";
                }
            }
            return null;
        }

        private string VerifyRandom(ParamSelection sel, int topRow, int startCol, Dictionary<string, object> args, out string sampleText)
        {
            sampleText = null;
            if (!args.TryGetValue("min", out var minRaw) || !args.TryGetValue("max", out var maxRaw)
                || !args.TryGetValue("decimals", out var decRaw))
            {
                return null;
            }
            var min = Convert.ToDouble(minRaw, CultureInfo.InvariantCulture);
            var max = Convert.ToDouble(maxRaw, CultureInfo.InvariantCulture);
            var decimals = Convert.ToInt32(decRaw);
            var values = ReadBlockValues(sel.SheetName, topRow, startCol, sel.Rows, sel.Cols);
            var fail = RibbonCommandPlanner.VerifyRandomRange(values, min, max, decimals);
            if (fail != null)
            {
                return fail;
            }
            // 随机类验收锚点：附 1-3 个实际样本，用户可核对「这次生成的值」。
            var samples = RibbonCommandPlanner.PickSampleValues(values, 3);
            if (samples.Count > 0)
            {
                sampleText = "样本：" + string.Join("，", samples);
            }
            return null;
        }

        private object[,] ReadBlockValues(string sheetName, int top, int left, int rows, int cols)
        {
            dynamic sheet = null;
            dynamic block = null;
            try
            {
                sheet = ResolveSheet(sheetName);
                block = BuildSheetRange(sheet, top, left, rows, cols);
                return ReadMatrix(block, rows, cols);
            }
            finally
            {
                SafeRelease((object)block);
                SafeRelease((object)sheet);
            }
        }

        // ================= 选区解析 =================

        private bool TryResolveSelection(out ParamSelection sel, out object[,] matrix, out int topRow, out int startCol, out SelectionStatus status)
        {
            sel = null;
            matrix = null;
            topRow = 0;
            startCol = 0;
            status = SelectionStatus.ReadError;
            dynamic selection = null;
            dynamic range = null;
            try
            {
                selection = application.Selection;
                if (selection == null)
                {
                    status = SelectionStatus.Empty;
                    return false;
                }

                var areas = 0;
                try { areas = Convert.ToInt32(selection.Areas.Count); }
                catch { areas = 0; }
                var selectionAddr = Convert.ToString(selection.Address);
                if (areas != 1 || string.IsNullOrWhiteSpace(selectionAddr) || selectionAddr.IndexOf(',') >= 0)
                {
                    status = SelectionStatus.NonRect;
                    return false;
                }

                var rows = Convert.ToInt32(selection.Rows.Count);
                var cols = Convert.ToInt32(selection.Columns.Count);
                if (rows <= 0 || cols <= 0)
                {
                    status = SelectionStatus.Empty;
                    return false;
                }

                range = selection; // Areas==1，selection 即单一 Range
                matrix = ReadMatrix(range, rows, cols);
                topRow = Convert.ToInt32(selection.Row);
                startCol = Convert.ToInt32(selection.Column);
                var sheetName = Convert.ToString(selection.Worksheet.Name);

                sel = new ParamSelection { SheetName = sheetName, Address = selectionAddr, Rows = rows, Cols = cols };
                for (var c = 0; c < cols; c++)
                {
                    var v = matrix[0, c];
                    var name = v == null || v is DBNull ? string.Empty : Convert.ToString(v);
                    sel.ColumnNames.Add(string.IsNullOrWhiteSpace(name) ? $"列{c + 1}" : name);
                }

                status = SelectionStatus.Ok;
                return true;
            }
            catch (Exception exception)
            {
                DiagnosticsLog.Write($"TryResolveSelection failed: {exception.GetType().Name}: {exception.Message}");
                return false;
            }
            finally
            {
                SafeRelease((object)range);
                SafeRelease((object)selection);
            }
        }

        private static object[,] ReadMatrix(dynamic range, int rows, int cols)
        {
            var result = new object[rows, cols];
            var obj = range.Value2;
            if (obj is object[,] m)
            {
                // Excel COM 多格 Value2 返回 1 基数组（GetLowerBound=1），必须按真实下界读取；
                // 避免用固定 [1,1] 索引导致单格/空选区（Value2 标量）时越界。
                var lo0 = m.GetLowerBound(0);
                var lo1 = m.GetLowerBound(1);
                var rn = m.GetLength(0);
                var cn = m.GetLength(1);
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        var v = (r < rn && c < cn) ? m[lo0 + r, lo1 + c] : DBNull.Value;
                        result[r, c] = v is DBNull ? null : v;
                    }
                }
            }
            else
            {
                // 单格选区：Value2 返回标量（数值/文本/null），仅填充 (0,0)。
                result[0, 0] = obj;
            }

            return result;
        }

        // ================= 目标带数据守卫（统一标准） =================

        private bool IsTargetOccupied(string toolName, ParamSelection sel, object[,] matrix, int topRow, int startCol,
            Dictionary<string, object> args, out string hint)
        {
            hint = null;
            if (!RibbonCommandPlanner.TryGetTargetRect(toolName, sel, matrix, args, topRow, startCol,
                out var sheetName, out var top, out var left, out var rows, out var cols))
            {
                return false; // 无可静态预算的目标矩形（动态目标/原位写回），不做守卫
            }

            dynamic sheet = null;
            dynamic block = null;
            try
            {
                sheet = ResolveSheet(sheetName);
                block = BuildSheetRange(sheet, top, left, rows, cols);
                if (BlockHasData(block, rows, cols))
                {
                    hint = "目标写入位置已有数据，可能被覆盖；请先清空该区域或换用其他位置，已取消执行。";
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                DiagnosticsLog.Write("[LogCover] RibbonCommandService.cs:477  silent-catch covered (no exception variable)");
                return false; // 守卫是保护性检查：读取失败不误拦，由工具自身校验兜底
            }
            finally
            {
                SafeRelease((object)block);
                SafeRelease((object)sheet);
            }
        }

        private dynamic ResolveSheet(string sheetName)
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    return application.ActiveSheet;
                }
                return workbook.Worksheets[sheetName];
            }
            finally
            {
                SafeRelease((object)workbook);
            }
        }

        private static dynamic BuildSheetRange(dynamic sheet, int top, int left, int rowCount, int colCount)
        {
            if (rowCount == 1 && colCount == 1)
            {
                return sheet.Cells[top, left];
            }

            dynamic r1 = null;
            dynamic r2 = null;
            try
            {
                r1 = sheet.Cells[top, left];
                r2 = sheet.Cells[top + rowCount - 1, left + colCount - 1];
                return sheet.Range[r1, r2];
            }
            finally
            {
                SafeRelease((object)r1);
                SafeRelease((object)r2);
            }
        }

        private static bool BlockHasData(dynamic block, int rows, int cols)
        {
            if (block == null)
            {
                return false;
            }

            var obj = block.Value2;
            if (obj == null)
            {
                return false;
            }

            if (obj is object[,] m)
            {
                // Excel COM 多格 Value2 返回 1 基数组（GetLowerBound=1）：归一化为 0 基矩阵后交给纯函数判定。
                var lo0 = m.GetLowerBound(0);
                var lo1 = m.GetLowerBound(1);
                var rn = m.GetLength(0);
                var cn = m.GetLength(1);
                var normalized = new object[rn, cn];
                for (var r = 0; r < rn; r++)
                {
                    for (var c = 0; c < cn; c++)
                    {
                        var v = m[lo0 + r, lo1 + c];
                        normalized[r, c] = v is DBNull ? null : v;
                    }
                }

                return RibbonCommandPlanner.MatrixHasData(normalized);
            }

            return !string.IsNullOrWhiteSpace(Convert.ToString(obj));
        }

        // ================= 删除行数预算 =================

        private static bool IsTrueArg(Dictionary<string, object> args, string key)
        {
            return args != null && args.TryGetValue(key, out var v)
                && string.Equals(Convert.ToString(v), "true", StringComparison.OrdinalIgnoreCase);
        }

        private static int CountRemovedRows(object[,] matrix, Dictionary<string, object> args)
        {
            var headerRows = Convert.ToInt32(args.TryGetValue("header_rows", out var hr) ? hr : 1);
            var keepLast = string.Equals(Convert.ToString(args.TryGetValue("keep", out var kp) ? kp : "first"), "last", StringComparison.OrdinalIgnoreCase);
            var colsList = new List<int>();
            if (args.TryGetValue("columns", out var raw) && raw is System.Collections.IEnumerable en && !(raw is string))
            {
                foreach (var item in en)
                {
                    if (int.TryParse(Convert.ToString(item), out var cn))
                    {
                        colsList.Add(cn);
                    }
                }
            }

            return RibbonCommandPlanner.CountRemovedRows(matrix, headerRows, colsList, keepLast, IsTrueArg(args, "case_sensitive"));
        }

        // ================= 宿主 =================

        private ExcelToolHost GetHost()
        {
            if (cachedHost != null)
            {
                return cachedHost;
            }

            cachedHost = SidebarBridge.ExcelHostFactory?.Invoke(application);
            if (cachedHost == null)
            {
                var repoPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAiAssistant",
                    "operations.db");
                var repo = new OperationRepository(repoPath);
                cachedHost = new ExcelToolHost(application, repo, ActiveWorkbookHash(), sessionId);
            }

            return cachedHost;
        }

        private string ActiveWorkbookHash()
        {
            dynamic workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                return ExcelToolHost.CreateWorkbookHash(Convert.ToString(workbook?.FullName) ?? string.Empty);
            }
            finally
            {
                SafeRelease((object)workbook);
            }
        }

        // ================= 辅助 =================

        private static string QuoteIfNeeded(string name)
        {
            return name.IndexOfAny(new[] { ' ', '-', '(', ')', '!', '\'', '&', '#' }) >= 0
                ? "'" + name + "'"
                : name;
        }

        private static void SafeRelease(object com)
        {
            if (com != null && Marshal.IsComObject(com))
            {
                try { Marshal.FinalReleaseComObject(com); }
                catch { }
            }
        }
    }
}