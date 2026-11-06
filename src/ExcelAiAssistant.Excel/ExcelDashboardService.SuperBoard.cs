using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using ExcelAiAssistant.Core.Diagnostics;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 超级看板（Super Dashboard，需求-超级看板.md v1.0 P1）：
    /// 任意明细一键生成「数据表底座(超级数据表) + 汇总透视层(staging) + 固定四宫格看板」三层结构。
    /// - 指标选择器（FR-SB-7）：看板下拉切换「销售/利润/自定义」视图，KPI 卡与图表全板联动（SB0 探针 P3 验证的 CHOOSE(MATCH) 镜像机制）；
    /// - KPI 卡视觉升级（FR-SB-9）：圆角矩形+阴影形状 + 文本框绑定 staging 镜像格（SB0 P2 验证的纯引用 Formula 路径，无等号）；
    /// - extra_kpis 派生指标（FR-SB-6）：基础指标四则（如 销售总额/销售量），公式引用基础 KPI 单元格，IFERROR 包裹可核验；
    /// - 透视分析开关（FR-SB-4）：可选用户可见透视 sheet，与看板主链隔离。
    /// 主链仍走 staging 公式（Q1 拍板 B），深度核验/守恒对账全兼容。
    /// </summary>
    public sealed partial class ExcelDashboardService
    {
        public ToolResult RenderSuperDashboard(Dictionary<string, object> args)
        {
            var swLog = System.Diagnostics.Stopwatch.StartNew();
            var sourceAddress = Required(args, "source_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 1);
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 2);
            var monthColumn = ExcelToolHostHelpers.GetIntSafe(args, "month_column", 3);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var quantityColumn = ExcelToolHostHelpers.GetIntSafe(args, "quantity_column", 5);
            var profitColumn = ExcelToolHostHelpers.GetIntSafe(args, "profit_column", 0);   // 0=无利润列（利润视图降级隐藏）
            var dashboardSheet = Optional(args, "dashboard_sheet") ?? "超级看板";
            var superSheet = (Optional(args, "super_sheet") ?? "超级数据表").Trim();
            var theme = (Optional(args, "theme") ?? "orange").Trim().ToLowerInvariant();
            if (!ThemeColors.TryGetValue(theme, out var themeColors))
            {
                return Invalid("INVALID_THEME", $"theme 仅支持 orange/blue/graygreen，当前为 {theme}。");
            }
            var pivotAnalysis = string.Equals(Optional(args, "pivot_analysis"), "true", StringComparison.OrdinalIgnoreCase);
            // Q3-A 拍板：透视字段配置——行维度（逗号分隔 基础字段名 月份/品类/地区）与值字段（金额/数量/利润），非空即构建真·可拖拽数据透视表
            var pivotRowsRaw = (Optional(args, "pivot_rows") ?? "").Trim();
            var pivotMeasuresRaw = (Optional(args, "pivot_measures") ?? "").Trim();
            var pivotRows = string.IsNullOrWhiteSpace(pivotRowsRaw)
                ? new List<string>()
                : pivotRowsRaw.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
            var pivotMeasures = string.IsNullOrWhiteSpace(pivotMeasuresRaw)
                ? new List<string>()
                : pivotMeasuresRaw.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
            if (pivotAnalysis && pivotRowsRaw.Length > 0 && pivotRows.Count == 0)
            {
                return Invalid("INVALID_PIVOT_ROWS", $"pivot_rows 无有效维度：{pivotRowsRaw}（可用：月份/品类/地区）。");
            }
            var slicerEnabled = string.Equals(Optional(args, "slicer_enabled"), "true", StringComparison.OrdinalIgnoreCase);

            var colParams = new[] { labelColumn, categoryColumn, monthColumn, valueColumn, quantityColumn };
            if (colParams.Any(c => c < 1))
            {
                return Invalid("INVALID_COLUMN", "各列参数需 ≥1（1 基）。");
            }

            // ---- extra_kpis 解析（FR-SB-6）：name/formula/format，formula=基础指标名四则 ----
            // SB1-T4 实证：Schema 声明 extra_kpis 为 JArray 字符串（AI 侧实际传参形态），
            // 解析必须先 Parse 字符串，只认 IEnumerable<object> 会让 extra_kpis 静默失效（视图清单缺自定义项→#N/A）
            var extraKpis = new List<SuperKpi>();
            if (args.TryGetValue("extra_kpis", out var extraRaw) && !(extraRaw is string s0 && string.IsNullOrWhiteSpace(s0)) && !(extraRaw is IEnumerable<object> e0 && IsEmpty(e0)))
            {
                IEnumerable<object> extraList = null;
                if (extraRaw is string extraStr && !string.IsNullOrWhiteSpace(extraStr))
                {
                    try
                    {
                        var parsed = Newtonsoft.Json.Linq.JArray.Parse(extraStr);
                        extraList = parsed.Select(t => (object)t).ToList();
                    }
                    catch (Newtonsoft.Json.JsonReaderException)
                    {
                        return Invalid("INVALID_EXTRA_KPI", "extra_kpis 不是合法的 JArray：" + extraStr);
                    }
                }
                else if (extraRaw is IEnumerable<object> direct)
                {
                    extraList = direct;
                }
                if (extraList == null)
                {
                    return Invalid("INVALID_EXTRA_KPI", "extra_kpis 需为 JArray（字符串或数组）。");
                }
                foreach (var item in extraList)
                {
                    string name = null, formula = null, format = null;
                    if (item is Newtonsoft.Json.Linq.JObject jo)
                    {
                        name = jo.Value<string>("name");
                        formula = jo.Value<string>("formula");
                        format = jo.Value<string>("format");
                    }
                    else if (item is Dictionary<string, object> dict)
                    {
                        name = dict.TryGetValue("name", out var n1) ? Convert.ToString(n1) : null;
                        formula = dict.TryGetValue("formula", out var f1) ? Convert.ToString(f1) : null;
                        format = dict.TryGetValue("format", out var f2) ? Convert.ToString(f2) : null;
                    }
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(formula))
                    {
                        return Invalid("INVALID_EXTRA_KPI", "extra_kpis 元素需含 name 与 formula（formula 如「销售总额/销售量」，操作数为内置基础指标名：销售总额/销售量/成交笔数/利润额）。");
                    }
                    if (extraKpis.Count >= 5)
                    {
                        return Invalid("TOO_MANY_EXTRA_KPIS", "extra_kpis 最多 5 个（自定义视图一屏 5 卡）。");
                    }
                    extraKpis.Add(new SuperKpi { Name = name.Trim(), Formula = formula.Trim(), Format = string.IsNullOrWhiteSpace(format) ? "#,##0" : format.Trim() });
                }
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic rangeSource = null;
            dynamic sheetBoard = null;
            dynamic sheetStaging = null;
            dynamic rangeTmp = null;
            dynamic validation = null;
            object bulkGuard = null;
            string dataSheetName = null;
            int monthCount = 0, catCount = 0, labelCount = 0;
            bool hasProfit = profitColumn >= 1;
            string slicerText = "看板内置下拉（品类 C4 / 地区 E4）";
            AppLog.Info("超级看板", $"render 开始 source={sourceAddress} profit={profitColumn} extraKpis={extraKpis.Count} slicer={slicerEnabled} theme={theme} pivot={pivotAnalysis}");
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var autoRename = string.Equals(Optional(args, "auto_rename"), "true", StringComparison.OrdinalIgnoreCase);
                dashboardSheet = ResolveDashboardName(workbook, dashboardSheet, autoRename);
                var sourceSheetName = Optional(args, "sheet");
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                rangeSource = sheetSource.Range[sourceAddress];
                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                var allCols = new[] { labelColumn, categoryColumn, monthColumn, valueColumn, quantityColumn, profitColumn };
                if (allCols.Where(c => c >= 1).Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE", $"列参数超界：区域列数 {srcCols}。");
                }
                string sourceSheetNameReal = Convert.ToString(sheetSource.Name);

                // ---- 步骤①：数据底座（FR-SB-2）：内部自动建超级数据表（Q5 拍板 B 行为 A 实现）----
                var superArgs = new Dictionary<string, object>(args)
                {
                    ["super_sheet"] = superSheet
                };
                var superResult = CreateSuperTable(superArgs);
                if (!superResult.Ok)
                {
                    return superResult;   // 底座失败透传（SUPER_TABLE_NAME_TAKEN 等既有错误码）
                }
                ExcelTrace.Info("render_super_dashboard", "底座超级数据表完成 sheet=" + superSheet);
                // 底座建好后，看板主链数据源切到超级表明细区（A1:E{n+1}，5 列结构固定）
                dataSheetName = superSheet;
                dynamic sheetSuper = workbook.Worksheets.Item[superSheet];
                int detailRows = 0;
                dynamic usedTmp = sheetSuper.UsedRange;
                detailRows = Convert.ToInt32(usedTmp.Rows.Count) - 1;
                Release((object)usedTmp);
                if (detailRows < 1) { return Invalid("NO_DATA", "超级数据表明细行为 0。"); }
                if (detailRows > 99998)
                {
                    return Invalid("ROWS_OUT_OF_RANGE",
                        $"超级数据表明细 {detailRows} 行超出超级看板 staging 公式引用上限 99998 行，请拆分数据源。");
                }

                ExcelTrace.Info("render_super_dashboard", "明细行数=" + detailRows);
                // ---- 步骤②：幂等删旧（看板/staging/透视分析）----
                foreach (var name in new[] { dashboardSheet, "_" + dashboardSheet + "_calc", "透视分析-" + dashboardSheet })
                {
                    dynamic sheetOld = null;
                    if (SheetExists(workbook, name))
                    {
                        sheetOld = workbook.Worksheets.Item[name];
                        try { sheetOld.Delete(); } catch { }
                        Release((object)sheetOld);
                    }
                    sheetOld = null;
                }

                sheetBoard = workbook.Worksheets.Add();
                sheetBoard.Name = dashboardSheet;
                sheetStaging = workbook.Worksheets.Add(After: sheetBoard);
                sheetStaging.Name = "_" + dashboardSheet + "_calc";

                // ---- 步骤③：staging 构建（全指标区 + 镜像区；公式全部引用超级表明细 + 看板控制格）----
                string escSuper = EscapeSheet(dataSheetName);
                string dRef = $"{escSuper}!$D$2:$D$99999";    // 金额（超级表 D 列）
                string qRef = $"{escSuper}!$E$2:$E$99999";    // 数量（超级表 E 列）
                string lRef = $"{escSuper}!$A$2:$A$99999";    // 地区/label
                string cRef = $"{escSuper}!$B$2:$B$99999";    // 品类
                string mRef = $"{escSuper}!$C$2:$C$99999";    // 月份
                // 利润列直接引源明细（不扩超级表 5 列结构；SB 设计 Q1：利润为可选扩展指标）
                // SB1-T3 实证：利润 SUMIFS 的 sum_range 与 criteria_range 必须同源同尺寸——
                // pRef(源明细 16 行) × cRef(超级表 65536 行) 尺寸不匹配恒 0，利润链全部用源明细同源维度引用
                string pRef = null, psCatRef = null, psLblRef = null, psMonthRef = null;
                if (hasProfit)
                {
                    var addr = Convert.ToString(rangeSource.Address, CultureInfo.InvariantCulture);
                    var letters = ColumnLetters(addr, labelColumn, profitColumn, out int hrP, out int lrP);
                    pRef = $"{EscapeSheet(sourceSheetNameReal)}!${letters[1]}${hrP + 1}:${letters[1]}${lrP}";
                    var catL = ColumnLetters(addr, labelColumn, categoryColumn, out int hc1, out int hl1)[1];
                    var lblL = ColumnLetters(addr, labelColumn, labelColumn, out int hc2, out int hl2)[1];
                    var monL = ColumnLetters(addr, labelColumn, monthColumn, out int hc3, out int hl3)[1];
                    psCatRef = $"{EscapeSheet(sourceSheetNameReal)}!${catL}${hc1 + 1}:${catL}${hl1}";
                    psLblRef = $"{EscapeSheet(sourceSheetNameReal)}!${lblL}${hc2 + 1}:${lblL}${hl2}";
                    psMonthRef = $"{EscapeSheet(sourceSheetNameReal)}!${monL}${hc3 + 1}:${monL}${hl3}";
                }

                // 看板控制格：C2=品类下拉、E2=人员下拉、H2=指标视图下拉（版式见 RenderSuperDashboardBoard）
                string boardEsc = "'" + EscapeSheetInner(dashboardSheet) + "'!";
                string catCond = $"IF({boardEsc}$C$4=\"全部\",\"*\",{boardEsc}$C$4)";
                string lblCond = $"IF({boardEsc}$E$4=\"全部\",\"*\",{boardEsc}$E$4)";
                string viewRef = $"{boardEsc}$H$4";

                // 维度去重（块读，超级表 A:C 列，NFR-2 纪律）
                dynamic dimRange = sheetSuper.Range["$A$2:$C$" + (detailRows + 1)];
                object[,] dimBlock = dimRange.Value2;
                Release((object)dimRange);
                var monthSet = new SortedSet<string>();
                var catSet = new SortedSet<string>();
                var labelSet = new SortedSet<string>();
                for (int r = 1; r <= detailRows; r++)
                {
                    if (dimBlock[r, 1] != null) { labelSet.Add(Convert.ToString(dimBlock[r, 1])); }
                    if (dimBlock[r, 2] != null) { catSet.Add(Convert.ToString(dimBlock[r, 2])); }
                    if (dimBlock[r, 3] != null) { monthSet.Add(Convert.ToString(dimBlock[r, 3])); }
                }
                Release((object)dimBlock);
                if (monthSet.Count == 0 || catSet.Count == 0 || labelSet.Count == 0)
                {
                    return Invalid("NO_DIMENSION_DATA",
                        $"月份 {monthSet.Count}/品类 {catSet.Count}/地区 {labelSet.Count} 存在空维度，与长表明细口径不符。请确认数据源是一行一条的明细记录。");
                }
                monthCount = monthSet.Count; catCount = catSet.Count; labelCount = labelSet.Count;

                bulkGuard = detailRows >= 5000 ? BeginBulkFormulaWrite() : null;

                // ---- staging A 列：视图清单（指标选择器下拉源）----
                var views = new List<string> { "销售视图" };
                if (hasProfit) { views.Add("利润视图"); }
                if (extraKpis.Count > 0) { views.Add("自定义视图"); }
                for (int i = 0; i < views.Count; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 1, 1];
                    rangeTmp.Value2 = views[i];
                    Release((object)rangeTmp);
                }

                // ---- staging C:E 月度销售矩阵（金额/数量，双维度掩码）----
                rangeTmp = sheetStaging.Range["C1"]; rangeTmp.Value2 = "月份"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["D1"]; rangeTmp.Value2 = "金额"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["E1"]; rangeTmp.Value2 = "数量"; Release((object)rangeTmp);
                int mi = 2;
                foreach (var month in monthSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[mi, 3];
                    if (month is string) { rangeTmp.NumberFormat = "@"; }   // 文本月份锁型（类型漂移教训）
                    rangeTmp.Value2 = month;
                    Release((object)rangeTmp);
                    dynamic c1 = sheetStaging.Cells.Item[mi, 4];
                    c1.Formula = $"=SUMIFS({dRef},{mRef},C{mi},{cRef},{catCond},{lRef},{lblCond})";
                    Release((object)c1);
                    dynamic c2 = sheetStaging.Cells.Item[mi, 5];
                    c2.Formula = $"=SUMIFS({qRef},{mRef},C{mi},{cRef},{catCond},{lRef},{lblCond})";
                    Release((object)c2);
                    mi++;
                }

                // ---- staging G:H 品类-金额、I:J 地区-金额 ----
                rangeTmp = sheetStaging.Range["G1"]; rangeTmp.Value2 = "品类"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["H1"]; rangeTmp.Value2 = "金额"; Release((object)rangeTmp);
                int ci2 = 2;
                foreach (var cat in catSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[ci2, 7]; rangeTmp.Value2 = cat; Release((object)rangeTmp);
                    dynamic c3 = sheetStaging.Cells.Item[ci2, 8];
                    c3.Formula = $"=SUMIFS({dRef},{cRef},G{ci2},{lRef},{lblCond})";
                    Release((object)c3);
                    ci2++;
                }
                rangeTmp = sheetStaging.Range["I1"]; rangeTmp.Value2 = "地区"; Release((object)rangeTmp);
                rangeTmp = sheetStaging.Range["J1"]; rangeTmp.Value2 = "金额"; Release((object)rangeTmp);
                int li2 = 2;
                foreach (var lbl in labelSet)
                {
                    rangeTmp = sheetStaging.Cells.Item[li2, 9]; rangeTmp.Value2 = lbl; Release((object)rangeTmp);
                    dynamic c4 = sheetStaging.Cells.Item[li2, 10];
                    c4.Formula = $"=SUMIFS({dRef},{lRef},I{li2},{cRef},{catCond})";
                    Release((object)c4);
                    li2++;
                }

                // ---- staging KPI 库：T 列销售视图 5 指标、U 列利润视图 5 指标（并排对齐 CHOOSE）----
                // 掩码双维度 AND；比率类 IFERROR；最大单笔走 SUMPRODUCT 惯用法（数组式 #VALUE! 教训）
                string sumAll = $"SUMIFS({dRef},{cRef},{catCond},{lRef},{lblCond})";
                string cntAll = $"COUNTIFS({cRef},{catCond},{lRef},{lblCond})";
                var tKpis = new[]
                {
                    ("销售总额", $"=IFERROR({sumAll},0)", "$#,##0"),
                    ("成交笔数", $"=IFERROR({cntAll},0)", "#,##0"),
                    ("平均客单价", $"=IFERROR({sumAll}/{cntAll},0)", "¥#,##0.0"),
                    ("最大单笔", $"=IFERROR(SUMPRODUCT(MAX(IF({cRef}={catCond},1,IF({lRef}={lblCond},1,1))*{dRef})),0)", "$#,##0"),
                    ("销售量", $"=IFERROR(SUMIFS({qRef},{cRef},{catCond},{lRef},{lblCond}),0)", "#,##0")
                };
                for (int i = 0; i < tKpis.Length; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 20];   // T2:T6
                    rangeTmp.Formula = tKpis[i].Item2;
                    Release((object)rangeTmp);
                }
                var uKpis = hasProfit
                    ? new[]
                    {
                        ("利润总额", $"=IFERROR(SUMIFS({pRef},{psCatRef},{catCond},{psLblRef},{lblCond}),0)", "$#,##0"),
                        ("利润占比", $"=IFERROR(SUMIFS({pRef},{psCatRef},{catCond},{psLblRef},{lblCond})/{sumAll},0)", "0.0%"),
                        ("百万利润", $"=IFERROR(SUMIFS({pRef},{psCatRef},{catCond},{psLblRef},{lblCond})/1000000,0)", "0.00"),
                        ("笔均利润", $"=IFERROR(SUMIFS({pRef},{psCatRef},{catCond},{psLblRef},{lblCond})/{cntAll},0)", "¥#,##0.0"),
                        ("最大利润单", $"=IFERROR(SUMPRODUCT(MAX({pRef})),0)", "$#,##0")
                    }
                    : new[]
                    {
                        ("数量总额", $"=IFERROR(SUMIFS({qRef},{cRef},{catCond},{lRef},{lblCond}),0)", "#,##0"),
                        ("平均单笔数量", $"=IFERROR(SUMIFS({qRef},{cRef},{catCond},{lRef},{lblCond})/{cntAll},0)", "#,##0.0"),
                        ("百万金额", $"=IFERROR({sumAll}/1000000,0)", "0.00"),
                        ("最大单笔数量", $"=IFERROR(SUMPRODUCT(MAX({qRef})),0)", "#,##0"),
                        ("平均订单金额", $"=IFERROR({sumAll}/{cntAll},0)", "¥#,##0.0")
                    };
                for (int i = 0; i < uKpis.Length; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 21];   // U2:U6
                    rangeTmp.Formula = uKpis[i].Item2;
                    Release((object)rangeTmp);
                }

                // ---- staging 镜像区（指标选择器驱动，全板联动核心，SB0 P3 验证机制）----
                string viewMatch = $"MATCH({viewRef},$A$1:$A${views.Count},0)";
                // V:X 月度镜像（柱/线随视图：销售=C:D、利润=K:M 列的利润+占比、自定义=C:D）
                // K:M 利润月度矩阵（有利润时）：K 月份 L 利润 M 占比——先建利润月度矩阵
                int lastMonthRow = monthCount + 1;
                // K:M 利润月度矩阵（利润视图镜像源；K 月份 / L 利润 / M 占比）
                if (hasProfit)
                {
                    rangeTmp = sheetStaging.Range["K1"]; rangeTmp.Value2 = "月份"; Release((object)rangeTmp);
                    rangeTmp = sheetStaging.Range["L1"]; rangeTmp.Value2 = "利润"; Release((object)rangeTmp);
                    rangeTmp = sheetStaging.Range["M1"]; rangeTmp.Value2 = "利润占比"; Release((object)rangeTmp);
                    int kmi = 2;
                    foreach (var month in monthSet)
                    {
                        rangeTmp = sheetStaging.Cells.Item[kmi, 11];
                        if (month is string) { rangeTmp.NumberFormat = "@"; }
                        rangeTmp.Value2 = month;
                        Release((object)rangeTmp);
                        dynamic p1 = sheetStaging.Cells.Item[kmi, 12];
                        p1.Formula = $"=IFERROR(SUMIFS({pRef},{psMonthRef},K{kmi},{psCatRef},{catCond},{psLblRef},{lblCond}),0)";
                        Release((object)p1);
                        dynamic p2 = sheetStaging.Cells.Item[kmi, 13];
                        p2.Formula = $"=IFERROR(L{kmi}/D{kmi},0)";
                        Release((object)p2);
                        kmi++;
                    }
                }
                for (int r = 2; r <= lastMonthRow; r++)
                {
                    rangeTmp = sheetStaging.Cells.Item[r, 22];   // V 月份
                    rangeTmp.Formula = hasProfit ? $"=CHOOSE({viewMatch},C{r},K{r},C{r})" : $"=C{r}";
                    Release((object)rangeTmp);
                    rangeTmp = sheetStaging.Cells.Item[r, 23];   // W 柱值
                    rangeTmp.Formula = hasProfit ? $"=CHOOSE({viewMatch},D{r},L{r},D{r})" : $"=D{r}";
                    Release((object)rangeTmp);
                    rangeTmp = sheetStaging.Cells.Item[r, 24];   // X 线值
                    rangeTmp.Formula = hasProfit ? $"=CHOOSE({viewMatch},E{r},M{r},E{r})" : $"=E{r}";
                    Release((object)rangeTmp);
                }
                // Z 列 KPI 镜像 5 槽（卡值文本框纯引用绑定目标；自定义分支引 AD 派生列）
                for (int i = 0; i < 5; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 26];
                    rangeTmp.Formula = $"=CHOOSE({viewMatch},T{i + 2},U{i + 2}" + (extraKpis.Count > 0 ? ",AD" + (i + 2) : "") + ")";
                    Release((object)rangeTmp);
                }
                // AA 列 KPI 标题镜像（TEXT 串随视图换）
                var tTitles = tKpis.Select(k => k.Item1).ToArray();
                var uTitles = uKpis.Select(k => k.Item1).ToArray();
                var extraTitles = extraKpis.Select(k => k.Name).ToArray();
                for (int i = 0; i < 5; i++)
                {
                    string tPart = $"\"{tTitles[i]}\"";
                    string uPart = $"\"{uTitles[i]}\"";
                    string ePart = extraKpis.Count > 0 ? $",\"{(i < extraTitles.Length ? extraTitles[i] : "—")}\"" : "";
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 27];
                    rangeTmp.Formula = $"=CHOOSE({viewMatch},{tPart},{uPart}{ePart})";
                    Release((object)rangeTmp);
                }
                // AB 列 KPI 格式串镜像（卡值 NumberFormat 不便动态改，用 TEXT 呈现在镜像值——Z 列直取数值，
                // 卡值文本框绑定 AA 标题 + Z 数值；格式由卡值格 NumberFormat 固定货币——v1 取舍：格式以视图主流货币/数量混排取 General）
                // ---- staging 自定义视图派生行 W2:W6（extra_kpis：引用基础 KPI 格四则，IFERROR 包裹）----
                if (extraKpis.Count > 0)
                {
                    var baseCells = new Dictionary<string, string>
                    {
                        ["销售总额"] = "$T$2", ["成交笔数"] = "$T$3", ["平均客单价"] = "$T$4",
                        ["最大单笔"] = "$T$5", ["销售量"] = "$T$6",
                        ["利润总额"] = "$U$2", ["利润占比"] = "$U$3", ["百万利润"] = "$U$4",
                        ["笔均利润"] = "$U$5", ["最大利润单"] = "$U$6"
                    };
                    for (int i = 0; i < extraKpis.Count; i++)
                    {
                        string expr = extraKpis[i].Formula;
                        foreach (var kv in baseCells)
                        {
                            expr = System.Text.RegularExpressions.Regex.Replace(
                                expr, System.Text.RegularExpressions.Regex.Escape(kv.Key), kv.Value);
                        }
                        var unknown = System.Text.RegularExpressions.Regex.Matches(expr, @"[\u4e00-\u9fa5]+");
                        if (unknown.Count > 0)
                        {
                            return Invalid("INVALID_EXTRA_KPI",
                                $"extra_kpis「{extraKpis[i].Name}」含未知指标「{unknown[0].Value}」。可用基础指标：{string.Join("、", baseCells.Keys)}。");
                        }
                        // SB1-T4 实证：W 列被月度镜像(V:X)占用——自定义派生值放独立列 AD，禁止与镜像区重叠
                        rangeTmp = sheetStaging.Cells.Item[i + 2, 30];   // AD2:AD6 自定义派生值
                        rangeTmp.Formula = $"=IFERROR({expr},0)";
                        Release((object)rangeTmp);
                    }
                }

                // ---- 步骤④：看板绘制（版式+形状 KPI 卡+四宫格）----
                // v1.1 用户文件实证：控制标签硬编码「品类/地区」与实际维度语义不符（label 可能是销售员等），
                // 从源明细表头读真实维度名
                int srcFirstCol = Convert.ToInt32(rangeSource.Column);
                int headerRow = Convert.ToInt32(rangeSource.Row);
                string catHeader = ReadHeaderText(sheetSource, headerRow, srcFirstCol + categoryColumn - 1, "品类");
                string lblHeader = ReadHeaderText(sheetSource, headerRow, srcFirstCol + labelColumn - 1, "地区");
                string monthHeader = ReadHeaderText(sheetSource, headerRow, srcFirstCol + monthColumn - 1, "月份");
                string valueHeader = ReadHeaderText(sheetSource, headerRow, srcFirstCol + valueColumn - 1, "金额");
                string quantityHeader = ReadHeaderText(sheetSource, headerRow, srcFirstCol + quantityColumn - 1, "数量");
                string profitHeader = hasProfit ? ReadHeaderText(sheetSource, headerRow, srcFirstCol + profitColumn - 1, "利润") : "利润";
                RenderSuperDashboardBoard(dashboardSheet, sheetBoard, sheetStaging, themeColors,
                    monthCount, catCount, labelCount, views, extraKpis, hasProfit, catHeader, lblHeader);

                // ---- 步骤⑤：透视分析开关（FR-SB-4，与主链隔离；Q3-A：字段可配置真透视表）----
                if (pivotAnalysis)
                {
                    // 维度/值语义名 → 源明细真实表头名（透视表字段名必须与数据源表头一致）
                    var pivotRowFields = new List<string>();
                    foreach (var dimName in pivotRows)
                    {
                        string field = dimName == "月份" ? monthHeader : dimName == "品类" ? catHeader : dimName == "地区" ? lblHeader : dimName;
                        pivotRowFields.Add(field);
                    }
                    var pivotMeasureFields = new List<string>();
                    foreach (var measureName in pivotMeasures)
                    {
                        string field = measureName == "金额" ? valueHeader : measureName == "数量" ? quantityHeader : measureName == "利润" ? profitHeader : measureName;
                        pivotMeasureFields.Add(field);
                    }
                    BuildPivotAnalysisSheet(workbook, dataSheetName, detailRows, dashboardSheet, sourceSheetNameReal,
                        sourceAddress, pivotRowFields, pivotMeasureFields);
                }

                // ---- 步骤⑥：切片器（FR-SB-8 表单开关；native 多选，label 维度；不可用回退看板内置下拉）----
                if (slicerEnabled)
                {
                    var linkResult = SetLinkage(new Dictionary<string, object>
                    {
                        ["source_address"] = "A1:E" + (detailRows + 1),
                        ["sheet"] = dataSheetName,
                        ["dashboard_sheet"] = dashboardSheet,
                        ["mode"] = "native",
                        ["field"] = "label"
                    });
                    slicerText = linkResult.Ok
                        ? "native 切片器（多选）已挂载，与内置下拉并存"
                        : "native 切片器挂载失败（已回退内置下拉）：" + Convert.ToString(linkResult.Data);
                    ExcelTrace.Info("render_super_dashboard", "slicer ok=" + linkResult.Ok);
                }
            }
            catch (InvalidOperationException ex)
            {
                ExcelTrace.Error("render_super_dashboard", "InvalidOperation: " + ex.Message);
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("SUPER_BOARD_RENDER_FAILED", ex.Message);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ExcelTrace.Error("render_super_dashboard", "COM HRESULT=0x" + ex.HResult.ToString("X8") + ": " + ex.Message);
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("SUPER_BOARD_RENDER_FAILED",
                    "超级看板渲染 COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            catch (Exception ex)
            {
                ExcelTrace.Error("render_super_dashboard", ex.GetType().Name + ": " + ex.Message);
                TryDeleteSheet(sheetBoard);
                TryDeleteSheet(sheetStaging);
                return Invalid("SUPER_BOARD_RENDER_FAILED",
                    "超级看板渲染失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally
            {
                EndBulkFormulaWrite(bulkGuard, dashboardSheet, "_" + dashboardSheet + "_calc");
                Release((object)validation);
                Release((object)rangeTmp);
                Release((object)rangeSource);
                Release((object)sheetSource);
            }

            // ---- 渲染后自检（守恒对账：镜像 KPI vs 源明细直和）----
            var selfcheck = RunPostRenderSelfCheck(dashboardSheet);
            ExcelTrace.Info("render_super_dashboard", "自检=" + selfcheck.Verdict + " elapsedMs=" + swLog.ElapsedMilliseconds);
            recorder?.Invoke("render_super_dashboard", args, "dashboard-render",
                new
                {
                    template = "super_board",
                    sheet = dashboardSheet,
                    base_sheet = dataSheetName,
                    months = monthCount,
                    categories = catCount,
                    labels = labelCount,
                    views = hasProfit ? 2 : 1,
                    extra_kpis = extraKpis.Count,
                    pivot_analysis = pivotAnalysis,
                    selfcheck = selfcheck.Verdict
                },
                new { Sheet = dashboardSheet, Template = "super_board" });

            AppLog.Info("超级看板", "render 完成 sheet=" + dashboardSheet + " elapsedMs=" + swLog.ElapsedMilliseconds + " 自检=" + selfcheck.Verdict);
            return new ToolResult
            {
                Ok = true,
                Data = new
                {
                    template = "super_board",
                    dashboard_sheet = dashboardSheet,
                    base_sheet = dataSheetName,
                    staging_sheet = "_" + dashboardSheet + "_calc",
                    views = hasProfit ? new[] { "销售视图", "利润视图" } : new[] { "销售视图" },
                    extra_kpi_count = extraKpis.Count,
                    months = monthCount,
                    categories = catCount,
                    labels = labelCount,
                    pivot_analysis = pivotAnalysis,
                    pivot_rows = pivotRows,
                    pivot_measures = pivotMeasures,
                    slicer = slicerText,
                    usage = "看板 H4 切换指标视图；C4/E4 维度下拉筛选；后续可用 verify_dashboard_deep 深度核验",
                    selfcheck = (selfcheck.Verdict == "ok" ? "自检通过 ✓ " : selfcheck.Verdict == "mismatch" ? "自检异常 ✗ " : "") + selfcheck.Detail,
                    selfcheck_verdict = selfcheck.Verdict,
                    theme = theme,
                    created = true
                }
            };
        }

        /// <summary>
        /// 分步工具②（向导 Step2）：独立生成透视分析 sheet——按勾选的行维度与值字段构建
        /// 可见、可拖拽的原生数据透视表（数据源=源明细区域，字段名=源表头）；失败降级静态快照。
        /// 与超级看板解耦：可先建数据表/透视表，最后再生成看板（用户"一步步来"拍板）。
        /// </summary>
        public ToolResult BuildPivotSheet(Dictionary<string, object> args)
        {
            var swLog = System.Diagnostics.Stopwatch.StartNew();
            var sourceAddress = Required(args, "source_address");
            var labelColumn = ExcelToolHostHelpers.GetIntSafe(args, "label_column", 1);
            var categoryColumn = ExcelToolHostHelpers.GetIntSafe(args, "category_column", 2);
            var monthColumn = ExcelToolHostHelpers.GetIntSafe(args, "month_column", 3);
            var valueColumn = ExcelToolHostHelpers.GetIntSafe(args, "value_column", 4);
            var quantityColumn = ExcelToolHostHelpers.GetIntSafe(args, "quantity_column", 5);
            var profitColumn = ExcelToolHostHelpers.GetIntSafe(args, "profit_column", 0);
            var pivotRowsRaw = (Optional(args, "pivot_rows") ?? "").Trim();
            var pivotMeasuresRaw = (Optional(args, "pivot_measures") ?? "").Trim();
            var pivotSheetName = (Optional(args, "pivot_sheet_name") ?? "透视分析").Trim();
            AppLog.Info("超级看板", $"build_pivot_sheet 开始 rows={pivotRowsRaw} measures={pivotMeasuresRaw} sheet={pivotSheetName}");
            if (string.IsNullOrWhiteSpace(pivotRowsRaw) || string.IsNullOrWhiteSpace(pivotMeasuresRaw))
            {
                return Invalid("INVALID_PIVOT_FIELDS", "pivot_rows 与 pivot_measures 均必填（逗号分隔，如「月份,品类」与「金额,利润」）。");
            }

            dynamic workbook = null;
            dynamic sheetSource = null;
            dynamic rangeSource = null;
            try
            {
                workbook = application.ActiveWorkbook
                           ?? throw new InvalidOperationException("当前没有打开的工作簿。");
                var sourceSheetName = Optional(args, "sheet");
                sheetSource = string.IsNullOrWhiteSpace(sourceSheetName)
                    ? workbook.ActiveSheet
                    : workbook.Worksheets.Item[sourceSheetName];
                rangeSource = sheetSource.Range[sourceAddress];
                int srcCols = Convert.ToInt32(rangeSource.Columns.Count);
                int srcFirstCol = Convert.ToInt32(rangeSource.Column);
                int headerRow = Convert.ToInt32(rangeSource.Row);
                var allCols = new[] { labelColumn, categoryColumn, monthColumn, valueColumn, quantityColumn, profitColumn };
                if (allCols.Where(c => c >= 1).Any(c => c > srcCols))
                {
                    return Invalid("COLUMN_OUT_OF_RANGE", $"列参数超界：区域列数 {srcCols}。");
                }
                bool hasProfit = profitColumn >= 1;

                // ---- 维度/值语义名 → 源表头真实名（透视字段名必须与数据源表头一致）----
                string HeaderOf(int column, string fallback)
                {
                    dynamic cell = null;
                    try
                    {
                        cell = sheetSource.Cells.Item[headerRow, srcFirstCol + column - 1];
                        var text = Convert.ToString(cell.Value2);
                        return string.IsNullOrWhiteSpace(text) ? fallback : text.Trim();
                    }
                    finally { Release((object)cell); }
                }
                var dimMap = new Dictionary<string, int>
                {
                    ["月份"] = monthColumn, ["日期"] = monthColumn,
                    ["品类"] = categoryColumn, ["类别"] = categoryColumn,
                    ["地区"] = labelColumn, ["销售员"] = labelColumn, ["人员"] = labelColumn
                };
                var measureMap = new Dictionary<string, int>
                {
                    ["金额"] = valueColumn, ["销售额"] = valueColumn,
                    ["数量"] = quantityColumn, ["销量"] = quantityColumn,
                    ["利润"] = hasProfit ? profitColumn : -1
                };
                var rowFields = new List<string>();
                foreach (var raw in pivotRowsRaw.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var name = raw.Trim();
                    if (!dimMap.TryGetValue(name, out var col))
                    {
                        // 允许直接传表头名：在表头里找
                        int found = -1;
                        for (int c = 1; c <= srcCols; c++)
                        {
                            if (HeaderOf(c, "") == name) { found = c; break; }
                        }
                        if (found < 0)
                        {
                            return Invalid("INVALID_PIVOT_ROWS", $"未知行维度「{name}」。可用：月份/品类/地区（或直接传表头列名）。");
                        }
                        col = found;
                    }
                    rowFields.Add(HeaderOf(col, name));
                }
                var measureFields = new List<string>();
                foreach (var raw in pivotMeasuresRaw.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var name = raw.Trim();
                    if (!measureMap.TryGetValue(name, out var col))
                    {
                        int found = -1;
                        for (int c = 1; c <= srcCols; c++)
                        {
                            if (HeaderOf(c, "") == name) { found = c; break; }
                        }
                        if (found < 0)
                        {
                            return Invalid("INVALID_PIVOT_MEASURES", $"未知值字段「{name}」。可用：金额/数量/利润（或直接传表头列名）。");
                        }
                        col = found;
                    }
                    if (col < 1) { return Invalid("INVALID_PIVOT_MEASURES", $"值字段「{name}」需要利润列（本源未提供 profit_column）。"); }
                    measureFields.Add(HeaderOf(col, name));
                }
                if (rowFields.Distinct().Count() != rowFields.Count || measureFields.Distinct().Count() != measureFields.Count)
                {
                    return Invalid("INVALID_PIVOT_FIELDS", "行维度/值字段存在重复。");
                }

                // ---- 构建真透视表（与超级看板透视分析同源逻辑；失败降级静态快照）----
                var sheetDstList = new List<string>();
                var buildResult = BuildPivotTableIntoNewSheet(workbook, Convert.ToString(sheetSource.Name),
                    Convert.ToString(rangeSource.Address, CultureInfo.InvariantCulture), pivotSheetName,
                    rowFields, measureFields);
                ExcelTrace.Info("build_pivot_sheet", "build ok=" + buildResult.Ok + " elapsedMs=" + swLog.ElapsedMilliseconds);
                return buildResult;
            }
            catch (InvalidOperationException ex)
            {
                ExcelTrace.Error("build_pivot_sheet", "InvalidOperation: " + ex.Message);
                return Invalid("PIVOT_SHEET_FAILED", ex.Message);
            }
            catch (Exception ex)
            {
                // 真透视表失败 → 返回可读错误（静态快照降级链在 BuildPivotAnalysisSheet，不在此处）
                ExcelTrace.Error("build_pivot_sheet", "build failed: " + ex.Message);
                return Invalid("PIVOT_SHEET_FAILED",
                    "透视分析构建失败：" + ex.Message + "。已回滚，源数据未受影响。");
            }
            finally

            {

                Release((object)rangeSource);

                Release((object)sheetSource);
            }
        }

        private static string sourceSheetNameReal(dynamic sheetSource)
        {
            return Convert.ToString(sheetSource.Name);
        }

        /// <summary>
        /// 共用核心：在新建 sheet 上构建原生数据透视表（行字段 xlRowField + 值字段 xlSum + xlTabularForm）。
        /// 供 render_super_dashboard 透视分析与分步工具 build_pivot_sheet 复用。失败抛异常（调用方决定降级）。
        /// 隐藏实例注意：不需要 Range.Select（SB3 实证对非活动表 Select 会炸且无必要）。
        /// </summary>
        private ToolResult BuildPivotTableIntoNewSheet(
            dynamic workbook, string sourceSheetName, string sourceAddress, string dstSheetName,
            List<string> rowFields, List<string> measureFields)
        {
            dynamic sheetSrc = null;
            dynamic rangeSrc = null;
            dynamic pivotCache = null;
            dynamic pivotTable = null;
            dynamic sheetDst = null;
            try
            {
                sheetSrc = workbook.Worksheets.Item[sourceSheetName];
                rangeSrc = sheetSrc.Range[sourceAddress];
                sheetDst = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
                sheetDst.Name = dstSheetName;
                pivotCache = workbook.PivotCaches().Create(SourceType: 1, SourceData: rangeSrc);   // xlDatabase
                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: "'" + EscapeSheetInner(dstSheetName) + "'!R3C1",
                    TableName: "透视分析表_" + Guid.NewGuid().ToString("N").Substring(0, 6));
                foreach (var fieldName in rowFields)
                {
                    dynamic pf = pivotTable.PivotFields(fieldName);
                    pf.Orientation = 1;   // xlRowField
                    Release((object)pf);
                }
                foreach (var fieldName in measureFields)
                {
                    dynamic pf = pivotTable.PivotFields(fieldName);
                    pf.Orientation = 4;   // xlDataField
                    pf.Function = -4157;  // xlSum
                    pf.Caption = fieldName + " 汇总";
                    Release((object)pf);
                }
                ExcelTrace.Info("BuildPivotTableIntoNewSheet", "rows=" + string.Join("/", rowFields) + " measures=" + string.Join("/", measureFields));
                return new ToolResult
                {
                    Ok = true,
                    Data = new { sheet = dstSheetName, rows = rowFields, measures = measureFields, created = true }
                };
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                return new ToolResult { Ok = false, ErrorCode = "PIVOT_BUILD_FAILED", Data = "COM 失败（HRESULT=0x" + ex.HResult.ToString("X8") + "）：" + ex.Message };
            }
            catch (Exception ex)
            {
                return new ToolResult { Ok = false, ErrorCode = "PIVOT_BUILD_FAILED", Data = ex.GetType().Name + " " + ex.Message };
            }
            finally
            {
                Release((object)pivotTable);
                Release((object)pivotCache);
                Release((object)rangeSrc);
                Release((object)sheetDst);
                Release((object)sheetSrc);
            }
        }

        /// <summary>超级 KPI 定义（extra_kpis）：formula=基础指标名四则表达式。</summary>
        private sealed class SuperKpi
        {
            public string Name;
            public string Formula;
            public string Format;
        }

        /// <summary>
        /// 超级看板绘制：控制区（品类/人员/指标视图三下拉）+ 形状 KPI 卡 5 张（圆角矩形+阴影+文本框绑定镜像格）
        /// + 固定四宫格（趋势组合[随视图]/品类排行/地区构成/月度双轴全景）+ 关网格线。构建失败以异常冒泡，外层统一回滚。
        /// </summary>
        private void RenderSuperDashboardBoard(
            string dashboardSheet, dynamic sheetBoard, dynamic sheetStaging, ThemePalette themeColors,
            int monthCount, int catCount, int labelCount, List<string> views, List<SuperKpi> extraKpis, bool hasProfit,
            string catHeader, string lblHeader)
        {
            dynamic rangeTmp = null;
            dynamic validation = null;
            dynamic shape = null;
            try
            {
                string stagingEsc = "'" + EscapeSheetInner("_" + dashboardSheet + "_calc") + "'!";

                // ---- 控制区（单元格下拉×3）----
                rangeTmp = sheetBoard.Range["B4"]; rangeTmp.Value2 = catHeader + "："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["C4"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                validation = sheetBoard.Range["C4"].Validation;
                validation.Delete();
                validation.Add(Type: 3, Formula1: $"={stagingEsc}$AF$1:$AF${catCount + 1}");
                Release((object)validation); validation = null;
                rangeTmp = sheetBoard.Range["D4"]; rangeTmp.Value2 = lblHeader + "："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["E4"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                validation = sheetBoard.Range["E4"].Validation;
                validation.Delete();
                validation.Add(Type: 3, Formula1: $"={stagingEsc}$AG$1:$AG${labelCount + 1}");
                Release((object)validation); validation = null;
                rangeTmp = sheetBoard.Range["G4"]; rangeTmp.Value2 = "指标视图："; Release((object)rangeTmp);
                rangeTmp = sheetBoard.Range["H4"]; rangeTmp.Value2 = views[0]; Release((object)rangeTmp);
                validation = sheetBoard.Range["H4"].Validation;
                validation.Delete();
                validation.Add(Type: 3, Formula1: $"={stagingEsc}$A$1:$A${views.Count}");
                Release((object)validation); validation = null;
                // 控制格视觉标记（截图视觉闭环跟进项：数据验证下拉在静态截图中不显眼，加底色+边框）
                foreach (var ctrlAddr in new[] { "C4", "E4", "H4" })
                {
                    dynamic ctrl = sheetBoard.Range[ctrlAddr];
                    ctrl.Interior.Color = 0x00F5E6D0;   // 浅主题金底
                    ctrl.Borders.LineStyle = 1;
                    ctrl.Borders.Color = 0x00B8D8EE;
                    ctrl.HorizontalAlignment = -4108;
                    ctrl.Font.Bold = true;
                    Release((object)ctrl);
                }

                // staging 下拉辅助清单：AF=全部+品类、AG=全部+地区（"全部"首项语义沿用 sales_overview）
                rangeTmp = sheetStaging.Range["AF1"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                for (int i = 0; i < catCount; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 32];
                    rangeTmp.Formula = $"=G{i + 2}";
                    Release((object)rangeTmp);
                }
                rangeTmp = sheetStaging.Range["AG1"]; rangeTmp.Value2 = "全部"; Release((object)rangeTmp);
                for (int i = 0; i < labelCount; i++)
                {
                    rangeTmp = sheetStaging.Cells.Item[i + 2, 33];
                    rangeTmp.Formula = $"=I{i + 2}";
                    Release((object)rangeTmp);
                }

                // ---- 标题横条（圆角矩形，主题色）----
                shape = sheetBoard.Shapes.AddShape(5, 8, 8, 1100, 36);
                shape.Fill.ForeColor.RGB = themeColors.CardBg;
                shape.Line.Visible = 0;
                shape.TextFrame2.TextRange.Text = "销售数据看板（超级看板）";
                shape.TextFrame2.TextRange.Font.Size = 18;
                shape.TextFrame2.TextRange.Font.Bold = -1;
                shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2;   // 居中
                Release(shape); shape = null;

                // ---- KPI 卡 5 张：圆角矩形底 + 标题/值两个文本框（值绑定 staging 镜像格，纯引用 Formula）----
                for (int i = 0; i < 5; i++)
                {
                    float x = 10 + i * 222;
                    dynamic card = sheetBoard.Shapes.AddShape(5, x, 95, 210, 80);
                    card.Fill.ForeColor.RGB = themeColors.CardBg;
                    card.Line.Visible = 0;
                    dynamic sh = card.Shadow;
                    sh.ForeColor.RGB = 0x000000;
                    sh.Transparency = 0.65;
                    sh.Blur = 12;
                    sh.OffsetX = 0; sh.OffsetY = 4;
                    Release(sh);
                    Release(card); card = null;

                    dynamic titleTb = sheetBoard.Shapes.AddTextbox(1, x + 8, 101, 194, 22);
                    dynamic titleDraw = titleTb.DrawingObject;
                    titleDraw.Formula = $"{stagingEsc}$AA${i + 2}";   // 标题随视图换（SB0 P2 纯引用路径）
                    titleTb.TextFrame2.TextRange.Font.Size = 11;
                    titleTb.TextFrame2.TextRange.Font.Bold = -1;
                    titleTb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = themeColors.TitleText;
                    titleTb.Fill.Visible = 0; titleTb.Line.Visible = 0;
                    Release(titleDraw); Release(titleTb); titleDraw = null; titleTb = null;

                    dynamic valTb = sheetBoard.Shapes.AddTextbox(1, x + 8, 125, 194, 44);
                    dynamic valDraw = valTb.DrawingObject;
                    valDraw.Formula = $"{stagingEsc}$Z${i + 2}";      // 值随视图+切片联动
                    valTb.TextFrame2.TextRange.Font.Size = 20;
                    valTb.TextFrame2.TextRange.Font.Bold = -1;
                    valTb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = themeColors.ValueText;
                    valTb.Fill.Visible = 0; valTb.Line.Visible = 0;
                    Release(valDraw); Release(valTb); valDraw = null; valTb = null;
                }

                // ---- 四宫格图表 ----
                // 宫1 月度趋势组合（随视图：镜像 V:X）；宫2 品类排行条形；宫3 地区构成圆环；宫4 月度金额+数量双轴（恒销售全景）
                AddComboChart(sheetBoard, sheetStaging, $"V1:X{monthCount + 1}", 10, 210, 540, 230,
                    "指标趋势（随指标视图切换）", true);
                AddBarChart(sheetBoard, sheetStaging, $"G1:H{catCount + 1}", 570, 210, 540, 230, catHeader + "销售排行");
                AddDoughnutChart(sheetBoard, sheetStaging, $"I1:J{labelCount + 1}", 10, 455, 540, 230, lblHeader + "销售构成");
                AddComboChart(sheetBoard, sheetStaging, $"C1:E{monthCount + 1}", 570, 455, 540, 230,
                    "月度金额与数量（销售全景）", true);

                try { application.Windows.Item(1).DisplayGridlines = false; } catch { }
                sheetStaging.Visible = 0;   // xlSheetHidden
            }
            finally
            {
                Release((object)shape);
                Release((object)validation);
                Release((object)rangeTmp);
            }
        }

        /// <summary>组合图（柱+折线挂次轴，仿 sales_overview 趋势图）。</summary>
        private void AddComboChart(dynamic sheetBoard, dynamic sheetStaging, string srcAddr, double x, double y, double w, double h, string title, bool smoothLine)
        {
            dynamic obj = null; dynamic chart = null; dynamic series = null;
            try
            {
                obj = sheetBoard.ChartObjects().Add(x, y, w, h);
                chart = obj.Chart;
                chart.ChartType = 51;   // xlColumnClustered
                chart.SetSourceData(sheetStaging.Range[srcAddr]);
                series = chart.SeriesCollection();
                if (Convert.ToInt32(series.Count) >= 2)
                {
                    dynamic s2 = series.Item(2);
                    s2.AxisGroup = 2;   // xlSecondary
                    s2.ChartType = 4;   // xlLine
                    if (smoothLine) { try { s2.Smooth = true; } catch { } }
                    Release((object)s2);
                }
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;
            }
            finally
            {
                Release((object)series);
                Release((object)chart);
                Release((object)obj);
            }
        }

        /// <summary>条形排行图。</summary>
        private void AddBarChart(dynamic sheetBoard, dynamic sheetStaging, string srcAddr, double x, double y, double w, double h, string title)
        {
            dynamic obj = null; dynamic chart = null;
            try
            {
                obj = sheetBoard.ChartObjects().Add(x, y, w, h);
                chart = obj.Chart;
                chart.ChartType = 57;   // xlBarClustered
                chart.SetSourceData(sheetStaging.Range[srcAddr]);
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;
            }
            finally
            {
                Release((object)chart);
                Release((object)obj);
            }
        }

        /// <summary>圆环构成图（hole 50%，视频同款）。</summary>
        private void AddDoughnutChart(dynamic sheetBoard, dynamic sheetStaging, string srcAddr, double x, double y, double w, double h, string title)
        {
            dynamic obj = null; dynamic chart = null;
            try
            {
                obj = sheetBoard.ChartObjects().Add(x, y, w, h);
                chart = obj.Chart;
                chart.ChartType = -4120;   // xlDoughnut
                chart.SetSourceData(sheetStaging.Range[srcAddr]);
                try { chart.ChartGroups().Item(1).DoughnutHoleSize = 50; } catch { }
                try { chart.HasLegend = true; } catch { }
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;
            }
            finally
            {
                Release((object)chart);
                Release((object)obj);
            }
        }

        /// <summary>
        /// 透视分析开关（FR-SB-4，Q3-A 拍板 A）：勾选行维度与值字段后生成「可见、可拖拽」的原生数据透视表
        /// （数据源=源明细区域，字段名=源表头）；真透视表构建失败自动降级为静态快照并说明。
        /// 与看板主链完全隔离（AC-4）。
        /// </summary>
        private void BuildPivotAnalysisSheet(
            dynamic workbook, string dataSheetName, int detailRows, string dashboardSheet,
            string sourceSheetNameReal, string sourceAddress,
            List<string> pivotRowFields, List<string> pivotMeasureFields)
        {
            dynamic sheetSrc = null;
            dynamic sheetDst = null;
            dynamic rangeSrc = null;
            dynamic pivotCache = null;
            dynamic pivotTable = null;
            dynamic rangeTmp = null;
            try
            {
                string dstName = "透视分析-" + dashboardSheet;
                // 工具分类审查总复盘修复：合法看板名（≤31 字符）拼接「透视分析-」后可能超长——截断到 31。
                if (dstName.Length > 31) { dstName = dstName.Substring(0, 31); }
                sheetSrc = workbook.Worksheets.Item[sourceSheetNameReal];
                rangeSrc = sheetSrc.Range[sourceAddress];

                sheetDst = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
                sheetDst.Name = dstName;

                if (pivotRowFields.Count > 0 && pivotMeasureFields.Count > 0)
                {
                    // ---- 真·数据透视表（可见可拖拽；多行字段必须 xlTabularForm——压缩式会把字段挤一列）----
                    // 隐藏实例对非活动 sheet 的 Range.Select 会炸（同 ChartObjects 教训）——且 Create(SourceData:=range) 无需 Select
                    pivotCache = workbook.PivotCaches().Create(SourceType: 1, SourceData: rangeSrc);   // xlDatabase
                    pivotTable = pivotCache.CreatePivotTable(
                        TableDestination: "'" + EscapeSheetInner(dstName) + "'!R5C1",
                        TableName: "透视分析表");
                    foreach (var fieldName in pivotRowFields)
                    {
                        dynamic pf = pivotTable.PivotFields(fieldName);
                        pf.Orientation = 1;   // xlRowField
                        Release((object)pf);
                    }
                    foreach (var fieldName in pivotMeasureFields)
                    {
                        dynamic pf = pivotTable.PivotFields(fieldName);
                        pf.Orientation = 4;   // xlDataField
                        pf.Function = -4157;  // xlSum
                        pf.Caption = fieldName + " 汇总";
                        Release((object)pf);
                    }
                    try { pivotTable.RowAxisLayout(1); } catch { }   // xlTabularForm
                    try { pivotTable.ColumnGrand = true; pivotTable.RowGrand = true; } catch { }
                    rangeTmp = sheetDst.Range["A1"];
                    rangeTmp.Value2 = "透视分析（可拖拽字段重新布局；与超级看板主链隔离，修改不影响看板数值）";
                    rangeTmp.Font.Bold = true;
                    Release((object)rangeTmp);
                    ExcelTrace.Info("render_super_dashboard", "pivot rows=" + string.Join("/", pivotRowFields) + " measures=" + string.Join("/", pivotMeasureFields));
                    return;
                }

                // ---- 未勾选字段：静态快照兜底 ----
                rangeTmp = sheetDst.Range["A1"];
                rangeTmp.Value2 = "透视分析（静态快照，与看板隔离）：如需自由拖拽分析，请选中下方数据后插入→数据透视表。";
                rangeTmp.Font.Bold = true;
                Release((object)rangeTmp);
                dynamic used = sheetSrc.UsedRange;
                dynamic target = sheetDst.Range["A3"];
                used.Copy(target);
                Release((object)target);
                Release((object)used);
            }
            catch (Exception ex)
            {
                // 真透视表构建失败 → 自动降级为静态快照（删除半成品 sheet 重建；TryDeleteSheet 判空安全）
                TryDeleteSheet(sheetDst);
                sheetDst = workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
                sheetDst.Name = "透视分析-" + dashboardSheet;
                rangeTmp = sheetDst.Range["A1"];
                rangeTmp.Value2 = "透视分析（静态快照，与看板隔离）——数据透视表构建失败已降级：" + ex.Message;
                rangeTmp.Font.Bold = true;
                Release((object)rangeTmp);
                dynamic used2 = sheetSrc.UsedRange;
                dynamic target2 = sheetDst.Range["A3"];
                used2.Copy(target2);
                Release((object)target2);
                Release((object)used2);
            }
            finally
            {
                Release((object)pivotTable);
                Release((object)pivotCache);
                Release((object)rangeSrc);
                Release((object)rangeTmp);
                Release((object)sheetDst);
                Release((object)sheetSrc);
            }
        }

        /// <summary>读取源明细表头文本（空表头回退缺省名）。</summary>
        private static string ReadHeaderText(dynamic sheetSource, int headerRow, int absoluteColumn, string fallback)
        {
            dynamic cell = null;
            try
            {
                cell = sheetSource.Cells.Item[headerRow, absoluteColumn];
                var text = Convert.ToString(cell.Value2);
                return string.IsNullOrWhiteSpace(text) ? fallback : text.Trim();
            }
            catch
            {
                AppLog.Warn("ExcelDashboardService.SuperBoard", "[LogCover] ExcelDashboardService.SuperBoard.cs:1088  silent-catch covered (no exception variable)");
                return fallback;
            }
            finally
            {
                Release((object)cell);
            }
        }
    }
}
