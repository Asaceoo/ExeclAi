using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelToolHostRegistration
    {
        public static void RegisterM3Tools(
            ToolRegistry registry,
            Func<Dictionary<string, object>, ToolResult> clean,
            Func<Dictionary<string, object>, ToolResult> filter,
            ExcelStructureService structure,
            ExcelPivotChartService pivotChart,
            ExcelDashboardService dashboard = null)
        {
            registry.Register(new ToolDescriptor("clean_excel_data", "清洗指定区域：去空行、修日期、填充空白。", RiskLevel.NormalWrite, clean, M3ToolSchemas.Clean()));
            registry.Register(new ToolDescriptor("filter_excel_data", "按列条件筛选数据，并把匹配结果写入目标区域。", RiskLevel.NormalWrite, filter, M3ToolSchemas.Filter()));
            registry.Register(new ToolDescriptor("add_excel_worksheet", "新增工作表，可指定名称和插入位置。", RiskLevel.StructuralChange, args => Execute(() => structure.AddWorksheet(args)), M3ToolSchemas.AddSheet()));
            registry.Register(new ToolDescriptor("rename_excel_worksheet", "重命名工作表。", RiskLevel.StructuralChange, args => Execute(() => structure.RenameWorksheet(args)), M3ToolSchemas.RenameSheet()));
            registry.Register(new ToolDescriptor("delete_excel_worksheet", "删除指定工作表。", RiskLevel.Dangerous, args => Execute(() => structure.DeleteWorksheet(args)), M3ToolSchemas.DeleteSheet()));
            registry.Register(new ToolDescriptor("create_excel_table", "把区域转为智能表（ListObject）：结构化引用、排序筛选、透选取数的基石。WPS 表格同样支持。", RiskLevel.StructuralChange, args => Execute(() => structure.CreateTable(args)), M3ToolSchemas.CreateTable()));
            registry.Register(new ToolDescriptor("insert_image", "在指定锚点单元格左上角插入本地图片（随工作簿保存，不链接源文件）。WPS 表格同样支持。", RiskLevel.StructuralChange, args => Execute(() => structure.InsertImage(args)), M3ToolSchemas.InsertImage()));
            registry.Register(new ToolDescriptor("freeze_panes", "冻结/解冻窗口窗格：锚点单元格上方所有行与左侧所有列被冻结（滚动表头/首列不动的标准做法）。", RiskLevel.StructuralChange, args => Execute(() => structure.FreezePanes(args)), M3ToolSchemas.FreezePanes()));
            registry.Register(new ToolDescriptor("manage_hyperlinks", "超链接管理：list 枚举整表超链接（可按 address 过滤）/add 新建（外部网址 link_url 或内部位置 link_subaddress 二选一）/update 修改/remove 删除（支持区域）/remove_all 清空整表。写动作前置工作表保护检查。", RiskLevel.NormalWrite, args => Execute(() => structure.ManageHyperlinks(args)), M3ToolSchemas.ManageHyperlinks()));
            registry.Register(new ToolDescriptor("create_excel_pivot_summary", "按指定标签列和数值列生成汇总透视结果。", RiskLevel.NormalWrite, args => Execute(() => pivotChart.CreatePivot(args)), M3ToolSchemas.Pivot()));
            registry.Register(new ToolDescriptor("create_excel_chart", "根据区域创建基础图表；柱/条/折线/雷达在系列≤3 且每系列≤12 点时自动打开数据标签（饼/环显示百分比），data_labels=on/off 可强制覆盖，返回值 data_labels 字段回显实际状态。", RiskLevel.StructuralChange, args => Execute(() => pivotChart.CreateChart(args)), M3ToolSchemas.Chart()));
            registry.Register(new ToolDescriptor("create_combo_chart", "创建双轴组合图（柱形主轴 + 指定系列折线次轴）。适用于量纲不同的双指标趋势对比（如销售额柱形 + 销量折线）。", RiskLevel.StructuralChange, args => Execute(() => pivotChart.CreateComboChart(args)), M3ToolSchemas.ComboChart()));
            registry.Register(new ToolDescriptor("create_doughnut", "创建环形图（可选分离型，孔径可调 10-90）。适用于构成占比展示（≤5 类效果最佳）。", RiskLevel.StructuralChange, args => Execute(() => pivotChart.CreateDoughnut(args)), M3ToolSchemas.Doughnut()));
            registry.Register(new ToolDescriptor("create_truncated_column", "特大值截断柱：某值远超其余（>均值 N 倍）时截断显示并在柱顶标注真实值，避免其余分类被压扁。", RiskLevel.StructuralChange, args => Execute(() => pivotChart.CreateTruncatedColumn(args)), M3ToolSchemas.TruncatedColumn()));
            if (dashboard != null)
            {
                registry.Register(new ToolDescriptor("create_dashboard_kpi_brief", "创建 KPI 速览看板（kpi_brief 模板）：隐藏 SUMIFS 辅助区 + 4-8 张 KPI 卡片（数值+环比位+IFERROR 兜底）。源数据只读；重生成幂等覆盖。", RiskLevel.StructuralChange, args => Execute(() => dashboard.CreateKpiBrief(args)), M3ToolSchemas.DashboardKpiBrief()));
                registry.Register(new ToolDescriptor("verify_dashboard", "看板自检：看板 sheet 与隐藏辅助区是否存在、辅助区 KPI 重算值是否全部为有效数值（守恒对账）。fail-closed。", RiskLevel.ReadOnly, args => Execute(() => dashboard.VerifyDashboard(args)), M3ToolSchemas.VerifyDashboard()));
                registry.Register(new ToolDescriptor("detect_dashboard_data_readiness", "数据门槛探测（生成看板前必跑）：合计行/合并单元格/文本型数字/空列硬拦截，ready=false 时给可读理由与清洗建议。只读不改数据。", RiskLevel.ReadOnly, args => Execute(() => dashboard.DetectDataReadiness(args)), M3ToolSchemas.DetectDataReadiness()));
                registry.Register(new ToolDescriptor("render_dashboard", "渲染 sales_overview 看板（视频同款）：地区下拉联动 + 4 KPI 卡 + 双轴趋势 + 品类环形，一次调用完整生成。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderSalesOverview(args)), M3ToolSchemas.DashboardSalesOverview()));
                registry.Register(new ToolDescriptor("render_super_dashboard", "渲染超级看板（三层结构一键生成）：自动建超级数据表底座 + 汇总透视层 + 固定四宫格看板（5 张形状 KPI 卡随指标视图切换、品类/地区下拉联动、透视分析可选）。支持 extra_kpis 派生指标（如利润占比=销售总额/销售量）。生成后须调 verify_dashboard_deep 核验与 screenshot_dashboard 截图。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderSuperDashboard(args)), M3ToolSchemas.DashboardSuperBoard()));
                registry.Register(new ToolDescriptor("build_pivot_sheet", "分步生成透视分析 sheet（超级看板向导 Step2）：按勾选行维度与值字段构建可见、可拖拽的原生数据透视表，失败降级静态快照。与看板主链隔离。", RiskLevel.StructuralChange, args => Execute(() => dashboard.BuildPivotSheet(args)), M3ToolSchemas.BuildPivotSheet()));
                registry.Register(new ToolDescriptor("set_linkage", "为已有看板设置联动筛选：native 模式创建切片器（多选/直观），cell 模式创建下拉（兼容 Excel 2013）。须先 render_dashboard 生成看板。", RiskLevel.StructuralChange, args => Execute(() => dashboard.SetLinkage(args)), M3ToolSchemas.SetLinkage()));
                registry.Register(new ToolDescriptor("check_chart_version", "图表版本徽章探测（只读）：检测当前 Excel 是否支持版本敏感图表（树状图/旭日图/漏斗图需 2021+，直方图/排列图/瀑布图需 2016+），不支持时给替代建议。", RiskLevel.ReadOnly, args => Execute(() => dashboard.CheckChartVersion(args)), M3ToolSchemas.CheckChartVersion()));
                registry.Register(new ToolDescriptor("get_dashboard_sources", "看板数据源识别（只读）：列出当前工作簿内全部看板（看板 sheet + 隐藏辅助表 + 明细数据源引用）。AI 理解看板作用范围、回答「这个看板的数据来自哪里」时调用。", RiskLevel.ReadOnly, args => Execute(() => dashboard.GetDashboardSources(args)), M3ToolSchemas.DashboardSources()));
                registry.Register(new ToolDescriptor("render_dim_cross", "渲染多维交叉看板：行×列维度交叉矩阵 + 双下拉联动（行/列筛选，(全部) 支持全量视图）+ 矩阵条形图。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderDimCross(args)), M3ToolSchemas.DashboardDimCross()));
                registry.Register(new ToolDescriptor("create_super_table", "超级数据表（看板数据底座）：把明细源规范化为一张可见 sheet——明细长表（类型规整：月份锁文本、文本数字数值化）+ 派生汇总区（KPI/月度/品类/地区，同表 SUMIF 公式），看板渲染与多切片器共享同一数据源。", RiskLevel.StructuralChange, args => Execute(() => dashboard.CreateSuperTable(args)), M3ToolSchemas.SuperTable()));
                registry.Register(new ToolDescriptor("delete_dashboard", "一键删除看板：删除看板 sheet 与隐藏辅助表并清理联动切片器。删除前回读断言看板结构，普通工作表会被拒绝（防误删）。", RiskLevel.Dangerous, args => Execute(() => dashboard.DeleteDashboard(args)), M3ToolSchemas.DeleteDashboard()));
                registry.Register(new ToolDescriptor("create_chart_advanced", "技巧型图表族单入口（Excel 2013+）：type=bullet 子弹图/waffle 华夫图/slider 滑珠图/funnel 漏斗图/butterfly 蝴蝶图/gantt 甘特图；2016+ 原生图表：waterfall 瀑布图/treemap 树状图/sunburst 旭日图/pareto 排列图。", RiskLevel.StructuralChange, args => Execute(() => pivotChart.CreateChartAdvanced(args)), M3ToolSchemas.ChartAdvanced()));
                // verify_chart 已由 ExcelThirdPhaseRegistration 注册（ExcelChartService.VerifyChart，含像素级核验），
                // 新图表类型支持通过扩展其 ChartTypeMap/NormalizeChartFamily 提供，不再重复注册。
                registry.Register(new ToolDescriptor("render_pnl_brief", "财务损益看板：收入/成本/费用 KPI 卡 + 科目明细数据条 + 净利润。数据三列：科目/金额/类型（revenue/cost）。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderPnlBrief(args)), M3ToolSchemas.DashboardPnlBrief()));
                registry.Register(new ToolDescriptor("render_target_track", "目标达成看板：总体达成率/达标项数/最佳指标 KPI 卡 + 逐指标完成率数据条。数据三列：指标/实际/目标。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderTargetTrack(args)), M3ToolSchemas.DashboardTargetTrack()));
                registry.Register(new ToolDescriptor("render_rank_dynamic", "排名动态看板：时间下拉联动 + 榜首/Top N 合计/全量合计 KPI 卡 + Top N 排行条形图。数据三列：类别/时间/数值。", RiskLevel.StructuralChange, args => Execute(() => dashboard.RenderRankDynamic(args)), M3ToolSchemas.DashboardRankDynamic()));
                registry.Register(new ToolDescriptor("suggest_chart_mapping", "图表选型建议（只读启发式）：分析选区各列画像（数值比/去重基数/日期特征/负值），给出图表或看板工具推荐清单，每条附理由与数据要求。生成图表/看板前先调用，把建议复述给用户并确认数据源/行/列映射。", RiskLevel.ReadOnly, args => Execute(() => dashboard.SuggestChartMapping(args)), M3ToolSchemas.ChartMappingSuggestion()));
                registry.Register(new ToolDescriptor("list_dashboard_charts", "看板图表清单（只读）：枚举看板上全部图表（名称/类型/锚点/系列数）与当前主题。删图/换主题前必调。", RiskLevel.ReadOnly, args => Execute(() => dashboard.ListDashboardCharts(args)), M3ToolSchemas.ListDashboardCharts()));
                registry.Register(new ToolDescriptor("edit_dashboard_theme", "看板换主题：只重刷卡片底色/字色等特征配色，不动数值、公式与图表（支持销售概览/KPI 摘要看板）。", RiskLevel.StructuralChange, args => Execute(() => dashboard.EditDashboardTheme(args)), M3ToolSchemas.EditDashboardTheme()));
                registry.Register(new ToolDescriptor("delete_dashboard_chart", "删单图：删除看板上指定图表对象（chart_name 逐字取自 list_dashboard_charts），强制数值守恒断言。切片器/下拉清理走 delete_dashboard。", RiskLevel.StructuralChange, args => Execute(() => dashboard.DeleteDashboardChart(args)), M3ToolSchemas.DeleteDashboardChart()));
                registry.Register(new ToolDescriptor("verify_dashboard_deep", "看板深度核验（只读）：六项检查各带期望/实际/依据——看板与辅助表存在、辅助表 KPI 重算全数值、看板卡片显示值守恒（显示值==辅助表值）、联动控件、数值守恒对账（全量视图 KPI==源列直和）、图表渲染像素核验。最终回复必须逐项转述。", RiskLevel.ReadOnly, args => Execute(() => dashboard.VerifyDashboardDeep(args)), M3ToolSchemas.VerifyDashboardDeep()));
            registry.Register(new ToolDescriptor("screenshot_dashboard", "截图工具（只读）：把看板整体、指定图表或任意选区（address）导出 PNG，返回文件路径与像素统计，供用户直接查看视觉效果。仅用于「看板整体预览」场景（含图表/形状的版面）；生成或修改看板后调用并把路径告知用户。普通区域看图请优先用 render_range_image（只读渲染，不碰剪贴板、支持 image 内容块，多模态模型可直接看）；需要像素级外观时用 take_screenshot（写入级）。三个工具不要重复调用。", RiskLevel.ReadOnly, args => Execute(() => dashboard.ScreenshotDashboard(args)), M3ToolSchemas.ScreenshotDashboard()));
            }
        }

        /// <summary>P0-2 L2 通用结构操作层（对齐 OfficeCLI L2 DOM）：统一元素模型六操作。</summary>
        public static void RegisterElementOps(
            ToolRegistry registry,
            ExcelElementOps elementOps,
            ExcelDraftService draftService)
        {
            registry.Register(new ToolDescriptor("query_elements", "按类型枚举结构元素，列出名称与属性。element_type 取值：sheets（工作表）/ shapes（形状）/ charts（图表）/ usedrange_rows（已用区域）。发现未知结构时先调用。", RiskLevel.ReadOnly, args => Execute(() => elementOps.QueryElements(args)), ElementOpSchemas.QueryElements()));
            registry.Register(new ToolDescriptor("get_element", "读取任意区域/单元格的值与公式（values 二维矩阵）。", RiskLevel.ReadOnly, args => Execute(() => elementOps.GetElement(args)), ElementOpSchemas.GetElement()));
            registry.Register(new ToolDescriptor("set_element", "设置单元格的值（单格，before/after 回显）。仅支持标量值；写区域/二维矩阵请用 write_excel_range（values 参数）。", RiskLevel.NormalWrite, args => Execute(() => elementOps.SetElement(args)), ElementOpSchemas.SetElement()));
            registry.Register(new ToolDescriptor("add_element", "新增元素：仅 textbox 文本框（可指定 top/left/width/height/text）。图表请用 create_excel_chart（基础）/ create_chart_advanced（技巧型）。", RiskLevel.StructuralChange, args => Execute(() => elementOps.AddElement(args)), ElementOpSchemas.AddElement()));
            registry.Register(new ToolDescriptor("remove_element", "删除元素：kind=shape（形状）或 chart（图表），按 name 删除（先用 query_elements 查名）。不可逆破坏，危险级需确认。整表删除请用 delete_excel_worksheet。", RiskLevel.Dangerous, args => Execute(() => elementOps.RemoveElement(args)), ElementOpSchemas.RemoveElement()));
            registry.Register(new ToolDescriptor("move_element", "移动元素：shape 或 chart（按名称改 top/left）。", RiskLevel.StructuralChange, args => Execute(() => elementOps.MoveElement(args)), ElementOpSchemas.MoveElement()));
            registry.Register(new ToolDescriptor("start_draft", "草稿模式第一步：复制源表为隐藏草稿副本（返回 draft_sheet 名）。后续写操作必须把 draft_sheet 当 sheet 参数传给 write/set 等工具；改完调 diff_draft 看差异，再 apply_draft 采纳或 discard_draft 放弃。", RiskLevel.StructuralChange, args => Execute(() => draftService.StartDraft(args)), ElementOpSchemas.DraftStart()));
            registry.Register(new ToolDescriptor("diff_draft", "草稿差异预览：对比草稿与源表的单元格级变更列表，供人工审阅。依赖 start_draft（需 --allow-write 模式）。", RiskLevel.ReadOnly, args => Execute(() => draftService.DiffDraft(args)), ElementOpSchemas.DraftDiff()));
            registry.Register(new ToolDescriptor("apply_draft", "采纳草稿：把草稿变更整表落盘到源工作表并删除草稿（危险级：需确认门，先 diff_draft 给用户看差异）。", RiskLevel.Dangerous, args => Execute(() => draftService.ApplyDraft(args)), ElementOpSchemas.DraftApply()));
            registry.Register(new ToolDescriptor("discard_draft", "放弃草稿：删除草稿不做任何修改。", RiskLevel.StructuralChange, args => Execute(() => draftService.DiscardDraft(args)), ElementOpSchemas.DraftDiscard()));
        }

        /// <summary>P1-3 外部数据源接入（对齐 Data-Analysis-Agent 多源）：只读源→临时 sheet。</summary>
        public static void RegisterExternalData(
            ToolRegistry registry,
            ExcelElementOps elementOps,
            Action<string, object, string, object, object> recordStructuralUndo = null)  // M27R3：结构撤销录制委托（宿主注入）
        {
            // FIX（M27R3 评审）：import_csv 建的是新工作表——必须留结构撤销（原只回 affected_ranges 无 operation_id，AI 无法 undo）。
            registry.Register(new ToolDescriptor("import_csv", "导入 CSV 文件到新工作表（自动探测编码 BOM→UTF8→GBK，来自 TextEncodingDetector）。", RiskLevel.StructuralChange, args => Execute(() =>
            {
                var importResult = elementOps.ImportCsv(args);
                if (importResult != null && importResult.Ok)
                {
                    try
                    {
                        string importedSheet = null;
                        try { importedSheet = Convert.ToString(((dynamic)importResult.Data).sheet); } catch { }
                        if (!string.IsNullOrWhiteSpace(importedSheet))
                        {
                            recordStructuralUndo?.Invoke("import_csv", args, "sheet-insert", new { sheet = importedSheet }, new { sheets = new[] { importedSheet } });
                        }
                    }
                    catch (Exception undoEx) { AppLog.Warn("host", "import_csv 结构撤销录制失败（导入本身已成功）: " + undoEx.Message); }
                }
                return importResult;
            }), ElementOpSchemas.ImportCsv()));
        }

        private static ToolResult Execute(Func<ToolResult> action)
        {
            try
            {
                return action();
            }
            catch (ArgumentException aex)
            {
                // 参数校验错误不是 COM 错误：独立错误码引导 AI 核对参数名/格式后重试
                //（真机实证：缺 source_address 曾报 EXCEL_COM_ERROR，误导 AI 以为是 Excel 侧故障）。
                // 与 ExcelToolHost.Compensable 的 M16 收敛口径对齐。
                return new ToolResult { Ok = false, ErrorCode = "INVALID_ARGS", Data = "参数校验未通过：" + aex.Message };
            }
            catch (KeyNotFoundException keyEx)
            {
                // Y4 修复：附可用表名指引，AI 拿到即可自愈，省一轮 get_excel_sheet_list 往返。
                var sheetHint = string.Empty;
                try
                {
                    // keyEx.TargetSite 不可靠，此处从当前宿主实例拿不到 application（静态方法）——
                    // 表名清单由调用方上下文可得时补充；拿不到时保持原语义只加通用指引。
                    sheetHint = "。可调用 get_excel_sheet_list 查询当前工作簿可用表名后重试";
                }
                catch { }
                return new ToolResult { Ok = false, ErrorCode = "INVALID_SHEET", Data = keyEx.Message + sheetHint };
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return new ToolResult { Ok = false, ErrorCode = "EXCEL_COM_ERROR", Data = ex.Message };
            }
        }
    }
}
