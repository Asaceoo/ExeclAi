using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    public static class M3ToolSchemas
    {
        public static Dictionary<string, object> Clean()
        {
            return Object(
                Names("address", "operations"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "源数据区域。"),
                Array("operations", "顺序执行的清洗操作。",
                    Object(null,
                        EnumProperty("tool", new[] { "remove_blank_rows", "fix_dates", "fill_blanks" }),
                        Number("column"),
                        Property("date_format", "string", "可选日期格式。"),
                        Property("value", "string", "填充值。"))));
        }

        public static Dictionary<string, object> Filter()
        {
            return Object(
                Names("address", "column", "operator"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "完整数据区域。"),
                Number("column"),
                EnumProperty("operator", new[] { "equals", "not_equals", "contains", "gt", "gte", "lt", "lte", "empty", "not_empty" }),
                Property("value", "string", "筛选值。"),
                Property("target_address", "string", "结果输出左上角地址。"));
        }

        public static Dictionary<string, object> AddSheet()
        {
            return Object(Names("name"),
                Property("name", "string", "新工作表名称。"),
                Property("before_sheet", "string", "在该工作表前插入。"),
                Property("after_sheet", "string", "在该工作表后插入。"));
        }

        public static Dictionary<string, object> RenameSheet()
        {
            return Object(Names("name"),
                Property("sheet", "string", "要重命名的工作表，默认活动工作表。"),
                Property("name", "string", "新名称。"));
        }

        public static Dictionary<string, object> DeleteSheet()
        {
            return Object(Names("sheet"), Property("sheet", "string", "要删除的工作表名称。"));
        }

        public static Dictionary<string, object> CreateTable()
        {
            return Object(Names("address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("address", "string", "转为智能表的区域地址（如 A1:C20），首行默认按表头处理。"),
                Property("name", "string", "智能表名称（可选；缺省由宿主自动命名，如 表1/Table1）。"),
                Property("has_headers", "boolean", "区域首行是否为表头（默认 true）。"),
                Property("table_style", "string", "表样式名称（可选，如 TableStyleMedium9；缺省不动样式）。"));
        }

        public static Dictionary<string, object> InsertImage()
        {
            return Object(Names("image_path", "anchor"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("image_path", "string", "本地图片文件完整路径（png/jpg/gif 等；不链接源文件，随工作簿保存）。"),
                Property("anchor", "string", "图片左上角锚定的单元格地址（如 F2）。"),
                Number("width", "图片宽度（磅，可选；缺省按图片原始尺寸插入）。"),
                Number("height", "图片高度（磅，可选；缺省按图片原始尺寸插入）。"));
        }

        public static Dictionary<string, object> FreezePanes()
        {
            return Object(Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表（冻结为窗口级操作，会切换到该工作表）。"),
                Property("cell", "string", "冻结锚点单元格：其上方所有行与左侧所有列被冻结（如 B2 冻结第 1 行和 A 列；A2 仅冻结行；B1 仅冻结列；A1 无效）。unfreeze 为 true 时可省略。"),
                Property("unfreeze", "boolean", "true 时解除冻结（可省略 cell）。"));
        }

        public static Dictionary<string, object> ManageHyperlinks()
        {
            return Object(Names(),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                EnumProperty("action", new[] { "list", "add", "update", "remove", "remove_all" }),
                Property("address", "string", "目标单元格/区域：add/update 需单个单元格（如 B2）；remove 可为区域（如 A1:A5）；list 可选作过滤；remove_all 不需要。"),
                Property("link_url", "string", "外部网址（add 必填其一；update 可选，与 link_subaddress 互斥）。"),
                Property("link_subaddress", "string", "工作簿内部位置（add 必填其一，如 目录!A1；update 可选，与 link_url 互斥）。"),
                Property("display_text", "string", "显示文本（add/update 可选；缺省保留单元格现有文本）。"),
                Property("tooltip", "string", "悬停提示（add 可选）。"));
        }

        public static Dictionary<string, object> Pivot()
        {
            return Object(Names("source_address", "target_address", "label_column", "value_column"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "含标题行的数据区域；缺省回退到当前多格选区。"),
                Property("target_address", "string", "汇总输出左上角地址；缺省自动放在源区域下方 2 行。"),
                Number("label_column"),
                Number("value_column"));
        }

        public static Dictionary<string, object> Chart()
        {
            return Object(Names("source_address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "图表数据区域（如 \"A2:B10\" 或 \"Sheet1!A2:B10\"；首行为表头、首列为类别）。"),
                EnumProperty("chart_type", new[] { "ColumnClustered", "BarClustered", "Line", "Pie", "Area", "Scatter", "Histogram", "Doughnut", "DoughnutExploded", "Radar", "RadarFilled", "RadarMarkers" }),
                Property("title", "string", "图表标题；缺省使用中文默认标题。契约要求带单位与口径（如\"销售额（万元）· 2025 年 · 不含退货\"）。"),
                EnumProperty("data_labels", new[] { "auto", "on", "off" }),
                Property("anchor", "string", "图表左上角锚点单元格（如 F2）；缺省自动放到数据区右侧。"),
                Number("width", "图表宽度（像素），默认 420。"),
                Number("height", "图表高度（像素），默认 260。"));
        }

        /// <summary>M-D1：双轴组合图（柱形主轴 + 指定系列折线次轴）。</summary>
        public static Dictionary<string, object> ComboChart()
        {
            return Object(Names("source_address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "图表数据区域；首列为类别，需含 ≥2 个数值系列。"),
                Property("secondary_series", "string", "挂次轴并改为折线的系列序号（1 基，如 \"2\" 或 \"2,3\"）；缺省最后一个系列。"),
                Property("title", "string", "图表标题；缺省使用「工作表名 组合分析」。"),
                Property("anchor", "string", "图表左上角锚点单元格；缺省自动放到数据区右侧。"),
                Number("width", "图表宽度（像素），默认 480。"),
                Number("height", "图表高度（像素），默认 280。"));
        }

        /// <summary>M-D1：环形图（含分离型），hole_size 10-90。</summary>
        public static Dictionary<string, object> Doughnut()
        {
            return Object(Names("source_address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "图表数据区域；首列类别 + 单数值列构成占比。"),
                EnumProperty("chart_type", new[] { "doughnut", "doughnutexploded" }),
                Number("hole_size", "环形孔径百分比 10-90，默认 50。"),
                Property("title", "string", "图表标题；缺省使用「工作表名 构成」。"),
                Property("anchor", "string", "图表左上角锚点单元格；缺省自动放到数据区右侧。"),
                Number("width", "图表宽度（像素），默认 360。"),
                Number("height", "图表高度（像素），默认 300。"));
        }

        /// <summary>M-D1：KPI 速览看板（kpi_brief）。</summary>
        public static Dictionary<string, object> DashboardKpiBrief()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（首行表头），如 A1:F500。须先通过数据门槛探测。"),
                Number("label_column", "聚合维度列（1 基，相对区域首列），默认 2。"),
                Number("value_column", "数值列（1 基，相对区域首列），默认 4。"),
                Array("kpis", "KPI 卡定义（最多 8 张）；缺省为总额/笔数/均价/最大单笔四卡。",
                    Object(null,
                        Property("title", "string", "卡片标题，如「销售总额」。"),
                        EnumProperty("aggregation", new[] { "sum", "count", "average", "max", "min" }),
                        Property("format", "string", "数字格式，如 ¥#,##0。"),
                        Property("note", "string", "卡片脚注（可选）。"))),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「KPI看板」；已存在则幂等重建。"),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）；缺省 false 走幂等重生成覆盖旧看板。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }));
        }

        /// <summary>M-D1：数据门槛探测（ReadOnly，生成看板前必跑）。</summary>
        public static Dictionary<string, object> DetectDataReadiness()
        {
            return Object(Names("source_address"),
                Property("sheet", "string", "数据工作表名，默认活动工作表；重生成/其他表激活后必须显式指定。"),
                Property("source_address", "string", "明细数据区域（含表头行），如 A1:F500。"),
                Number("value_column", "数值列（1 基，相对区域首列），默认 4。"),
                Number("sample_rows", "扫描行数上限（大表抽样），默认 500。"));
        }

        /// <summary>M-D2：sales_overview 看板渲染。</summary>
        public static Dictionary<string, object> DashboardSalesOverview()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（首行表头），如 A1:F500。须先通过数据门槛探测。"),
                Property("sheet", "string", "明细数据工作表名，默认活动工作表；重生成/连续生成时必须显式指定（渲染会激活看板表）。"),
                Number("label_column", "地区列（1 基，相对区域首列），默认 1。"),
                Number("category_column", "品类列（1 基），默认 2。"),
                Number("month_column", "月份列（1 基），默认 3。"),
                Number("value_column", "销售额列（1 基），默认 4。"),
                Number("quantity_column", "销量列（次轴折线用，1 基），默认 5。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「销售看板」；已存在则幂等重建。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）；缺省 false 走幂等重生成覆盖旧看板。"),
                Property("auto_slicer", "string", "缺省 true：渲染后自动挂 native 切片器（不可用时回退下拉并说明）；传 false 保持纯下拉联动。"));
        }

        /// <summary>超级看板（需求-超级看板.md v1.0）：三层结构一键生成+指标选择器+extra_kpis。</summary>
        public static Dictionary<string, object> DashboardSuperBoard()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（首行表头），如 A1:F500；内部自动建超级数据表底座（已有则复用）。"),
                Property("sheet", "string", "明细数据工作表名，默认活动工作表。"),
                Number("label_column", "地区/人员列（1 基），默认 1。"),
                Number("category_column", "品类列（1 基），默认 2。"),
                Number("month_column", "月份列（1 基），默认 3。"),
                Number("value_column", "金额列（1 基），默认 4。"),
                Number("quantity_column", "数量列（1 基），默认 5。"),
                Number("profit_column", "利润列（1 基，可选）：传入后启用「利润视图」（利润总额/利润占比/百万利润等）。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「超级看板」；已存在则幂等重建。"),
                Property("super_sheet", "string", "数据底座 sheet 名，默认「超级数据表」。"),
                Property("extra_kpis", "string", "JArray 字符串，派生指标 [{name:\"利润占比\", formula:\"销售总额/销售量\", format:\"0.0%\"}]；formula 操作数限内置基础指标名（销售总额/销售量/成交笔数/平均客单价/最大单笔/利润总额等），≤5 个，非空时追加「自定义视图」。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("pivot_analysis", "string", "传 true 额外生成「透视分析-看板名」sheet；配合 pivot_rows/pivot_measures 生成可拖拽原生数据透视表（字段名=源表头），失败自动降级静态快照；与看板主链隔离。"),
                Property("pivot_rows", "string", "透视行维度，逗号分隔，如「品类,地区」（可用：月份/品类/地区）。pivot_analysis=true 时生效。"),
                Property("pivot_measures", "string", "透视值字段，逗号分隔，如「金额,利润」（可用：金额/数量/利润，求和）。pivot_analysis=true 时生效。"),
                Property("slicer_enabled", "string", "传 true 渲染后挂 native 切片器（原生支持单选/Ctrl 多选/多选模式按钮）；不可用时回退看板内置维度下拉。"),
                Property("auto_rename", "string", "传 true 时看板名重名自动递增共存；缺省 false 幂等重建。"),
                Property("auto_slicer", "string", "预留参数（v1 维度筛选用看板内置三下拉，切片器挂载为后续增强）。"));
        }

        /// <summary>分步工具②（超级看板向导 Step2）：独立生成透视分析 sheet（可拖拽原生透视表）。</summary>
        public static Dictionary<string, object> BuildPivotSheet()
        {
            return Object(Names("source_address", "pivot_rows", "pivot_measures"),
                Property("source_address", "string", "明细数据区域（首行表头），如 A1:F500。"),
                Property("sheet", "string", "明细数据工作表名，默认活动工作表。"),
                Number("label_column", "地区/人员列（1 基），默认 1。"),
                Number("category_column", "品类列（1 基），默认 2。"),
                Number("month_column", "月份/日期列（1 基），默认 3。"),
                Number("value_column", "金额列（1 基），默认 4。"),
                Number("quantity_column", "数量列（1 基），默认 5。"),
                Number("profit_column", "利润列（1 基，可选）。"),
                Property("pivot_rows", "string", "行维度，逗号分隔（可用：月份/品类/地区/销售员，或直接传表头列名）。"),
                Property("pivot_measures", "string", "值字段，逗号分隔（可用：金额/数量/利润，求和聚合）。"),
                Property("pivot_sheet_name", "string", "生成的透视分析 sheet 名，默认「透视分析」。"));
        }

        /// <summary>FR-ED：verify_dashboard_deep 深度核验（只读，六项检查各带期望/实际/依据）。</summary>
        public static Dictionary<string, object> VerifyDashboardDeep()
        {
            return Object(Names("dashboard_sheet"),
                Property("dashboard_sheet", "string", "看板 sheet 名（必填）。"),
                Property("check_pixels", "string", "传 false 跳过图表渲染像素核验（默认开启）。"));
        }

        /// <summary>FR-6.5：screenshot_dashboard 看板截图（只读）。</summary>
        public static Dictionary<string, object> ScreenshotDashboard()
        {
            var schema = Object(Names("dashboard_sheet"),
                Property("dashboard_sheet", "string", "可选：看板工作表名（缺省为活动工作表）。"),
                Property("address", "string", "可选：任意区域地址（如 C13:P25）——用户要求截图当前选区/数据时传此参数。"),
                Property("sheet", "string", "可选：address 模式的目标工作表，默认活动工作表。"),
                Property("chart_name", "string", "可选：指定图表名则只截该图表。"),
                Property("out_dir", "string", "可选：PNG 输出目录，缺省工作簿目录下 _看板截图/（未保存时用 %TEMP%）。"));
            schema["description"] =
                "返回 screenshot_path/route/non_blank_ratio/image_data_uri；image_data_uri 为缩略图，AI 可直接读图视觉自审，发现问题先 diagnose/repair 再重截。";
            return schema;
        }

        /// <summary>FR-ED 看板编辑器：list / theme / 删单图。</summary>
        public static Dictionary<string, object> ListDashboardCharts()
        {
            return Object(Names("dashboard_sheet"),
                Property("dashboard_sheet", "string", "看板 sheet 名（必填）；须为本助手生成的看板。"));
        }

        public static Dictionary<string, object> EditDashboardTheme()
        {
            return Object(Names("dashboard_sheet", "theme"),
                Property("dashboard_sheet", "string", "看板 sheet 名（必填）。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }));
        }

        public static Dictionary<string, object> DeleteDashboardChart()
        {
            return Object(Names("dashboard_sheet", "chart_name"),
                Property("dashboard_sheet", "string", "看板 sheet 名（必填）。"),
                Property("chart_name", "string", "图表名（必填），逐字取自 list_dashboard_charts 清单。"));
        }

        /// <summary>FR-6.2：suggest_chart_mapping 图表选型建议（只读启发式）。</summary>
        public static Dictionary<string, object> ChartMappingSuggestion()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "数据区域（首行表头）；建议先让用户选中数据再调用。"),
                Property("sheet", "string", "数据工作表名，默认活动工作表。"),
                Number("max_suggestions", "最多返回建议条数，默认 5。"));
        }

        /// <summary>M-D2：特大值截断柱。</summary>
        public static Dictionary<string, object> TruncatedColumn()
        {
            return Object(Names("source_address"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "数据区域：首列类别 + 1 数值列。"),
                Number("truncate_factor", "截断阈值倍数（超出均值 N 倍判定为特大值），默认 3.0。"),
                Property("title", "string", "图表标题。"),
                Property("anchor", "string", "锚点单元格；缺省放数据区右侧。"),
                Number("width", "宽度像素，默认 460。"),
                Number("height", "高度像素，默认 280。"));
        }

        /// <summary>M-D3：看板联动设置（native 切片器 / cell 下拉）。</summary>
        public static Dictionary<string, object> SetLinkage()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（含表头行），与生成看板时的数据源一致。"),
                Property("sheet", "string", "源数据工作表名；建议显式指定（render_dashboard 后活动表会漂移到看板 sheet）。"),
                EnumProperty("mode", new[] { "native", "cell" }),
                Number("field_column", "筛选维度列（1 基，相对区域首列），默认 1。"),
                Number("value_column", "数值列（1 基，相对区域首列），默认 4。"),
                Property("dashboard_sheet", "string", "目标看板 sheet 名，默认「销售看板」。"),
                Property("control_cell", "string", "cell 模式的控制单元格，默认 B2。"));
        }

        /// <summary>M-D3：图表版本徽章探测（只读，AC-4）。</summary>
        public static Dictionary<string, object> CheckChartVersion()
        {
            return Object(Names("chart_type"),
                EnumProperty("chart_type", new[] { "treemap", "sunburst", "funnel", "histogram", "pareto", "waterfall", "boxwhisker" }));
        }

        /// <summary>超级数据表（看板数据底座，2026-09-12）：明细规范化+派生汇总一张表。</summary>
        public static Dictionary<string, object> SuperTable()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（含表头行）：维度列+数值列长表，一行一条记录。"),
                Property("sheet", "string", "源数据工作表名，默认活动工作表。"),
                Number("label_column", "地区/筛选维度列（1 基，相对区域首列），默认 1。"),
                Number("category_column", "品类列（1 基），默认 2。"),
                Number("month_column", "月份列（1 基），默认 3。"),
                Number("value_column", "金额列（1 基），默认 4。"),
                Number("quantity_column", "数量列（1 基），默认 5。"),
                Property("super_sheet", "string", "超级数据表 sheet 名，默认「超级数据表」；同名旧超级表幂等重建，被其他表占用时报错。"));
        }

        /// <summary>M-D3：dim_cross 多维交叉看板。</summary>
        public static Dictionary<string, object> DashboardDimCross()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "明细数据区域（含表头行）。须先通过数据门槛探测。"),
                Property("sheet", "string", "源数据工作表名，默认活动工作表。"),
                Number("row_dim_column", "行维度列（1 基，相对区域首列），默认 1。"),
                Number("col_dim_column", "列维度列（1 基），默认 2；不能与行维度相同。"),
                Number("value_column", "数值列（1 基），默认 4。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「多维交叉看板」；已存在则幂等重建。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）；缺省 false 走幂等重生成覆盖旧看板。"));
        }

        /// <summary>M-D3：一键删除看板（Dangerous，AC-10 回读断言防误删）。</summary>
        public static Dictionary<string, object> DeleteDashboard()
        {
            return Object(Names("dashboard_sheet"),
                Property("dashboard_sheet", "string", "要删除的看板 sheet 名，默认「销售看板」。删除前回读断言看板结构，普通工作表会被拒绝。"));
        }

        /// <summary>M-D3 Q2：create_chart_advanced 技巧型图表族单入口。</summary>
        public static Dictionary<string, object> ChartAdvanced()
        {
            return Object(Names("source_address", "type"),
                Property("sheet", "string", "目标工作表，默认活动工作表。"),
                Property("source_address", "string", "数据区域：首列类别标签 + 数值列（各 type 所需列不同：bullet/slider/gantt 需 2 数值列，histogram 仅需 1 列原始观测值、无需类别，其余 1 列）。"),
                EnumProperty("type", new[] { "bullet", "waffle", "slider", "funnel", "butterfly", "gantt", "waterfall", "treemap", "sunburst", "pareto", "histogram" }),
                Property("title", "string", "图表标题，缺省按工作表名生成。"),
                Property("anchor", "string", "锚点单元格；缺省放数据区右侧。"),
                Number("width", "宽度像素，默认 520。"),
                Number("height", "高度像素，默认 320。"));
        }

        /// <summary>FR-5.1：pnl_brief 财务损益看板。</summary>
        public static Dictionary<string, object> DashboardPnlBrief()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "数据区域三列：科目 | 金额 | 类型（revenue/cost），首行表头。"),
                Property("sheet", "string", "源数据工作表名，默认活动工作表。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「财务损益看板」；已存在则幂等重建。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）。"));
        }

        /// <summary>FR-5.1：target_track 目标达成看板。</summary>
        public static Dictionary<string, object> DashboardTargetTrack()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "数据区域三列：指标名 | 实际值 | 目标值，首行表头。"),
                Property("sheet", "string", "源数据工作表名，默认活动工作表。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「目标达成看板」；已存在则幂等重建。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）。"));
        }

        /// <summary>FR-5.1：rank_dynamic 排名动态看板（首批六类收口）。</summary>
        public static Dictionary<string, object> DashboardRankDynamic()
        {
            return Object(Names("source_address"),
                Property("source_address", "string", "数据区域三列：类别 | 时间 | 数值，首行表头（列位可传参重指）。"),
                Property("sheet", "string", "源数据工作表名，默认活动工作表。"),
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「排名动态看板」；已存在则幂等重建。"),
                Number("top_n", "Top N 排行条数，默认 10（1~100）。"),
                Number("category_column", "类别列（1 基，相对区域首列），默认 1。"),
                Number("time_column", "时间列（1 基），默认 2。"),
                Number("value_column", "数值列（1 基），默认 3。"),
                EnumProperty("theme", new[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" }),
                Property("auto_rename", "string", "传 true 时看板名重名自动「名 2/名 3」递增共存（多看板）。"));
        }

        /// <summary>需求①：get_dashboard_sources 看板数据源识别（只读，无参数）。</summary>
        public static Dictionary<string, object> DashboardSources()
        {
            return Object(null);
        }

        /// <summary>M-D1：看板自检（fail-closed）。</summary>
        public static Dictionary<string, object> VerifyDashboard()
        {
            return Object(null,
                Property("dashboard_sheet", "string", "看板 sheet 名，默认「KPI看板」。"),
                Property("staging_sheet", "string", "辅助汇总表名；缺省为 _<看板名>_calc。"),
                Number("expected_kpi_count", "预期 KPI 数量下限；辅助区数量少于该值判 mismatch。"));
        }

private static string[] Names(params string[] names)
        {
            return names;
        }

        private static Dictionary<string, object> Object(string[] required, params Dictionary<string, object>[] properties)
        {
            var map = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                map[Convert.ToString(property["name"])] = property;
            }

            var schema = new Dictionary<string, object>
            {
                ["type"] = "object",
                ["properties"] = map
            };

            if (required != null && required.Length > 0)
            {
                schema["required"] = required;
            }

            return schema;
        }

        private static Dictionary<string, object> Array(string name, string description, Dictionary<string, object> items)
        {
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "array",
                ["description"] = description,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> Property(string name, string type, string description = null)
        {
            var result = new Dictionary<string, object> { ["name"] = name, ["type"] = type };
            if (description != null)
            {
                result["description"] = description;
            }

            return result;
        }

        private static Dictionary<string, object> Number(string name)
        {
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "integer",
                ["minimum"] = 1,
                ["description"] = "从 1 开始的列号。"
            };
        }

        private static Dictionary<string, object> Number(string name, string description)
        {
            return new Dictionary<string, object>
            {
                ["name"] = name,
                ["type"] = "integer",
                ["minimum"] = 1,
                ["description"] = description
            };
        }

        private static Dictionary<string, object> EnumProperty(string name, string[] values)
        {
            return new Dictionary<string, object> { ["name"] = name, ["type"] = "string", ["enum"] = values };
        }
    }
}



