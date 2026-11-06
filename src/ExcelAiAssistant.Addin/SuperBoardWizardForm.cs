using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Addin
{
    /// <summary>
    /// 超级看板三步向导（v3.7.1.54 布局成熟方案重构）：
    /// - 布局弃用全部绝对坐标（new Point 清零），改微软官方自适应三件套：
    ///   TableLayoutPanel 百分比行列 + Dock + FlowLayoutPanel——任意窗口尺寸与 DPI 下不遮挡、不裁切；
    ///   三层保险：TLP 百分比 / Dock=Fill / 面板 AutoScroll（内容超出出滚动条，控件永不丢失）。
    /// - 底部说明区 = 矢量示意图（WizardDiagram 随步骤切换）+ 双层文字说明（通俗/专业/怎么操作）。
    /// - 全控件双层 tooltip（【是什么】【通俗讲】【怎么操作】）；全路径日志打点（切步/生成/一键链/失败分支）。
    /// 三步：①数据表（create_super_table）②透视表（build_pivot_sheet）③超级看板（render_super_dashboard）。
    /// </summary>
    public sealed class SuperBoardWizardForm : Form
    {
        private readonly IReadOnlyList<string> _columnNames;
        private readonly Func<string, Dictionary<string, object>, ToolResult> _executeTool;
        private readonly string _logPrefix = "超级看板向导";
        private readonly ToolTip _wizardTip = new ToolTip { InitialDelay = 300, ReshowDelay = 80, AutoPopDelay = 8000 };

        private TextBox _addressBox;
        private ComboBox _cboLabel, _cboCategory, _cboMonth, _cboValue, _cboQuantity, _cboProfit;
        private DataGridView _gridExtra;
        private CheckBox _chkSlicer, _chkAutoRename, _chkSkipPivot;
        private ComboBox _cboTheme;
        private CheckedListBox _clbPivotRows, _clbPivotMeasures, _clbPivotVals;
        private FlowLayoutPanel _flowBase;

        private readonly Panel[] _stepPanels = new Panel[3];
        private readonly Label[] _stepIndicators = new Label[3];
        private readonly Label[] _stepStatus = new Label[3];
        private readonly Button[] _stepButtons = new Button[3];
        private readonly string[] _stepDescriptions = new string[3];
        private TableLayoutPanel _root;
        private Panel _stepHost;
        private WizardDiagram _diagram;
        private Label _stepDesc;
        private int _currentStep;

        private bool _step1Done, _step2Done;
        private string _superSheetName = "超级数据表";
        private string _pivotSheetName = "透视分析";
        private string _sourceSheet, _sourceRange;

        /// <summary>执行器：RibbonCommandService 注入（工具名 + 实参 → 结果）。</summary>
        public SuperBoardWizardForm(
            string sheetName, string address, IReadOnlyList<string> columnNames,
            Func<string, Dictionary<string, object>, ToolResult> executeTool)
        {
            _columnNames = columnNames ?? new List<string>();
            _executeTool = executeTool ?? throw new ArgumentNullException(nameof(executeTool));
            _sourceSheet = sheetName ?? "";
            _sourceRange = address ?? "";

            Text = "超级看板 · 三步向导";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            MinimumSize = new Size(640, 760);
            ClientSize = new Size(640, 760);
            Font = new Font("Microsoft YaHei UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;

            BuildRoot();
            BuildStepIndicator();
            BuildStep1();
            BuildStep2();
            BuildStep3();
            BuildBottomArea();
            BuildNavButtons();
            ShowStep(0);
            DiagnosticsLog.Write(_logPrefix + " 打开(TLP自适应布局) columnCount=" + _columnNames.Count + " addr=" + sheetName + "!" + address);
            AppLog.Info(_logPrefix, "向导打开 step=1/3 列数=" + _columnNames.Count);
        }

        /// <summary>根布局：单列四行 TableLayoutPanel（指示条/步骤区/图示+说明/导航），Dock=Fill 全自适应。</summary>
        private void BuildRoot()
        {
            _root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(6, 2, 6, 2), Margin = new Padding(0) };
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 步骤指示条
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // 三步表单容器
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 图示 + 说明区
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 导航按钮
            Controls.Add(_root);

            _stepHost = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 2) };
            _root.Controls.Add(_stepHost, 0, 1);
        }

        /// <summary>顶部步骤指示条：①数据表 → ②透视表 → ③超级看板（当前步高亮、已完成打 ✓）。</summary>
        private void BuildStepIndicator()
        {
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 2) };
            string[] names = { "① 数据表", "② 透视表", "③ 超级看板" };
            for (int i = 0; i < 3; i++)
            {
                _stepIndicators[i] = new Label
                {
                    Text = names[i],
                    Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(120, 120, 120),
                    AutoSize = true,
                    Margin = new Padding(4, 4, 2, 2)
                };
                bar.Controls.Add(_stepIndicators[i]);
                if (i < 2)
                {
                    bar.Controls.Add(new Label { Text = "──▶", AutoSize = true, ForeColor = Color.Silver, Margin = new Padding(2, 6, 14, 2) });
                }
            }
            _root.Controls.Add(bar, 0, 0);
        }

        /// <summary>三个同区域步骤面板（Dock=Fill + Visible 切换，AutoScroll 兜底）。标题由各 BuildStepN 作为内容首行加入。</summary>
        private Panel NewStepPanel(int index, string title, string description)
        {
            _stepDescriptions[index] = description;
            var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Visible = index == 0, Margin = new Padding(0) };
            _stepPanels[index] = panel;
            _stepHost.Controls.Add(panel);
            return panel;
        }

        private static Label MakeTitleLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 4)
            };
        }

        /// <summary>步骤内容单列纵排容器（Dock=Top + AutoSize，行全部 AutoSize）。</summary>
        private static TableLayoutPanel NewStack()
        {
            var tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Margin = new Padding(0) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            return tlp;
        }

        /// <summary>向内容容器追加一行（Auto 行高）。</summary>
        private static void AddRow(TableLayoutPanel tlp, Control c, Padding? margin = null)
        {
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowCount++;
            c.Margin = margin ?? new Padding(0, 2, 0, 2);
            tlp.Controls.Add(c, 0, tlp.RowCount - 1);
        }

        /// <summary>双列映射网格：标签列 AutoSize + 下拉列百分比（38%），自适应宽度。</summary>
        private static TableLayoutPanel NewPairGrid()
        {
            var g = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Margin = new Padding(0) };
            g.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            g.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            g.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            g.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            return g;
        }

        private static void AddPair(TableLayoutPanel g, int row, Label l1, Control c1, Label l2, Control c2)
        {
            while (g.RowCount <= row) { g.RowStyles.Add(new RowStyle(SizeType.AutoSize)); g.RowCount++; }
            l1.Margin = new Padding(0, 7, 6, 2);
            g.Controls.Add(l1, 0, row);
            c1.Dock = DockStyle.Fill;
            c1.Margin = new Padding(0, 3, 10, 3);
            g.Controls.Add(c1, 1, row);
            if (l2 != null)
            {
                l2.Margin = new Padding(0, 7, 6, 2);
                g.Controls.Add(l2, 2, row);
            }
            if (c2 != null)
            {
                c2.Dock = DockStyle.Fill;
                c2.Margin = new Padding(0, 3, 0, 3);
                g.Controls.Add(c2, 3, row);
            }
        }

        private static Label TLabel(string text, Color? color = null)
        {
            return new Label { Text = text, AutoSize = true, ForeColor = color ?? Color.FromArgb(60, 70, 85), Margin = new Padding(0, 4, 0, 2) };
        }

        private static Button MakeStepButton(string text)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                MinimumSize = new Size(140, 32),
                BackColor = Color.FromArgb(68, 114, 196),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
        }

        private ComboBox MakeColumnCombo(params string[] keywords)
        {
            var cbo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            for (int i = 0; i < _columnNames.Count; i++) { cbo.Items.Add(_columnNames[i]); }
            if (_columnNames.Count > 0)
            {
                int hit = -1;
                for (int i = 0; i < _columnNames.Count && hit < 0; i++)
                {
                    foreach (var kw in keywords)
                    {
                        if (_columnNames[i].Contains(kw)) { hit = i; break; }
                    }
                }
                cbo.SelectedIndex = hit >= 0 ? hit : 0;
            }
            return cbo;
        }

        /// <summary>操作行：步骤执行按钮 + 状态字（FlowLayoutPanel 自动排列）。</summary>
        private static FlowLayoutPanel NewActionRow(Button btn, Label status)
        {
            var row = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Margin = new Padding(0, 6, 0, 2) };
            row.Controls.Add(btn);
            status.Margin = new Padding(12, 10, 0, 0);
            row.Controls.Add(status);
            return row;
        }

        private const string Desc1 = "【通俗】这一步把你的原始数据复制成规整的「原材料表」：文字格式的数字转真数字、日期防错乱，并自动算好月度/品类/地区三张汇总——原表一点不动。\r\n【专业】create_super_table：选区类型规整 + 月度/品类/地区三维汇总派生，后续透视表与看板统一取数。\r\n【怎么操作】对照下方示意图核对六个下拉与你的列对应无误 → 点「生成数据表」→ 看到 ✓ 后进下一步；鼠标停在任一下拉框上有逐字段说明。";

        /// <summary>Step1 数据表：数据区 + 列映射 → create_super_table。</summary>
        private void BuildStep1()
        {
            var panel = NewStepPanel(0, "第 1 步 · 生成数据表（超级数据表）", Desc1);
            var stack = NewStack();
            AddRow(stack, MakeTitleLabel("第 1 步 · 生成数据表（超级数据表）"));
            var grid = NewPairGrid();

            _addressBox = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 3) };
            _addressBox.Text = string.IsNullOrEmpty(_sourceSheet) ? "" : $"{_sourceSheet}!{_sourceRange}";
            AddPair(grid, 0, TLabel("数据区域："), _addressBox, null, null);
            grid.SetColumnSpan(_addressBox, 3);

            _cboLabel = MakeColumnCombo("地区", "区域", "销售员", "人员");
            _cboCategory = MakeColumnCombo("品类", "类别", "产品", "分类");
            AddPair(grid, 1, TLabel("地区/人员："), _cboLabel, TLabel("品类："), _cboCategory);
            _cboMonth = MakeColumnCombo("月份", "月", "日期");
            _cboValue = MakeColumnCombo("金额", "销售额", "收入");
            AddPair(grid, 2, TLabel("月份/日期："), _cboMonth, TLabel("金额："), _cboValue);
            _cboQuantity = MakeColumnCombo("数量", "销量");
            _cboProfit = MakeColumnCombo("利润");
            _cboProfit.Items.Insert(0, "（无）");
            _cboProfit.SelectedIndex = 0;
            AddPair(grid, 3, TLabel("数量："), _cboQuantity, TLabel("利润（可选）："), _cboProfit);
            AddRow(stack, grid);

            _stepButtons[0] = MakeStepButton("生成数据表");
            _stepButtons[0].Click += OnGenerateDataSheet;
            _stepStatus[0] = new Label { Text = "未生成", AutoSize = true, ForeColor = Color.Gray };
            AddRow(stack, NewActionRow(_stepButtons[0], _stepStatus[0]));
            panel.Controls.Add(stack);

            TrySetDefault(_cboLabel, "地区", "区域", "销售员", "人员");
            TrySetDefault(_cboCategory, "品类", "类别", "产品", "分类");
            TrySetDefault(_cboMonth, "月份", "月", "日期");
            TrySetDefault(_cboValue, "金额", "销售额", "收入");
            TrySetDefault(_cboQuantity, "数量", "销量");
            TrySetDefault(_cboProfit, "利润");

            // ---- 全控件双层说明（v3.7.1.54）：【是什么】专业口径 +【通俗讲】+【怎么操作】 ----
            _wizardTip.SetToolTip(_addressBox, "【是什么】要分析的数据范围（工作表名!区域，如 Sheet1!$B$10:$S$27），第一行必须是表头。\r\n【通俗讲】告诉助手「从哪块格子取数」；只想分析一部分就改小范围（比如去掉合计行）。\r\n【怎么操作】保持默认即可；要改就按 Excel 地址格式手动输入。");
            _wizardTip.SetToolTip(_cboLabel, "【是什么】看板第一分类维度（离散维度字段）。\r\n【通俗讲】告诉助手哪一列是「地区/人员」这类分类：销售表选地区或销售员，库存表选仓库，工资表选部门。\r\n【怎么操作】下拉里选你表里对应的列名；选错看板分类就会张冠李戴。");
            _wizardTip.SetToolTip(_cboCategory, "【是什么】第二分类维度，用于占比结构（饼图）。\r\n【通俗讲】销售表选产品类别；库存表选物品大类。不能和「地区/人员」选同一列。\r\n【怎么操作】下拉选列名；两列选重了点「生成数据表」会弹窗提醒。");
            _wizardTip.SetToolTip(_cboMonth, "【是什么】时间维度，趋势折线图的横轴。\r\n【通俗讲】选你表里的日期列或月份列；「2026-01」这种文本也可以，助手会自动转成日期。\r\n【怎么操作】下拉选列名即可。");
            _wizardTip.SetToolTip(_cboValue, "【是什么】核心度量（专业：主度量列，KPI 销售总额=本列求和）。\r\n【通俗讲】最值钱的那个数字列：销售额/收入/库存金额。\r\n【怎么操作】下拉选列名；这列决定看板上最大的那张数字卡。");
            _wizardTip.SetToolTip(_cboQuantity, "【是什么】数量度量（成交笔数/销量来源）。\r\n【通俗讲】记录个数或件数的列；没有合适列就选最接近的。\r\n【怎么操作】下拉选列名。");
            _wizardTip.SetToolTip(_cboProfit, "【是什么】利润度量（可选），利润视角与利润指标的数据源。\r\n【通俗讲】表里有利润列就选上；没有就保持（无），只是少一个利润视角，不影响其他功能。\r\n【怎么操作】下拉选利润列或保持（无）。");
            _wizardTip.SetToolTip(_stepButtons[0], "【是什么】第一步执行按钮：生成「超级数据表」。\r\n【通俗讲】点一下，助手把你的数据复制规整成标准格式并算好三张汇总小表，原表一点不动。\r\n【怎么操作】检查六个下拉 → 点这里 → 看到 ✓ 和提示后点「是」进入下一步。");
        }

        private const string Desc2 = "【通俗】透视表=能拖拽的分析表：勾选「按什么分类」（行维度）和「算什么数」（值字段），生成后可随时在 Excel 里拖拽重排。\r\n【专业】build_pivot_sheet：原生 PivotTable，行字段=勾选维度，数据字段=勾选度量求和，可挂接看板切片器。\r\n【怎么操作】勾行维度 + 值字段 → 点「生成透视表」；不需要分析表就勾「跳过这一步」，不影响看板成品。";

        /// <summary>Step2 透视表：行维度勾选 + 值字段勾选 → build_pivot_sheet（可跳过）。</summary>
        private void BuildStep2()
        {
            var panel = NewStepPanel(1, "第 2 步 · 生成透视表（可选，可跳过）", Desc2);
            var stack = NewStack();
            AddRow(stack, MakeTitleLabel("第 2 步 · 生成透视表（可选，可跳过）"));

            // 行维度 / 值字段双列各 50%（TableLayoutPanel，任意宽度/DPI 不挤出可视区）
            var tlp = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 2, Margin = new Padding(0) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 116));
            var lblRows = new Label { Text = "行维度（可多选）：", AutoSize = true, Margin = new Padding(0, 4, 4, 2) };
            var lblVals = new Label { Text = "值字段（可多选）：", AutoSize = true, Margin = new Padding(12, 4, 4, 2) };
            _clbPivotRows = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Margin = new Padding(0, 2, 8, 2) };
            _clbPivotVals = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, Margin = new Padding(0) };
            _clbPivotRows.Items.Add(SafeName(_cboMonth));
            _clbPivotRows.Items.Add(SafeName(_cboCategory));
            _clbPivotRows.Items.Add(SafeName(_cboLabel));
            _clbPivotRows.SetItemChecked(0, true);
            _clbPivotMeasures = _clbPivotVals;
            _clbPivotMeasures.Items.Add("金额");
            _clbPivotMeasures.Items.Add("数量");
            bool hasProfit = _cboProfit.SelectedIndex > 0;
            if (hasProfit) { _clbPivotMeasures.Items.Add("利润"); _clbPivotMeasures.SetItemChecked(0, true); }
            else { _clbPivotMeasures.SetItemChecked(0, true); }
            tlp.Controls.Add(lblRows, 0, 0);
            tlp.Controls.Add(lblVals, 1, 0);
            tlp.Controls.Add(_clbPivotRows, 0, 1);
            tlp.Controls.Add(_clbPivotVals, 1, 1);
            AddRow(stack, tlp);

            _chkSkipPivot = new CheckBox { Text = "跳过这一步（不生成透视表）", AutoSize = true, Margin = new Padding(0, 4, 0, 2) };
            AddRow(stack, _chkSkipPivot);

            _stepButtons[1] = MakeStepButton("生成透视表");
            _stepButtons[1].Click += OnGeneratePivot;
            _stepStatus[1] = new Label { Text = "未生成", AutoSize = true, ForeColor = Color.Gray };
            AddRow(stack, NewActionRow(_stepButtons[1], _stepStatus[1]));
            panel.Controls.Add(stack);

            _wizardTip.SetToolTip(_clbPivotRows, "【是什么】透视表的行标签维度（可多选）。\r\n【通俗讲】按什么分类看：勾「月份」看每月一行；勾「地区」看每个地区一行；两个都勾=每月再分地区。\r\n【怎么操作】点条目打勾/取消；至少留一个。");
            _wizardTip.SetToolTip(_clbPivotMeasures, "【是什么】透视表的数值区字段（可多选，自动求和）。\r\n【通俗讲】格子里算什么数：勾「金额」就是销售额表；再勾「数量」就多一列销量。\r\n【怎么操作】点条目打勾；至少留一个。");
            _wizardTip.SetToolTip(_chkSkipPivot, "【是什么】跳过本步开关。\r\n【通俗讲】不想要分析表、只要看板？勾它直接进第 3 步，成品不受影响。\r\n【怎么操作】不需要透视表就勾上，再点「下一步」。");
            _wizardTip.SetToolTip(_stepButtons[1], "【是什么】第二步执行按钮：生成 Excel 原生数据透视表（新表「透视分析」）。\r\n【通俗讲】得到一张能随意拖拽、刷新的分析表，还能和看板切片器联动。\r\n【怎么操作】勾好行维度和值字段 → 点这里 → ✓ 后进下一步。");
        }

        private const string Desc3 = "【通俗】最后一步出成品：5 张大数字卡 + 四宫格图表 + 筛选按钮。想自己算新指标（如 利润占比=利润额/销售总额）就点彩色按钮组合公式；配色和筛选方式也在这里定。\r\n【专业】render_super_dashboard：KPI 卡 + 四宫格图表 + 原生 Slicer，extra_kpis 逐行派生派生指标。\r\n【怎么操作】（可选）加自定义指标 → 选主题 → 点「生成超级看板」；分步嫌麻烦就点右下角「一键完成」。";

        /// <summary>Step3 超级看板：自定义指标 + 切片器 + 主题 → render_super_dashboard。</summary>
        private void BuildStep3()
        {
            var panel = NewStepPanel(2, "第 3 步 · 生成超级看板", Desc3);
            var stack = NewStack();
            AddRow(stack, MakeTitleLabel("第 3 步 · 生成超级看板"));

            AddRow(stack, new Label { Text = "自定义派生指标（可选，≤5 行）      可用基础指标（点击自动填入公式）：", AutoSize = true, Margin = new Padding(0, 4, 0, 2) });

            // 可用基础指标「点击填入」按钮区（FlowLayoutPanel 自动换行，不遮挡）
            var flowBase = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Margin = new Padding(0) };
            foreach (var baseName in new[] { "销售总额", "销售量", "成交笔数", "平均客单价", "最大单笔", "利润额" })
            {
                var chip = new Button
                {
                    Text = baseName,
                    AutoSize = true,
                    MinimumSize = new Size(0, 26),
                    BackColor = Color.FromArgb(238, 242, 255),
                    FlatStyle = FlatStyle.Flat,
                    Tag = baseName,
                    Margin = new Padding(0, 1, 6, 1)
                };
                chip.FlatAppearance.BorderColor = Color.FromArgb(199, 210, 254);
                chip.Click += (s, e) => InsertBaseKpiIntoGrid((string)((Button)s).Tag);
                flowBase.Controls.Add(chip);
            }
            panel.Controls.Add(stack);          // 先挂容器再继续加行（顺序无关，Dock=Top 由 z-order 决定）
            _flowBase = flowBase;
            AddRow(stack, flowBase);
            AddRow(stack, TLabel("写法：直接点上方按钮插入指标名，再用 + - * / 组合（如 利润额/销售总额）。除零与错误值已自动保护。", Color.FromArgb(90, 100, 115)));

            _gridExtra = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 132,
                AllowUserToAddRows = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 30 },
                ColumnHeadersHeight = 32,
                Margin = new Padding(0, 2, 0, 2)
            };
            _gridExtra.Columns.Add("name", "指标名（如 利润占比）");
            _gridExtra.Columns.Add("formula", "公式（基础指标四则，如 利润额/销售总额）");
            _gridExtra.Columns.Add("format", "格式（如 0.0%）");
            _gridExtra.Columns[0].FillWeight = 25;
            _gridExtra.Columns[1].FillWeight = 55;
            _gridExtra.Columns[2].FillWeight = 20;
            AddRow(stack, _gridExtra);

            var optRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 2) };
            _chkSlicer = new CheckBox { Text = "启用原生切片器（单选 / Ctrl 多选 / 多选模式按钮）", AutoSize = true, Checked = true };
            _cboTheme = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
            _cboTheme.Items.AddRange(new object[] { "orange", "blue", "graygreen", "purple", "teal", "crimson", "gold", "slate" });
            _cboTheme.SelectedIndex = 0;
            optRow.Controls.Add(_chkSlicer);
            optRow.Controls.Add(new Label { Text = "   主题：", AutoSize = true, Margin = new Padding(8, 8, 2, 0) });
            optRow.Controls.Add(_cboTheme);
            AddRow(stack, optRow);

            _chkAutoRename = new CheckBox { Text = "重名自动递增共存", AutoSize = true, Margin = new Padding(0, 2, 0, 2) };
            AddRow(stack, _chkAutoRename);

            _stepButtons[2] = MakeStepButton("生成超级看板");
            _stepButtons[2].Click += OnGenerateBoard;
            _stepStatus[2] = new Label { Text = "未生成", AutoSize = true, ForeColor = Color.Gray };
            AddRow(stack, NewActionRow(_stepButtons[2], _stepStatus[2]));

            _wizardTip.SetToolTip(_flowBase, "【是什么】六个基础指标按钮（销售总额/销售量/成交笔数/平均客单价/最大单笔/利润额）。\r\n【通俗讲】点任一按钮=把该指标名自动填进下方表格的「公式」列，不用手打也不会拼错字。\r\n【怎么操作】先在下方表格点一格，再点按钮；组合公式如 利润额/销售总额。");
            _gridExtra.Columns[0].HeaderCell.ToolTipText = "【是什么】新指标的名字，显示在看板「指标视图」下拉里。\r\n【通俗讲】起个好懂的短名，如「利润占比」。\r\n【怎么操作】直接打字；留空=这一行不生效。";
            _gridExtra.Columns[1].HeaderCell.ToolTipText = "【是什么】派生指标公式=基础指标四则运算（+ 加 - 减 * 乘 / 除）。\r\n【通俗讲】不会写？点上方彩色按钮自动填指标名，中间加运算符即可。\r\n【怎么操作】示例：利润额/销售总额=利润占比；销售总额/10000=万元口径。除零已自动保护。";
            _gridExtra.Columns[2].HeaderCell.ToolTipText = "【是什么】数字显示格式（Excel 格式代码）。\r\n【通俗讲】想让 0.53 显示成 53.0% 就填 0.0%；留空=千分位整数。\r\n【怎么操作】常用：0.0% 百分比 / #,##0 千分位 / #,##0.00 两位小数。";
            _wizardTip.SetToolTip(_chkSlicer, "【是什么】原生切片器开关（专业：SlicerCache 绑定数据透视表）。\r\n【通俗讲】看板上的一排筛选按钮，点一下只看选中项；按住 Ctrl 点可多选。\r\n【怎么操作】建议保持勾选；不勾则用普通下拉筛选。");
            _wizardTip.SetToolTip(_cboTheme, "【是什么】看板配色主题（8 种）。\r\n【通俗讲】看板的配色风格，纯外观随时可换。\r\n【怎么操作】下拉选一个；生成后说「换 XX 主题」也能切换。");
            _wizardTip.SetToolTip(_chkAutoRename, "【是什么】同名看板冲突策略。\r\n【通俗讲】已存在同名看板时：勾选=新看板自动叫「xx2」两个并存；不勾=提示先删旧的。\r\n【怎么操作】一般保持勾选防覆盖。");
            _wizardTip.SetToolTip(_stepButtons[2], "【是什么】最终执行按钮：生成超级看板（新工作表）。\r\n【通俗讲】5 张大数字卡 + 四宫格图表 + 筛选器一次生成，用上本页所有配置。\r\n【怎么操作】配好自定义指标/切片器/主题 → 点这里 → ✓ 后按提示查看成品。");
        }

        /// <summary>底部区：矢量示意图（随步骤切换）+ 双层文字说明，位于导航按钮上方。</summary>
        private void BuildBottomArea()
        {
            var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 1, Margin = new Padding(0) };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _diagram = new WizardDiagram { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 2) };
            bottom.Controls.Add(_diagram, 0, 0);

            _stepDesc = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.FromArgb(90, 100, 115),
                Margin = new Padding(2, 2, 2, 4)
            };
            bottom.Controls.Add(_stepDesc, 0, 1);
            _root.Controls.Add(bottom, 0, 2);
        }

        /// <summary>导航：上一步 / 下一步 / 一键完成（TLP 4 列，右列弹性）。</summary>
        private void BuildNavButtons()
        {
            var nav = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 4, Margin = new Padding(0) };
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            nav.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var btnPrev = new Button { Text = "◀ 上一步", AutoSize = true, MinimumSize = new Size(96, 28), Enabled = false, Margin = new Padding(0, 2, 6, 2) };
            var btnNext = new Button { Text = "下一步 ▶", AutoSize = true, MinimumSize = new Size(96, 28), Margin = new Padding(0, 2, 6, 2) };
            var btnAll = new Button
            {
                Text = "⚡ 一键完成（推荐）",
                AutoSize = true,
                MinimumSize = new Size(200, 28),
                BackColor = Color.FromArgb(230, 126, 34),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(0, 2, 0, 2)
            };
            btnPrev.Click += (_, __) => { if (_currentStep > 0) { ShowStep(_currentStep - 1); } };
            btnNext.Click += (_, __) => { if (_currentStep < 2) { ShowStep(_currentStep + 1); } };
            btnAll.Click += OnAllInOne;
            _wizardTip.SetToolTip(btnPrev, "【是什么】返回上一步。\r\n【通俗讲】改配置用；已生成的表不会删。\r\n【怎么操作】直接点。");
            _wizardTip.SetToolTip(btnNext, "【是什么】进入下一步。\r\n【通俗讲】本步可以先不生成直接跳过去，也可以生成完再走。\r\n【怎么操作】直接点。");
            _wizardTip.SetToolTip(btnAll, "【是什么】串行一键执行三步。\r\n【通俗讲】懒人模式：数据表→透视表→看板一路自动跑完，透视失败也会跳过继续，全程不用管。\r\n【怎么操作】不想分步就点它。");
            nav.Controls.Add(btnPrev, 0, 0);
            nav.Controls.Add(btnNext, 1, 0);
            nav.Controls.Add(btnAll, 3, 0);
            _root.Controls.Add(nav, 0, 3);
            _navPrev = btnPrev;
            _navNext = btnNext;
        }

        private Button _navPrev, _navNext;

        private void TrySetDefault(ComboBox cbo, params string[] keywords)
        {
            // 按关键词猜默认列（表头名匹配），未命中保持现状
            if (cbo.Items.Count == 0) { return; }
            foreach (var kw in keywords)
            {
                for (int i = 0; i < _columnNames.Count; i++)
                {
                    if (_columnNames[i].Contains(kw)) { cbo.SelectedIndex = i; return; }
                }
            }
        }

        private string SafeName(ComboBox cbo)
        {
            return cbo.SelectedIndex >= 0 ? Convert.ToString(cbo.SelectedItem) : "";
        }

        private void ShowStep(int index)
        {
            _currentStep = index;
            for (int i = 0; i < 3; i++)
            {
                _stepPanels[i].Visible = i == index;
                _stepIndicators[i].ForeColor = i == index ? Color.FromArgb(200, 90, 20) : _stepStatus[i].Text.StartsWith("✓") ? Color.FromArgb(40, 140, 60) : Color.FromArgb(120, 120, 120);
            }
            if (_stepDesc != null) { _stepDesc.Text = _stepDescriptions[index] ?? ""; }
            if (_diagram != null)
            {
                _diagram.Kind = (WizardDiagram.DiagramKind)index;
                _diagram.Invalidate();
            }
            DiagnosticsLog.Write(_logPrefix + " 切步 step=" + (index + 1) + "/3");
            _navPrev.Enabled = index > 0;
            _navNext.Enabled = index < 2;
        }

        /// <summary>探针专用：离屏渲染时切换步骤页（SB5 W5）。</summary>
        public void WizardShowStep(int index)
        {
            ShowStep(Math.Max(0, Math.Min(2, index)));
            Application.DoEvents();
        }

        private void MarkDone(int step, string extra)
        {
            _stepStatus[step].Text = "✓ 已生成 " + extra;
            _stepStatus[step].ForeColor = Color.FromArgb(40, 140, 60);
            _stepIndicators[step].Text = _stepIndicators[step].Text.Split(' ')[0] + " ✓";
            AppLog.Info(_logPrefix, "第" + (step + 1) + "步完成：" + extra);
        }

        // ---- 执行逻辑 ----

        private Dictionary<string, object> BuildBaseArgs()
        {
            var addr = (_addressBox.Text ?? "").Trim();
            return new Dictionary<string, object>
            {
                ["source_address"] = addr.Split('!').Last().Trim(),
                ["sheet"] = addr.Split('!').First().Trim('\'', '[', ']'),
                ["label_column"] = _cboLabel.SelectedIndex + 1,
                ["category_column"] = _cboCategory.SelectedIndex + 1,
                ["month_column"] = _cboMonth.SelectedIndex + 1,
                ["value_column"] = _cboValue.SelectedIndex + 1,
                ["quantity_column"] = _cboQuantity.SelectedIndex + 1,
                ["super_sheet"] = _superSheetName
            };
        }

        private ToolResult RunTool(string tool, Dictionary<string, object> args)
        {
            var sw = Stopwatch.StartNew();
            DiagnosticsLog.Write(_logPrefix + " " + tool + " args=" + Newtonsoft.Json.JsonConvert.SerializeObject(args));
            var result = _executeTool(tool, args);
            sw.Stop();
            DiagnosticsLog.Write(_logPrefix + " " + tool + " ok=" + result.Ok + " code=" + result.ErrorCode + " elapsedMs=" + sw.ElapsedMilliseconds);
            if (result.Ok) { AppLog.Info(_logPrefix, tool + " 完成 " + sw.ElapsedMilliseconds + "ms"); }
            else { AppLog.Error(_logPrefix, tool + " 失败 code=" + result.ErrorCode); }
            return result;
        }

        /// <summary>点击基础指标按钮 → 填入当前行公式单元格（光标处追加）。</summary>
        private void InsertBaseKpiIntoGrid(string baseName)
        {
            var cell = _gridExtra.CurrentCell;
            int rowIndex;
            if (cell == null || cell.RowIndex < 0)
            {
                rowIndex = _gridExtra.Rows.Add();
                DiagnosticsLog.Write(_logPrefix + " 基础指标填入 新行 row=" + rowIndex + " kpi=" + baseName);
            }
            else
            {
                rowIndex = cell.RowIndex;               // 指标名列点击也填入公式列（用户意图=写公式）
                DiagnosticsLog.Write(_logPrefix + " 基础指标填入 row=" + rowIndex + " kpi=" + baseName);
            }
            var existing = Convert.ToString(_gridExtra.Rows[rowIndex].Cells[1].Value);
            _gridExtra.Rows[rowIndex].Cells[1].Value = string.IsNullOrWhiteSpace(existing) ? baseName : existing + baseName;
            _gridExtra.CurrentCell = _gridExtra.Rows[rowIndex].Cells[1];
        }

        private void OnGenerateDataSheet(object sender, EventArgs e)
        {
            if (_cboLabel.SelectedIndex == _cboCategory.SelectedIndex
                || _cboLabel.SelectedIndex == _cboMonth.SelectedIndex
                || _cboCategory.SelectedIndex == _cboMonth.SelectedIndex)
            {
                DiagnosticsLog.Write(_logPrefix + " Step1 校验拒绝：地区/品类/月份同列 label=" + _cboLabel.SelectedIndex + " category=" + _cboCategory.SelectedIndex + " month=" + _cboMonth.SelectedIndex);
                MessageBox.Show("地区/品类/月份不能映射到同一列。", "超级看板向导", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var args = BuildBaseArgs();
            var result = RunTool("create_super_table", args);
            if (!result.Ok)
            {
                MessageBox.Show(Convert.ToString(result.Data), "数据表生成失败（" + result.ErrorCode + "）", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _step1Done = true;
            MarkDone(0, "超级数据表已生成");
            if (MessageBox.Show("数据表已生成 ✓\n\n进入下一步（透视表）吗？", "超级看板向导", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
            {
                ShowStep(1);
            }
        }

        private void OnGeneratePivot(object sender, EventArgs e)
        {
            if (_chkSkipPivot.Checked)
            {
                DiagnosticsLog.Write(_logPrefix + " Step2 用户跳过透视表");
                _stepStatus[1].Text = "— 已跳过";
                ShowStep(2);
                return;
            }
            var rows = _clbPivotRows.CheckedItems.Cast<string>().Select(t => t.Split('（')[0]).ToList();
            var measures = _clbPivotMeasures.CheckedItems.Cast<string>().ToList();
            if (rows.Count == 0 || measures.Count == 0)
            {
                DiagnosticsLog.Write(_logPrefix + " Step2 校验拒绝：rows=" + rows.Count + " measures=" + measures.Count);
                MessageBox.Show("请至少勾选一个行维度与一个值字段（或勾选跳过）。", "超级看板向导", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var args = BuildBaseArgs();
            args["pivot_rows"] = string.Join(",", rows);
            args["pivot_measures"] = string.Join(",", measures);
            args["pivot_sheet_name"] = _pivotSheetName;
            var result = RunTool("build_pivot_sheet", args);
            if (!result.Ok)
            {
                MessageBox.Show(Convert.ToString(result.Data), "透视表生成失败（" + result.ErrorCode + "）", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _step2Done = true;
            MarkDone(1, "透视分析已生成");
            if (MessageBox.Show("透视表已生成 ✓\n\n进入最后一步（超级看板）吗？", "超级看板向导", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
            {
                ShowStep(2);
            }
        }

        private void OnGenerateBoard(object sender, EventArgs e)
        {
            var args = BuildBaseArgs();
            if (_cboProfit.SelectedIndex > 0) { args["profit_column"] = _cboProfit.SelectedIndex + 1; }
            args["theme"] = Convert.ToString(_cboTheme.SelectedItem);
            args["auto_rename"] = _chkAutoRename.Checked ? "true" : "false";
            args["slicer_enabled"] = _chkSlicer.Checked ? "true" : "false";

            var extraRows = new List<Dictionary<string, object>>();
            foreach (DataGridViewRow row in _gridExtra.Rows)
            {
                if (row.IsNewRow) { continue; }
                var name = Convert.ToString(row.Cells[0].Value)?.Trim();
                var formula = Convert.ToString(row.Cells[1].Value)?.Trim();
                var format = Convert.ToString(row.Cells[2].Value)?.Trim();
                if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(formula)) { continue; }
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(formula))
                {
                    DiagnosticsLog.Write(_logPrefix + " Step3 校验拒绝：自定义指标第 " + (row.Index + 1) + " 行缺名或公式");
                    MessageBox.Show($"自定义指标第 {row.Index + 1} 行需同时填写指标名与公式（或整行留空）。", "超级看板向导", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                extraRows.Add(new Dictionary<string, object> { ["name"] = name, ["formula"] = formula, ["format"] = string.IsNullOrWhiteSpace(format) ? "#,##0" : format });
            }
            if (extraRows.Count > 0) { args["extra_kpis"] = Newtonsoft.Json.JsonConvert.SerializeObject(extraRows); }
            DiagnosticsLog.Write(_logPrefix + " Step3 配置 theme=" + args["theme"] + " slicer=" + args["slicer_enabled"] + " extraKpis=" + extraRows.Count);

            // 向导模式：透视分析已在 Step2 独立生成（"透视分析"），看板内部不再重复建（pivot_analysis 缺省 false）
            var result = RunTool("render_super_dashboard", args);
            if (!result.Ok)
            {
                MessageBox.Show(Convert.ToString(result.Data), "超级看板生成失败（" + result.ErrorCode + "）", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _step2Done = true;
            MarkDone(2, "超级看板已生成");
            MessageBox.Show("超级看板已生成 ✓\n\n· 看板上「指标视图」下拉可切换销售/利润/自定义指标\n· 品类与地区下拉可筛选\n· 建议打开侧边栏让 AI 深度核验并截图存档", "超级看板向导", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnAllInOne(object sender, EventArgs e)
        {
            DiagnosticsLog.Write(_logPrefix + " 一键完成 链路开始");
            AppLog.Info(_logPrefix, "一键完成开始");
            ShowStep(0);
            OnGenerateDataSheet(sender, e);
            if (!_step1Done)
            {
                DiagnosticsLog.Write(_logPrefix + " 一键完成 中止于 Step1");
                return;
            }
            ShowStep(1);
            if (!_chkSkipPivot.Checked)
            {
                OnGeneratePivot(sender, e);
                if (!_step2Done && _stepStatus[1].Text != "— 已跳过")
                {
                    DiagnosticsLog.Write(_logPrefix + " 一键完成 中止于 Step2（透视失败，跳过开关未勾）");
                    return;
                }
            }
            ShowStep(2);
            OnGenerateBoard(sender, e);
            DiagnosticsLog.Write(_logPrefix + " 一键完成 链路结束 step1Done=" + _step1Done + " boardDone=" + _stepStatus[2].Text.StartsWith("✓"));
        }
    }
}
