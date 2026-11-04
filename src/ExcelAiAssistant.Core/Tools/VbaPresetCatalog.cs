using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Tools
{
    /// <summary>一个预置 VBA 模板条目（分类 preset，随产品内置）。</summary>
    public sealed class VbaPresetTemplate
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Code { get; set; }
        public string Version { get; set; }
        public IReadOnlyList<string> Tags { get; set; }
    }

    /// <summary>
    /// 内置预置 VBA 模板库（三源之一：预置片段）。
    /// 仅作为「种子数据」被 <see cref="CustomToolStore.SeedPresets"/> 首次导入到统一存储；
    /// 前端模板下拉框不再硬编码，统一从 CustomToolStore 动态加载。
    /// </summary>
    public static class VbaPresetCatalog
    {
        public static IReadOnlyList<VbaPresetTemplate> All { get; } = Build();

        private static List<VbaPresetTemplate> Build()
        {
            return new List<VbaPresetTemplate>
            {
                new VbaPresetTemplate
                {
                    Name = "sum_column",
                    Description = "汇总一列数值",
                    Version = "1.0",
                    Tags = new[] { "数值", "求和", "汇总" },
                    Code = Produce("Sub 汇总一列()",
                        "  Dim s As Double",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择要汇总的数值区域\", Type:=8)",
                        "  s = Application.WorksheetFunction.Sum(rng)",
                        "  MsgBox \"合计：\" & s",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "fill_down",
                    Description = "向下填充公式",
                    Version = "1.0",
                    Tags = new[] { "公式", "填充", "FillDown" },
                    Code = Produce("Sub 向下填充公式()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择首格已写公式、下方要填充的区域（含首格）\", Type:=8)",
                        "  rng.FillDown",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "format_header",
                    Description = "表头加粗+底色",
                    Version = "1.0",
                    Tags = new[] { "表头", "格式", "加粗" },
                    Code = Produce("Sub 表头加粗底色()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择表头区域\", Type:=8)",
                        "  With rng",
                        "    .Font.Bold = True",
                        "    .Interior.Color = RGB(221, 235, 247)",
                        "    .HorizontalAlignment = xlCenter",
                        "  End With",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "remove_duplicates",
                    Description = "整行去重",
                    Version = "1.0",
                    Tags = new[] { "去重", "重复值", "清洗" },
                    Code = Produce("Sub 整行去重()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择要去重的区域（含表头）\", Type:=8)",
                        "  rng.RemoveDuplicates Columns:=Array(1), Header:=xlYes",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "trim_cells",
                    Description = "批量去空格",
                    Version = "1.0",
                    Tags = new[] { "空格", "文本", "清洗" },
                    Code = Produce("Sub 批量去空格()",
                        "  Dim rng As Range",
                        "  Dim c As Range",
                        "  Set rng = Application.InputBox(\"选择要去空格的区域\", Type:=8)",
                        "  For Each c In rng",
                        "    If Not IsEmpty(c) And VarType(c.Value) = vbString Then",
                        "      c.Value = Application.WorksheetFunction.Trim(c.Value)",
                        "    End If",
                        "  Next c",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "auto_width",
                    Description = "自动调整列宽",
                    Version = "1.0",
                    Tags = new[] { "列宽", "格式", "布局" },
                    Code = Produce("Sub 自动调整列宽()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择要调整列宽的区域\", Type:=8)",
                        "  rng.EntireColumn.AutoFit",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "highlight_duplicates",
                    Description = "高亮重复值",
                    Version = "1.0",
                    Tags = new[] { "重复值", "高亮", "条件格式" },
                    Code = Produce("Sub 高亮重复值()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择要检查重复值的区域\", Type:=8)",
                        "  rng.FormatConditions.AddUniqueValues",
                        "  rng.FormatConditions(rng.FormatConditions.Count).DupUnique = xlDuplicate",
                        "  rng.FormatConditions(rng.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "export_pdf",
                    Description = "导出为 PDF",
                    Version = "1.0",
                    Tags = new[] { "导出", "PDF", "打印" },
                    Code = Produce("Sub 导出为PDF()",
                        "  Dim path As String",
                        "  path = ThisWorkbook.Path & \"\\\" & ThisWorkbook.Name & \"_导出.pdf\"",
                        "  ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path",
                        "  MsgBox \"已导出：\" & path",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "hr_dept_count",
                    Description = "按部门统计人数",
                    Version = "1.0",
                    Tags = new[] { "人事", "统计", "部门" },
                    Code = Produce("Sub 按部门统计人数()",
                        "  Dim rng As Range",
                        "  Dim dict As Object",
                        "  Dim ws As Worksheet",
                        "  Dim s As Long",
                        "  Dim dept As Variant",
                        "  Set dict = CreateObject(\"Scripting.Dictionary\")",
                        "  Set rng = Application.InputBox(\"选择部门列（第1行为表头）\", Type:=8)",
                        "  For s = 2 To rng.Rows.Count",
                        "    dept = Trim(CStr(rng.Cells(s, 1).Value))",
                        "    If dept <> \"\" Then",
                        "      If dict.Exists(dept) Then dict(dept) = dict(dept) + 1 Else dict(dept) = 1",
                        "    End If",
                        "  Next s",
                        "  Set ws = Sheets.Add(After:=Worksheets(Worksheets.Count))",
                        "  ws.Cells(1, 1).Value = \"部门\"",
                        "  ws.Cells(1, 2).Value = \"人数\"",
                        "  s = 2",
                        "  For Each dept In dict.Keys",
                        "    ws.Cells(s, 1).Value = dept",
                        "    ws.Cells(s, 2).Value = dict(dept)",
                        "    s = s + 1",
                        "  Next dept",
                        "  MsgBox \"已统计 \" & dict.Count & \" 个部门，结果在新表 \" & ws.Name",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "hr_id_extract",
                    Description = "身份证提取出生日期+性别",
                    Version = "1.0",
                    Tags = new[] { "人事", "身份证", "性别" },
                    Code = Produce("Sub 身份证提取出生日期性别()",
                        "  Dim rng As Range",
                        "  Dim s As Long",
                        "  Dim id As String",
                        "  Set rng = Application.InputBox(\"选择身份证号所在列（第1行为表头）\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"出生日期\"",
                        "  rng.Cells(1, 3).Value = \"性别\"",
                        "  For s = 2 To rng.Rows.Count",
                        "    id = Trim(CStr(rng.Cells(s, 1).Value))",
                        "    If Len(id) >= 18 Then",
                        "      rng.Cells(s, 2).Value = Mid(id, 7, 4) & \"-\" & Mid(id, 11, 2) & \"-\" & Mid(id, 13, 2)",
                        "      rng.Cells(s, 3).Value = IIf(CLng(Mid(id, 17, 1)) Mod 2 = 1, \"男\", \"女\")",
                        "    End If",
                        "  Next s",
                        "  MsgBox \"已提取 \" & (rng.Rows.Count - 1) & \" 行\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "hr_age_from_birth",
                    Description = "出生日期算年龄",
                    Version = "1.0",
                    Tags = new[] { "人事", "年龄", "出生日期" },
                    Code = Produce("Sub 出生日期算年龄()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Set rng = Application.InputBox(\"选择含表头的出生日期列\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"年龄\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsDate(rng.Cells(i, 1).Value) Then rng.Cells(i, 2).Value = DateDiff(\"yyyy\", rng.Cells(i, 1).Value, Date)",
                        "  Next i",
                        "  MsgBox \"年龄计算完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "hr_due_alert",
                    Description = "合同证件到期提醒",
                    Version = "1.0",
                    Tags = new[] { "人事", "到期", "提醒" },
                    Code = Produce("Sub 合同到期提醒()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Dim d As Date",
                        "  Set rng = Application.InputBox(\"选择含表头的到期日期列\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"状态\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsDate(rng.Cells(i, 1).Value) Then",
                        "      d = rng.Cells(i, 1).Value",
                        "      If d < Date Then",
                        "        rng.Cells(i, 2).Value = \"已到期\"",
                        "      ElseIf d <= Date + 30 Then",
                        "        rng.Cells(i, 2).Value = \"30天内到期\"",
                        "      Else",
                        "        rng.Cells(i, 2).Value = \"正常\"",
                        "      End If",
                        "      If d <= Date Then rng.Cells(i, 2).Interior.Color = RGB(255, 199, 206)",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"到期检查完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "da_clean_numbers",
                    Description = "文本数字批量转数值",
                    Version = "1.0",
                    Tags = new[] { "数据分析", "清洗", "类型转换" },
                    Code = Produce("Sub 文本数字转数值()",
                        "  Dim rng As Range",
                        "  Dim c As Range",
                        "  Set rng = Application.InputBox(\"选择文本数字区域\", Type:=8)",
                        "  For Each c In rng",
                        "    If VarType(c.Value) = vbString And IsNumeric(c.Value) Then",
                        "      c.Value = CLng(c.Value)",
                        "    End If",
                        "  Next c",
                        "  MsgBox \"转换完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "da_pivot",
                    Description = "一键生成透视表",
                    Version = "1.0",
                    Tags = new[] { "数据分析", "透视表", "汇总" },
                    Code = Produce("Sub 一键生成透视表()",
                        "  Dim src As Range",
                        "  Dim ws As Worksheet",
                        "  Dim pc As PivotCache",
                        "  Dim pt As PivotTable",
                        "  Set src = Application.InputBox(\"选择数据区域（含表头）\", Type:=8)",
                        "  Set ws = Sheets.Add(After:=Worksheets(Worksheets.Count))",
                        "  Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=src.Address(True, True, xlA1, True))",
                        "  Set pt = pc.CreatePivotTable(TableDestination:=ws.Range(\"A3\"), TableName:=\"透视表与分析\")",
                        "  pt.AddDataField pt.PivotFields(1), \"计数\", xlCount",
                        "  MsgBox \"已生成透视表：\" & ws.Name",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "da_correlation",
                    Description = "两列相关系数",
                    Version = "1.0",
                    Tags = new[] { "数据分析", "相关性", "回归" },
                    Code = Produce("Sub 两列相关系数()",
                        "  Dim rng As Range",
                        "  Dim r As Double",
                        "  Set rng = Application.InputBox(\"选择两列等长数值区域（不含表头）\", Type:=8)",
                        "  r = Application.WorksheetFunction.Correl(rng.Columns(1), rng.Columns(2))",
                        "  MsgBox \"相关系数 r = \" & Round(r, 4)",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "da_zscore",
                    Description = "标准化（Z分数）",
                    Version = "1.0",
                    Tags = new[] { "数据分析", "标准化", "Z分数" },
                    Code = Produce("Sub 标准化Z分数()",
                        "  Dim rng As Range",
                        "  Dim data As Range",
                        "  Dim i As Long",
                        "  Dim m As Double",
                        "  Dim sd As Double",
                        "  Set rng = Application.InputBox(\"选择含表头的数值列\", Type:=8)",
                        "  Set data = rng.Offset(1, 0).Resize(rng.Rows.Count - 1)",
                        "  m = Application.WorksheetFunction.Average(data)",
                        "  sd = Application.WorksheetFunction.StDev(data)",
                        "  rng.Cells(1, 2).Value = \"Z分数\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsNumeric(rng.Cells(i, 1).Value) And sd <> 0 Then rng.Cells(i, 2).Value = (CDbl(rng.Cells(i, 1).Value) - m) / sd",
                        "  Next i",
                        "  MsgBox \"均值 \" & Round(m, 4) & \"，标准差 \" & Round(sd, 4)",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "fin_vat_split",
                    Description = "价税分离（13%）",
                    Version = "1.0",
                    Tags = new[] { "财务", "增值税", "含税" },
                    Code = Produce("Sub 价税分离()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Dim taxIn As Double",
                        "  Set rng = Application.InputBox(\"选择含税金额列（含表头，按13%增值税）\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"不含税金额\"",
                        "  rng.Cells(1, 3).Value = \"税额\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsNumeric(rng.Cells(i, 1).Value) Then",
                        "      taxIn = CDbl(rng.Cells(i, 1).Value)",
                        "      rng.Cells(i, 2).Value = Round(taxIn / 1.13, 2)",
                        "      rng.Cells(i, 3).Value = Round(taxIn - taxIn / 1.13, 2)",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"价税分离完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "fin_currency_format",
                    Description = "金额千分位格式",
                    Version = "1.0",
                    Tags = new[] { "财务", "金额", "格式" },
                    Code = Produce("Sub 金额千分位格式()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择金额区域\", Type:=8)",
                        "  rng.NumberFormat = \"#,##0.00\"",
                        "  rng.HorizontalAlignment = xlRight",
                        "  MsgBox \"格式已应用\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "fin_negative_red",
                    Description = "负数标红",
                    Version = "1.0",
                    Tags = new[] { "财务", "负数", "高亮" },
                    Code = Produce("Sub 负数标红()",
                        "  Dim rng As Range",
                        "  Dim c As Range",
                        "  Set rng = Application.InputBox(\"选择数值区域\", Type:=8)",
                        "  For Each c In rng",
                        "    If IsNumeric(c.Value) And CDbl(c.Value) < 0 Then c.Font.Color = RGB(200, 0, 0)",
                        "  Next c",
                        "  MsgBox \"负数已标红\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "att_punch_dedupe",
                    Description = "考勤打卡同人同日去重",
                    Version = "1.0",
                    Tags = new[] { "考勤", "去重", "打卡" },
                    Code = Produce("Sub 考勤打卡去重()",
                        "  Dim rng As Range",
                        "  Dim dict As Object",
                        "  Dim i As Long",
                        "  Dim key As String",
                        "  Set dict = CreateObject(\"Scripting.Dictionary\")",
                        "  Set rng = Application.InputBox(\"选择打卡记录（第1列工号、第2列日期，含表头）\", Type:=8)",
                        "  For i = rng.Rows.Count To 2 Step -1",
                        "    key = Trim(CStr(rng.Cells(i, 1).Value)) & \"|\" & Trim(CStr(rng.Cells(i, 2).Value))",
                        "    If dict.Exists(key) Then",
                        "      rng.Rows(i).EntireRow.Delete",
                        "    Else",
                        "      dict(key) = 1",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"去重完成，保留 \" & dict.Count & \" 条（同人同年月日仅留一次）\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "att_late_flag",
                    Description = "签到迟到标记",
                    Version = "1.0",
                    Tags = new[] { "考勤", "迟到", "标记" },
                    Code = Produce("Sub 迟到标记()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Dim t As Date",
                        "  Set rng = Application.InputBox(\"选择含表头的签到时间列（形如 9:00）\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"状态\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsDate(rng.Cells(i, 1).Value) Then",
                        "      t = TimeValue(rng.Cells(i, 1).Value)",
                        "      rng.Cells(i, 2).Value = IIf(t > TimeValue(\"09:00\"), \"迟到\", \"正常\")",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"迟到检查完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "att_overtime_hours",
                    Description = "加班时长（小时）",
                    Version = "1.0",
                    Tags = new[] { "考勤", "加班", "时长" },
                    Code = Produce("Sub 计算加班时长()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Set rng = Application.InputBox(\"选择含表头的下班打卡时间列（18:00 后计加班）\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"加班小时\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsDate(rng.Cells(i, 1).Value) Then",
                        "      rng.Cells(i, 2).Value = WorksheetFunction.Max(0, (TimeValue(rng.Cells(i, 1).Value) - TimeValue(\"18:00\")) * 24)",
                        "      rng.Cells(i, 2).NumberFormat = \"0.0\"",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"加班时长计算完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "stat_freq_table",
                    Description = "频数分布表",
                    Version = "1.0",
                    Tags = new[] { "统计", "频数", "直方图" },
                    Code = Produce("Sub 生成频数分布表()",
                        "  Dim rng As Range",
                        "  Dim bins As Variant",
                        "  Dim out As Variant",
                        "  Dim ws As Worksheet",
                        "  Dim k As Long",
                        "  bins = Array(10, 20, 30, 40, 50, 60)",
                        "  Set rng = Application.InputBox(\"选择数值数据列（不含表头）\", Type:=8)",
                        "  out = Application.WorksheetFunction.Frequency(rng, bins)",
                        "  Set ws = Sheets.Add(After:=Worksheets(Worksheets.Count))",
                        "  ws.Range(\"A1\").Value = \"分组上限\"",
                        "  ws.Range(\"B1\").Value = \"频数\"",
                        "  For k = 1 To UBound(bins) + 1",
                        "    ws.Cells(k + 1, 1).Value = bins(k - 1)",
                        "    ws.Cells(k + 1, 2).Value = out(k, 1)",
                        "  Next k",
                        "  ws.Cells(UBound(bins) + 3, 1).Value = \">\" & bins(UBound(bins))",
                        "  ws.Cells(UBound(bins) + 3, 2).Value = out(UBound(bins) + 2, 1)",
                        "  MsgBox \"已生成频数分布表：\" & ws.Name",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "stat_share_pct",
                    Description = "占比统计",
                    Version = "1.0",
                    Tags = new[] { "统计", "占比", "百分比" },
                    Code = Produce("Sub 计算占比()",
                        "  Dim rng As Range",
                        "  Dim data As Range",
                        "  Dim i As Long",
                        "  Dim total As Double",
                        "  Set rng = Application.InputBox(\"选择含表头的数值列\", Type:=8)",
                        "  Set data = rng.Offset(1, 0).Resize(rng.Rows.Count - 1)",
                        "  total = Application.WorksheetFunction.Sum(data)",
                        "  rng.Cells(1, 2).Value = \"占比\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsNumeric(rng.Cells(i, 1).Value) And total <> 0 Then",
                        "      rng.Cells(i, 2).Value = CDbl(rng.Cells(i, 1).Value) / total",
                        "      rng.Cells(i, 2).NumberFormat = \"0.0%\"",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"占比计算完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "an_growth_ratio",
                    Description = "环比增长率",
                    Version = "1.0",
                    Tags = new[] { "分析", "增长率", "环比" },
                    Code = Produce("Sub 计算环比增长率()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Set rng = Application.InputBox(\"选择数值列（不含表头，每行一个期间）\", Type:=8)",
                        "  rng.Offset(0, 1).Cells(1, 1).Value = \"环比增长\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsNumeric(rng.Cells(i - 1, 1).Value) And Val(rng.Cells(i - 1, 1).Value) <> 0 Then",
                        "      rng.Cells(i, 2).Value = (CDbl(rng.Cells(i, 1).Value) - CDbl(rng.Cells(i - 1, 1).Value)) / CDbl(rng.Cells(i - 1, 1).Value)",
                        "      rng.Cells(i, 2).NumberFormat = \"0.0%\"",
                        "    End If",
                        "  Next i",
                        "  MsgBox \"增长率计算完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "an_rank_fill",
                    Description = "生成排名（RANK）",
                    Version = "1.0",
                    Tags = new[] { "分析", "排名", "排序" },
                    Code = Produce("Sub 生成排名()",
                        "  Dim rng As Range",
                        "  Dim i As Long",
                        "  Set rng = Application.InputBox(\"选择含表头的数值列\", Type:=8)",
                        "  rng.Cells(1, 2).Value = \"排名\"",
                        "  For i = 2 To rng.Rows.Count",
                        "    If IsNumeric(rng.Cells(i, 1).Value) Then rng.Cells(i, 2).Value = Application.WorksheetFunction.Rank(rng.Cells(i, 1).Value, rng.Columns(1))",
                        "  Next i",
                        "  MsgBox \"排名生成完成\"",
                        "End Sub")
                },
                new VbaPresetTemplate
                {
                    Name = "an_topn_highlight",
                    Description = "TOP10 高亮",
                    Version = "1.0",
                    Tags = new[] { "分析", "TOP", "高亮" },
                    Code = Produce("Sub 高亮TOP10()",
                        "  Dim rng As Range",
                        "  Set rng = Application.InputBox(\"选择数值区域\", Type:=8)",
                        "  rng.FormatConditions.AddTop10",
                        "  With rng.FormatConditions(rng.FormatConditions.Count)",
                        "    .TopBottom = xlTop10Top",
                        "    .Rank = 10",
                        "    .Interior.Color = RGB(255, 235, 153)",
                        "  End With",
                        "  MsgBox \"已高亮 TOP10\"",
                        "End Sub")
                },
                // ===== Python 预置模板（带 python 标签，Python 面板模板库过滤展示）=====
                new VbaPresetTemplate
                {
                    Name = "py_sum_range",
                    Description = "读取区域求和（xlwings·Excel 模式）",
                    Version = "1.0",
                    Tags = new[] { "python", "求和", "读取" },
                    Code = Produce(
                        "import xlwings as xw",
                        "rng = xw.books.active.sheets.active.range(\"A1:A20\")",
                        "vals = [v for v in rng.value if isinstance(v, (int, float))]",
                        "xai_result({'sum': sum(vals), 'count': len(vals)})")
                },
                new VbaPresetTemplate
                {
                    Name = "py_fill_formula",
                    Description = "向下填充公式（xlwings·Excel 模式）",
                    Version = "1.0",
                    Tags = new[] { "python", "公式", "填充" },
                    Code = Produce(
                        "import xlwings as xw",
                        "sht = xw.books.active.sheets.active",
                        "rng = sht.range(\"B2:B10\")",
                        "rng.formula = \"=ROUND(A2*1.2, 2)\"",
                        "xai_result({'filled': str(rng.address)})")
                },
                new VbaPresetTemplate
                {
                    Name = "py_clean_text",
                    Description = "批量去空格（xlwings·Excel 模式）",
                    Version = "1.0",
                    Tags = new[] { "python", "清洗", "文本" },
                    Code = Produce(
                        "import xlwings as xw",
                        "sht = xw.books.active.sheets.active",
                        "rng = sht.range(\"A1:A100\")",
                        "raw = rng.value or []",
                        "cleaned = [[(str(c).strip() if c is not None else None) for c in row] for row in raw]",
                        "rng.value = cleaned",
                        "xai_result({'cleaned': True})")
                },
                new VbaPresetTemplate
                {
                    Name = "py_chart_selection",
                    Description = "按选区画图（支持类型与标题）",
                    Version = "1.1",
                    Tags = new[] { "python", "图表", "可视化" },
                    Code = Produce(
                        "import xlwings as xw",
                        "import re",
                        "sht = xw.books.active.sheets.active",
                        "sel_range = globals().get('sel_range') or 'A1:B6'",
                        "chart_type = globals().get('chart_type') or 'column_clustered'",
                        "chart_title = globals().get('chart_title') or ''",
                        "# 多选区取首块，避免把逗号串整段传给 range() 解析错位",
                        "addr = str(sel_range).split(',')[0].strip()",
                        "# 解析带表名地址：'表名'!A1:B6 → 定位到对应表再取区（跨表画图），否则用活动表",
                        "m = re.match(r\"^'([^']+)'!(.+)$\", addr) or re.match(r'^([^!]+)!(.+)$', addr)",
                        "if m:",
                        "    data = xw.books.active.sheets[m.group(1)].range(m.group(2).strip())",
                        "else:",
                        "    data = sht.range(addr)",
                        "chart = sht.charts.add()",
                        "chart.set_source_data(data)",
                        "chart.chart_type = chart_type",
                        "if chart_title:",
                        "    ch_obj = chart.api[0].Chart  # api 是 (ChartObject, Chart) 元组，[0].Chart 为图表对象",
                        "    ch_obj.HasTitle = True",
                        "    ch_obj.ChartTitle.Text = str(chart_title)",
                        "xai_result({'chart_name': chart.name, 'chart_type': chart_type, 'used_range': data.address, 'title_set': bool(chart_title)})")
                },
                new VbaPresetTemplate
                {
                    Name = "py_calc_stats",
                    Description = "纯计算均值/中位数（纯计算模式）",
                    Version = "1.0",
                    Tags = new[] { "python", "统计", "纯计算" },
                    Code = Produce(
                        "import statistics",
                        "data = [12, 34, 56, 78, 90]",
                        "xai_result({'mean': statistics.mean(data), 'median': statistics.median(data), 'max': max(data)})")
                }
            };
        }

        private static string Produce(params string[] lines)
        {
            return string.Join("\n", lines);
        }
    }
}