using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 结果自检工具注册（只读，不修改数据）：覆盖写入/公式/图表/透视/排序筛选/结构/清洗/
    /// 单元格格式/条件格式/标记/迷你图/数据验证/合并/智能表/保护/图片/冻结窗格/超链接 18 个维度。
    /// 复用 verify_chart 的「预期值 vs 实际值」自检模式，AI 操作后必须调用对应自检工具形成闭环。
    /// </summary>
    public static class ExcelVerifyRegistration
    {
        public static void RegisterVerifyTools(
            ToolRegistry registry,
            dynamic application,
            OperationRecorder recorder)
        {
            var verify = new ExcelVerifyService(application);

            Register(registry, "verify_cell_format",
                "单元格格式自检：读取目标区域的数字格式/对齐/字体（名称/字号/颜色/加粗/斜体）/边框/填充/锁定属性，与 AI 操作时传入的预期值（expected_*）逐项比对，返回 verdict/issues/selfcheck。设置任何单元格格式后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyCellFormat, VerifySchemas.VerifyCellFormat());

            Register(registry, "verify_conditional_format",
                "条件格式自检：读取目标区域的条件格式规则（类型/公式/运算符/填充色/字体色），与预期值（expected_rule/expected_formula1/expected_operator 等）比对，返回 verdict/issues/selfcheck。设置任何条件格式后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyConditionalFormat, VerifySchemas.VerifyConditionalFormat());

            Register(registry, "verify_sort_filter",
                "排序筛选自检：基于目标区域数据矩阵的实际值验证——检查排序列（expected_sort_column/direction）从表头行起的身体行是否按预期方向单调，以及每个数据行是否都满足筛选谓词（expected_filter_field/criteria1/operator）。与内存 sort_range / filter_excel_data 工具语义一致（不依赖原生 Sort/AutoFilter 对象）。返回 verdict/issues/selfcheck。执行排序或筛选后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifySortFilter, VerifySchemas.VerifySortFilter());

            Register(registry, "verify_formula",
                "公式自检：检查目标单元格是否含公式、是否返回错误值（#REF!/#VALUE! 等）、公式内容与预期是否一致、计算值与预期是否相符（数值按容差比较），返回 verdict/issues/selfcheck。写入任何公式后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyFormula, VerifySchemas.VerifyFormula());

            Register(registry, "verify_write_result",
                "写入结果自检：读取目标区域实际值，与预期值矩阵（expected_values）逐格比对（数值按容差、文本忽略大小写与空白），并可按 expected_count 核对非空单元格数，返回 verdict/issues/selfcheck。执行 write_excel_cell/write_excel_cells_batch/write_excel_range 等写入后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyWriteResult, VerifySchemas.VerifyWriteResult());

            Register(registry, "verify_data_clean",
                "数据清洗自检：对区域身体行（跳过表头）逐检查项扫描——no_blank_rows(无整行空白)/no_blanks(指定列无空白)/no_duplicates(无重复记录)/dates_valid(日期列文本可解析)/no_text_numbers(无文本型数字)，任一发现问题即报 issue，返回 verdict/issues/selfcheck。执行 clean_excel_data/dedup_range 等清洗后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyDataClean, VerifySchemas.VerifyDataClean());

            Register(registry, "verify_pivot",
                "透视表自检：读取透视表汇总输出区域（首行为表头 Row Labels/Total），按 expected_rows 中的标签逐行抽查汇总值，并可按 expected_count 核对数据行数，返回 verdict/issues/selfcheck。执行 create_excel_pivot_summary 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyPivot, VerifySchemas.VerifyPivot());

            Register(registry, "verify_marker",
                "标记自检：扫描目标区域单元格的非默认填充色（被标记），按 expected_count 核对标记总数、按 expected_fill_colors 核对实际标记色集合、按 expected_marked 核对期望被标记地址是否命中，返回 verdict/issues/selfcheck。执行 flag_cells 等标记操作后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyMarker, VerifySchemas.VerifyMarker());

            Register(registry, "verify_structure",
                "结构自检：读取工作簿工作表名单与可选区域行列数，与期望比对（expected_sheet_exists/expected_sheet_absent/expected_row_count/expected_column_count），返回 verdict/issues/selfcheck。新增/重命名/删除工作表等结构变更后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyStructure, VerifySchemas.VerifyStructure());

            Register(registry, "verify_sparkline",
                "迷你图自检：读取目标区域的 SparklineGroups（类型/落点/数据源），与预期（expected_type/expected_data_source）比对，返回 verdict/issues/selfcheck。执行 create_excel_sparkline 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifySparkline, VerifySchemas.VerifySparkline());

            Register(registry, "verify_data_validation",
                "数据验证自检：读取目标区域的验证规则（类型/来源/运算符/边界），与预期（expected_type/expected_source/expected_operator/expected_value1/value2）比对，返回 verdict/issues/selfcheck。执行 data_validation 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyDataValidation, VerifySchemas.VerifyDataValidation());

            Register(registry, "verify_merge",
                "合并状态自检：读取目标区域 MergeCells 实际状态（true 整区已合并/false 无合并/null 混合），与 expected_merged 比对，返回 verdict/issues/selfcheck。执行 merge_or_split_cells/merge_duplicate_cells 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyMerge, VerifySchemas.VerifyMerge());

            Register(registry, "verify_table",
                "智能表自检：按名称或地址定位 ListObject，核对存在性/表头/行列数（expected_headers/expected_row_count/expected_column_count），返回 verdict/issues/selfcheck。执行 create_excel_table 及拆表/建表类操作后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyTable, VerifySchemas.VerifyTable());

            Register(registry, "verify_protection",
                "保护状态自检（只读探测）：回报工作簿结构保护/工作表保护/单元格锁定与筛选排序许可，支持 expected_protected 与 expected_writable 期望比对；写入被拒（0x800A03EC）时的首选排查工具。",
                RiskLevel.ReadOnly, () => verify.VerifyProtection, VerifySchemas.VerifyProtection());

            Register(registry, "verify_image",
                "图片自检：枚举工作表中的图片（不含图表/文本框），按 expected_count 核对数量、按 anchor 定位锚点单元格内的图片并按 expected_width/expected_height 核对尺寸（磅，容差 2），返回 verdict/issues/selfcheck。执行 insert_image 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyImage, VerifySchemas.VerifyImage());

            Register(registry, "verify_freeze_panes",
                "冻结窗格自检：读取目标工作表视图的冻结状态与分割行列号，与 expected_frozen/expected_cell 比对，返回 verdict/issues/selfcheck。执行 freeze_panes 后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyFreezePanes, VerifySchemas.VerifyFreezePanes());

            Register(registry, "verify_file",
                "文件核验（只读）：只读取文件头 512 字节判断类型，不返回文件内容。核对导出/生成文件是否存在、大小是否达下限、魔数与扩展名是否匹配（pdf/xlsx/zip/png/jpeg/gif/text；xlsx 实为 zip 容器）。export_sheets_to_pdf / export_excel_chart_image 成功后自动带出本自检；也可手动核验本地绝对路径文件：仅读取存在性/大小/文件头 512B 类型元数据（不返回内容，拒绝 UNC/网络、设备、相对/.. 与符号链接路径）。返回 verdict（ok/missing/mismatch/not_a_file）与 issues；ok=false 表示判定未通过。",
                RiskLevel.ReadOnly, () => verify.VerifyFile, VerifySchemas.VerifyFile());

            Register(registry, "verify_hyperlink",
                "超链接自检：枚举工作表超链接（锚点/网址/内部位置/显示文本），按 expected_count 核对整表总数、按 address 定位目标单元格/区域并按 expected_url/expected_subaddress/expected_text 核对内容，或按 expected_present=false 断言删除生效，返回 verdict/issues/selfcheck。执行 manage_hyperlinks 的写动作后都应调用一次。",
                RiskLevel.ReadOnly, () => verify.VerifyHyperlink, VerifySchemas.VerifyHyperlink());
        }

        private static void Register(
            ToolRegistry registry,
            string name,
            string description,
            RiskLevel risk,
            Func<Func<Dictionary<string, object>, ToolResult>> target,
            Dictionary<string, object> parameters)
        {
            var execute = target();
            registry.Register(new ToolDescriptor(name, description, risk, args => ExcelToolHost.Execute(() => execute(args)), parameters));
        }
    }
}
