using System;
using System.Collections.Generic;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Operations;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// JYYJtool 借鉴工具注册：③ 文本语义 / 数字 / 行列表格 / 工作表 / 复制粘贴增强。
    /// </summary>
    public static class ExcelJyyjtoolRegistration
    {
        private static readonly Func<Dictionary<string, object>, object[,], string, string, ToolResult> NoWriter =
            (_1, _2, _3, _4) => throw new InvalidOperationException("缺少写入回调。");

        public static void RegisterJyyjtoolTools(
            ToolRegistry registry,
            dynamic application,
            Func<Dictionary<string, object>, object[,], string, string, ToolResult> writer,
            OperationRecorder recorder)
        {
            var srv = new ExcelJyyjtoolService(application, writer ?? NoWriter, recorder);

            // ---- ③ 文本语义 ----
            Register(registry, "extract_text_characters", "提取单元格文本中的指定类型字符（中文/英文/数字/字母数字/特殊符号）。", RiskLevel.NormalWrite,
                () => srv.ExtractTextCharacters, JyyjtoolSchemas.ExtractTextCharacters());
            Register(registry, "insert_text_into_cell", "按字符位置或某个分隔符第 N 次出现的位置，在文本中插入子串。", RiskLevel.NormalWrite,
                () => srv.InsertTextIntoCell, JyyjtoolSchemas.InsertTextIntoCell());
            Register(registry, "delete_text_matches", "按正则表达式删除匹配文本，其余保留。", RiskLevel.NormalWrite,
                () => srv.DeleteTextMatches, JyyjtoolSchemas.DeleteTextMatches());
            Register(registry, "convert_text_to_pinyin", "把中文文本转成拼音全拼或首字母。", RiskLevel.NormalWrite,
                () => srv.ConvertTextToPinyin, JyyjtoolSchemas.ConvertTextToPinyin());
            Register(registry, "convert_chinese_variant", "简繁中文之间互相转换。", RiskLevel.NormalWrite,
                () => srv.ConvertChineseVariant, JyyjtoolSchemas.ConvertChineseVariant());
            Register(registry, "amount_to_uppercase", "把数值金额转换为中文大写金额（壹贰叁…元角分）。", RiskLevel.NormalWrite,
                () => srv.AmountToUppercase, JyyjtoolSchemas.AmountToUppercase());

            // ---- ③ 数字 ----
            Register(registry, "bankers_round_values", "按银行家舍入（四舍六入五取偶）保留指定小数位。", RiskLevel.NormalWrite,
                () => srv.BankersRoundValues, JyyjtoolSchemas.BankersRoundValues());
            Register(registry, "pick_numbers_to_sum", "凑数：从候选数字中找一组合尽量接近目标值（支持精确组合与贪心近似）。", RiskLevel.ReadOnly,
                () => srv.PickNumbersToSum, JyyjtoolSchemas.PickNumbersToSum());
            Register(registry, "fill_random_numbers", "向指定区域填充随机数，可设值域与小数位。", RiskLevel.NormalWrite,
                () => srv.FillRandomNumbers, JyyjtoolSchemas.FillRandomNumbers());
            Register(registry, "fill_number_sequence", "向下或向右填充公差为 step 的等差序列。", RiskLevel.NormalWrite,
                () => srv.FillNumberSequence, JyyjtoolSchemas.FillNumberSequence());
            Register(registry, "generate_permutation_list", "生成 n 个元素的排列（或两两组合）并写入单元格。", RiskLevel.NormalWrite,
                () => srv.GeneratePermutationList, JyyjtoolSchemas.GeneratePermutationList());

            // ---- ③ 行列表格 ----
            Register(registry, "insert_blank_rows_every_n", "每隔 N 行插入若干空行（工资条空格/分段）。", RiskLevel.StructuralChange,
                () => srv.InsertBlankRowsEveryN, JyyjtoolSchemas.InsertBlankRowsEveryN());
            Register(registry, "delete_blank_rows_range", "删除区域内整行为空或首列为空的行。", RiskLevel.StructuralChange,
                () => srv.DeleteBlankRowsRange, JyyjtoolSchemas.DeleteBlankRowsRange());
            Register(registry, "make_payslip_rows", "把表头与数据区生成为工资条：每行数据上方重复表头。", RiskLevel.NormalWrite,
                () => srv.MakePayslipRows, JyyjtoolSchemas.MakePayslipRows());
            Register(registry, "fill_blank_cells_down", "向下填充空单元格：沿用上方同列非空值（或无公式）。", RiskLevel.NormalWrite,
                () => srv.FillBlankCellsDown, JyyjtoolSchemas.FillBlankCellsDown());
            Register(registry, "merge_duplicate_cells", "合并一列内相邻的相同单元格（垂直合并同类项）。", RiskLevel.StructuralChange,
                () => srv.MergeDuplicateCells, JyyjtoolSchemas.MergeDuplicateCells());
            Register(registry, "fuzzy_match_columns", "两列做相似度模糊匹配，返回最佳配对与阈值判定。", RiskLevel.ReadOnly,
                () => srv.FuzzyMatchColumns, JyyjtoolSchemas.FuzzyMatchColumns());
            Register(registry, "wrap_formulas_with_iferror", "给区域内的公式加上 IFERROR 包裹，避免错误值残留。", RiskLevel.NormalWrite,
                () => srv.WrapFormulasWithIfError, JyyjtoolSchemas.WrapFormulasWithIfError());
            Register(registry, "convert_2d_to_1d", "宽表转长表：保留前 N 列标识，其余列拆成多行（逆透视）。", RiskLevel.NormalWrite,
                () => srv.Convert2dTo1d, JyyjtoolSchemas.Convert2dTo1d());
            Register(registry, "convert_1d_to_2d", "一维数据按每行固定列数折叠为二维表。", RiskLevel.NormalWrite,
                () => srv.Convert1dTo2d, JyyjtoolSchemas.Convert1dTo2d());
            Register(registry, "create_pivot_table", "对明细数据做轻量透视（按行字段分组并对数值列求和/计数/平均）。", RiskLevel.NormalWrite,
                () => srv.CreatePivotTable, JyyjtoolSchemas.CreatePivotTable());
            Register(registry, "cross_reference_columns", "跨区键值匹配回填：以源区域键列去查找区域匹配，命中则把查找区域值列的值回填到源区域目标列（免 VLOOKUP 公式）。", RiskLevel.NormalWrite,
                () => srv.CrossReferenceColumns, JyyjtoolSchemas.CrossReferenceColumns());

            // ---- ③ 工作表 ----
            Register(registry, "split_sheet_by_field", "按字段取值把一张表拆分成多个新工作表（各含子集数据）。", RiskLevel.StructuralChange,
                () => srv.SplitSheetByField, JyyjtoolSchemas.SplitSheetByField());
            Register(registry, "create_sheet_by_value", "以指定单元格的值作为新工作表明称（可复制某表内容）。", RiskLevel.StructuralChange,
                () => srv.CreateSheetByValue, JyyjtoolSchemas.CreateSheetByValue());
            Register(registry, "compare_two_sheets", "逐格对比两张表，输出新增/删除/变更差异。", RiskLevel.ReadOnly,
                () => srv.CompareTwoSheets, JyyjtoolSchemas.CompareTwoSheets());
            Register(registry, "protect_sheet_watermark", "锁定/解锁工作表，或在指定单元格写入水印文本。", RiskLevel.StructuralChange,
                () => srv.ProtectSheetWatermark, JyyjtoolSchemas.ProtectSheetWatermark());

            // ---- ③ 复制粘贴增强 ----
            Register(registry, "copy_range_as_image", "把区域复制为图片（静态快照）并粘贴到锚点。", RiskLevel.NormalWrite,
                () => srv.CopyRangeAsImage, JyyjtoolSchemas.CopyRangeAsImage());
            Register(registry, "copy_visible_range_only", "仅复制筛选后可见的行到目标区域，跳过隐藏行。", RiskLevel.NormalWrite,
                () => srv.CopyVisibleRangeOnly, JyyjtoolSchemas.CopyVisibleRangeOnly());
            Register(registry, "smart_paste_delimiter", "智能粘贴：把一列含分隔符的文本拆成多列（自动识别分隔符）。", RiskLevel.NormalWrite,
                () => srv.SmartPasteDelimiter, JyyjtoolSchemas.SmartPasteDelimiter());
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
            registry.Register(new ToolDescriptor(name, description, risk, args => Execute(() => execute(args)), parameters));
        }

        private static ToolResult Execute(Func<ToolResult> action)
        {
            return ExcelToolHost.Execute(action);
        }
    }
}