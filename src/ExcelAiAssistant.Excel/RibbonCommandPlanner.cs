using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAiAssistant.Core.Models;

namespace ExcelAiAssistant.Excel
{
    /// <summary>
    /// 「JYYJ快捷」功能区按钮 → 既有本地工具 的纯逻辑映射层（无 COM、无 UI，可单测）。
    /// 负责：安全分级判定、参数框字段模型、从参数值组装工具实参、去重删行数预算、结果摘要。
    /// 与既有工具执行体零耦合——只生成会被既有工具执行的参数字典，绝不改动工具本身。
    /// </summary>
    public enum CommandSafety
    {
        /// <summary>安全写回：不删行、不覆盖用户其他数据的本区写回（清洗/文本/透视写新区）。</summary>
        SafeWrite = 0,

        /// <summary>覆盖写回：结果覆盖源区数据（排序/筛选原地覆盖；转置原地重排）。</summary>
        Overwrite = 1,

        /// <summary>删行类结构性操作：物理删行（去重删重复行），需要二次确认。</summary>
        DeleteRows = 2
    }

    /// <summary>参数框下拉单项：Value 为对工具的实参值，Display 为界面展示文本。</summary>
    public sealed class ParamOption
    {
        public string Value { get; set; }
        public string Display { get; set; }
    }

    /// <summary>
    /// 参数框字段模型（纯数据，供静态 HTML 通用渲染）。
    /// Type 取值：select（单选下拉）、multiselect（勾选多选）、text（文本框）、checkbox（开关）、static（只读说明）。
    /// </summary>
    public sealed class ParamField
    {
        public string Name { get; set; }

        public string Label { get; set; }

        public string Type { get; set; }

        public List<ParamOption> Options { get; } = new List<ParamOption>();

        /// <summary>标量默认值（select/text 为字符串值，checkbox 为 "true"/"false"）。</summary>
        public string Default { get; set; }

        /// <summary>多选默认选中的 Value 列表。</summary>
        public List<string> Defaults { get; } = new List<string>();

        public bool Required { get; set; }

        public bool ReadOnly { get; set; }

        /// <summary>参数框内该字段下方的小字说明（告诉用户这个配置是做什么的）。</summary>
        public string Hint { get; set; }
    }

    /// <summary>当前矩形选区描述（地址=去空格的绝对地址；ColumnNames=每列表头，无表头回退“列N”）。</summary>
    public sealed class ParamSelection
    {
        public string SheetName { get; set; }

        public string Address { get; set; }

        public int Rows { get; set; }

        public int Cols { get; set; }

        public List<string> ColumnNames { get; } = new List<string>();
    }

    /// <summary>
    /// JYYJ快捷 按钮 → 工具参数的纯逻辑规划器。所有方法为纯函数，便于对抗性审查与回归测试锁死判定。
    /// </summary>
    public static class RibbonCommandPlanner
    {
        /// <summary>需求文档 §5 锁定的安全分级（权威）：去重是唯一删行类；排序/筛选/转置原地覆盖为覆盖写回；其余安全写回。</summary>
        public static CommandSafety Classify(string toolName)
        {
            switch (toolName)
            {
                case "dedup_range":
                case "delete_blank_rows_range":
                    // 物理删行：需二次确认（删空行为首列空/整行全空即删）。
                    return CommandSafety.DeleteRows;
                case "sort_range":
                case "filter_excel_data":
                case "transpose_range":
                case "insert_blank_rows_every_n":
                case "merge_duplicate_cells":
                    // 结构性改写：插入空行 / 合并单元格原地改变行结构，参数框确定即确认门（与排序/筛选同级）。
                    return CommandSafety.Overwrite;
                case "clean_excel_data":
                case "create_excel_pivot_summary":
                case "text_split":
                case "text_merge":
                case "text_clean":
                case "fill_formula":
                case "insert_text_into_cell":
                case "delete_text_matches":
                case "fill_random_numbers":
                case "fill_number_sequence":
                case "fill_blank_cells_down":
                    return CommandSafety.SafeWrite;
                default:
                    throw new ArgumentException($"未知的 JYYJ快捷 工具：{toolName}", nameof(toolName));
            }
        }

        /// <summary>是否需要在参数框内二次确认（仅删行类）。</summary>
        public static bool RequiresSecondaryConfirm(CommandSafety safety)
        {
            return safety == CommandSafety.DeleteRows;
        }

        /// <summary>覆盖写回类：返回参数框内黄字提示文案，含被覆盖单元格预算（无则返回 null）。</summary>
        public static string OverlayHint(CommandSafety safety, ParamSelection sel)
        {
            if (safety != CommandSafety.Overwrite)
            {
                return null;
            }
            // 用 long 防 int 溢出（整列选区 1048576×16384 远超 int.MaxValue）。
            var cells = (long)Math.Max(1, sel.Rows) * Math.Max(1, sel.Cols);
            return $"将覆盖原区域（约 {cells} 个单元格）";
        }

        /// <summary>删行类：返回参数框内红色二次警示文案。</summary>
        public static string DeleteWarning(int rowsToDelete)
        {
            return $"将删除 {rowsToDelete} 行，确定？";
        }

        /// <summary>解析绝对地址（如 F1 / $F$1 / F$1）为 1 基行列；非法返回 false。</summary>
        public static bool TryParseAddress(string address, out int row, out int col)
        {
            row = 0;
            col = 0;
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            var s = address.Replace("$", string.Empty).Trim().ToUpperInvariant();
            var letters = new string(s.TakeWhile(char.IsLetter).ToArray());
            var digits = new string(s.SkipWhile(char.IsLetter).ToArray());
            if (letters.Length == 0 || !int.TryParse(digits, out row) || row < 1)
            {
                return false;
            }

            col = 0;
            foreach (var ch in letters)
            {
                col = col * 26 + (ch - 'A' + 1);
            }
            return col >= 1 && col <= 16384;
        }

        /// <summary>
        /// 计算目标写入矩形（绝对 1 基坐标），供「目标带数据守卫」使用。
        /// 仅对目标可静态预算的工具返回 true：text_split（拆分带区）、text_merge（结果列）、transpose_range（指定位置模式）。
        /// 动态目标（pivot/filter 输出行数未知）与写回源区（transpose 原位）返回 false，不做守卫。
        /// </summary>
        public static bool TryGetTargetRect(string toolName, ParamSelection sel, object[,] matrix, IReadOnlyDictionary<string, object> args,
            int topRow, int startCol, out string sheetName, out int top, out int left, out int rows, out int cols)
        {
            sheetName = null;
            top = left = rows = cols = 0;
            switch (toolName)
            {
                case "text_split":
                {
                    var sourceColumn = Convert.ToInt32(args["column"]);
                    var delimiter = Convert.ToString(args["delimiter"]);
                    var band = SplitTargetColumns(matrix, sourceColumn, delimiter);
                    if (band == null || band.Count == 0)
                    {
                        return false;
                    }
                    // 用户显式指定结果起始列时，守卫检查该列起的带区（缺省 band[0]=源列右侧相邻列）。
                    var bandStartRel = args.TryGetValue("target_column", out var tc) && Convert.ToInt32(tc) >= 1
                        ? Convert.ToInt32(tc)
                        : band[0];
                    sheetName = sel.SheetName;
                    top = topRow;
                    left = startCol + bandStartRel - 1;
                    rows = sel.Rows;
                    cols = band.Count;
                    return true;
                }
                case "text_merge":
                {
                    sheetName = sel.SheetName;
                    top = topRow;
                    left = startCol + Convert.ToInt32(args["target_column"]) - 1;
                    rows = sel.Rows;
                    cols = 1;
                    return true;
                }
                case "transpose_range":
                {
                    if (!args.TryGetValue("target_address", out var addr))
                    {
                        return false; // 原位写回源区不守卫（源区本就有数据）
                    }
                    if (!TryParseAddress(Convert.ToString(addr), out var targetRow, out var targetCol))
                    {
                        return false; // 解析失败交给工具层校验
                    }
                    sheetName = args.TryGetValue("target_sheet", out var ts) ? Convert.ToString(ts) : sel.SheetName;
                    top = targetRow;
                    left = targetCol;
                    rows = sel.Cols;
                    cols = sel.Rows;
                    return true;
                }
                default:
                    return false;
            }
        }

        /// <summary>执行后自检覆盖清单（单一来源，供 RibbonCommandService.PostVerify 与测试锁定）。
        /// 覆盖执行语义高风险/可强断言的工具；新增此类工具时必须在此表态，否则不加自检。</summary>
        public static bool NeedsPostVerify(string toolName)
        {
            switch (toolName)
            {
                case "transpose_range":     // 维度 + 锚点格回读
                case "dedup_range":         // 删除行数 = 执行前预估
                case "delete_blank_rows_range":
                case "fill_random_numbers": // 回读值域 ∈ [min,max]
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>转置自检：目标区维度应为源区行列互换，且抽查锚点格与源转置一致（null=通过，非 null=失败详情）。</summary>
        public static string VerifyTransposeResult(object[,] source, object[,] target)
        {
            if (source == null || target == null)
            {
                return "自检读取失败";
            }

            var srcRows = source.GetLength(0);
            var srcCols = source.GetLength(1);
            if (target.GetLength(0) != srcCols || target.GetLength(1) != srcRows)
            {
                return $"转置维度不符：目标 {target.GetLength(0)}×{target.GetLength(1)}，应为 {srcCols}×{srcRows}";
            }

            var anchors = new[] { (0, 0), (0, srcCols - 1), (srcRows - 1, 0), (srcRows - 1, srcCols - 1), (srcRows / 2, srcCols / 2) };
            foreach (var (sr, sc) in anchors.Distinct())
            {
                var expect = CellText(source[sr, sc]);
                var actual = CellText(target[sc, sr]);
                if (!string.Equals(expect, actual, StringComparison.Ordinal))
                {
                    return $"锚点 ({sr + 1},{sc + 1}) 值不符：应为「{expect}」，实际「{actual}」";
                }
            }
            return null;
        }

        /// <summary>随机数自检：回读值应全部落在 [min,max]（含舍入容差；null=通过，非 null=失败详情）。</summary>
        public static string VerifyRandomRange(object[,] values, double min, double max, int decimals)
        {
            if (values == null)
            {
                return "自检读取失败";
            }

            // 容差覆盖工具按 decimals 四舍五入的偏差（round 后可能超出非整边界半个刻度）。
            var tol = Math.Pow(10, -Math.Max(0, decimals)) / 2;
            for (var r = 0; r < values.GetLength(0); r++)
            {
                for (var c = 0; c < values.GetLength(1); c++)
                {
                    var v = values[r, c];
                    if (v == null || v is DBNull)
                    {
                        return $"单元格 ({r + 1},{c + 1}) 为空";
                    }
                    if (!double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)
                        || d < min - tol || d > max + tol)
                    {
                        return $"单元格 ({r + 1},{c + 1}) 值「{v}」不在 [{min},{max}] 范围内";
                    }
                }
            }
            return null;
        }

        /// <summary>目标占用判定（0 基矩阵）：任一单元格有非空白值即视为已有数据，用于写前守卫。</summary>
        public static bool MatrixHasData(object[,] matrix)
        {
            if (matrix == null)
            {
                return false;
            }

            for (var r = 0; r < matrix.GetLength(0); r++)
            {
                for (var c = 0; c < matrix.GetLength(1); c++)
                {
                    var v = matrix[r, c];
                    if (v != null && !(v is DBNull) && !string.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.InvariantCulture)))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>随机类验收锚点：提取矩阵前 maxCount 个非空单元格文本（不足则取现有全部）。</summary>
        public static List<string> PickSampleValues(object[,] matrix, int maxCount)
        {
            var samples = new List<string>();
            if (matrix == null || maxCount <= 0)
            {
                return samples;
            }

            for (var r = 0; r < matrix.GetLength(0) && samples.Count < maxCount; r++)
            {
                for (var c = 0; c < matrix.GetLength(1) && samples.Count < maxCount; c++)
                {
                    var v = matrix[r, c];
                    if (v != null && !(v is DBNull))
                    {
                        samples.Add(Convert.ToString(v, CultureInfo.InvariantCulture));
                    }
                }
            }

            return samples;
        }

        private static string CellText(object v)
        {
            if (v == null || v is DBNull)
            {
                return string.Empty;
            }
            return v is double d ? d.ToString("R", CultureInfo.InvariantCulture) : Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        /// <summary>为指定工具构建参数框字段（列下拉由选区列名注入；默认值对齐需求文档 §3.3）。</summary>
        public static List<ParamField> BuildFields(string toolName, ParamSelection sel)
        {
            var colOptions = ColumnOptions(sel);
            var fields = new List<ParamField>();
            switch (toolName)
            {
                case "clean_excel_data":
                    fields.Add(Checkbox("remove_blank_rows", "去空行", true,
                        "勾选后删除整行全空或首列为空的行"));
                    fields.Add(Checkbox("fix_dates", "修日期", true,
                        "把可识别的日期文本统一转为日期值"));
                    fields.Add(Select("date_column", "日期列（修日期目标列）", colOptions, "1",
                        "指定「修日期」作用的列"));
                    fields.Add(Text("date_format", "日期格式（留空自动识别）", string.Empty,
                        hint: "如 yyyy-MM-dd 或 yyyy/M/d；留空则自动识别常见格式"));
                    fields.Add(Checkbox("fill_blanks", "填空白", true,
                        "把空白单元格填上指定内容"));
                    fields.Add(Select("fill_column", "填空列（全部列=全选区填空）", AllColumnsWith(sel), "0",
                        "「填空白」作用范围；选「全部列」则全选区填空"));
                    fields.Add(Text("fill_value", "填空填充值", "待补",
                        hint: "「填空白」使用的填充内容"));
                    fields.Add(Checkbox("trim", "去首尾空格", false,
                        "去除每个单元格首尾的空格"));
                    fields.Add(Checkbox("remove_non_printable", "去控制字符", false,
                        "删除制表符、换行等控制字符"));
                    break;

                case "dedup_range":
                    fields.Add(Select("scope", "判重范围",
                        new[] { new ParamOption { Value = "whole", Display = "整行判重" }, new ParamOption { Value = "columns", Display = "指定列判重" } },
                        "whole",
                        "整行判重=整行内容相同才删；指定列判重=仅按勾选的列比较"));
                    fields.Add(MultiSelect("columns", "判重列（按列判重时勾选）", colOptions, new List<string> { "1" },
                        "仅勾选的列参与判重，其余列不影响"));
                    fields.Add(Select("keep", "保留", new[] { new ParamOption { Value = "first", Display = "首次出现" }, new ParamOption { Value = "last", Display = "末次出现" } }, "first",
                        "重复行保留首次出现还是末次出现"));
                    fields.Add(Checkbox("has_header", "含表头", true,
                        "首行是表头则不参与判重、不会被删除"));
                    fields.Add(Checkbox("case_sensitive", "区分大小写", false,
                        "勾选后 A 与 a 视为不同内容，不判为重复"));
                    break;

                case "sort_range":
                    fields.Add(Select("sort_column", "排序列", colOptions, "1",
                        "按哪一列排序"));
                    fields.Add(Select("direction", "排序方向", new[] { new ParamOption { Value = "asc", Display = "升序" }, new ParamOption { Value = "desc", Display = "降序" } }, "asc",
                        "升序=从小到大；降序=从大到小"));
                    fields.Add(Select("sort_column2", "次排序列（可无）", WithNoneOption(sel), "0",
                        "主排序列相同的数据行，再按此列排序"));
                    fields.Add(Select("direction2", "次排序列方向", new[] { new ParamOption { Value = "asc", Display = "升序" }, new ParamOption { Value = "desc", Display = "降序" } }, "asc",
                        "次排序列的排序方向"));
                    fields.Add(Checkbox("has_header", "含表头", true,
                        "首行是表头则排序时不移动表头"));
                    break;

                case "transpose_range":
                    fields.Add(Select("target_mode", "写入位置",
                        new[] { new ParamOption { Value = "top_left", Display = "原区域左上角" }, new ParamOption { Value = "custom", Display = "指定位置" } },
                        "top_left",
                        "原区域左上角=转置结果直接写回源区（覆盖原数据）；指定位置=写到新地址"));
                    fields.Add(Text("target_address", "目标左上角地址（选“指定位置”时填）", string.Empty,
                        hint: "如 F1；跨表写入请另填下方「目标工作表名」，不要把表名写进地址"));
                    fields.Add(Text("target_sheet", "目标工作表名（留空=同源表）", string.Empty,
                        hint: "跨表写入时填目标工作表名，如 Sheet2；配合「指定位置」使用"));
                    break;

                case "filter_excel_data":
                    fields.Add(Select("column", "条件列", colOptions, "1",
                        "按哪一列的条件筛选行"));
                    fields.Add(Select("operator", "运算符", OperatorOptions(), "not_empty",
                        "匹配条件：等于/包含/大于/小于/空/非空等"));
                    fields.Add(Text("value", "筛选值（“空/非空”无需填写）", string.Empty,
                        hint: "匹配值；运算符为「空/非空」时无需填写"));
                    fields.Add(Text("target_address", "结果写入地址（留空=原地覆盖）", string.Empty,
                        hint: "筛选结果写到该地址；留空则覆盖原选区数据（原数据会被替换）"));
                    break;

                case "create_excel_pivot_summary":
                    var valueOptions = new List<ParamOption> { new ParamOption { Value = "0", Display = "自动探测首个数值列" } };
                    valueOptions.AddRange(colOptions);
                    fields.Add(Select("label_column", "行标签列", colOptions, "1",
                        "作为分组行标签的列，按该列的值分组统计"));
                    fields.Add(Select("value_column", "数值列", valueOptions, "0",
                        "被统计的数值列；选「自动探测」则取选区内首个数值列"));
                    fields.Add(Text("target_address", "输出左上角地址（留空=自动）", string.Empty,
                        hint: "汇总结果写入的左上角地址；留空自动放在源区域下方 2 行"));
                    break;

                case "text_split":
                    fields.Add(Select("column", "源列", colOptions, "1",
                        "要拆分的源列"));
                    fields.Add(Text("delimiter", "分隔符", ",",
                        hint: "按此分隔符拆分，如逗号 、竖线 |"));
                    fields.Add(Text("target_column", "结果起始列号（留空=源列右侧）", string.Empty,
                        hint: "拆分结果从第几列开始写（相对选区，1 起）；留空则从源列右侧相邻列开始"));
                    fields.Add(Checkbox("keep_source", "保留源列", true,
                        "保留源列，拆分结果写到右侧新列"));
                    break;

                case "text_merge":
                    fields.Add(MultiSelect("columns", "参与合并列（多选）", colOptions, new List<string> { "1", "2" },
                        "勾选参与合并的列，按行顺序拼接成一个单元格"));
                    fields.Add(Select("target_column", "结果写入列", WithNewColumn(sel), (sel.Cols + 1).ToString(CultureInfo.InvariantCulture),
                        "合并结果写入哪一列；默认写到选区右侧新列"));
                    fields.Add(Text("delimiter", "分隔符", " ",
                        hint: "合并后单元格之间的连接符"));
                    fields.Add(Checkbox("skip_empty", "跳过空单元格", true,
                        "跳过空单元格，避免产生多余连接符"));
                    break;

                case "text_clean":
                    fields.Add(Checkbox("trim", "去首尾空格", true,
                        "去除每个单元格首尾的空格"));
                    fields.Add(Checkbox("collapse", "合并连续空白", true,
                        "把单元格内连续多个空格合并为一个"));
                    fields.Add(Checkbox("remove_non_printable", "去控制字符", false,
                        "删除制表符、换行等控制字符"));
                    fields.Add(Select("case", "大小写", new[] { new ParamOption { Value = "none", Display = "不变" }, new ParamOption { Value = "upper", Display = "转大写" }, new ParamOption { Value = "lower", Display = "转小写" } }, "none",
                        "统一字母大小写；选「不变」则不改动"));
                    fields.Add(Text("remove_chars", "删除指定字符（如 -（））", string.Empty,
                        hint: "删除单元格中出现的这些字符；留空则不删除"));
                    fields.Add(Text("find", "查找文本（可选）", string.Empty,
                        hint: "在单元格中查找该文本；配合下方「替换为」做查找替换"));
                    fields.Add(Text("replace", "替换为（留空=删除）", string.Empty,
                        hint: "查找到的文本替换成的内容；留空则直接删除匹配文本"));
                    break;

                case "fill_formula":
                    fields.Add(Text("formula", "公式模板（{row} 表示当前行号）", "=A{row}*B{row}", true,
                        "公式模板，{row} 会被替换为当前行号；如 =A{row}*B{row}"));
                    fields.Add(Text("start_row", "起始行号（留空自动）", string.Empty,
                        hint: "从第几行开始填充；留空则从选区首行开始"));
                    break;

                // ================= 常用组（ExcelJyyjtool） =================

                case "insert_text_into_cell":
                    fields.Add(Text("insertion", "要插入的子串", string.Empty, true,
                        "要插入到单元格文本中的子串"));
                    fields.Add(Select("when", "定位方式",
                        new[] { new ParamOption { Value = "char", Display = "按字符位置" }, new ParamOption { Value = "separator", Display = "按分隔符第 N 次出现" } },
                        "char",
                        "按字符位置定位，或按分隔符第 N 次出现的位置定位"));
                    fields.Add(Select("where", "插入位置",
                        new[] { new ParamOption { Value = "before", Display = "之前" }, new ParamOption { Value = "after", Display = "之后" } },
                        "after",
                        "在定位点之前或之后插入"));
                    fields.Add(Text("index", "字符下标（0 起）或分隔符第 N 次（1 起）", "0",
                        hint: "按字符定位填下标（0 起）；按分隔符定位填第 N 次（1 起）"));
                    fields.Add(Text("separator", "分隔符（按分隔符定位时填写）", string.Empty,
                        hint: "按分隔符定位时填写，如 - 或 /"));
                    break;

                case "delete_text_matches":
                    fields.Add(Text("pattern", "要删除的正则表达式（.NET 语法）", string.Empty, true,
                        "匹配即删除，.NET 正则语法；如 \\d+ 删除所有数字"));
                    break;

                case "fill_random_numbers":
                    fields.Add(Text("min", "最小值", "0",
                        hint: "随机数最小值（含）"));
                    fields.Add(Text("max", "最大值", "100",
                        hint: "随机数最大值（含）"));
                    fields.Add(Text("decimals", "小数位（0-6）", "0",
                        hint: "保留的小数位，范围 0-6"));
                    break;

                case "fill_number_sequence":
                    fields.Add(Text("start", "起始值（留空=用当前单元格值）", string.Empty,
                        hint: "序列起始值；留空则取选区当前值（多格选区取首行首列）"));
                    fields.Add(Text("step", "公差（步长）", "1",
                        hint: "相邻两个数的差，可为负数"));
                    fields.Add(Text("count", "填充数量", "10",
                        hint: "填充个数，上限 10000"));
                    fields.Add(Select("direction", "填充方向",
                        new[] { new ParamOption { Value = "down", Display = "向下" }, new ParamOption { Value = "right", Display = "向右" } },
                        "down",
                        "向下填充（整列）或向右填充（整行）"));
                    break;

                case "insert_blank_rows_every_n":
                    fields.Add(Text("interval", "每隔 N 行", "1",
                        hint: "每隔 N 行插入一次空行"));
                    fields.Add(Text("count", "插入空行数", "1",
                        hint: "每次插入几行空行"));
                    break;

                case "delete_blank_rows_range":
                    fields.Add(Checkbox("only_fully_blank", "仅删除整行全为空的行", false,
                        "勾选后仅删除整行全空的行；不勾选则删除首列为空的行"));
                    break;

                case "fill_blank_cells_down":
                    fields.Add(Checkbox("copy_value", "沿用上方值（否=复制公式引用）", true,
                        "勾选=空单元格沿用上方单元格的值；不勾选=复制上方公式引用"));
                    break;

                case "merge_duplicate_cells":
                    fields.Add(Checkbox("first_only", "仅保留每组首个值", true,
                        "合并相邻相同单元格；勾选时把每组首个值写入合并格"));
                    break;

                default:
                    throw new ArgumentException($"未知的 JYYJ快捷 工具：{toolName}", nameof(toolName));
            }

            return fields;
        }

        /// <summary>
        /// 校验参数框采集值并组装可执行的工具实参。返回 ToolResult 风格错误（Ok=false, Data=错误文案）或正常参数字典。
        /// 约定 values[name] = string（select/text）或 List&lt;string&gt;（multiselect）。
        /// </summary>
        public static object BuildArgs(string toolName, ParamSelection sel, IReadOnlyDictionary<string, object> values)
        {
            string S(string name, string def)
            {
                return values != null && values.TryGetValue(name, out var v) ? Convert.ToString(v, CultureInfo.InvariantCulture) ?? def : def;
            }

            bool B(string name, bool def)
            {
                if (values != null && values.TryGetValue(name, out var v) && v is string sv && bool.TryParse(sv, out var b))
                {
                    return b;
                }
                return def;
            }

            string[] M(string name)
            {
                if (values != null && values.TryGetValue(name, out var v))
                {
                    if (v is List<string> list)
                    {
                        return list.ToArray();
                    }
                    if (v is string s)
                    {
                        return string.IsNullOrWhiteSpace(s) ? new string[0] : new[] { s };
                    }
                }
                return new string[0];
            }

            int ColIndex(string raw, string fieldName)
            {
                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
                {
                    return n;
                }
                return -1;
            }

            int headerRows = B("has_header", true) ? 1 : 0;

            var args = new Dictionary<string, object>
            {
                ["address"] = sel.Address,
                ["sheet"] = sel.SheetName
            };

            switch (toolName)
            {
                case "clean_excel_data":
                {
                    var ops = new List<Dictionary<string, object>>();
                    if (B("remove_blank_rows", true))
                    {
                        ops.Add(new Dictionary<string, object> { ["tool"] = "remove_blank_rows" });
                    }
                    if (B("fix_dates", true))
                    {
                        var dc = ColIndex(S("date_column", "1"), "date_column");
                        if (dc < 1 || dc > sel.Cols)
                        {
                            return "日期列超出选区列范围。";
                        }
                        var op = new Dictionary<string, object> { ["tool"] = "fix_dates", ["column"] = dc };
                        var dateFormat = S("date_format", string.Empty).Trim();
                        if (!string.IsNullOrEmpty(dateFormat))
                        {
                            op["date_format"] = dateFormat;
                        }
                        ops.Add(op);
                    }
                    if (B("fill_blanks", true))
                    {
                        var fc = ColIndex(S("fill_column", "0"), "fill_column");
                        if (fc < 0 || fc > sel.Cols)
                        {
                            return "填空列超出选区列范围。";
                        }
                        var op = new Dictionary<string, object> { ["tool"] = "fill_blanks", ["value"] = S("fill_value", "待补") };
                        if (fc > 0)
                        {
                            // fill_blanks 的 column 缺省为 null=全部列；显式传 0 会得到 start=-1 越界，故仅 >0 时注入
                            op["column"] = fc;
                        }
                        ops.Add(op);
                    }
                    if (B("trim", false))
                    {
                        ops.Add(new Dictionary<string, object> { ["tool"] = "trim" });
                    }
                    if (B("remove_non_printable", false))
                    {
                        ops.Add(new Dictionary<string, object> { ["tool"] = "remove_non_printable" });
                    }
                    if (ops.Count == 0)
                    {
                        return "请至少勾选一项清洗操作。";
                    }
                    args["operations"] = ops;
                    return args;
                }

                case "dedup_range":
                {
                    var scope = S("scope", "whole");
                    if (string.Equals(scope, "columns", StringComparison.OrdinalIgnoreCase))
                    {
                        var cols = M("columns").Select(c => ColIndex(c, "columns")).Where(c => c >= 1).ToList();
                        if (cols.Count == 0)
                        {
                            return "请至少勾选一个判重列。";
                        }
                        if (cols.Any(c => c > sel.Cols))
                        {
                            return "判重列超出选区列范围。";
                        }
                        args["columns"] = cols;
                    }

                    args["keep"] = S("keep", "first");
                    args["header_rows"] = headerRows;
                    args["case_sensitive"] = B("case_sensitive", false);
                    return args;
                }

                case "sort_range":
                {
                    var c = ColIndex(S("sort_column", "1"), "sort_column");
                    if (c < 1 || c > sel.Cols)
                    {
                        return "排序列超出选区列范围。";
                    }
                    args["sort_column"] = c;
                    var direction = S("direction", "asc");
                    if (direction != "asc" && direction != "desc")
                    {
                        return "排序方向无效（应为 asc 或 desc）。";
                    }
                    args["direction"] = direction;
                    var c2 = ColIndex(S("sort_column2", "0"), "sort_column2");
                    if (c2 < 0 || c2 > sel.Cols)
                    {
                        return "次排序列超出选区列范围。";
                    }
                    if (c2 > 0)
                    {
                        var d2 = S("direction2", "asc");
                        if (d2 != "asc" && d2 != "desc")
                        {
                            return "次排序列方向无效（应为 asc 或 desc）。";
                        }
                        args["sort_column2"] = c2;
                        args["direction2"] = d2;
                    }
                    args["header_rows"] = headerRows;
                    return args;
                }

                case "transpose_range":
                {
                    var customMode = string.Equals(S("target_mode", "top_left"), "custom", StringComparison.OrdinalIgnoreCase);
                    var targetSheet = S("target_sheet", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(targetSheet) && !customMode)
                    {
                        // 「原区域左上角」写回源区，填目标工作表名无意义且工具层行为未定义，规划层直接拦截。
                        return "目标工作表名需配合「指定位置」模式使用。";
                    }
                    if (customMode)
                    {
                        var addr = S("target_address", string.Empty);
                        if (string.IsNullOrWhiteSpace(addr))
                        {
                            return "请填写目标左上角地址。";
                        }
                        args["target_address"] = addr.Trim();
                    }
                    if (!string.IsNullOrEmpty(targetSheet))
                    {
                        args["target_sheet"] = targetSheet;
                    }
                    return args;
                }

                case "filter_excel_data":
                {
                    var c = ColIndex(S("column", "1"), "column");
                    if (c < 1 || c > sel.Cols)
                    {
                        return "条件列超出选区列范围。";
                    }
                    var op = S("operator", "not_empty");
                    if (op != "equals" && op != "not_equals" && op != "contains" && op != "gt" && op != "gte" && op != "lt" && op != "lte" && op != "empty" && op != "not_empty")
                    {
                        return "运算符无效（应为 equals/not_equals/contains/gt/gte/lt/lte/empty/not_empty）。";
                    }
                    var value = S("value", string.Empty);
                    bool needsValue = !(op == "empty" || op == "not_empty");
                    if (needsValue && string.IsNullOrEmpty(value))
                    {
                        return "请填写筛选值。";
                    }
                    args["column"] = c;
                    args["operator"] = op;
                    args["value"] = value;
                    var targetAddress = S("target_address", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(targetAddress))
                    {
                        args["target_address"] = targetAddress;
                    }
                    return args;
                }

                case "create_excel_pivot_summary":
                {
                    var label = ColIndex(S("label_column", "1"), "label_column");
                    if (label < 1 || label > sel.Cols)
                    {
                        return "行标签列超出选区列范围。";
                    }
                    var value = ColIndex(S("value_column", "0"), "value_column");
                    if (value < 0 || value > sel.Cols)
                    {
                        return "数值列超出选区列范围。";
                    }
                    var pivotArgs = new Dictionary<string, object>
                    {
                        ["source_address"] = sel.Address,
                        ["label_column"] = label,
                        ["value_column"] = value
                    };
                    var pivotTarget = S("target_address", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(pivotTarget))
                    {
                        pivotArgs["target_address"] = pivotTarget;
                    }
                    return pivotArgs;
                }

                case "text_split":
                {
                    var c = ColIndex(S("column", "1"), "column");
                    if (c < 1 || c > sel.Cols)
                    {
                        return "源列超出选区列范围。";
                    }
                    var delimiter = S("delimiter", ",");
                    if (string.IsNullOrEmpty(delimiter))
                    {
                        return "请填写分隔符。";
                    }
                    var targetColumnRaw = S("target_column", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(targetColumnRaw))
                    {
                        if (!int.TryParse(targetColumnRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var targetColumn) || targetColumn < 1 || targetColumn > 16384)
                        {
                            return "结果起始列号无效（应为 1..16384 的整数）。";
                        }
                        args["target_column"] = targetColumn;
                    }
                    args["column"] = c;
                    args["delimiter"] = delimiter;
                    args["keep_source"] = B("keep_source", true);
                    return args;
                }

                case "text_merge":
                {
                    var cols = M("columns").Select(c => ColIndex(c, "columns")).Where(c => c >= 1).ToList();
                    if (cols.Count == 0)
                    {
                        return "请至少勾选一个参与合并的列。";
                    }
                    if (cols.Any(c => c > sel.Cols))
                    {
                        return "参与合并的列超出选区列范围。";
                    }
                    var target = ColIndex(S("target_column", (sel.Cols + 1).ToString(CultureInfo.InvariantCulture)), "target_column");
                    if (target < 1 || target > 16384)
                    {
                        return "结果写入列号无效（应在 1..16384）。";
                    }
                    args["columns"] = cols;
                    args["target_column"] = target;
                    args["delimiter"] = S("delimiter", " ");
                    args["skip_empty"] = B("skip_empty", true);
                    return args;
                }

                case "text_clean":
                {
                    args["trim"] = B("trim", true);
                    args["collapse_whitespace"] = B("collapse", true);
                    args["remove_non_printable"] = B("remove_non_printable", false);
                    args["case"] = S("case", "none");
                    var removeChars = S("remove_chars", string.Empty);
                    if (!string.IsNullOrEmpty(removeChars))
                    {
                        args["remove_chars"] = removeChars;
                    }
                    var find = S("find", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(find))
                    {
                        args["find"] = find;
                    }
                    var replace = S("replace", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(replace))
                    {
                        args["replace"] = replace;
                    }
                    return args;
                }

                case "fill_formula":
                {
                    var formula = S("formula", string.Empty);
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        return "请填写公式模板。";
                    }
                    args["formula"] = formula.Trim();
                    var startRaw = S("start_row", string.Empty);
                    if (int.TryParse(startRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var startRow) && startRow > 0)
                    {
                        args["start_row"] = startRow;
                    }
                    return args;
                }

                // ================= 常用组（ExcelJyyjtool） =================

                case "insert_text_into_cell":
                {
                    var insertion = S("insertion", string.Empty).Trim();
                    if (string.IsNullOrEmpty(insertion))
                    {
                        return "请填写要插入的子串。";
                    }
                    var when = S("when", "char");
                    if (when != "char" && when != "separator")
                    {
                        return "定位方式无效（应为 char 或 separator）。";
                    }
                    var where = S("where", "after");
                    if (where != "before" && where != "after")
                    {
                        return "插入位置无效（应为 before 或 after）。";
                    }
                    var indexText = S("index", when == "char" ? "0" : "1");
                    if (!int.TryParse(indexText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) || index < 0)
                    {
                        return "字符下标/出现次数应为非负整数。";
                    }
                    args["insertion"] = insertion;
                    args["when"] = when;
                    args["where"] = where;
                    args["index"] = index;
                    if (when == "separator")
                    {
                        var separator = S("separator", string.Empty);
                        if (string.IsNullOrEmpty(separator))
                        {
                            return "请填写分隔符（when=separator）。";
                        }
                        args["separator"] = separator;
                    }
                    return args;
                }

                case "delete_text_matches":
                {
                    var pattern = S("pattern", string.Empty).Trim();
                    if (string.IsNullOrEmpty(pattern))
                    {
                        return "请填写要删除的正则表达式。";
                    }
                    try
                    {
                        // 与工具层一致的编译参数预检：非法正则（如 "["）提前拦截，避免静默不删除。
                        _ = new Regex(pattern, RegexOptions.CultureInvariant | RegexOptions.Compiled, TimeSpan.FromMilliseconds(2000));
                    }
                    catch (ArgumentException ex)
                    {
                        return $"正则表达式无效：{ex.Message}";
                    }
                    args["pattern"] = pattern;
                    return args;
                }

                case "fill_random_numbers":
                {
                    if (!double.TryParse(S("min", "0"), NumberStyles.Float, CultureInfo.InvariantCulture, out var min) || double.IsNaN(min) || double.IsInfinity(min))
                    {
                        return "最小值无效。";
                    }
                    if (!double.TryParse(S("max", "100"), NumberStyles.Float, CultureInfo.InvariantCulture, out var max) || double.IsNaN(max) || double.IsInfinity(max))
                    {
                        return "最大值无效。";
                    }
                    if (min > max)
                    {
                        return "最小值不能大于最大值。";
                    }
                    if (!int.TryParse(S("decimals", "0"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var decimals) || decimals < 0 || decimals > 6)
                    {
                        return "小数位应为 0-6 的整数。";
                    }
                    args["min"] = min;
                    args["max"] = max;
                    args["decimals"] = decimals;
                    return args;
                }

                case "fill_number_sequence":
                {
                    if (!double.TryParse(S("step", "1"), NumberStyles.Float, CultureInfo.InvariantCulture, out var step) || double.IsNaN(step) || double.IsInfinity(step))
                    {
                        return "公差（步长）无效。";
                    }
                    if (!int.TryParse(S("count", "10"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var count) || count < 1)
                    {
                        return "填充数量应为正整数。";
                    }
                    if (count > 10000)
                    {
                        return "填充数量过大（上限 10000）。";
                    }
                    var direction = S("direction", "down");
                    if (direction != "down" && direction != "right")
                    {
                        return "填充方向无效（应为 down 或 right）。";
                    }
                    var startRaw = S("start", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(startRaw))
                    {
                        if (!double.TryParse(startRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out var start) || double.IsNaN(start) || double.IsInfinity(start))
                        {
                            return "起始值无效。";
                        }
                        args["start"] = start;
                    }
                    args["step"] = step;
                    args["count"] = count;
                    args["direction"] = direction;
                    return args;
                }

                case "insert_blank_rows_every_n":
                {
                    if (!int.TryParse(S("interval", "1"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var interval) || interval < 1)
                    {
                        return "间隔 N 应为正整数。";
                    }
                    if (!int.TryParse(S("count", "1"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var insertCount) || insertCount < 1)
                    {
                        return "插入空行数应为正整数。";
                    }
                    // 工具层无上限：每隔 N 行插 count 行，按选区行数预算插入总量，超限拒绝（Excel 行上限 1048576，留余量）。
                    var budget = (sel.Rows / interval + 1) * insertCount;
                    if (budget > 100000)
                    {
                        return $"插入空行数过多（预算约 {budget} 行，上限 100000 行），请调小「每隔 N 行」或「插入空行数」。";
                    }
                    args["interval"] = interval;
                    args["count"] = insertCount;
                    return args;
                }

                case "delete_blank_rows_range":
                    args["only_fully_blank"] = B("only_fully_blank", false);
                    return args;

                case "fill_blank_cells_down":
                    args["copy_value"] = B("copy_value", true);
                    return args;

                case "merge_duplicate_cells":
                    args["first_only"] = B("first_only", true);
                    return args;

                default:
                    throw new ArgumentException($"未知的 JYYJ快捷 工具：{toolName}", nameof(toolName));
            }
        }

        /// <summary>
        /// 预算去重将删除的行数（不含表头），与 ExcelDataToolsService.DedupRange 的 DedupKey 语义严格一致。
        /// matrix 为选区原始值（含表头；数字为 double，空为 null/DBNull）。columns 为 1 基判重列；空集合=整行判重。
        /// </summary>
        public static int CountRemovedRows(object[,] matrix, int headerRows, IReadOnlyList<int> columns, bool keepLast, bool caseSensitive)
        {
            if (matrix == null)
            {
                return 0;
            }

            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            var colSet = new HashSet<int>(columns ?? new int[0]);
            Func<int, bool> useCol = c => colSet.Count == 0 || colSet.Contains(c + 1);
            var bodyStart = Math.Min(Math.Max(0, headerRows), rows);
            var bodyCount = rows - bodyStart;
            if (bodyCount <= 0)
            {
                return 0;
            }

            Func<int, string> key = i =>
            {
                var sb = new System.Text.StringBuilder();
                for (var c = 0; c < cols; c++)
                {
                    if (!useCol(c))
                    {
                        continue;
                    }
                    var v = matrix[i, c];
                    string s;
                    if (v == null || v is DBNull)
                    {
                        s = "\u0000";
                    }
                    else if (v is double d)
                    {
                        s = d.ToString("R", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        s = Convert.ToString(v, CultureInfo.InvariantCulture);
                    }
                    if (!caseSensitive)
                    {
                        s = s.ToLowerInvariant();
                    }
                    sb.Append(s).Append('\u0001');
                }
                return sb.ToString();
            };

            int survivors;
            if (keepLast)
            {
                var lastPos = new Dictionary<string, int>(StringComparer.Ordinal);
                for (var i = bodyStart; i < rows; i++)
                {
                    lastPos[key(i)] = i;
                }
                survivors = lastPos.Count;
            }
            else
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                survivors = 0;
                for (var i = bodyStart; i < rows; i++)
                {
                    if (seen.Add(key(i)))
                    {
                        survivors++;
                    }
                }
            }

            return Math.Max(0, bodyCount - survivors);
        }

        /// <summary>
        /// 预算「删除空行」（delete_blank_rows_range）将删除的行数，与 ExcelJyyjtoolService.DeleteBlankRowsRange 语义严格一致：
        /// onlyFullyBlank=true 仅删整行全空（任一列非空即保留）；false 删首列为空的行。matrix 为选区原始值（含表头）。
        /// </summary>
        public static int CountRemovedBlankRows(object[,] matrix, bool onlyFullyBlank)
        {
            if (matrix == null)
            {
                return 0;
            }

            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            var removed = 0;
            for (var r = 0; r < rows; r++)
            {
                bool blank;
                if (onlyFullyBlank)
                {
                    blank = true;
                    for (var c = 0; c < cols; c++)
                    {
                        var v = matrix[r, c];
                        if (v != null && !(v is DBNull) && !string.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.InvariantCulture)))
                        {
                            blank = false;
                            break;
                        }
                    }
                }
                else
                {
                    var v0 = matrix[r, 0];
                    blank = v0 == null || v0 is DBNull || string.IsNullOrWhiteSpace(Convert.ToString(v0, CultureInfo.InvariantCulture));
                }

                if (blank)
                {
                    removed++;
                }
            }

            return removed;
        }

        /// <summary>
        /// 文本拆分（text_split）写入的目标列区间 [outCol, outCol+maxParts-1]，用于右键额部的“预留空列”守卫。
        /// 返回第 1 基目标列集合。源矩阵含表头。
        /// </summary>
        public static IReadOnlyList<int> SplitTargetColumns(object[,] matrix, int sourceColumn, string delimiter)
        {
            var cols = matrix.GetLength(1);
            if (sourceColumn < 1 || sourceColumn > cols || string.IsNullOrEmpty(delimiter))
            {
                return Array.Empty<int>();
            }

            var maxParts = 1;
            var rows = matrix.GetLength(0);
            for (var r = 0; r < rows; r++)
            {
                var v = matrix[r, sourceColumn - 1];
                var text = v == null || v is DBNull ? string.Empty : Convert.ToString(v, CultureInfo.InvariantCulture);
                var parts = text.Split(new[] { delimiter }, StringSplitOptions.None);
                if (parts.Length > maxParts)
                {
                    maxParts = parts.Length;
                }
            }

            var outCol = sourceColumn + 1;
            var result = new List<int>();
            for (var c = outCol; c < outCol + maxParts; c++)
            {
                result.Add(c);
            }
            return result;
        }

        /// <summary>给定工具执行结果，生成人可读摘要（首行为处理规模，次行为结果地址/计数）。</summary>
        public static string BuildSummary(string toolName, ToolResult result, ParamSelection sel)
        {
            if (result == null)
            {
                return "无返回结果。";
            }

            if (!result.Ok)
            {
                switch (result.ErrorCode)
                {
                    case "CONSERVATION_VIOLATED":
                        return "排序行守恒校验失败，已拒绝写入，数据未被修改。";
                    case "INVALID_COLUMN":
                    case "INVALID_OPERATOR":
                    case "INVALID_KEEP":
                    case "INVALID_DIRECTION":
                    case "INVALID_VALUE":
                    case "INVALID_WIDTH":
                    case "INVALID_FORMULA":
                    case "INVALID_COLUMNS":
                    case "INVALID_RANGE":
                    case "TOO_MANY_CELLS":
                    case "INVALID_ADDRESS":
                    case "EXCEL_MEMORY_ERROR":
                    case "EXCEL_COM_ERROR":
                        return $"执行失败：{result.ErrorCode} — {result.Data}";
                    case "INVALID_PATTERN":
                    case "INVALID_CASE":
                        // 工具层这些错误码带完整中文文案（如正则过长），优先展示文案而非裸码。
                        return result.Data is string p && !string.IsNullOrWhiteSpace(p)
                            ? $"执行失败：{p}"
                            : $"执行失败：{result.ErrorCode}";
                    case "WRITER_UNAVAILABLE":
                        return "写入器未挂接，无法落盘结果。";
                    default:
                        return result.Data is string s2 && !string.IsNullOrWhiteSpace(s2) ? $"执行失败：{s2}" : $"执行失败：{result.ErrorCode}";
                }
            }

            var scope = $"处理 {sel.Rows} 行 × {sel.Cols} 列";
            switch (toolName)
            {
                case "clean_excel_data":
                    return $"{scope}，清洗完成。";
                case "dedup_range":
                {
                    var dict = DataObject(result);
                    var removed = IntValue(dict, "deduped_rows_removed");
                    return removed.HasValue ? $"去重完成，删除重复 {removed.Value} 行。" : $"{scope}，去重完成。";
                }
                case "sort_range":
                    return $"{scope}，排序完成。";
                case "transpose_range":
                {
                    var dict = DataObject(result);
                    var target = dict != null && dict.TryGetValue("result", out var r) ? Convert.ToString(r, CultureInfo.InvariantCulture) : null;
                    return string.IsNullOrEmpty(target) ? $"{scope}，已转置。" : $"{scope}，已转置（新维度 {target}）。";
                }
                case "filter_excel_data":
                {
                    var dict = DataObject(result);
                    var matched = IntValue(dict, "matched_rows");
                    var output = Value(dict, "output_address");
                    return matched.HasValue && !string.IsNullOrEmpty(output)
                        ? $"筛选完成，命中 {matched.Value} 行，输出到 {output}。"
                        : $"{scope}，筛选完成。";
                }
                case "create_excel_pivot_summary":
                {
                    var dict = DataObject(result);
                    var target = Value(dict, "target_address");
                    var rowsCount = IntValue(dict, "row_count");
                    return rowsCount.HasValue && !string.IsNullOrEmpty(target)
                        ? $"透视汇总完成，{rowsCount.Value} 组，输出到 {target}。"
                        : "透视汇总完成，输出到新区域。";
                }
                case "text_split":
                {
                    var dict = DataObject(result);
                    var target = IntValue(dict, "target_column");
                    var depth = IntValue(dict, "split_columns");
                    return target.HasValue ? $"拆分完成，从第 {target.Value} 列起共 {depth ?? 0} 列。" : "拆分完成。";
                }
                case "text_merge":
                {
                    var dict = DataObject(result);
                    var target = IntValue(dict, "target_column");
                    return target.HasValue ? $"合并完成，写入第 {target.Value} 列。" : "合并完成。";
                }
                case "text_clean":
                {
                    var dict = DataObject(result);
                    var cleaned = IntValue(dict, "cleaned_cells");
                    return cleaned.HasValue ? $"文本清洗完成，处理 {cleaned.Value} 格。" : "文本清洗完成。";
                }
                case "fill_formula":
                {
                    var dict = DataObject(result);
                    var filled = IntValue(dict, "filled_cells");
                    return filled.HasValue ? $"公式填充完成，共 {filled.Value} 格。" : "公式填充完成。";
                }

                // ================= 常用组（ExcelJyyjtool） =================

                case "insert_text_into_cell":
                    return "已插入文本，处理 1 格。";

                case "delete_text_matches":
                {
                    var dict = DataObject(result);
                    var changed = IntValue(dict, "changed_cells");
                    return changed.HasValue ? $"已删除匹配文本，共改动 {changed.Value} 格。" : "已删除匹配文本。";
                }

                case "fill_random_numbers":
                {
                    var dict = DataObject(result);
                    var rows = IntValue(dict, "rows");
                    var columns = IntValue(dict, "columns");
                    var min = Value(dict, "min");
                    var max = Value(dict, "max");
                    return rows.HasValue && !string.IsNullOrEmpty(min) && !string.IsNullOrEmpty(max)
                        ? $"已填充 {rows.Value}×{columns ?? 0} 个随机数（{min}–{max}）。"
                        : $"{scope}，已填充随机数。";
                }

                case "fill_number_sequence":
                {
                    var dict = DataObject(result);
                    var count = IntValue(dict, "count");
                    var start = Value(dict, "start");
                    var step = Value(dict, "step");
                    var direction = Value(dict, "direction");
                    return count.HasValue && !string.IsNullOrEmpty(start)
                        ? $"已填充等差序列：起 {start}，步长 {step}，共 {count.Value} 个，{direction ?? "down"}向。"
                        : $"{scope}，已填充等差序列。";
                }

                case "insert_blank_rows_every_n":
                {
                    var dict = DataObject(result);
                    var inserted = IntValue(dict, "inserted_rows");
                    var interval = Value(dict, "interval");
                    return inserted.HasValue
                        ? $"已每隔 {interval ?? "?"} 行插入 {inserted.Value} 个空行。"
                        : $"{scope}，已插入空行。";
                }

                case "delete_blank_rows_range":
                {
                    var dict = DataObject(result);
                    var deleted = IntValue(dict, "deleted_rows");
                    return deleted.HasValue ? $"已删除 {deleted.Value} 个空行。" : $"{scope}，已删除空行。";
                }

                case "fill_blank_cells_down":
                {
                    var dict = DataObject(result);
                    var filled = IntValue(dict, "filled_cells");
                    return filled.HasValue ? $"已向下填充 {filled.Value} 个空单元格。" : $"{scope}，已向下填充。";
                }

                case "merge_duplicate_cells":
                {
                    var dict = DataObject(result);
                    var merged = dict != null && dict.TryGetValue("merged", out var mv) && Convert.ToBoolean(mv, CultureInfo.InvariantCulture);
                    return merged ? "已合并相邻相同单元格。" : $"已合并相邻相同单元格（{scope}）。";
                }

                default:
                    return $"{scope}，操作完成。";
            }
        }

        private static ParamField Select(string name, string label, IReadOnlyList<ParamOption> options, string def, string hint = null)
        {
            var field = new ParamField { Name = name, Label = label, Type = "select", Default = def, Hint = hint };
            foreach (var o in options)
            {
                field.Options.Add(o);
            }
            return field;
        }

        private static ParamField MultiSelect(string name, string label, IReadOnlyList<ParamOption> options, IReadOnlyList<string> defaults, string hint = null)
        {
            var field = new ParamField { Name = name, Label = label, Type = "multiselect", Hint = hint };
            foreach (var o in options)
            {
                field.Options.Add(o);
            }
            foreach (var d in defaults)
            {
                field.Defaults.Add(d);
            }
            return field;
        }

        private static ParamField Text(string name, string label, string def, bool required = false, string hint = null)
        {
            return new ParamField { Name = name, Label = label, Type = "text", Default = def, Required = required, Hint = hint };
        }

        private static ParamField Checkbox(string name, string label, bool def, string hint = null)
        {
            return new ParamField { Name = name, Label = label, Type = "checkbox", Default = def ? "true" : "false", Hint = hint };
        }

        private static List<ParamOption> ColumnOptions(ParamSelection sel)
        {
            var result = new List<ParamOption>();
            for (var i = 0; i < sel.Cols; i++)
            {
                var name = sel.ColumnNames.Count > i ? sel.ColumnNames[i] : $"列{i + 1}";
                result.Add(new ParamOption
                {
                    Value = (i + 1).ToString(CultureInfo.InvariantCulture),
                    Display = string.IsNullOrWhiteSpace(name) ? $"列{i + 1}" : $"{i + 1}·{name}"
                });
            }
            return result;
        }

        private static List<ParamOption> OperatorOptions()
        {
            return new List<ParamOption>
            {
                new ParamOption { Value = "equals", Display = "=" },
                new ParamOption { Value = "not_equals", Display = "≠" },
                new ParamOption { Value = "contains", Display = "包含" },
                new ParamOption { Value = "gt", Display = ">" },
                new ParamOption { Value = "gte", Display = "≥" },
                new ParamOption { Value = "lt", Display = "<" },
                new ParamOption { Value = "lte", Display = "≤" },
                new ParamOption { Value = "empty", Display = "空" },
                new ParamOption { Value = "not_empty", Display = "非空" }
            };
        }

        private static List<ParamOption> WithNewColumn(ParamSelection sel)
        {
            var result = ColumnOptions(sel);
            result.Add(new ParamOption { Value = (sel.Cols + 1).ToString(CultureInfo.InvariantCulture), Display = "选区右侧新列" });
            return result;
        }

        /// <summary>列下拉 + 顶部「无」选项（值为 0），用于可选次排序列。</summary>
        private static List<ParamOption> WithNoneOption(ParamSelection sel)
        {
            var result = new List<ParamOption> { new ParamOption { Value = "0", Display = "无" } };
            result.AddRange(ColumnOptions(sel));
            return result;
        }

        /// <summary>列下拉 + 顶部「全部列」选项（值为 0），用于按列执行的操作缺省作用于全部列。</summary>
        private static List<ParamOption> AllColumnsWith(ParamSelection sel)
        {
            var result = new List<ParamOption> { new ParamOption { Value = "0", Display = "全部列" } };
            result.AddRange(ColumnOptions(sel));
            return result;
        }

        private static Dictionary<string, object> DataObject(ToolResult result)
        {
            if (result == null || result.Data == null)
            {
                return null;
            }
            if (result.Data is Dictionary<string, object> map)
            {
                return map;
            }
            if (result.Data is Newtonsoft.Json.Linq.JObject j)
            {
                return j.ToObject<Dictionary<string, object>>();
            }
            try
            {
                // 工具层常用匿名对象（new { ... }）作 Data：转成字典以便提取统计字段。
                return Newtonsoft.Json.Linq.JObject.FromObject(result.Data).ToObject<Dictionary<string, object>>();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? IntValue(Dictionary<string, object> map, string key)
        {
            if (map == null || !map.TryGetValue(key, out var v) || v == null)
            {
                return null;
            }
            return long.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out var l) ? (int?)l : null;
        }

        private static string Value(Dictionary<string, object> map, string key)
        {
            return map != null && map.TryGetValue(key, out var v) ? Convert.ToString(v, CultureInfo.InvariantCulture) : null;
        }
    }
}