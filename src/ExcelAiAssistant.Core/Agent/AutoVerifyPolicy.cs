using System;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>一次自动自检的执行计划：目标 verify 工具与其参数。</summary>
    public sealed class AutoVerifyPlan
    {
        public string VerifyTool { get; set; }

        public Dictionary<string, object> Args { get; set; }

        /// <summary>触发原因（用于观察消息与日志，说明这条自检由哪个主操作自动带出）。</summary>
        public string Reason { get; set; }
    }

    /// <summary>
    /// 执行层自动自检策略：主写操作成功后，服务端自动调用配套 verify_* 工具做回读核对，
    /// 不依赖模型自觉调用。预期参数从主工具实参直接推导（写入值即预期值），保证「预期 vs 实际」闭环。
    /// 只读、纯映射推导，不接触 COM；verify 工具本身由注册方提供，缺失或非只读时静默跳过。
    /// </summary>
    public static class AutoVerifyPolicy
    {
        public static bool IsVerifyTool(string tool)
        {
            return !string.IsNullOrWhiteSpace(tool) && tool.StartsWith("verify_", StringComparison.Ordinal);
        }

        /// <summary>
        /// 自检通过判定（中央单点）：verify 工具的 Ok 只表示"执行成功"，verdict 才是判定结论。
        /// verdict 缺失时退回 Ok（兼容非标准结果）；verdict 非 ok（no_expected/mismatch/…）一律视为未通过。
        /// verify_file 自身已用 Ok 表达判定（Ok=verdict=="ok"），同样适用。
        /// 供 AgentRunner 自动自检、run_self_check / apply_correction 工具共用，避免多份判定漂移。
        /// </summary>
        public static bool VerifyPassed(ToolResult result)
        {
            if (result == null || !result.Ok)
            {
                return false;
            }

            var verdict = ExtractVerdict(result);
            return string.IsNullOrEmpty(verdict)
                || string.Equals(verdict, "ok", StringComparison.Ordinal);
        }

        /// <summary>从 verify 工具结果 Data 中提取 verdict；取不到返回 null。
        /// Data 形态兼容：匿名对象（verify 家族主流形态，经 JToken.FromObject）、JObject、JSON 字符串。</summary>
        public static string ExtractVerdict(ToolResult result)
        {
            try
            {
                var data = result?.Data;
                if (data == null)
                {
                    return null;
                }

                // JToken.FromObject 覆盖匿名对象/字典/JObject 全形态（原 AgentRunner 同款口径）；
                // 仅字符串走解析分支（JSON 文本）。
                var token = data as JToken ?? (data is string text
                    ? (text.TrimStart().Length > 0 && text.TrimStart()[0] == '{' ? TryParseObject(text) : null)
                    : JToken.FromObject(data));
                return (token as JObject)?.Value<string>("verdict");
            }
            catch (Newtonsoft.Json.JsonException)
            {
                AppLog.Warn("AutoVerifyPolicy", "ExtractVerdict：Data 不是合法 JSON 对象，按无 verdict 处理");
                return null;
            }
            catch (ArgumentException)
            {
                // Data 形态非 JSON（纯文本/标量）：无 verdict 可言，交由 Ok 兼容语义处理。
                return null;
            }
        }

        private static JObject TryParseObject(string text)
        {
            try
            {
                return JObject.Parse(text);
            }
            catch (Newtonsoft.Json.JsonException)
            {
                return null;
            }
        }

        /// <summary>批量回写逐格自动自检的上限：超过则整体跳过，避免一次操作生成几十次回读拖慢回合。</summary>
        public const int MaxAutoVerifyCellsPerCall = 5;

        /// <summary>
        /// 自动自检映射覆盖的主工具全集（与 TryPlan/TryPlanAll 的 switch 分支一一对应）。
        /// 发布门禁据此审计：所有写能力工具（NormalWrite/StructuralChange/Dangerous）必须在这里、
        /// 或在 DocumentedWriteToolExemptions 豁免清单里，防止新工具带着「永不自检」上线。
        /// </summary>
        public static readonly IReadOnlyList<string> MappedPrimaryTools = new[]
        {
            "write_excel_cell",
            "write_excel_range",
            "write_excel_cells_batch",
            "append_excel_rows",
            "clean_excel_data",
            "dedup_range",
            "sort_range",
            "format_excel_range",
            "format_excel_conditional",
            "create_excel_pivot_summary",
            "flag_cells",
            "create_excel_chart",
            "create_combo_chart",
            "create_doughnut",
            "create_dashboard_kpi_brief",
            "render_dashboard",
            "render_super_dashboard",
            "build_pivot_sheet",
            "set_linkage",
            "render_dim_cross",
            "create_super_table",
            "repair_chart",
            "delete_dashboard",
            "create_chart_advanced",
            "render_pnl_brief",
            "render_target_track",
            "render_rank_dynamic",
            "edit_dashboard_theme",
            "delete_dashboard_chart",
            "create_truncated_column",
            "insert_recommended_chart",
            "add_excel_worksheet",
            "rename_excel_worksheet",
            "delete_excel_worksheet",
            "clear_excel_range",
            "delete_text_matches",
            "data_validation",
            "merge_or_split_cells",
            "merge_duplicate_cells",
            "create_excel_sparkline",
            "create_excel_table",
            "insert_image",
            "freeze_panes",
            "manage_hyperlinks",
            // M26：L2 元素层 / 外部数据源（有确定性期望，可自检）
            "set_element",
            "import_csv",
            // M28：导出文件核验（verify_file 成功后自动带出）
            "export_sheets_to_pdf",
            "export_excel_chart_image"
    };

        /// <summary>
        /// 工具是否豁免于自动自检：精确豁免清单命中，或 <c>ct_</c> 前缀的用户自定义沉淀工具
        /// （产出由用户意图核对，无法预置确定性期望）。覆盖审计两处（真机 harness / 全功能巡检）共用本判定。
        /// </summary>
        public static bool IsExemptedFromAutoVerify(string tool)
        {
            if (string.IsNullOrWhiteSpace(tool))
            {
                return false;
            }

            if (DocumentedWriteToolExemptions.ContainsKey(tool))
            {
                return true;
            }

            // 用户自定义沉淀工具（CustomToolStore 注册名以 ct_ 开头）：行为由沉淀时的脚本/模板决定。
            return tool.StartsWith("ct_", StringComparison.Ordinal);
        }

        /// <summary>
        /// 有写能力但刻意不做自动自检的豁免清单（工具名 → 理由）。豁免必须有明确理由并随评审变更，
        /// 未列出的写能力工具会被覆盖审计拦下。
        /// </summary>
        public static readonly IReadOnlyDictionary<string, string> DocumentedWriteToolExemptions =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["open_excel_external_book"] = "外部工作簿只读句柄：只读打开/关闭，不修改活动工作簿单元格；打开失败由工具自身错误码表达，关闭由 close_external_book 与登记表兜底。",
                ["filter_excel_data"] = "verify_sort_filter 断言区域内每行都满足谓词，而筛选源区域保留不匹配行；输出区域范围只有模型知道，由模型按输出区显式核对。",
                ["undo_excel_operation"] = "回滚操作本身由回滚后状态与撤销历史保证，且撤销目标即上一次快照。",
                ["run_vba_snippet"] = "脚本执行结果形态不可预知，由探针回读与模型按脚本意图核对。",
                ["run_python_code"] = "同 run_vba_snippet。",
                ["run_python_excel"] = "同 run_vba_snippet。",
                ["run_cli_command"] = "进程级命令，无工作簿状态可回读核对。",
                ["install_python_packages"] = "Python 包安装：结果以 pip 退出码与 import 自检表达，无工作簿状态可回读核对。",
                ["list_python_packages"] = "同 run_cli_command（只读环境枚举，Defense-in-depth 双保险豁免）。",
                ["python_setup_guide"] = "只读教程返回，无任何写副作用（双保险豁免）。",
                ["run_workflow"] = "工作流内部步骤已有独立的步骤级自检（step.Verify），外层不再重复包装。",
                ["apply_correction"] = "组合修正工具：每条 attempt 由自带的只读 verify_tool 判定并回传 verdict，外层不重复自检；且明确拒绝 Dangerous 子工具。",
                ["connect_ai_client"] = "写的是第三方客户端 MCP 配置文件（默认仅预览，落盘前备份 *.jyyj-bak，只增改 jyyj-excel 一项），不产生任何工作簿状态，无 verify_* 可回读；连通性由 test_ai_client_connection 实测握手验证。",
                ["save_vba_tool"] = "只写本地工具库文件，不触碰工作簿状态。",
                // M28/R3：文件落盘动作，无单元格期望可回读核对；由 file_path 存在性与后续读取表达。
                ["save_workbook"] = "文件落盘动作，无单元格期望可回读；由 file_path 存在性与后续读取表达。",
                ["merge_workbooks"] = "跨工作簿合并，产出由模型按合并意图核对。",
                ["split_workbook"] = "同 merge_workbooks。",
                ["import_delimited_file"] = "导入内容由源文件决定，模型按导入意图核对。",
                ["add_insight_card"] = "洞察卡为展示形状，写入后已内嵌 TextFrame2 文本读回自检（verdict=ok 才交付）；形状不在单元格快照回滚体系，replace_existing=true 可幂等重写。",
                ["import_excel_file_sheet"] = "同 import_delimited_file。",
                ["merge_delimited_files"] = "同 import_delimited_file。",
                ["rewrite_formulas"] = "批量公式文本改写（相对/绝对/IFERROR 包装），改写后公式可经 eval_formula/diagnose_formula 按意图核对；无预置确定性期望。",
                ["export_selection_to_markdown"] = "只读导出为文本。",
                ["import_markdown_to_excel"] = "导入内容由源文本决定，模型按导入意图核对。",
                ["copy_visible_rows_to_sheet"] = "复制源行到新表，内容守恒由源区域与目标区域只读比对表达，交模型核对。",
                ["generate_analysis_report"] = "生成报告输出（工作表/Markdown/HTML），内容由模型按报告意图核对。",
                ["transform_excel_formulas"] = "公式批量变换，由模型按变换规则核对。",
                ["convert_formula_references"] = "同 transform_excel_formulas。",
                ["apply_page_setup"] = "页面设置不改变数据。",
                ["replace_external_links"] = "外部链接维护，不改变单元格值语义。",
                ["reset_cell_value_type"] = "重置值类型，由模型按重置意图核对。",
                ["find_replace_range"] = "替换后的期望集合无法从实参自动推导，由模型按替换意图核对。",
                ["transpose_range"] = "确定性变换类：期望可由模型从源数据推导，逐步补自动映射。",
                ["text_split"] = "确定性变换类，同上。",
                ["text_merge"] = "确定性变换类，同上。",
                ["text_clean"] = "确定性变换类，同上。",
                ["date_normalize"] = "确定性变换类，同上。",
                ["number_pad"] = "确定性变换类，同上。",
                ["convert_2d_to_1d"] = "确定性变换类，同上。",
                ["convert_1d_to_2d"] = "确定性变换类，同上。",
                ["bankers_round_values"] = "确定性变换类，同上。",
                ["amount_to_uppercase"] = "确定性变换类，同上。",
                ["fill_blank_cells_down"] = "确定性变换类，同上。",
                ["apply_fill_blank_cells"] = "确定性变换类，同上。",
                ["cross_reference_columns"] = "确定性变换类，同上。",
                ["wrap_formulas_with_iferror"] = "公式类变换：写入后由模型按提示词补 verify_formula。",
                ["convert_text_to_pinyin"] = "文本变换类，由模型按变换意图核对。",
                ["convert_chinese_variant"] = "文本变换类，同上。",
                ["extract_text_characters"] = "文本变换类，同上。",
                ["insert_text_into_cell"] = "文本插入类，同上。",
                ["insert_excel_dimension"] = "行列结构变更：变更后维度由模型按插入意图核对（可调 verify_structure 断言行列数）。",
                ["delete_excel_dimension"] = "行列结构变更：删除后期望维度依赖删除前状态（内容整体位移），无确定性期望可自动推导，由模型按删除意图核对（可调 verify_structure 断言行列数）。",
                ["insert_blank_rows_every_n"] = "行列结构变更，同上。",
                ["delete_blank_rows_range"] = "行列结构变更，同上。",
                ["split_sheet_by_field"] = "拆表产出由模型按拆分意图核对。",
                ["create_sheet_by_value"] = "同 split_sheet_by_field。",
                ["create_toc_sheet"] = "目录表生成，由模型核对目录项与实际表名。",
                ["merge_sheets"] = "合并产出由模型按合并意图核对。",
                ["make_payslip_rows"] = "模板化生成，由模型按模板意图核对。",
                ["copy_paste_excel_range"] = "复制粘贴内容守恒可由源/目标只读比对表达，交模型核对。",
                ["smart_paste_delimiter"] = "同 copy_paste_excel_range。",
                ["copy_visible_range_only"] = "同 copy_paste_excel_range。",
                ["copy_range_as_image"] = "产出为图片文件，不改变单元格值。",
                ["protect_sheet_watermark"] = "水印/保护为呈现层，不改变数据语义。",
                ["manage_excel_filter"] = "AutoFilter 对象状态，verify 家族无对应维度。",
                ["manage_excel_named_range"] = "命名区域管理，verify 家族无对应维度。",
                ["fill_random_numbers"] = "随机生成无确定性期望，由模型按生成意图核对。",
                ["fill_number_sequence"] = "序列生成期望可推导，逐步补自动映射。",
                ["generate_permutation_list"] = "排列生成，由模型按生成意图核对。",
                ["rate_change_analysis"] = "分析产出，由模型按分析意图核对。",
                ["create_pivot_table"] = "原生透视表对象（非汇总块），verify_pivot 面向汇总块语义；由模型核对。",
                ["fill_formula"] = "公式区间：回读计算值与公式文本错位，由模型按提示词逐格补 verify_formula。",
                ["append_calculation_scratchpad"] = "草稿板区域，不影响业务数据语义。",
                ["clear_calculation_scratchpad"] = "草稿板区域，同上。",
                ["cell_comment"] = "批注操作，不改变单元格值。",
                ["run_exe_command"] = "外部进程命令，无工作簿状态可回读核对。",
                ["preview_external_file"] = "外部文件预览，不修改工作簿。",
                ["validate_vba_snippet"] = "静态校验工具，不修改工作簿状态。",
                ["ct_format_answer_sheet"] = "用户自定义模板工具（Dangerous，默认关闭），产出由用户意图核对。",
                ["ct_proc_FormatAlignment"] = "同 ct_format_answer_sheet。",
                ["ct_proc_FormatAnswerSheet"] = "同 ct_format_answer_sheet。",
                ["ct_vba_tool"] = "用户沉淀的 VBA 工具，产出由脚本意图核对。",
                ["ct_vba_tool_2"] = "同 ct_vba_tool。",
                ["ct_vba_tool_3"] = "同 ct_vba_tool。",
            // M26 审查修复：L2 元素层 / 草稿模式 / 智能体桥接（无可测确定性期望，逐条带理由）
            ["add_element"] = "新增文本框属装饰性元素，无业务数据期望可核；版面判断优于机械断言形状存在。",
            ["remove_element"] = "删除形状/图表为结构操作，删除后无期望可回读；事后用 query_elements 确认即可。",
            ["move_element"] = "仅改 Top/Left 版面坐标，无数据期望可测。",
            ["start_draft"] = "草稿表是临时隔离副本，断言其存在无业务意义；真正落盘期望由 apply_draft 承担。",
            ["apply_draft"] = "回写内容已由 diff_draft 事先展示并经人工采纳，工具层无独立期望值；AffectedRanges 精确回显支撑人工核对。",
            ["discard_draft"] = "清理动作（删除草稿表），无业务数据期望可核。",
            ["buddy_append_memory"] = "文件层追加（workbuddy 记忆），非 Excel 数据面写入；可用 buddy_read_memory 自验。",
            };

        /// <summary>
        /// 出全部自动自检计划。多数工具至多一个计划（等价 TryPlan）；
        /// write_excel_cells_batch 逐格出计划（≤ MaxAutoVerifyCellsPerCall 格），公式格与值格分别映射。
        /// 返回 false 表示本参数组合不适合自动自检。
        /// </summary>
        public static bool TryPlanAll(string tool, IReadOnlyDictionary<string, object> args, List<AutoVerifyPlan> plans)
        {
            if (string.IsNullOrWhiteSpace(tool) || plans == null || IsVerifyTool(tool))
            {
                return false;
            }

            if (tool == "write_excel_cells_batch")
            {
                return TryPlanBatch(args, plans);
            }

            if (TryPlan(tool, args, out var plan))
            {
                plans.Add(plan);
                return true;
            }

            return false;
        }

        /// <summary>批量回写：逐格映射（值格→verify_write_result，公式格→verify_formula），超上限整体跳过。</summary>
        private static bool TryPlanBatch(IReadOnlyDictionary<string, object> map, List<AutoVerifyPlan> plans)
        {
            if (!map.TryGetValue("cells", out var raw) || raw == null)
            {
                return false;
            }

            var cells = new List<Dictionary<string, object>>();
            if (raw is System.Collections.IEnumerable sequence && !(raw is string))
            {
                foreach (var item in sequence)
                {
                    var dict = item as Dictionary<string, object> ?? AsDict(item);
                    if (dict == null)
                    {
                        return false;
                    }

                    cells.Add(dict);
                }
            }
            else
            {
                return false;
            }

            if (cells.Count == 0 || cells.Count > MaxAutoVerifyCellsPerCall)
            {
                return false;
            }

            foreach (var cell in cells)
            {
                var address = GetString(cell, "address");
                if (string.IsNullOrWhiteSpace(address) || !cell.TryGetValue("value", out var value) || value == null)
                {
                    return false;
                }

                var valueMode = GetString(cell, "value_mode") ?? "value";
                if (valueMode == "formula" || StartsWithFormula(value))
                {
                    plans.Add(new AutoVerifyPlan
                    {
                        VerifyTool = "verify_formula",
                        Reason = "批量公式写入后自动核对公式内容与计算值",
                        Args = Build(map,
                            ("address", address),
                            ("expected_formula", Convert.ToString(value)))
                    });
                }
                else
                {
                    plans.Add(new AutoVerifyPlan
                    {
                        VerifyTool = "verify_write_result",
                        Reason = "批量写入后自动回读核对实际值",
                        Args = Build(map,
                            ("address", address),
                            ("expected_values", new List<object> { new List<object> { value } }),
                            ("type_aware", true))
                    });
                }
            }

            return plans.Count > 0;
        }

        /// <summary>
        /// 为成功执行的主工具推导自动自检计划。返回 false 表示无映射或本参数组合不适合自动自检
        /// （如公式模式区间写入——回读计算值与公式文本必然不同，交给 verify_formula 由模型按格断言）。
        /// </summary>
        public static bool TryPlan(string tool, IReadOnlyDictionary<string, object> args, out AutoVerifyPlan plan)
        {
            plan = null;
            if (string.IsNullOrWhiteSpace(tool) || IsVerifyTool(tool))
            {
                return false;
            }

            var map = args as Dictionary<string, object> ?? (args == null
                ? new Dictionary<string, object>()
                : args.ToDictionary(kv => kv.Key, kv => kv.Value));

            switch (tool)
            {
                case "write_excel_cell":
                    return TryPlanCellWrite(map, out plan);
                case "write_excel_range":
                    return TryPlanRangeWrite(map, out plan);
                case "clean_excel_data":
                    return TryPlanCleanChecks(map, CleanChecksFromOperations(map), out plan);
                case "dedup_range":
                    return TryPlanCleanChecks(map, new List<object> { new Dictionary<string, object> { ["type"] = "no_duplicates" } }, out plan);
                case "sort_range":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_sort_filter",
                        Reason = "排序后自动核对排序列与方向",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("expected_sort_column", GetInt(map, "sort_column") ?? 1),
                            ("expected_sort_direction", GetString(map, "direction") ?? "asc"),
                            ("header_rows", GetInt(map, "header_rows")))
                    };
                    return true;

                // filter_excel_data 刻意不做自动自检：ComputeFilterVerdict 断言区域内每个数据行都满足谓词，
                // 而筛选的源区域本就保留不匹配行，输出区域范围只有模型知道——由模型按输出区域显式核对。
                case "format_excel_range":
                    return TryPlanCellFormat(map, out plan);
                case "format_excel_conditional":
                    return TryPlanConditional(map, out plan);
                case "create_excel_pivot_summary":
                    // 透视自动自检必须配合主工具结果富化（EnrichFromResult）：
                    // 输出地址可能未传（自动落点）、数据行数只有结果知道——无断言时 verify_pivot 会判 no_expected 误报。
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_pivot",
                        Reason = "透视汇总后自动核对汇总块结构与明细/汇总守恒",
                        Args = Build(map,
                            ("target_address", GetString(map, "target_address")),
                            ("source_address", GetString(map, "source_address")),
                            ("source_value_column", GetInt(map, "value_column")),
                            ("source_label_column", GetInt(map, "label_column")))
                    };
                    return true;
                case "append_excel_rows":
                    // 目标地址要等执行后由锚点列末行推得（结果 AffectedRanges），此处只带期望值矩阵；
                    // EnrichFromResult 富化出 address/sheet 后 verify_write_result 才可执行。
                    var appendValues = map.ContainsKey("values") ? map["values"] : null;
                    if (appendValues == null || ContainsFormulaText(appendValues))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_write_result",
                        Reason = "追加行后自动回读核对新行内容",
                        Args = Build(map,
                            ("expected_values", appendValues),
                            ("type_aware", true))
                    };
                    return true;
                // M28：导出类工具——文件落盘后由 verify_file 核对存在性/大小/魔数（xlsx 实为 zip 容器）
                case "export_sheets_to_pdf":
                case "export_excel_chart_image":
                    var exportFile = GetString(map, "file_path");
                    if (string.IsNullOrWhiteSpace(exportFile))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_file",
                        Reason = "导出文件后自动核验存在性/大小/魔数",
                        Args = Build(map, ("file_path", exportFile))
                    };
                    return true;
                case "add_excel_worksheet":
                    if (string.IsNullOrWhiteSpace(GetString(map, "name")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "新增工作表后自动核对表存在性",
                        Args = Build(map, ("expected_sheet_exists", GetString(map, "name")))
                    };
                    return true;
                case "rename_excel_worksheet":
                    if (string.IsNullOrWhiteSpace(GetString(map, "name")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "重命名工作表后自动核对新名存在、旧名消失",
                        Args = Build(map,
                            ("expected_sheet_exists", GetString(map, "name")),
                            ("expected_sheet_absent", GetString(map, "sheet")))
                    };
                    return true;
                case "delete_excel_worksheet":
                    var deletedSheet = GetString(map, "sheet");
                    if (string.IsNullOrWhiteSpace(deletedSheet))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "删除工作表后自动核对表已不存在",
                        Args = Build(map, ("expected_sheet_absent", deletedSheet))
                    };
                    return true;
                case "flag_cells":
                    // 无 fill_colors 时没有任何判定依据（Build 会丢弃 null）→ verify_marker 只会返回 no_expected。
                    // 宁缺勿误报：不生成自动计划，标记结果由模型按需显式调 verify_marker 核对。
                    if (!map.ContainsKey("fill_colors") || map["fill_colors"] == null)
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_marker",
                        Reason = "标记后自动核对标记分布与颜色",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("expected_fill_colors", map.ContainsKey("fill_colors") ? map["fill_colors"] : null))
                    };
                    return true;
                case "create_excel_chart":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "图表生成后自动核对类型/标题/数据源有效性",
                        Args = Build(map,
                            ("expected_chart_type", GetString(map, "chart_type")),
                            ("expected_title", GetString(map, "title")),
                            ("expected_data_source", GetString(map, "source_address")))
                    };
                    return true;
                case "create_combo_chart":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "双轴组合图生成后自动核对类型/系列/次轴构成",
                        Args = Build(map,
                            ("expected_title", GetString(map, "title")),
                            ("expected_data_source", GetString(map, "source_address")),
                            ("expected_secondary_axis", 1))
                    };
                    return true;
                case "create_doughnut":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "环形图生成后自动核对类型/标题/数据源有效性",
                        Args = Build(map,
                            ("expected_chart_type", GetString(map, "chart_type") ?? "doughnut"),
                            ("expected_title", GetString(map, "title")),
                            ("expected_data_source", GetString(map, "source_address")))
                    };
                    return true;
                case "create_dashboard_kpi_brief":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "看板渲染后自动核对 sheet/辅助区存在与 KPI 重算守恒",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet")))
                    };
                    return true;
                case "render_dashboard":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "sales_overview 渲染后自动核对看板结构与 KPI 守恒",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "销售看板"))
                    };
                    return true;
                case "build_pivot_sheet":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "透视分析 sheet 生成后自动核对工作簿结构完好（透视表内容为运行时状态不做静态断言）",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "超级看板"))
                    };
                    return true;
                case "render_super_dashboard":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "超级看板渲染后自动核对看板结构与 KPI 守恒（三视图 staging 全指标区）",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "超级看板"))
                    };
                    return true;
                case "create_truncated_column":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "截断柱生成后自动核对类型/标题/数据源",
                        Args = Build(map,
                            ("expected_title", GetString(map, "title")),
                            ("expected_data_source", GetString(map, "source_address")))
                    };
                    return true;
                case "set_linkage":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "联动设置后自动核对看板结构完好（fail-closed 兜底，切片器项为运行时状态不做静态断言）",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "销售看板"))
                    };
                    return true;
                case "repair_chart":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "图表修复后自动核对类型/系列/点数与渲染",
                        Args = Build(map,
                            ("sheet", GetString(map, "sheet")),
                            ("chart_name", GetString(map, "chart_name")),
                            ("expected_data_source", GetString(map, "source_address")))
                    };
                    return true;
                case "create_super_table":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "超级数据表生成后自动核对 sheet 存在",
                        Args = Build(map,
                            ("expected_sheet_exists", GetString(map, "super_sheet") ?? "超级数据表"))
                    };
                    return true;
                case "render_dim_cross":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "dim_cross 渲染后自动核对看板与辅助 sheet 存在性（矩阵数值语义由数据门槛+SUMIFS 口径保证）",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "多维交叉看板"))
                    };
                    return true;
                case "delete_dashboard":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "删除看板后自动核对看板 sheet 与辅助表均已消失（AC-10 反向断言，同 delete_excel_worksheet 口径）",
                        Args = Build(map,
                            ("expected_sheet_absent", GetString(map, "dashboard_sheet") ?? "销售看板"))
                    };
                    return true;
                case "render_pnl_brief":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "损益看板渲染后自动核对看板与辅助 sheet 存在性",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "财务损益看板"))
                    };
                    return true;
                case "render_target_track":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "目标达成看板渲染后自动核对看板与辅助 sheet 存在性",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "目标达成看板"))
                    };
                    return true;
                case "render_rank_dynamic":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "排名动态看板渲染后自动核对看板与辅助 sheet 存在性",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "排名动态看板"))
                    };
                    return true;
                case "edit_dashboard_theme":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "换主题后自动核对看板与辅助 sheet 完整性",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "销售看板"))
                    };
                    return true;
                case "delete_dashboard_chart":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_dashboard",
                        Reason = "删单图后自动核对看板与辅助 sheet 完整性",
                        Args = Build(map,
                            ("dashboard_sheet", GetString(map, "dashboard_sheet") ?? "销售看板"))
                    };
                    return true;
                case "create_chart_advanced":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "技巧型图表插入后自动核对图表存在且非空",
                        Args = Build(map)
                    };
                    return true;
                case "insert_recommended_chart":
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_chart",
                        Reason = "推荐图表插入后自动核对图表是否真正生成且非空",
                        Args = Build(map)
                    };
                    return true;
                case "clear_excel_range":
                    // 清空的确定性期望=全区域为空：verify_data_clean all_blank（含表头行全矩阵扫描）。
                    // 区域过大时放弃自动自检（读回矩阵代价高，模型可显式核对），避免拖慢回合。
                    if (string.IsNullOrWhiteSpace(GetString(map, "address")) || !IsBoundedSmallRange(GetString(map, "address"), 20000))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_data_clean",
                        Reason = "清空后自动核对区域已全部为空",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("checks", new List<object> { new Dictionary<string, object> { ["type"] = "all_blank" } }))
                    };
                    return true;
                case "delete_text_matches":
                    // 删除文本的确定性期望=区域（或输出区）不再残留匹配 pattern 的文本。
                    var deleteTarget = GetString(map, "target_address") ?? GetString(map, "address");
                    var deletePattern = GetString(map, "pattern");
                    if (string.IsNullOrWhiteSpace(deleteTarget) || string.IsNullOrWhiteSpace(deletePattern)
                        || !IsBoundedSmallRange(deleteTarget, 20000))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_data_clean",
                        Reason = "删除匹配文本后自动核对区域无残留匹配",
                        Args = Build(map,
                            ("address", deleteTarget),
                            ("checks", new List<object> { new Dictionary<string, object> { ["type"] = "no_text_matches", ["pattern"] = deletePattern } }))
                    };
                    return true;
                case "data_validation":
                    if (string.IsNullOrWhiteSpace(GetString(map, "address")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_data_validation",
                        Reason = "数据验证设置后自动核对规则类型/来源/边界",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("expected_type", GetString(map, "type") ?? "list"),
                            ("expected_source", GetString(map, "source")),
                            ("expected_operator", GetString(map, "operator")),
                            ("expected_value1", GetString(map, "value1")),
                            ("expected_value2", GetString(map, "value2")))
                    };
                    return true;
                case "merge_or_split_cells":
                    if (string.IsNullOrWhiteSpace(GetString(map, "address")))
                    {
                        return false;
                    }

                    var mergeAction = GetString(map, "action");
                    var expectMerged = mergeAction == "merge" || mergeAction == "merge_similar";
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_merge",
                        Reason = "合并/拆分后自动核对区域 MergeCells 状态",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("expected_merged", expectMerged ? "true" : "false"))
                    };
                    return true;
                case "merge_duplicate_cells":
                    if (string.IsNullOrWhiteSpace(GetString(map, "address")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_merge",
                        Reason = "同类项合并后自动核对区域出现合并单元格（混合状态按通过+提示处理）",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("expected_merged", "true"))
                    };
                    return true;
                case "create_excel_sparkline":
                    if (string.IsNullOrWhiteSpace(GetString(map, "dest_range")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_sparkline",
                        Reason = "迷你图生成后自动核对落点/类型/数据源",
                        Args = Build(map,
                            ("dest_range", GetString(map, "dest_range")),
                            ("expected_type", GetString(map, "type") ?? "line"),
                            ("expected_data_source", GetString(map, "data_range")))
                    };
                    return true;
                case "create_excel_table":
                    if (string.IsNullOrWhiteSpace(GetString(map, "address")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_table",
                        Reason = "智能表创建后自动核对表存在性与落点区域",
                        Args = Build(map,
                            ("address", GetString(map, "address")),
                            ("name", GetString(map, "name")))
                    };
                    return true;
                case "insert_image":
                    if (string.IsNullOrWhiteSpace(GetString(map, "anchor")))
                    {
                        return false;
                    }

                    // 审查 F2：width/height ≤ 0 视为「按原生尺寸」哨兵或缺省，不透传 expected_*，避免尺寸必然误报。
                    var imageWidth = GetDouble(map, "width");
                    var imageHeight = GetDouble(map, "height");
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_image",
                        Reason = "图片插入后自动核对锚点与尺寸",
                        Args = Build(map,
                            ("anchor", GetString(map, "anchor")),
                            ("expected_width", imageWidth.HasValue && imageWidth.Value > 0 ? imageWidth.Value : (object)null),
                            ("expected_height", imageHeight.HasValue && imageHeight.Value > 0 ? imageHeight.Value : (object)null))
                    };
                    return true;
                case "freeze_panes":
                    var freezeUnfreeze = GetBool(map, "unfreeze");
                    if (freezeUnfreeze == true)
                    {
                        plan = new AutoVerifyPlan
                        {
                            VerifyTool = "verify_freeze_panes",
                            Reason = "解冻后自动核对冻结状态已解除",
                            Args = Build(map,
                                ("expected_frozen", false))
                        };
                        return true;
                    }

                    if (string.IsNullOrWhiteSpace(GetString(map, "cell")))
                    {
                        return false;
                    }

                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_freeze_panes",
                        Reason = "冻结后自动核对冻结状态与位置",
                        Args = Build(map,
                            ("expected_frozen", true),
                            ("expected_cell", GetString(map, "cell")))
                    };
                    return true;
                // M26：L2 元素层 / 外部数据源（有确定性期望）
                case "set_element":
                    // FIX（M27 审查 A6）：参数名必须对齐 verify_write_result 真实契约（address/expected_values）——
                    // 原传 expected_address 会让自检 Required 抛错 → fail-closed 拦掉每一次写入。
                    // set_element 参数集与 write_excel_cell 完全一致，直接复用同一 plan 构造。
                    return TryPlanCellWrite(map, out plan);
                case "import_csv":
                    // FIX（M27 审查 A6）：verify_structure 真实参数是 expected_sheet_exists（原传 expected_sheet 自检必失败）
                    plan = new AutoVerifyPlan
                    {
                        VerifyTool = "verify_structure",
                        Reason = "导入后自动核对新工作表已创建",
                        Args = Build(map, ("expected_sheet_exists", ResolveCsvSheetName(map)))
                    };
                    return true;

                case "manage_hyperlinks":
                    var hyperlinkAction = GetString(map, "action") ?? "list";
                    var hyperlinkAddress = GetString(map, "address");
                    switch (hyperlinkAction)
                    {
                        case "add":
                        case "update":
                            if (string.IsNullOrWhiteSpace(hyperlinkAddress))
                            {
                                return false;
                            }

                            var hyperlinkUrl = GetString(map, "link_url");
                            var hyperlinkSubAddress = GetString(map, "link_subaddress");
                            // 审查 F3：仅改显示文本的 update 同样出计划（expected_text 断言），不留自检盲区。
                            var hyperlinkDisplayText = GetString(map, "display_text");
                            if (hyperlinkUrl == null && hyperlinkSubAddress == null && hyperlinkDisplayText == null)
                            {
                                return false;
                            }

                            plan = new AutoVerifyPlan
                            {
                                VerifyTool = "verify_hyperlink",
                                Reason = hyperlinkAction == "add" ? "添加超链接后自动核对目标与内容" : "更新超链接后自动核对目标与内容",
                                Args = Build(map,
                                    ("address", hyperlinkAddress),
                                    ("expected_url", hyperlinkUrl),
                                    ("expected_subaddress", hyperlinkSubAddress),
                                    ("expected_text", hyperlinkDisplayText))
                            };
                            return true;
                        case "remove":
                            if (string.IsNullOrWhiteSpace(hyperlinkAddress))
                            {
                                return false;
                            }

                            plan = new AutoVerifyPlan
                            {
                                VerifyTool = "verify_hyperlink",
                                Reason = "删除超链接后自动核对目标已无链接",
                                Args = Build(map,
                                    ("address", hyperlinkAddress),
                                    ("expected_present", false))
                            };
                            return true;
                        case "remove_all":
                            plan = new AutoVerifyPlan
                            {
                                VerifyTool = "verify_hyperlink",
                                Reason = "清空超链接后自动核对整表归零",
                                Args = Build(map,
                                    ("expected_count", 0))
                            };
                            return true;
                        default:
                            // list 为只读动作，无需自动自检。
                            return false;
                    }
                default:
                    return false;
            }
        }

        private static bool TryPlanCellWrite(Dictionary<string, object> map, out AutoVerifyPlan plan)
        {
            plan = null;
            var address = GetString(map, "address");
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            var value = map.ContainsKey("value") ? map["value"] : null;
            var valueMode = GetString(map, "value_mode") ?? "auto";
            var isFormula = valueMode == "formula" || (valueMode == "auto" && StartsWithFormula(value));

            if (isFormula)
            {
                plan = new AutoVerifyPlan
                {
                    VerifyTool = "verify_formula",
                    Reason = "公式写入后自动核对公式内容与计算值",
                    Args = Build(map,
                        ("address", address),
                        ("expected_formula", value == null ? null : Convert.ToString(value)))
                };
                return true;
            }

            plan = new AutoVerifyPlan
            {
                VerifyTool = "verify_write_result",
                Reason = "写入后自动回读核对实际值",
                Args = Build(map,
                    ("address", address),
                    ("expected_values", new List<object> { new List<object> { value } }),
                    ("type_aware", true))
            };
            return true;
        }

        private static bool TryPlanRangeWrite(Dictionary<string, object> map, out AutoVerifyPlan plan)
        {
            plan = null;
            var address = GetString(map, "address");
            var values = map.ContainsKey("values") ? map["values"] : null;
            if (string.IsNullOrWhiteSpace(address) || values == null)
            {
                return false;
            }

            // 公式模式（或 auto 下含公式文本）回读的是计算值而非公式文本，预期比对必然错位，不做自动自检。
            var valueMode = GetString(map, "value_mode") ?? "auto";
            if (valueMode == "formula")
            {
                return false;
            }

            if (valueMode == "auto" && ContainsFormulaText(values))
            {
                return false;
            }

            plan = new AutoVerifyPlan
            {
                VerifyTool = "verify_write_result",
                Reason = "写入后自动回读逐格核对实际值",
                Args = Build(map,
                    ("address", address),
                    ("expected_values", values),
                    ("type_aware", true))
            };
            return true;
        }

        private static bool TryPlanCleanChecks(Dictionary<string, object> map, List<object> checks, out AutoVerifyPlan plan)
        {
            plan = null;
            if (checks.Count == 0 || string.IsNullOrWhiteSpace(GetString(map, "address")))
            {
                return false;
            }

            plan = new AutoVerifyPlan
            {
                VerifyTool = "verify_data_clean",
                Reason = "清洗后自动核对清洗目标是否达成",
                Args = Build(map,
                    ("address", GetString(map, "address")),
                    ("checks", checks))
            };
            return true;
        }

        private static List<object> CleanChecksFromOperations(Dictionary<string, object> map)
        {
            var checks = new List<object>();
            if (!map.TryGetValue("operations", out var raw) || !(raw is System.Collections.IEnumerable sequence) || raw is string)
            {
                return checks;
            }

            foreach (var op in sequence)
            {
                var dict = op as Dictionary<string, object> ?? AsDict(op);
                var opName = dict == null ? null : GetString(dict, "tool");
                if (string.IsNullOrWhiteSpace(opName))
                {
                    continue;
                }

                switch (opName)
                {
                    case "remove_blank_rows":
                        checks.Add(new Dictionary<string, object> { ["type"] = "no_blank_rows" });
                        break;
                    case "fill_blanks":
                        var check = new Dictionary<string, object> { ["type"] = "no_blanks" };
                        if (dict.ContainsKey("column"))
                        {
                            check["columns"] = new List<object> { dict["column"] };
                        }

                        checks.Add(check);
                        break;
                    case "fix_dates":
                        checks.Add(new Dictionary<string, object> { ["type"] = "dates_valid" });
                        break;
                }
            }

            return checks;
        }

        private static bool TryPlanCellFormat(Dictionary<string, object> map, out AutoVerifyPlan plan)
        {
            plan = null;
            var expected = Build(map,
                ("expected_number_format", GetString(map, "number_format")),
                ("expected_bold", GetBool(map, "bold")),
                ("expected_italic", GetBool(map, "italic")),
                ("expected_font_color", GetString(map, "font_color")),
                ("expected_back_color", GetString(map, "back_color")));

            // 没有任何可断言的预期属性时不自动自检，避免产生无断言的空核验。
            if (expected.Count == 0)
            {
                return false;
            }

            var pairs = new List<(string key, object value)> { ("address", GetString(map, "address")) };
            pairs.AddRange(expected.Select(kv => (kv.Key, kv.Value)));
            plan = new AutoVerifyPlan
            {
                VerifyTool = "verify_cell_format",
                Reason = "格式设置后自动核对实际格式与预期一致",
                Args = Build(map, pairs.ToArray())
            };
            return true;
        }

        private static bool TryPlanConditional(Dictionary<string, object> map, out AutoVerifyPlan plan)
        {
            plan = null;
            var rule = GetString(map, "rule");
            if (string.IsNullOrWhiteSpace(rule) || rule == "clear")
            {
                return false;
            }

            // data_bar 的颜色写在 BarColor 上而非 Interior/Font，verify 比对填充色会误报 mismatch，
            // 因此 data_bar 只断言规则类型；其余规则传填充/字体色断言。
            var isDataBar = rule == "data_bar";
            plan = new AutoVerifyPlan
            {
                VerifyTool = "verify_conditional_format",
                Reason = "条件格式设置后自动核对规则类型与颜色",
                Args = Build(map,
                    ("address", GetString(map, "address")),
                    ("expected_rule", rule),
                    ("expected_formula1", isDataBar ? null : map.ContainsKey("value") ? Convert.ToString(map["value"]) : null),
                    ("expected_formula2", isDataBar ? null : map.ContainsKey("value2") ? Convert.ToString(map["value2"]) : null),
                    ("expected_fill_color", isDataBar ? null : GetString(map, "fill_color")),
                    ("expected_font_color", isDataBar ? null : GetString(map, "text_color")))
            };
            return true;
        }

        /// <summary>
        /// 用主工具的执行结果富化自检计划（在主工具成功后、执行 verify 前调用），返回 false 表示计划应被丢弃：
        /// 透视汇总——从结果取实际输出地址（含未传 target_address 时的自动落点）、数据行数与守恒对账参数；
        /// 追加行——目标地址由锚点列末行推得，只能从结果 AffectedRanges 取，取不到则放弃本次自检。
        /// </summary>
        public static bool EnrichFromResult(
            AutoVerifyPlan plan, string tool, object resultData, IReadOnlyList<string> affectedRanges = null)
        {
            if (plan == null)
            {
                return false;
            }

            switch (tool)
            {
                case "create_excel_pivot_summary":
                    if (resultData == null)
                    {
                        return true; // 无结果可富化，保持原计划；verify_pivot 缺断言会如实给出 no_expected
                    }

                    if (JToken.FromObject(resultData) is JObject data)
                    {
                        var target = data.Value<string>("target_address");
                        if (!string.IsNullOrWhiteSpace(target))
                        {
                            plan.Args["target_address"] = target;
                        }

                        var rowCount = data.Value<long?>("row_count");
                        if (rowCount.HasValue && rowCount.Value > 0)
                        {
                            plan.Args["expected_count"] = (int)rowCount.Value;
                        }
                    }

                    return true;
                case "export_sheets_to_pdf":
                case "export_excel_chart_image":
                    // M28 审查自纠：plan 用用户实参 file_path，但工具经 PathGuard.GuardAndComplete
                    // 会补全扩展名/绝对路径（x → x.pdf）——必须用结果里的落盘路径覆盖，否则自检
                    // 对"未补全路径"判 missing，把成功导出误判失败。
                    if (resultData != null && JToken.FromObject(resultData) is JObject exportData)
                    {
                        var completedPath = exportData.Value<string>("file_path");
                        if (!string.IsNullOrWhiteSpace(completedPath))
                        {
                            plan.Args["file_path"] = completedPath;
                            return true;
                        }
                    }

                    return false; // 拿不到落盘路径 → 丢弃自检（宁缺勿误报）
                case "append_excel_rows":
                    var (sheetName, address) = SplitAffectedRange(affectedRanges);
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        return false;
                    }

                    plan.Args["address"] = address;
                    if (!string.IsNullOrWhiteSpace(sheetName))
                    {
                        plan.Args["sheet"] = sheetName;
                    }

                    return true;
                default:
                    return true;
            }
        }

        /// <summary>拆 "SheetName!$A$5:$B$7" 形态的受影响区域（按最后一个 '!' 拆，地址部分不含 '!'）。</summary>
        private static (string sheetName, string address) SplitAffectedRange(IReadOnlyList<string> affectedRanges)
        {
            if (affectedRanges == null)
            {
                return (null, null);
            }

            foreach (var range in affectedRanges)
            {
                if (string.IsNullOrWhiteSpace(range))
                {
                    continue;
                }

                var bang = range.LastIndexOf('!');
                if (bang > 0 && bang < range.Length - 1)
                {
                    return (range.Substring(0, bang), range.Substring(bang + 1));
                }
            }

            return (null, null);
        }

        /// <summary>
        /// M27 审查 B：import_csv 的 sheet_name 是可选参数，引擎缺省时按 CSV_文件名 推导表名。
        /// 自检 plan 必须复刻同一推导——否则 plan 无期望值 → verify_structure 判 no_expected → fail-closed
        /// 把「未传 sheet_name 的合法导入」误判失败。
        /// </summary>
        private static string ResolveCsvSheetName(IReadOnlyDictionary<string, object> map)
        {
            // FIX（M27R2 审查 B）：引擎对最终表名无条件清洗——显式名同样要走
            // 「非法字符→_、31 截断、空回退」，否则自检核对到的名字与真实建表不同。
            var raw = GetString(map, "sheet_name");
            if (string.IsNullOrWhiteSpace(raw))
            {
                var filePath = GetString(map, "file_path");
                if (string.IsNullOrWhiteSpace(filePath)) { return null; }
                raw = "CSV_" + System.IO.Path.GetFileNameWithoutExtension(filePath);
            }

            var name = raw;
            foreach (var bad in new[] { '[', ']', ':', '*', '?', '/', '\\' }) { name = name.Replace(bad, '_'); }
            if (name.Length > 31) { name = name.Substring(0, 31); }
            if (name.Length == 0) { name = "CSV_import"; }
            return name;
        }

        private static Dictionary<string, object> Build(
            IReadOnlyDictionary<string, object> source, params (string key, object value)[] pairs)
        {
            var result = new Dictionary<string, object>();
            foreach (var (key, value) in pairs)
            {
                if (value != null && !(value is JToken jt && jt.Type == JTokenType.Null))
                {
                    result[key] = value;
                }
            }

            // sheet 等上下文参数原样透传，verify 工具按同一目标工作表回读。
            if (source.TryGetValue("sheet", out var sheet) && sheet != null)
            {
                result["sheet"] = sheet;
            }

            return result;
        }

        private static bool StartsWithFormula(object value)
        {
            var text = value == null ? null : Convert.ToString(value);
            return !string.IsNullOrWhiteSpace(text) && text.TrimStart().StartsWith("=", StringComparison.Ordinal);
        }

        /// <summary>
        /// 判断区域地址是否为可整体回读的小区域（行列可解析且格数 ≤ 上限）。
        /// 地址解析不了（如整列引用 A:C、非矩形）返回 false——宁可不自动自检，不做无依据断言。
        /// </summary>
        private static bool IsBoundedSmallRange(string address, int maxCells)
        {
            if (string.IsNullOrWhiteSpace(address) || address.Contains(':') == false)
            {
                return address != null && !address.Contains(':'); // 单格地址恒为小区域
            }

            var parts = address.Split(':');
            if (parts.Length != 2)
            {
                return false;
            }

            if (!TryCellIndex(parts[0], out var row1, out var col1) || !TryCellIndex(parts[1], out var row2, out var col2))
            {
                return false;
            }

            var rows = Math.Abs(row2 - row1) + 1L;
            var cols = Math.Abs(col2 - col1) + 1L;
            return rows * cols <= maxCells;
        }

        /// <summary>把 A1/$A$1 形态的单元格引用解析为行列号；整列（A:A）/整行（1:1）/非法引用返回 false。</summary>
        private static bool TryCellIndex(string cell, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(cell))
            {
                return false;
            }

            var text = cell.Trim().Replace("$", string.Empty);
            var split = 0;
            while (split < text.Length && char.IsLetter(text[split]))
            {
                split++;
            }

            if (split == 0 || split == text.Length || !int.TryParse(text.Substring(split), out row) || row < 1)
            {
                return false;
            }

            long col = 0;
            foreach (var ch in text.Substring(0, split).ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z')
                {
                    return false;
                }

                col = col * 26 + (ch - 'A' + 1);
                if (col > 16384)
                {
                    return false;
                }
            }

            column = (int)col;
            return true;
        }

        private static bool ContainsFormulaText(object values)
        {
            if (values == null)
            {
                return false;
            }

            // JValue 也实现了 IEnumerable（枚举为空），必须先于集合分支处理，否则标量值检测不到。
            if (values is string text)
            {
                return StartsWithFormula(text);
            }

            if (values is JValue jv)
            {
                return StartsWithFormula(jv.ToString());
            }

            if (values is System.Collections.IEnumerable sequence)
            {
                foreach (var item in sequence)
                {
                    if (ContainsFormulaText(item))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string GetString(IReadOnlyDictionary<string, object> map, string key)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            if (value is JValue jv)
            {
                return jv.Type == JTokenType.Null ? null : Convert.ToString(jv.Value);
            }

            return Convert.ToString(value);
        }

        private static int? GetInt(IReadOnlyDictionary<string, object> map, string key)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            try
            {
                if (value is JValue jv && jv.Type != JTokenType.Null)
                {
                    return Convert.ToInt32(jv.Value);
                }

                return Convert.ToInt32(value);
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
        }
        }

        private static double? GetDouble(IReadOnlyDictionary<string, object> map, string key)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            try
            {
                if (value is JValue jv && jv.Type != JTokenType.Null)
                {
                    return Convert.ToDouble(jv.Value);
                }

                return Convert.ToDouble(value);
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
        }
        }

        private static bool? GetBool(IReadOnlyDictionary<string, object> map, string key)
        {
            if (!map.TryGetValue(key, out var value) || value == null)
            {
                return null;
            }

            try
            {
                if (value is JValue jv && jv.Type != JTokenType.Null)
                {
                    return Convert.ToBoolean(jv.Value);
                }

                return Convert.ToBoolean(value);
            }
            catch (FormatException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
        }
        }

        private static Dictionary<string, object> AsDict(object op)
        {
            try
            {
                return JObject.FromObject(op).ToObject<Dictionary<string, object>>();
            }
            catch (ArgumentException)
            {
                return null;
            }
        }
    }
}
