# WPS 版工具分类审查报告（逐类明细，M6 收官）

> 版本：v1.0（2026-09-08）　基线：3.0.0.0（2.19.7.0 基线的 M6 审查结论经 3.0.0.0 full 复跑一致）　宿主：WPS 表格 12.1.0.26899 个人版（32 位）
> 数据源：WpsToolHarness full 模式真机执行（优先最新轮 F9b）+ battery 电池补充；执行记录见 `_dev_tmp/m6_full_f9b.txt` 等。

## 一、结论摘要

| 品类 | 工具数 | ✅ 通过 | 🔶 降级 | ⚪ 不适用 | ❌ 复核 | 结论 |
|---|---|---|---|---|---|---|
| 撤销与操作历史 | 1 | 0 | 1 | 0 | 0 | 正常 |
| 读取与查找 | 8 | 8 | 0 | 0 | 0 | 正常 |
| 写入与批量写 | 6 | 6 | 0 | 0 | 0 | 正常 |
| 筛选与命名区域 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 行列与结构操作 | 3 | 3 | 0 | 0 | 0 | 正常 |
| 工作表结构管理 | 3 | 3 | 0 | 0 | 0 | 正常 |
| 数据加工与条件格式 | 6 | 6 | 0 | 0 | 0 | 正常 |
| 清洗与合并工具族 | 9 | 9 | 0 | 0 | 0 | 正常 |
| JYYJ·文本处理 | 6 | 6 | 0 | 0 | 0 | 正常 |
| JYYJ·数值处理 | 5 | 5 | 0 | 0 | 0 | 正常 |
| JYYJ·行列与表格变换 | 11 | 11 | 0 | 0 | 0 | 正常 |
| JYYJ·工作表与对比 | 4 | 4 | 0 | 0 | 0 | 正常 |
| JYYJ·复制粘贴增强 | 3 | 3 | 0 | 0 | 0 | 正常 |
| AI 智能填充 | 2 | 2 | 0 | 0 | 0 | 正常 |
| AI 计算草稿 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 诊断标记 | 1 | 1 | 0 | 0 | 0 | 正常 |
| 批注与数据验证 | 3 | 3 | 0 | 0 | 0 | 正常 |
| 公式解释 | 1 | 1 | 0 | 0 | 0 | 正常 |
| 公式与引用 | 7 | 7 | 0 | 0 | 0 | 正常 |
| Markdown 互通 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 页面设置与 PDF 导出 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 工作簿合并拆分 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 外部链接 | 2 | 2 | 0 | 0 | 0 | 正常 |
| 外部文件导入 | 4 | 4 | 0 | 0 | 0 | 正常 |
| 统计分析与图表 | 11 | 11 | 0 | 0 | 0 | 正常 |
| 操作自检族（verify_*） | 9 | 9 | 0 | 0 | 0 | 正常 |
| Web 抓取 | 1 | 0 | 1 | 0 | 0 | 正常 |
| VBA / Python / 受控执行 | 7 | 3 | 4 | 0 | 0 | 正常 |
| 自定义沉淀样本（不适用） | 7 | 0 | 0 | 7 | 0 | 正常 |
| 其他 | 1 | 0 | 0 | 1 | 0 | 正常 |
| **合计** | **131** | 117 | 6 | 8 | 0 | 覆盖 124/131 |

**总体结论：WPS 版全部 30 个分类无真故障档（❌=0），逐类审查通过。** 合计 131 = 运行时注册 130 + run_workflow（静态存在、本环境未激活）。运行时 130 中 124 个有真机执行记录，覆盖 95.4%。

## 二、逐类明细

### 撤销与操作历史

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| undo_excel_operation | - | FAIL UNDO_UNAVAILABLE | 🔶 🔶 预期降级（harness 未注入操作历史，产品侧正常） | full-F9b |

### 读取与查找

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| find_excel_cells | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| get_current_excel_context | - | OK - | ✅ 真机通过 | full-F9b |
| get_current_selection_values_and_formulas | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| get_excel_range_values_and_formulas | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| get_excel_sheet_list | ReadOnly | OK - | ✅ 真机通过 | battery |
| preview_excel_range | - | OK - | ✅ 真机通过 | full-F9b |
| preview_excel_sheet | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| select_excel_range | ReadOnly | OK - | ✅ 真机通过 | full-F9b |

### 写入与批量写

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| append_excel_rows | - | OK - | ✅ 真机通过 | battery |
| clear_excel_range | - | OK - | ✅ 真机通过 | full-F9b |
| create_toc_sheet | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| write_excel_cell | - | OK - | ✅ 真机通过 | full-F9b |
| write_excel_cells_batch | NormalWrite | OK - | ✅ 真机通过 | battery |
| write_excel_range | - | OK - | ✅ 真机通过 | full-F9b |

### 筛选与命名区域

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| manage_excel_filter | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| manage_excel_named_range | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### 行列与结构操作

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| copy_paste_excel_range | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| delete_excel_dimension | Dangerous | OK - | ✅ 真机通过 | full-F9b |
| insert_excel_dimension | StructuralChange | OK - | ✅ 真机通过 | full-F9b |

### 工作表结构管理

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| add_excel_worksheet | StructuralChange | OK - | ✅ 真机通过 | battery |
| delete_excel_worksheet | Dangerous | OK - | ✅ 真机通过 | battery |
| rename_excel_worksheet | StructuralChange | OK - | ✅ 真机通过 | battery |

### 数据加工与条件格式

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| dedup_range | - | OK - | ✅ 真机通过 | battery |
| find_replace_range | - | OK - | ✅ 真机通过 | battery |
| format_excel_conditional | - | OK - | ✅ 真机通过 | battery |
| format_excel_range | - | OK - | ✅ 真机通过 | full-F9b |
| sort_range | - | OK - | ✅ 真机通过 | battery |
| transpose_range | - | OK - | ✅ 真机通过 | battery |

### 清洗与合并工具族

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| clean_excel_data | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| date_normalize | - | OK - | ✅ 真机通过 | full-F9b |
| fill_formula | - | OK - | ✅ 真机通过 | full-F9b |
| filter_excel_data | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| merge_sheets | - | OK - | ✅ 真机通过 | full-F9b |
| number_pad | - | OK - | ✅ 真机通过 | full-F9b |
| text_clean | - | OK - | ✅ 真机通过 | full-F9b |
| text_merge | - | OK - | ✅ 真机通过 | full-F9b |
| text_split | - | OK - | ✅ 真机通过 | full-F9b |

### JYYJ·文本处理

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| amount_to_uppercase | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| convert_chinese_variant | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| convert_text_to_pinyin | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| delete_text_matches | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| extract_text_characters | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| insert_text_into_cell | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### JYYJ·数值处理

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| bankers_round_values | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| fill_number_sequence | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| fill_random_numbers | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| generate_permutation_list | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| pick_numbers_to_sum | ReadOnly | OK - | ✅ 真机通过 | full-F9b |

### JYYJ·行列与表格变换

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| convert_1d_to_2d | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| convert_2d_to_1d | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| create_pivot_table | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| cross_reference_columns | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| delete_blank_rows_range | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| fill_blank_cells_down | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| fuzzy_match_columns | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| insert_blank_rows_every_n | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| make_payslip_rows | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| merge_duplicate_cells | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| wrap_formulas_with_iferror | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### JYYJ·工作表与对比

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| compare_two_sheets | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| create_sheet_by_value | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| protect_sheet_watermark | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| split_sheet_by_field | StructuralChange | OK - | ✅ 真机通过 | full-F9b |

### JYYJ·复制粘贴增强

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| copy_range_as_image | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| copy_visible_range_only | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| smart_paste_delimiter | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### AI 智能填充

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| ai_fill_blank_cells | - | OK - | ✅ 真机通过 | full-F9b |
| apply_fill_blank_cells | - | OK - | ✅ 真机通过 | full-F9b |

### AI 计算草稿

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| append_calculation_scratchpad | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| clear_calculation_scratchpad | StructuralChange | OK - | ✅ 真机通过 | full-F9b |

### 诊断标记

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| flag_cells | - | OK - | ✅ 真机通过 | full-F9b |

### 批注与数据验证

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| cell_comment | - | OK - | ✅ 真机通过 | full-F9b |
| data_validation | - | OK - | ✅ 真机通过 | battery |
| rate_change_analysis | - | OK - | ✅ 真机通过 | full-F9b |

### 公式解释

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| explain_excel_formula | - | OK - | ✅ 真机通过 | full-F9b |

### 公式与引用

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| convert_formula_references | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| copy_visible_rows_to_sheet | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| drill_excel_range | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| merge_or_split_cells | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| reset_cell_value_type | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| trace_formula_references | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| transform_excel_formulas | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### Markdown 互通

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| export_selection_to_markdown | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| import_markdown_to_excel | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### 页面设置与 PDF 导出

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| apply_page_setup | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| export_sheets_to_pdf | NormalWrite | OK - | ✅ 真机通过 | battery |

### 工作簿合并拆分

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| merge_workbooks | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| split_workbook | StructuralChange | OK - | ✅ 真机通过 | full-F9b |

### 外部链接

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| list_external_links | ReadOnly | OK - | ✅ 真机通过 | battery |
| replace_external_links | Dangerous | OK - | ✅ 真机通过 | full-F9b |

### 外部文件导入

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| import_delimited_file | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| import_excel_file_sheet | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| merge_delimited_files | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| preview_external_file | Dangerous | OK - | ✅ 真机通过 | full-F9b |

### 统计分析与图表

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| analyze_column_correlation | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| analyze_range_statistics | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| analyze_value_distribution | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| create_excel_chart | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| create_excel_pivot_summary | NormalWrite | OK - | ✅ 真机通过 | battery |
| create_excel_sparkline | NormalWrite | OK - | ✅ 真机通过 | battery |
| export_excel_chart_image | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| generate_analysis_report | NormalWrite | OK - | ✅ 真机通过 | full-F9b |
| insert_recommended_chart | StructuralChange | OK - | ✅ 真机通过 | full-F9b |
| recommend_excel_chart | ReadOnly | OK - | ✅ 真机通过 | full-F9b |
| verify_chart | ReadOnly | OK - | ✅ 真机通过 | full-F9b |

### 操作自检族（verify_*）

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| verify_cell_format | - | OK - | ✅ 真机通过 | full-F9b |
| verify_conditional_format | - | OK - | ✅ 真机通过 | full-F9b |
| verify_data_clean | - | OK - | ✅ 真机通过 | full-F9b |
| verify_formula | - | OK - | ✅ 真机通过 | full-F9b |
| verify_marker | - | OK - | ✅ 真机通过 | full-F9b |
| verify_pivot | - | OK - | ✅ 真机通过 | full-F9b |
| verify_sort_filter | - | OK - | ✅ 真机通过 | full-F9b |
| verify_structure | - | OK - | ✅ 真机通过 | full-F9b |
| verify_write_result | - | OK - | ✅ 真机通过 | full-F9b |

### Web 抓取

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| fetch_url | - | FAIL NETWORK_ERROR | 🔶 🔶 预期降级（测试网络目标不可达） | full-F9b |

### VBA / Python / 受控执行

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| run_cli_command | Dangerous | FAIL NOT_ALLOWLISTED | 🔶 🔶 预期降级（CLI/EXE 白名单未配置） | full-F9b |
| run_exe_command | Dangerous | FAIL NOT_ALLOWLISTED | 🔶 🔶 预期降级（CLI/EXE 白名单未配置） | full-F9b |
| run_python_code | Dangerous | OK - | ✅ 真机通过 | full-F9b |
| run_python_excel | Dangerous | FAIL PYTHON_BITNESS_MISMATCH | 🔶 🔶 预期降级（本机 Python 64 位与 32 位宿主不匹配） | full-F9b |
| run_vba_snippet | Dangerous | FAIL VBA_UNAVAILABLE | 🔶 🔶 预期降级（WPS 个人版无 VBA） | full-F9b |
| save_vba_tool | - | OK - | ✅ 真机通过 | full-F9b |
| validate_vba_snippet | NormalWrite | OK - | ✅ 真机通过 | full-F9b |

### 自定义沉淀样本（不适用）

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| ct_format_answer_sheet | - | FAIL VBA_UNAVAILABLE | ⚪ 不适用（自定义沉淀样本，VBA 执行体） | full-F9b |
| ct_proc_FixSalesData | - | - | ❓ 未执行（自定义样本） | - |
| ct_proc_FormatAlignment | - | - | ❓ 未执行（自定义样本） | - |
| ct_proc_FormatAnswerSheet | - | - | ❓ 未执行（自定义样本） | - |
| ct_vba_tool | - | - | ❓ 未执行（自定义样本） | - |
| ct_vba_tool_2 | - | - | ❓ 未执行（自定义样本） | - |
| ct_vba_tool_3 | - | - | ❓ 未执行（自定义样本） | - |

### 其他

| 工具 | 风险级 | 真机结果 | 判定 | 来源 |
|---|---|---|---|---|
| run_workflow | NormalWrite | - | ❓ 未执行（自定义样本） | - |

## 三、口径说明

1. **判定优先级**：同一工具多轮执行取最新一轮结果；full-F9b 为最终轮（含 4 项契约缺陷修复后的复跑），更早的 full-final2 与 battery 仅补位。
2. **🔶 降级档**均为环境/权限预期行为，错误码结构化、不崩溃、不静默失败，符合"硬缺口显式化"验收要求。
3. **⚪ 不适用**：ct_* 共 7 个为历史会话 save_vba_tool 沉淀的自定义工具样本，执行体为 VBA，WPS 个人版下由宿主感知过滤或 VBA_UNAVAILABLE 兜底，非产品静态注册工具；run_workflow 为条件注册工具，本环境未激活故未进入运行时名单（130 的口径不含它）。
4. **本轮审查修复的缺陷落点**：工作簿合并拆分类（merge_workbooks/split_workbook，数组契约）、外部文件导入类（merge_delimited_files，数组契约）、AI 计算草稿类（scratchpad 草稿表不存在即崩）、清洗合并类历史修复（ReadMatrix 空表越界）、写入类历史修复（NormalizeValues int[] 封送）——全部已双副本同步 + 872 单测回归通过。
5. 风险级 ReadOnly/NormalWrite/StructuralChange/Dangerous 取自注册源码，Dangerous 类在产品侧另有确认门。
6. 同步副本：`D:\execl1.0\docs\WPS工具分类审查报告.md`。

## 四、M6 复审（M7 专项，2026-09-08 下午）补充

M6 收官后对抗性源码复审发现 4 项缺陷并双副本修复（import `{name}` 契约漏修调用点、merge_workbooks/import 静默 OK×2、SidebarBridge save_vba_tool 提示词与清单矛盾），harness 新增 `m7` 专项模式真机验证全过。明细与证据见 `WPS工具兼容差异清单.md` §七、`踩坑记录与规避手册.md` §7.21。本报告 §一/§二 数据不受影响（缺陷均属"schema 对象项传参形态未被 harness 覆盖"的盲区，非已列档工具行为翻转）。

**3.0.0.0 发布复跑（2026-09-08 14:36）**：full 全量重跑 111 OK + 7 预期降级 + 0 EXCEPTION（`_dev_tmp/m8_full_3000.txt`），逐类结论与 §一/§二 完全一致；m7 契约专项 3/3。打包门禁三关全绿（单测 872 + Excel 真机 80/80 + WPS full），产物 `dist/installer/JYYJ助手-3.0.0.0-Setup-wps.exe`。

## 五、M10 审查（2026-09-08 16:00）：WPS 指南/函数库双视图 + 模型方案/Python/VBA 三链路专项

### 5.1 新增功能（用户需求）
1. **设置页「WPS 环境 · 指南」卡片**：新增桥接 `SidebarBridge.GetWpsEnvironment()`——探测本机 WPS 安装（三根候选取最高版 office6\et.exe，与 register-auto.ps1 同源口径）、et.exe 文件版本、PE 头位数、VBA 环境（WPS 宿主内 VBE 直测 / 进程外探测 VBA 7.1 运行库 VBE7.dll 落点）、助手注册状态（ET Addins 键 LoadBehavior）；个人版/专业版按版本系推断并明确标注。卡片含四个 details：安装说明（VBA 三条真实地址）、使用帮助、注意事项、WPS vs Excel 功能/工具差异（130 工具口径）。
2. **函数库 Excel/WPS 双视图**：`按 Excel / 按 WPS` 切换；WPS 视图逐函数打标（WPS·支持 / WPS·新内核 / WPS·随内核），规则基于引擎依赖分组（LAMBDA/动态数组/正则族=随内核），不逐函数虚构结论；详情页附「示例试运行」实测引导。

### 5.2 三视角对抗审查结论（模型方案 / Python / VBA）
| 视角 | 模型方案 | Python | VBA |
|---|---|---|---|
| 工具实现者 | DPAPI CurrentUser 加密 ApiKey；TestConnection 30s 超时 | probe 一次拿版本/位数/依赖；worker 常驻+reset；write_ranges 快照回滚 | 每轮独立运行模块+finally 整模块回收；探针重试自愈 0x800A9C68 |
| 审查者 | 错误结构化透传（"没有返回有效响应"不暴露内部）；profile 空值保留旧 Key | PYTHON_NOT_FOUND/VERSION_MISMATCH/BITNESS_MISMATCH/DEPS_INSTALL_FAILED 四级显式错误码；超时杀进程 | PreRunVbaGate 确定性拦截；VBA_NO_PROCEDURE/INVALID 显式化；截图级中文错误回喂 |
| AI 使用方 | 测试失败信息可直接转告用户 | input_data 传大数据替代 args 限流提示 | 报错附静态修复建议（BuildStaticAdvice） |
| **本轮新发现** | 无硬伤 | 无硬伤 | 无硬伤（架构已经 P1/P3 多轮沉淀，注释含决定性实验记录） |

### 5.3 工具闭环 / 真机验证证据
- **真实报错捕获**：平台构建暴露 CS0165（vbaEvidence 未赋值）——sln 构建走 AnyCPU 不会编平台输出，`-p:Platform=x64/x86` 显式构建才抓到；已修复（初始化 null）并沉淀新坑（见踩坑手册 §7.22）。
- **单测**：872/872 通过（构建 0 警告）。
- **WPS 真机 full**：111 OK + 7 环境预期降级 + 0 EXCEPTION（`_dev_tmp/m10_full.txt`，与 M8 基线一致）；m7 契约专项 3/3（`m10_m7.txt`）。
- **环境探测真机口径验证**：et.exe 12.1.0.26899 / 32 位 / VBE7.dll 未检出（与 harness VBA_UNAVAILABLE 相互印证）/ LoadBehavior=3（`_dev_tmp/wpsenv_probe.txt`）。
- **清单约束兜底**：GetWpsEnvironment 全部分支结构化错误码；个人版/专业版推断显式标注"推断"防 AI 过度声明；函数库 WPS 标注三条规则可解释、附实测路径。
- **交付**：双架构 restage → zip 重打 → ISCC 重出 `JYYJ助手-3.0.0.0-Setup-wps.exe` → 真机静默重装成功，安装落位 DLL 字节级验证含 GetWpsEnvironment。

## 六、M11 审查（2026-09-08 17:00）：自检补强 + 日志增强（v3.1.0.0）

### 6.1 自检体系扩容

| 维度 | 内容 |
|---|---|
| 新 verify 工具 ×3 | verify_sparkline（SparklineGroups 类型/落点/数据源）、verify_data_validation（Validation 类型/来源/运算符/边界）、verify_merge（MergeCells true/false/null 三态）；全部 ReadOnly + 判定器纯静态可单测 + 逐属性 try/catch 容错（WPS 兼容口径） |
| verify_data_clean 新检查项 | no_text_matches（删除文本的确定性期望，正则 500ms 超时护栏 + 非法 pattern fail-closed）、all_blank（清空的确定性期望，全矩阵含表头扫描） |
| 自动映射 +6 | clear_excel_range→all_blank、delete_text_matches→no_text_matches、data_validation→verify_data_validation、merge_or_split_cells（merge/merge_similar→true，unmerge/fill_merged→false）、merge_duplicate_cells→true、create_excel_sparkline→verify_sparkline；豁免清单 83→77 条，MappedPrimaryTools 16→22 类，覆盖审计（一致性锁测试）同步 |
| fail-safe 边界 | 整列引用/超大区域（>20000 格）不出自动自检计划，不做无依据断言 |

### 6.2 日志增强（成熟方案=既有 环形缓冲+文件 落盘模式延伸）

1. 自检失败明细：verify 未通过时按 Error 级落 issues 全文（此前只有错误码），定位"哪步核对没过"不再反推；
2. 回合自检汇总：`回合自检汇总 可自检=N 已执行=N 跳过=N` 落 [verify] 日志；
3. 文件轮转：`%TEMP%\ExcelAiAssistant.log` 超 5MB 滚出时间戳备份、保留最近 5 份，防无限膨胀；
4. 会话首行：版本+宿主+进程位数+.NET 版本一次留全。

### 6.3 对抗性审查与真机证据

- **R1（源码复审）**：修 2 处——迷你图落点匹配由子串包含改为边界感知匹配（防 dest="B1" 误匹配 location="AB1"）；豁免清单残留"已接入"说明行删除（一致性锁测试抓到）。
- **R2（真机闭环）**：full 模式暴露 ReadCleanChecks 只认 JArray、harness 进程内 object[] 传 checks 被判 no_checks——补 IEnumerable 归一路径（同口径）。
- **R3（清单兜底）**：grep 残留清零；映射↔switch 一致性锁全过。
- **门禁**：单测 **890/890**（新增 18 判据/映射用例）；Excel 真机 **80/80**（两次复跑）；WPS 真机 full：**134 注册（130+3 verify+run_workflow）→ 120 OK + 7 预期降级（错误码与基线一致）+ 0 EXCEPTION**；m7 契约专项 **3/3**。证据 `_dev_tmp/m11_full_r2.txt`。
- **发版**：版本号三处红线同步（csproj/README/iss）→ 3.1.0.0 → 双架构 restage（落位 DLL 字节级验证含 wpsEnvAdvice/新 verify）→ archiver 重打 zip → ISCC 重出 `JYYJ助手-3.1.0.0-Setup-wps.exe`（4.66MB）。

## §七 M12：智能表 + 保护探测（对标调研闭环，2026-09-08）

外部对标调研（docs/自检与工具完备度外部对标调研报告.md）确认 9 项工具缺口，本轮落地 ROI 最高的两项：

- **create_excel_table**（StructuralChange）：区域转智能表 ListObject，支持命名/表头/样式；受保护表显式报 `SHEET_PROTECTED`。自动映射接入 verify_table（MappedPrimaryTools 23 类）。
- **verify_table**（ReadOnly）：按名称/地址定位 ListObject，核对存在性/表头/行列数；地址匹配用 A1 坐标包含判断（对抗审查修正子串误报，B2:C4 能命中 A1:C4 表、Z99 不会命中 A1:C4）。
- **verify_protection**（ReadOnly）：只读探测工作簿结构保护/工作表保护/单元格锁定/筛选排序许可；expected_writable=true 且受保护+锁定时显式报「写入将被拒绝（0x800A03EC 的典型成因）」。

**真机发现（重要）**：WPS ET 的 `ListObjects.Add` 实测会把区域向下扩展一行（请求 A1:C5 → 实得 A1:C6，两轮复现），与 Excel 尊重给定区域不同——已记入差异清单，verify_table 在 WPS 下建议不断言行数。

**门禁**：902/902 单测（+12）· Excel 真机 80/80 · WPS full 125 OK/7 预期降级/0 EXCEPTION · m7 契约口径与基线一致。发版 v3.2.0.0。

## §八 M13：图片 + 冻结窗格（对标调研增量继续，2026-09-08）

对标调研缺口清单继续落地（insert_image 高频 / freeze_panes 一行 COM 调用）：

- **insert_image**（StructuralChange）：Shapes.AddPicture 在指定锚点单元格左上角插入本地图片（LinkToFile=msoFalse / SaveWithDocument=msoTrue，随工作簿保存不链接源文件）；支持显式宽高（磅）或缺省按图片原生尺寸；受保护表显式报 `SHEET_PROTECTED`，文件不存在显式报 `IMAGE_NOT_FOUND`。
- **verify_image**（ReadOnly）：枚举 Shapes 中图片类（msoPicture=13/msoLinkedPicture=11，图表/文本框不计入），按 expected_count 核数量、按 anchor（单元格页面坐标 Left/Top/Width/Height）定位图片左上角落点、按 expected_width/expected_height 核尺寸（容差 2 磅）；尺寸断言不配 anchor 时 fail-closed（size_without_anchor）。
- **freeze_panes**（StructuralChange）：冻结/解冻窗口窗格。**WPS 真机关键差异**：Excel 惯用的「Select 锚点 + FreezePanes=true」在 WPS ET 被静默忽略（赋值后读回 false），必须先设 `SplitRow/SplitColumn` 再 `FreezePanes=true`（探针四写法实测，已记入差异清单）；实现统一走 Split 路径双宿主通用。A1 锚点显式报 `INVALID_ANCHOR`；宿主未接受冻结显式报 `FREEZE_FAILED`（防静默成功）。
- **verify_freeze_panes**（ReadOnly）：激活目标表后读 ActiveWindow 的 FreezePanes/SplitRow/SplitColumn，按 expected_frozen / expected_cell（锚点行列-1=期望分割量）判定；expected_cell 与 expected_frozen=false 组合矛盾 fail-closed。

**门禁**：914/914 单测（+12 判定器/映射用例）· Excel 真机 86/86（+用例23 图片、用例24 冻结/解冻自动自检链路）· WPS full 136 OK/9 预期降级/0 EXCEPTION（M13 四工具 13 行全符合预期，含 IMAGE_NOT_FOUND/INVALID_ANCHOR 两条负路径）。发版 v3.3.0.0（dist/installer/JYYJ助手-3.3.0.0-Setup-wps.exe + 双架构 zip）。


## §九 M13-R：图片+冻结窗格对抗性审查（三视角第二轮收敛，2026-09-08）

M13 发版后按用户要求加跑一轮对抗性审查（工具实现者/审查者/AI 使用方三视角），四项真实发现全部修复并真机取证：

- **F1（实现者）**：insert_image 的 width/height 无校验——0/负数会插入不可见图片或触发宿主异常，非数字抛 FormatException 被包装成误导性 EXCEL_COM_ERROR。修复：double.TryParse + 正数校验，违规显式报 INVALID_ARGS。
- **F2（审查者）**：AutoVerifyPolicy 的 insert_image 映射把 width/height 原样透传为 expected_width/height——用户传 -1（原生尺寸哨兵）时 verify_image 尺寸必然误报。修复：≤0 的哨兵值映射层剔除；新增单测回归锁（Insert_image_mapping_skips_non_positive_size_sentinels）。
- **F4（AI 使用方，Excel 真机抓到的关键差异）**：5 字节垃圾文件被 Excel 的 AddPicture **宽容接受**（静默插入无效图片），WPS 则抛 E_FAIL——宿主过滤器行为不可控。修复：前置魔数白名单嗅探（PNG/JPEG/GIF/BMP/TIFF/EMF/WMF），未识别格式触碰宿主前显式报 IMAGE_INVALID，双宿主行为确定一致；AddPicture 残余 COM 异常同样包装为 IMAGE_INVALID。
- **F6（AI 使用方）**：verify_freeze_panes 激活目标表的副作用 schema 已有披露，核实关闭，无需修改。

**门禁（第二轮）**：915/915 单测（+1 回归锁）· Excel 真机 90/90（用例23 扩为 4 条负路径断言：INVALID_ARGS ×2/IMAGE_NOT_FOUND/IMAGE_INVALID）· WPS full 147 行 = 136 OK + 11 预期 FAIL（4 负路径 + 7 降级）+ 0 EXCEPTION。发版 v3.3.1.0。


## §十 M14：列质量画像（analyze_column_profile，2026-09-08）

对标调研第 7 项 verify 盲区（Power Query 列质量/分布/剖析单面板）落地，新工具 **analyze_column_profile**（ReadOnly，工具总数 141→142）：

- **三联单面板**：质量（空值率/错误率/唯一值/类型占比 number·text·date·bool·error·blank）+ 分布（Top N 高频值，错误码计入）+ 剖析（数值列 min/max/均值/中位数/样本标准差；文本列长度统计/纯空白/带空格填充数），附质量提示 flag（blank_heavy/has_errors/mixed_types/padded_text/constant）。
- **纯引擎可单测**：核心抽为 ColumnProfileEngine（无 COM 依赖），10 个 xUnit 用例覆盖类型判定顺序/边界/flag。
- **真机闭环两轮收敛，抓到两个真实缺陷**：① Excel 错误单元格经 Value2 读回 **Int32 CVErr 整数**（探针实证 #N/A → -2146826246），被误判 number 污染统计 → 修复：7 个 CVErr 码映射回错误码文本，且错误码计入唯一值/分布（Power Query 同口径）；② 错误值分类进 error 后分布缺失 → 错误码补入分布。另修 harness 断言键名大小写 bug（命名类 JSON 序列化为 PascalCase）。
- **门禁**：925/925 单测（+11）· Excel 真机 98/98（用例 25：写入「3 数值+1 错误+1 空值」混合列 → 空值/错误/唯一值/均值 16.6667/has_errors 全命中 + INVALID_COLUMN 负路径）· WPS full 137 OK/12 预期 FAIL（1 M14 负路径 + 4 M13 负路径 + 7 降级）/0 异常。发版 v3.4.0.0。

## §十一 M15：超链接链路（manage_hyperlinks + verify_hyperlink，2026-09-08）

补齐超链接操作与自检盲区，新工具 **manage_hyperlinks**（NormalWrite，list/add/update/remove/remove_all 五动作）与 **verify_hyperlink**（ReadOnly，verify 家族第 18 维度，工具总数 142→144）：

- **manage_hyperlinks**：list 枚举整表（可按 address 过滤，返回 cell/text/url/subaddress/tooltip）；add 新建（link_url 外部 / link_subaddress 内部二选一，互斥显式 INVALID_ARGS；锚点已被覆盖显式 HYPERLINK_EXISTS；多格地址拒绝）；update 修改目标单元格链接；remove 按区域相交删除；remove_all 清空整表（空表显式 HYPERLINK_NOT_FOUND）。写动作前置工作表保护检查。
- **verify_hyperlink**：expected_count 整表总数 / address 定位 + expected_url/expected_subaddress/expected_text 内容断言 / expected_present=false 断言删除生效；fail-closed：内容断言必须与 address 搭配、present=false 与内容断言矛盾均显式报 invalid_expectation。宽松比对：url 忽略大小写与末尾斜杠、subaddress 忽略 $ 与引号、text 忽略大小写空白。
- **自动自检映射（25→26 类）**：add/update→verify_hyperlink（期望从实参推导）；remove→expected_present=false；remove_all→expected_count=0；list 不自检。
- **对抗性审查修复（发版前）**：F1 多格链接判重/删除只看首格（handle 记录 RowSpan/ColSpan，add 按「覆盖」判重、remove 按「相交」删除）；F2 AddHyperlinkAt 成功路径 added COM 引用不释放；F3 update 仅改显示文本映射不出自检计划；F4 外链 Add 统一 3 参避免宿主默认参数绑定不确定性。
- **真机探针实证修复（F5）**：Excel 对超链接赋 `SubAddress=""` 会把外链打回「内链空目标」（Address 读回 ''）——update 初版在设完 Address 后又赋空 SubAddress 导致地址丢失。探针（_dev_tmp/HyperlinkProbe）实证：外→外改 Address、任意→内改 SubAddress 均可靠生效；内→外切换属性赋值语义不可控 → 统一走 Delete+Add 重建（保留显示文本）。详见兼容差异清单。
- **门禁**：938/938 单测（+13）· Excel 真机 116/116（用例 26：外链/内链 add + 自动自检、list 计数、update + 自检、remove + expected_present 自检、remove_all + expected_count 自检、负路径 HYPERLINK_EXISTS/HYPERLINK_NOT_FOUND×3/INVALID_ARGS）· WPS full 148 OK/19 预期 FAIL（6 M15 负路径 + 1 M14 负路径 + 4 M13 负路径 + 8 降级）/0 异常。发版 v3.5.0.0。

## §十二 M16：全量分类审查（31 类 144 工具，2026-09-08）

对本项目全部 31 类 144 个工具做一轮完整的三视角对抗审查（工具实现者/审查者/AI 使用方），以 M13–M15 沉淀的 10 类缺陷原型为兜底清单，7 个并行审查分组 + 人工逐条核实（trust but verify），共确认修复如下：

### 数据正确性（P0/P1）
- **[P0] import_excel_file_sheet / open_excel_external_book**：COM `Workbooks.Open` 对用户已打开的文件会返回既有实例，finally 里 `Close(false)` 把用户未保存的工作簿无提示关闭。修复：按全路径比对 Workbooks 集合，命中则复用"借用"实例且不负责关闭（ExternalBookRegistry 增加 userOwned 标记）。
- **[P0] copy_visible_rows_to_sheet / copy_visible_range_only**：缺省写回源区域 → 可见行覆盖到源区顶部、隐藏/被筛选行数据被静默破坏。修复：缺省 target 时显式报 INVALID_ARGS（copy_visible_rows_to_sheet 在 target_sheet 已给出时仍允许缺省同形地址）。
- **[P1] CVErr 系统性污染**：错误单元格经 Value2 读回 Int32（如 #N/A=-2146826246）被当 number 参与统计/排序/凑数/写回（错误值静默变数字）。修复：新增全项目单一来源 `ExcelErrorValues`（7 码映射），ColumnProfileEngine 及 ExcelDataToolsService/ExcelUtilityToolsService/ExcelAnalysisService 三统计工具/ExcelJyyjtoolService（ToText2/IsBlank/Normalize）/ExcelAdvancedEditService（FlattenValues）/ExcelPivotChartService（ExtractNumbers）全部接入。
- **[P1] transpose_range 非方阵缺省原地转置**：转置矩阵（cols×rows）与源区域（rows×cols）不重合，尾部残留旧值。修复：补一次"残留清空"写（右侧/底部块），与区域替换语义一致且留撤销快照。
- **[P1] amount_to_uppercase**：缺省目标行偏移 +1（A1:A10 写到 B2:B11）。修复为与源起始行对齐。
- **[P1] JYYJ Resolve 静默回退**：表名不存在时静默落到活动表（数据写到错误表）。修复为显式 INVALID_SHEET（Compensable 增加 KeyNotFoundException→INVALID_SHEET、ArgumentException→INVALID_ARGS 映射）。

### 安全（P1）
- **run_cli_command 白名单 basename 绕过**：只按文件名评估白名单却按原样执行 command（`C://temp//python.exe` 借 basename "python" 命中）。修复：带路径成分的 command 按完整路径评估，裸命令名（PATH 解析）才按 basename。
- **fetch_url SSRF 两条**：① IPv4-mapped IPv6（::ffff:127.0.0.1 等）绕过 16 字节分支 → 映射回 IPv4 复核；② 主机名不解析放行 DNS 重绑定/xip.io → 解析主机名全部地址逐一校验，localhost 显式拒绝，解析失败 fail-closed。错误码 BLOCKED_PRIVATE_ADDR（harness 由 NETWORK_ERROR 转为该码，仍为预期降级）。

### RCW 清理（12 处）
verify_cell_format（死释放变量→真正捕获 Font/Interior/Borders）、verify_conditional_format（fc.Interior/Font）、verify_marker（逐格 Interior，最坏 2 万个）、verify_table（item.Range 捕获一次）、FindName/FindSheet（ExcelEditService/ScratchpadService 枚举未命中项）、GetSheetList 枚举、split_sheet_by_field/create_sheet_by_value 新建表、ValidateVbaSnippet（components/module/codeModule）、CreateBinnedHistogram（直方图 Range/Cells）、read_external_range（ws/rng + 守卫前移）、ExternalBookRegistry.Close、trace EnumerateAreas、BuildTemporaryWorkbook firstSheet。**新增策略红线：工作表/工作簿级共享 RCW 一律不做强制释放**（M16 真机 WPS 实证：强制释放会"RCW 分离"）。

### 契约与观测性
- get_current_excel_context 输出 PascalCase → 统一 snake_case 投影（ToSnakeCaseData，SnakeCaseNamingStrategy）。
- get_excel_sheet_list VeryHidden 误报为可见 → visible=（Visible==-1），新增 hidden_state 字段。
- verify 判定器无效期望值静默放行统一修复：conditional_format expected_rule、sort expected_direction、cell_format 对齐/颜色/边框 → invalid_expectation（10 条回归锁）。**两档口径文档化**：断言型 verify 无期望=no_expected 失败；探测型 verify（freeze/protection/image/hyperlink 等）无期望=ok（观察用途）。
- verify_hyperlink 多格链接按矩形相交匹配（与写侧口径对齐）。
- import/export markdown 管道符往返：导入侧按未转义 | 拆分并反转义。
- write_excel_range 形状拟合上报 refit_source_cells；flag_cells 上报 planned_cells + 实际成功数 + warning。
- split_workbook 失败路径补 Close + PathGuard；merge_workbooks 源路径过 PathGuard；convert_formula_references mode 显式校验；QuoteArg 改 Windows CRT `""` 转义。

### 门禁
948/948 单测（+10 M16 回归锁）· Excel 真机 116/116 · WPS full 148 OK/19 预期 FAIL/0 异常（create_sheet_by_value 与 import_excel_file_sheet 在 WPS 复跑恢复 OK，证明修复引入的回归已被双宿主闭环捕获并收敛）。发版 v3.6.0.0。

## §十三 M17：写路径工具保护态守卫全覆盖（2026-09-09）

### 背景与目标

M16 收口时留有两条下一步：缺陷原型并入自动检测、SHEET_PROTECTED 全覆盖。本轮取后者：此前全项目仅 ExcelStructureService 三处（create_excel_table/insert_image/manage_hyperlinks）有显式保护检查，其余写路径工具遇到受保护工作表会把原始 COM 异常（0x800A03EC 等）散落给 AI 使用方。M17 把保护态检查收敛为**宿主层集中守卫**，一处包装、全量覆盖、带理由豁免。

### 实现要点

- **集中守卫** `SheetProtectionGuard.WrapWriteTools`：在 `ExcelToolHost.CreateRegistry` 收口处对所有非 ReadOnly 工具统一包装前置探测；命中保护态直接返回 `SHEET_PROTECTED` 结构化错误（消息含表名、工具名与 verify_protection 指引）。
- **判定口径**：普通写工具按目标表 `ProtectContents`；add/rename/delete_excel_worksheet 三个结构类工具按工作簿 `ProtectStructure`（delete 追加目标表 ProtectContents）。目标表解析缺省 `sheet` 参数，`copy_visible_rows_to_sheet`/`merge_sheets` 覆盖为 `target_sheet`（`sheet` 是只读源表）。
- **豁免清单**（27 条，全部带理由）：脚本/进程类（run_vba_snippet 等 9 条）、工作簿/文件级（split_workbook 等 11 条）、视图级 freeze_panes、产出为新表类（split_sheet_by_field 等 4 条）、ct_ 前缀用户沉淀工具按前缀规则豁免。
- **加法前置原则**：探测阶段任何异常一律放行回工具自身错误路径——守卫只把"必然失败"提前为结构化错误，不制造新失败模式。
- **测试接缝**：判定逻辑抽为纯函数 `Evaluate(tool, ISheetProtectionProbe, args)`，生产实现 `ComSheetProtectionProbe` 走 dynamic COM；单测用假探针覆盖判定矩阵，COM 层由双宿主真机覆盖。

### 对抗审查闭环（两轮收敛）

1. **动态绑定可见性**：第一版判定直接用 dynamic 打测试假对象——DLR 绑定器只对调用点程序集（Excel）引用的程序集可见，测试假对象全部抛 RuntimeBinderException 被守卫 catch 吞成放行，出现"8 例失败"。重构为探针接缝后判定可测。
2. **RCW 分离回归（WPS 真机抓到）**：守卫探测枚举命中项用完即 FinalRelease——WPS 下 COM RCW 按 IUnknown 身份缓存，该释放把 harness 跨调用持有的同一工作表 RCW 一并"分离"（Unprotect 抛"COM 对象与其基础 RCW 分开后就不能再使用"，解保失效）。这是 M16"工作表级共享 RCW 不得 FinalRelease"红线的延伸：**枚举命中项同样不可释放，探测类短命对象一律交 GC 回收**。修复后 WPS m17 专项 6/6 全过。
3. **豁免审计联动**：单测锁定 AutoVerifyPolicy.MappedPrimaryTools 的 26 个主写工具不得豁免守卫（唯一例外 freeze_panes：视图级操作保护态下原生可用，且自带 verify_freeze_panes 独立自检链路）。

### 门禁

962/962 单测（+14 M17 回归锁：判定矩阵 8 例、豁免纪律 4 例、包装元数据与守卫先行 2 例）· Excel 真机 119/119（用例16 强化为 SHEET_PROTECTED 精确断言 + 新增 flag_cells 负向、freeze_panes 豁免、解保恢复三断言）· WPS full 146 OK/21 预期 FAIL/0 异常（新增 2 条预期降级：apply_page_setup 与 manage_excel_named_range 在 HF5 被 protect_sheet_watermark 保护后场景顺序所致，守卫把原生必然失败提前为结构化错误码；WPS m17 专项 6/6）。发版 v3.7.0.0。

## §十四 M17.1：工程化收口——预期 FAIL 基线锁死 + 发版自动化（2026-09-09）

### 背景与目标

grill 复盘（M17 收口后）确认两个系统性弱点：① WPS full 的"预期 FAIL"长期靠人工解释，新 FAIL 混入只能靠人肉对比上一轮基线（M17 那两条 SHEET_PROTECTED 就是这么发现的）；② 发版版本号靠每轮手工 sed 三处（csproj/README/iss）+ 复制改名 pack 脚本，与"单一来源+一键升版"规范矛盾。本轮把两者都固化成机器断言。

### 实现

- **WPS full 结果基线**（`tools/release/expected_wps_full.tsv` + `check_wps_full.py`）：锁定 full 模式全部 RESULT 分布（131 组 (工具,状态,错误码) 多重集，OK 146 / FAIL 21，每条带中文理由）。校验器对日志做多重集 diff：新增、缺失、次数变化、OK 总数跌破下限、TOOL_COUNT≠144 任一命中即 exit 1。**结果：21 条历史 FAIL 全部有归属（负路径 14 组 + 环境能力 5 组 + M17 场景顺序降级 2 组），此后任何新 FAIL 首轮即亮红**。基线更新走 `--update` 重生成 + 人工逐行评审。篡改测试（注入新 FAIL / 删除 OK 场景）均正确亮红。
- **版本号单一来源**：`Directory.Build.props` 成为全仓库唯一版本定义点（Version/InformationalVersion），Addin.csproj 删除本地定义改为继承（tests/tools 同步继承，无副作用）。
- **iss 模板化**：`tools/installer/JYYJ助手-Setup-wps.iss` 固定文件名、零硬编码版本；`#ifndef AppVersion` 兜底 + ISCC `/DAppVersion=x.y.z.w` 命令行注入；PkgSrc/安装说明路径用 ISPP 字符串拼接派生。历史 28 个带版本号 iss 全部作废（留 tools/installer 存档）。
- **release.py**（`tools/release/`）：`bump --ver`（props+README 正则同步，拒绝回退/非法格式）与 `pack`（双架构 --no-incremental 重建 → 便携目录组装（说明文件自上一版继承+版本替换，--from 可指定或缺省自动探测 dist 最高旧版）→ 进程内 zip → 版本戳断言（新戳≥3/旧戳=0/SHEET_PROTECTED 符号≥2）→ ISCC）。

### 验证（以真实发版 3.7.1.0 走完全链）

bump 3.7.1.0 → 全量单测 962/962（props 继承无副作用）→ pack 全链 33s：戳断言 3.7.1.0:3 / 3.7.0.0 残留 0 / 守卫符号 2 → ISCC 产出 `JYYJ助手-3.7.1.0-Setup-wps.exe`。代码零变化（仅版本资源），M17 双宿主门禁结论平移。护栏实测：非法版本号、回退版本均被拒绝；dry-run 可预演。

### M17.1 续：一键门禁 gate.py + 仓库卫生（同日追加）

- **`tools/release/gate.py`**：一条命令完成 构建→单测→Excel 真机→WPS full→基线校验（`--pack` 直通发版），fail-fast。真跑验证：单测 962 / Excel 119 / WPS full 基线全过。
- **基线机制首轮实战抓到真实漂移**：`run_python_excel` 从 `PYTHON_NOT_FOUND` 漂移为 `PYTHON_BITNESS_MISMATCH`（本机新增 64 位 Python，32 位 Excel 宿主下 fail-closed 给精确错误码+安装指引）——环境改善非回归，按 `--update`+人工修订 reason 入基线。机制有效性实证：漂移首轮即亮红。
- **ISCC 暂态锁坑**：输出 exe 被杀软/预览句柄锁定（Error 32），release.py 已补 4 次退避重试（0/5/10/20s）；本会话因 artifact 预览长锁，用 `/O` 独立输出目录验证编译通过（注意 `/O` 后不带等号，`/O=` 会把等号并进路径）。
- **仓库卫生**：dist 旧版本产物（2.19–3.6 全部目录/zip/txt/exe）+ 28 个历史 iss 移入待删目录 `_dev_tmp/_trash_20260909/`（84 项，约 256MB），保留 3.7.0.0（release.py --from 参考）与 3.7.1.0（当前版）。**仅移动未删除**，待用户确认后手动清理。

### M17.3：M16 缺陷原型并入门禁自动审计 + 恢复路径文档化（同日续，v3.7.1.3）

grill 清单收尾两项：

- **⑤ 缺陷原型自动审计**（`M16DefectPatternAuditTests`，11 条事实）：把 M13–M16 原型从"文档清单+行为回归锁"升格为**源码级门禁**——R1 CVErr 魔数只许存在于 ExcelErrorValues.cs（注释剥离后扫描）；R2 七个消费文件接入在位；R3 Workbooks.Open 借用语义（ownedMates/userOwned）；R4 CLI 白名单全路径评估分支；R5 SSRF 双防护符号；R6 markdown 管道符往返配套；R7 PathGuard 三文件在位；R8 Application/Active 级共享 RCW 禁 FinalRelease（全项目正则）；R9 SheetProtectionGuard 零释放；R10 snake_case 投影+refit_source_cells；R11 hidden_state。实现要点：测试从 AppContext.BaseDirectory 向上定位仓库根读取生产源码；审计假阳性取"宁漏报不误报"口径（R8 变量名白名单制，R1 剥 // 注释后扫描）。**此后新增/重构代码丢弃任一防护基础设施，单测门禁首轮即红。**
- **⑥ 恢复路径文档化**（AI交接文档 §4A）：SHEET_PROTECTED 五步恢复闭环（verify_protection → 不盲试 → protect_sheet_watermark action=unprotect → 复验 → 重放）；守卫豁免四类工具的"失败时不要按 SHEET_PROTECTED 归因"指引；AutoVerify 断言型/探测型两档口径；高频错误码速查表。文档引用的工具名经源码核实（unprotect 是 protect_sheet_watermark 的 action 分支，非独立工具）。

门禁：979 单测（+17 审计锁）· Excel 真机 119/119 · WPS full 基线 131 组一致 · 戳断言 3.7.1.3:3 / 残留 0 · ISCC。**grill 清单 ①–⑥ 全部闭环。**


### M17.4：M17.1–M17.3 工程化代码的三视角对抗审查与变异测试（同日续，v3.7.1.4）

按标准纪律对 M17.1–M17.3 新增代码（审计测试、release.py、check_wps_full.py、gate.py、AI交接文档 §4A）走完整闭环：

**三视角审查发现与修复（5 项）**：
1. R4 断言不完整（只查反斜杠分支）→ 补 IndexOf('/') 与 GetFullPath 断言（正斜杠路径绕过会漏检）；
2. R8 正则只列 activeWorkbook/activeSheet → 加强为 active[a-z]* 任意 active 前缀共享对象全覆盖；
3. release.py ref=None 时**静默跳过安装说明生成** → fail-fast 显式报错（发版产物缺件不可静默）；
4. check_wps_full.py TOOL_COUNT=144 硬编码 → 提为 EXPECTED_TOOL_COUNT 常量 + --expect-tools 参数（有意设置的变更卡点）；
5. AI交接文档 §4A "SHEET_NOT_FOUND（守卫前）"误导行 → 修正为"守卫不越权拦截、放行回工具自身报 INVALID_ARGS/INVALID_SHEET"。

**变异测试（审查者视角闭环，真实报错）**：向生产源码注入缺陷原型再断言审计亮红——R1 注入 CVErr 魔数、R8/R9 注入 Application 级 FinalRelease，**3 条规则全部真实被抓**（exit 1 + 违规消息命中），注入后还原逐字校验一致。审计有效性实证：不是"永远绿灯"。

**变异测试附带抓到新坑（已固化）**：沙箱裸环境缺 TMP/TEMP → MSBuild 试图写 C:\WINDOWS\MSBuildTemp 报 UnauthorizedAccessException；缺 USERPROFILE → NuGet 直接报错。两处均已补进 release.py ENV_FIX。

**AI 使用方视角核查**：§4A 引用的 verify_protection / get_excel_sheet_list / CONSERVATION_VIOLATED / expected_writable / NOT_ALLOWLISTED / BLOCKED_PRIVATE_ADDR 等符号逐一在源码核实存在。

门禁：979 单测 · Excel 真机 119/119 · WPS full 基线 131 组一致 · 戳断言 3.7.1.4:3 / 残留 0 · ISCC。发版 v3.7.1.4。

### M17.5：安装器"管理员也装不上"根因修复与真机实证（v3.7.1.5）

用户真机反馈（截图）：安装包报"自动关闭失败：已多次强制结束，仍有 Excel / WPS 进程存活"。

**根因（三层，逐层实证）**：
1. 【弹窗误导】`PrivilegesRequired=lowest` 下 Inno 永远以普通权限运行——右键"以管理员身份运行"也被降权（安装日志 `User privileges: None` 活体实证）。旧弹窗指引"右键管理员运行"完全无效。
2. 【击杀名单缺员】缺 `wpsoffice.exe`（WPS 2019+ 主控/守护进程），单杀 et.exe 可能被拉起。
3. 【真凶】遗留 EXCEL.EXE ×2（创建于 9/8 20:59/21:51，恰为 M17.4 门禁真机 harness 运行时段；父进程=Services 会话 svchost，即 DCOM 拉起；子进程 WebView2 Not Responding）——**harness 被 SIGTERM/异常中断时 KillSpawnedExcel 兜底静默失效泄漏的跨 silo 僵尸进程**。实证：普通 3 轮 + UAC 提权限强杀共 8 轮（持久化日志 = 静默跑 4 轮 + 交互跑 4 轮）全部报"没有此任务的实例在运行"，任何权限不可杀（tskill 在 Win11 已移除，Task Manager 同样无解），只有重启能清。

**修复**：
- iss [Code] 重构：`KillHostTree(elevate, log)` 普通 3 轮 → `ShellExec('runas')` UAC 提权限强杀（唯一真实提权路径）；交互模式提权前弹窗预告 UAC；击杀名单补 wpsoffice.exe；击杀日志持久化到 `{localappdata}\XaiAssistant\logs\hostkill.log`（原 {tmp} 安装即删无据可查）；失败归因弹窗改真实原因并补第 4 条（服务/自动化拉起 → 重启后安装）；/SILENT 自动强杀（批量部署 + 自动化验证）。
- gate.py 泄漏断言：真机步前快照宿主进程、步后 `assert_no_host_leak` 断言无新增泄漏，泄漏当场亮红——堵死 harness 中断静默泄漏毒害安装的通路。

**真机验证（诱饵进程法 + 实跑）**：copy cmd.exe 改名 et.exe（/C ping -t 保活）+ 真实遗留 EXCEL 同场——/SILENT 实跑：诱饵 et.exe 被击杀 ✓（同权限击杀路径实证）；持久化击杀日志与失败归因弹窗均按设计工作 ✓；提权轮无效确证遗留进程跨 silo 不可杀（用户重启后安装即通）。附带实证：沙箱内 taskkill 会被 SIGTERM 拦截，杀进程类验证必须实跑安装器本体。

门禁与发版：见 CLAUDE.md 版本行。安装器专属修复 + gate 兜底，无 C# 业务代码变更。


### M17.7 诊断助手双宿主化 + 三件套集合进安装包（2026-09-09 晨，v3.7.1.6）

**背景**：随包安装说明.txt 仍写"不支持 WPS 表格"、diagnose.ps1 只查 Excel——与 WPS 版产品现状严重脱节（用户指出）。

**诊断助手（diagnose.ps1 重构）**：
- 新增 WPS 表格节：进程检测（wps/et/wpp/wpsoffice）、et.exe 探测（与 register-auto 同源三安装根逻辑）、PE 位数、文件版本、Kingsoft\Office\ET\Addins 双视图 + AddinsWL 白名单核对、KET.Application COM 探测（回落 ET.Application）、VBA 环境诚实指引（真机实证：无可远程判定的注册键/文件信号，只指引 Alt+F11 自查，不做断言）。
- 代码优化：Get-PEArch / Probe-OfficeCom 公共函数化（Excel 与 WPS 共用）；-NoOpen 参数供自动化验证；报告头标注"双宿主版"。
- 三视角审查修复 3 项：F1 32 位 Windows 无 Registry64 视图 → 双视图循环加 OS 位数守卫（否则误报"注册不完整"）；F2 `$wpsProcs.Count -ge 0` 恒真死条件删除；F3 Office 通道 Channel 为空时尾部 `（）` → "通道未上报"。
- 清单约束：诊断的 ProgId/CLSID/Kingsoft 键位与 register-auto.ps1 逐一核对同源。

**真机验证（两轮，真实 Excel x64 16.0 + WPS x86 12.1.0.26899 双宿主）**：报告完整生成；双视图 Kingsoft Addins LoadBehavior=3+白名单在；真实暴露本机 3 问题（TaskPaneHost/DeepSeekPaneHost LoadBehavior=2、Excel 禁用列表 2 项、用户手册.html 缺失）——诊断能力实证。第一轮→3 修复→第二轮复跑确认收敛。

**安装说明 txt 小白化重写（tools/installer/安装说明.txt 为单一来源）**：双宿主覆盖、第零步改"不关也能装（安装器帮你关+UAC 指引）"、新增"装不上（自动关闭失败）"三步排查（含 hostkill.log 路径）、Excel/WPS 分流找选项卡、FAQ 增 WPS VBA（指向手册 FAQ7）、系统要求更正（+WPS 2019+、Win10/11）。release.py pack 改为模板直拷（M17.7 起），杜绝随包版与 dist 版内容分叉；模板"适用版本"戳纳入 sync_manuals。

**三件套集合进安装包**：iss 增用户手册.html（pack 从 docs 版本化手册拷为 tools/installer/用户手册.html，文件名不带版本防 bump 断链；同步进双架构便携目录）+ 开始菜单"JYYJ助手 用户手册"入口；ISCC 编译冒烟通过（注：二进制明文探针断言无效——Inno LZMA 压缩，教训记录）。

**发版自动化增量**：release.py sync_manuals（M17.6）+ 安装说明模板戳；本版起 bump→gate --pack 全自动联动（手册改名/戳/引用一次到位）。
