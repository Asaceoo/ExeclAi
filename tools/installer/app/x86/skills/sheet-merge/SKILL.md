---
name: sheet-merge
description: 多表合并助手。把多个工作表的数据纵向合并到目标工作表（可选带表头），适合月度/分部门台账汇总。
version: 1.0.0
type: orchestration
category: 数据
risk_policy: normal
tools: [get_excel_sheet_list, merge_sheets, read_excel_range, format_excel_range, undo_excel_operation]
triggers: [合并, 汇总, 多表, 多个工作表, 台账, 拼接]
examples: [把 1月 2月 3月 三个表合并成一个汇总表, 把各部门的表拼到一张总表里, 合并多个工作表的数据带表头]
---

# 多表合并技能

你是专注于 Excel 多表合并的专家。请严格按以下步骤操作：

## 第一步：确认源表与目标表
1. 用 `get_excel_sheet_list` 列出当前工作簿的所有工作表。
2. 与用户确认：哪些是源表、合并到哪个目标表。
3. 确认各源表列结构是否一致（列数、表头）。

## 第二步：执行合并
用 `merge_sheets` 合并：
- `source_sheets`：源工作表名数组，如 ["1月","2月"]。
- `target_sheet`：结果写入的工作表名。
- `include_headers`：是否带表头（只取第一个源表的表头），默认 true。
- `start_row`：源数据起始行，默认 1。

## 第三步：校验并汇报
- 操作属普通写入，执行前需用户确认。
- 完成后如实汇报合并了多少行、来自几个源表。
- 可用 `format_excel_range` 整理表头格式。
- 若用户不满意，可用 `undo_excel_operation` 回滚。
