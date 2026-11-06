#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 从 ErrorCatalog.cs 生成 docs/错误码目录.md（M36-3；勿手改产物）
import os
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(ROOT)

import io, re
src = io.open(r"src\ExcelAiAssistant.Core\Errors\ErrorCatalog.cs", encoding="utf-8").read()
rows = re.findall(r'\["([A-Z0-9_]+)"\] = new ErrorEntry\("([A-Z0-9_]+)", "(.*?)", "(.*?)"\),', src)
md = ["# 错误码目录（M36-3）", "", "> 由 `src/ExcelAiAssistant.Core/Errors/ErrorCatalog.cs` 生成，勿手改。MCP / CLI / 侧边栏共用同一份解释。", "",
      "| 错误码 | 含义 | 下一步动作 |", "|---|---|---|"]
for code, _, msg, act in rows:
    md.append("| `%s` | %s | %s |" % (code, msg, act))
md.append("")
md.append("未登记的码不影响原有行为：调用方跳过富化，仍返回原始 `error_code`。")
io.open(r"docs\错误码目录.md","w",encoding="utf-8",newline="\n").write("\n".join(md)+"\n")
print("entries:", len(rows))
