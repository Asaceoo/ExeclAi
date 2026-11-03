# jyyj-mcp

JYYJ助手（Excel/WPS AI 加载项）的 MCP server 分发外壳。npm 只是「找得到/拉得下/配得好」三层壳，
运行时是原生 `ExcelAiMcpServer.exe`（.NET Framework 4.8，需本机装有 Excel 或 WPS ET）。

## 快速接入（干净机器）

```bash
# 1. 打印配置（本机没 exe 会自动从 GitHub Release 拉单文件便携版）
npx jyyj-mcp config

# 2. Claude Desktop：把输出片段粘进 claude_desktop_config.json 的 mcpServers 里
#    DSH / 其它 MCP 客户端：粘贴同一段 JSON 到 MCP 连接器
```

## 子命令

| 命令 | 作用 |
|---|---|
| `jyyj-mcp config` | 打印 MCP 客户端配置片段（默认）；`--ro` 只读模式（不暴露写工具） |
| `jyyj-mcp locate` | 定位本机 ExcelAiMcpServer.exe（安装目录 / ~/.jyyj-mcp 缓存 / dist） |
| `jyyj-mcp install` | 从 GitHub Release v3.8.5.0 下载单文件便携 zip 并解压到 ~/.jyyj-mcp/<版本>/ |
| `jyyj-mcp run` | 直接 stdio 启动 server（客户端 command 也可指 `npx -y jyyj-mcp run`） |

## 运行时要求

- Windows + 本机装有 **Excel 2013+** 或 **WPS ET**（COM 激活需要）
- .NET Framework 4.8（Win10 1903+ 自带）
- MCP 工具面：188 个（167 代码注册 + 智能体桥接 + 动态沉淀），写操作全量审计、可 `undo_excel_operation` 回滚，Dangerous 级走一次性确认码（码只落文件，AI 无法自答）

免责声明见主仓库 README。
