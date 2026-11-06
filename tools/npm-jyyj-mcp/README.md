# jyyj-mcp

JYYJ助手（Excel/WPS AI 加载项）的 MCP server 分发外壳。npm 只是「找得到/拉得下/配得好」三层壳，
运行时是原生 `ExcelAiMcpServer.exe`（.NET Framework 4.8，需本机装有 Excel 或 WPS ET）。

## 快速接入（干净机器）

```bash
# 1. 打印配置（本机没 exe 会自动从 GitHub Release 拉 mcp-server 便携包；确认门默认 stderr 带外硬门）
npx jyyj-mcp config

# 2. Claude Desktop：把输出片段粘进 claude_desktop_config.json 的 mcpServers 里
#    DSH / 其它 MCP 客户端：粘贴同一段 JSON 到 MCP 连接器
```

## 子命令

| 命令 | 作用 |
|---|---|
| `jyyj-mcp config` | 打印 MCP 客户端配置片段：**默认只读**（不暴露写工具）；需要写能力加 `--allow-write`（输出带 `--allow-write` 的配置） |
| `jyyj-mcp locate` | 定位本机 ExcelAiMcpServer.exe（安装目录 / ~/.jyyj-mcp 缓存 / dist） |
| `jyyj-mcp install` | 从 GitHub Release v3.9.7.0 下载 **mcp-server-x64.zip** 并解压到 ~/.jyyj-mcp/<版本>/（SHA256 清单校验，缺失即拒绝） |
| `jyyj-mcp run` | 直接 stdio 启动 server（**默认只读**，`--allow-write` 开放写工具；客户端 command 也可指 `npx -y jyyj-mcp run`） |
| `jyyj-mcp verify` | 手动校验下载产物：`--manifest SHA256SUMS-x.y.z.w.txt --file <文件>` |

## 下载链完整性（M28）

- `install` 必须在解压前通过 **SHA256SUMS-<版本>.txt** 校验（大小 + SHA256）；清单缺失即拒绝自动安装（fail-closed）
- 代理走 `HTTPS_PROXY` / `HTTP_PROXY` 的 CONNECT 隧道；单请求 60s 超时；失败/中断自动删除 `.part` 半途文件
- 解压只接受白名单资产名（`[A-Za-z0-9._-]+`），解压后仍校验 `ExcelAiMcpServer.exe` 存在

## 运行时要求

- Windows + 本机装有 **Excel 2013+** 或 **WPS ET**（COM 激活需要）
- .NET Framework 4.8（Win10 1903+ 自带）
- MCP 工具面：190 个（169 代码注册 + 智能体桥接 + 动态沉淀），写操作全量审计、可 `undo_excel_operation` 回滚，Dangerous 级走一次性确认码（默认 file 模式为**软门**：开放文件读取的宿主仍可读码；`--confirm-mode stderr` 为**带外硬门**，码只进 server stderr）

免责声明见主仓库 README。
