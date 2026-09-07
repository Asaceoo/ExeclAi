<img width="3199" height="1931" alt="13c346452c0bb71f" src="https://github.com/user-attachments/assets/926c721e-c881-46d3-a137-5e0de6b5bbdb" />
<img width="3199" height="1931" alt="0576e5c2efc96902" src="https://github.com/user-attachments/assets/d274be95-8222-4039-aded-9f61709cee93" />
<img width="3199" height="1916" alt="61a0fba95552aff3" src="https://github.com/user-attachments/assets/03c5f344-a2dd-4433-832f-932a50efe0e9" />

通过网盘分享的文件：JYYJ助手
链接: https://pan.baidu.com/s/1Ssg_czSMuOqmUK45z55SZA?pwd=1111 提取码: 1111

#JYYJ助手

Windows Excel 本机智能助手（v2.1.0.18）。基于 .NET Framework 4.8 + VSTO + WebView2 侧边栏，兼容 Excel 2016+ / Microsoft 365，提供 x64 / x86 双架构 Inno Setup 安装包。

JYYJ助手 是一个内嵌在 Excel 中的 **AI 办公助手**，通过侧边栏面板与工作表对话，理解自然语言需求并自动调用 Excel 工具完成数据读取、清洗、分析、填表、公式生成、图表绘制、报告撰写等操作。

## 已实现能力

- 读取当前工作簿、工作表、选区和保护状态
- 写入单元格和二维区域，支持值与公式两种模式
- 数据清洗：去空行、修日期、填充空白
- 条件筛选并输出匹配结果
- 区域格式化：数字格式、加粗/斜体、字体色、背景色、自动列宽
- 工作表新增、重命名、删除（删除为高危操作，默认阻断）
- 基础透视汇总与图表创建（柱形、条形、折线、饼图、面积图）
- 写入前后快照与差异回滚
- OpenAI 兼容模型配置（Base URL / API Key / 模型名）
- 原生 tool_calls + JSON 降级双协议模型集成
- 100+ 内置工具、批量宏、内置流程与技能包
- 公式解释、智能填充、标记清洗
- 一键自检与运行日志
- 跨工作簿只读引用
- WebView2 侧边栏多页 UI：聊天 / 公式 / VBA / 流程 / 设置

## 文档

| 文档 | 说明 |
|------|------|
| [用户手册](user-manual.md) | 面向最终用户：安装、功能、安全、FAQ |
| [技术手册](technical-manual.md) | 面向开发者：架构、Agent 引擎、工具层、构建与打包 |

## 安全分级

| 级别 | 行为 |
|---|---|
| ReadOnly | 直接执行 |
| NormalWrite | 默认确认 |
| StructuralChange | 必须确认 |
| Dangerous | 默认关闭，需显式启用 |

## 构建

```powershell
dotnet build ExcelAiAssistant.sln
```

## 测试

```powershell
dotnet test ExcelAiAssistant.sln
```
