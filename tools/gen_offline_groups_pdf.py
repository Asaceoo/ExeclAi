# -*- coding: utf-8 -*-
"""JYYJ助手 离线命令组调研报告 PDF 生成脚本"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Flowable, PageBreak, KeepTogether)

OUT = r"d:\execl1.0\功能区命令组调研_离线版.pdf"

# ---------- 中文字体 ----------
FONT_PATH = "C:/Windows/Fonts/msyh.ttc"
pdfmetrics.registerFont(TTFont("CJKFont", FONT_PATH, subfontIndex=0))

PRIMARY = HexColor("#1a365d")
ACCENT = HexColor("#2b6cb0")
INK = HexColor("#2d3748")
MUTED = HexColor("#718096")
TIER1 = HexColor("#1f8a5f")
TIER2 = HexColor("#b45309")
TIER3 = HexColor("#64748b")
HAVE = HexColor("#2563eb")

CONTENT_WIDTH = 6.9 * inch

# ---------- 样式 ----------
def st(name, **kw):
    base = dict(fontName="CJKFont", wordWrap="CJK")
    base.update(kw)
    return ParagraphStyle(name, **base)

S = {
    "title":   st("title", fontSize=26, leading=32, textColor=PRIMARY, alignment=TA_CENTER, spaceAfter=10),
    "subtitle": st("subtitle", fontSize=12.5, leading=17, textColor=MUTED, alignment=TA_CENTER, spaceAfter=6),
    "meta":    st("meta", fontSize=10, leading=14, textColor=MUTED, alignment=TA_CENTER),
    "h1":      st("h1", fontSize=16, leading=22, textColor=PRIMARY, spaceBefore=18, spaceAfter=8),
    "h2":      st("h2", fontSize=12.5, leading=18, textColor=ACCENT, spaceBefore=12, spaceAfter=5),
    "body":    st("body", fontSize=10.5, leading=16.5, textColor=INK, spaceAfter=6),
    "bullet":  st("bullet", fontSize=10.5, leading=16, textColor=INK, leftIndent=14, bulletIndent=4, spaceAfter=3),
    "note":    st("note", fontSize=9, leading=13, textColor=MUTED, spaceAfter=6),
    "caption": st("caption", fontSize=9, leading=12, textColor=MUTED, alignment=TA_CENTER, spaceBefore=4, spaceAfter=8),
    "cell":    st("cell", fontSize=9, leading=13, textColor=INK),
    "cellb":   st("cellb", fontSize=9.5, leading=13, textColor=colors.white),
    "cellh":   st("cellh", fontSize=10, leading=14, textColor=colors.white),
}

# ---------- 分隔线 ----------
class Divider(Flowable):
    def __init__(self, width, height, color, space_before=4, space_after=8):
        Flowable.__init__(self)
        self.width = width; self.height = height; self.color = color
        self.spaceBefore = space_before; self.spaceAfter = space_after
    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)

def divider(w, h, c, sb=4, sa=8):
    return Divider(w, h, c, sb, sa)

# ---------- 表格 ----------
def styled_table(data, widths, header_bg=ACCENT, font_size=9, align="LEFT"):
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "CJKFont"),
        ("FONTSIZE", (0, 0), (-1, 0), 9.5),
        ("FONTSIZE", (0, 1), (-1, -1), font_size),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [HexColor("#f7fafc"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#e2e8f0")),
    ]))
    return t

def pill(text, color):
    return Paragraph(f'<font color="{color.hexval()[2:] if hasattr(color,"hexval") else "1f8a5f"}"><b>{text}</b></font>', S["cell"])

def body(text):
    return Paragraph(text, S["body"])

def bullet(text):
    return Paragraph(text, S["bullet"], bulletText="•")

story = []

# ========== 封面 ==========
story.append(Spacer(1, 1.6 * inch))
story.append(Paragraph("JYYJ助手 · 功能区命令组调研报告", S["title"]))
story.append(divider(CONTENT_WIDTH * 0.5, 3, ACCENT, 4, 14))
story.append(Paragraph("离线版（非 AI）—— 纯本地高频工具命令组", S["subtitle"]))
story.append(Spacer(1, 6))
story.append(Paragraph("对标 Kutools / 方方格子 / ASAP Utilities / Ablebits 的经典加载项分组范式<br/>调研日期：2026-08-28", S["meta"]))
story.append(PageBreak())

# ========== 一、调研背景与现状 ==========
story.append(Paragraph("一、调研背景与现状", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))
story.append(body("功能区「JYYJ助手」Tab 现有 3 个命令：<b>打开面板 / 聚光灯 / 指针</b>。底层已有较完整的 AI 工具层（数据清洗、一键透视、生成图表、智能填充、VBA/Python 执行等），但全部需要联网、走侧栏对话或确认流程。"))
story.append(body("本次调研聚焦<b>离线（非 AI）命令组</b>：纯本地计算、确定性、一键即用、可撤销的批量工具。它们与 AI 工具形成互补——AI 解决「语义级」任务，离线工具解决「体力级」重复任务，且不受网络与模型可用性限制。"))

# ========== 二、调研发现 ==========
story.append(Paragraph("二、调研发现：经典加载项的分组范式", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))

story.append(Paragraph("2.1 Kutools for Excel（300+ 工具）", S["h2"]))
story.append(bullet("典型分组：<b>工作表 / 视图 / 选区与单元格 / 编辑 / 文本 / 数据 / 公式 / 合并拆分</b>"))
story.append(bullet("高频工具：合并多工作表、批量删除重复项、文本与数字格式互转、隐藏非选中区域、查找重复值、多区域复制、合并单元格保留数据"))

story.append(Paragraph("2.2 方方格子（国产，业界模板）", S["h2"]))
story.append(bullet("智能文本处理：批量提取/过滤数字·汉字·英文·符号、按位置截取、一键大小写、高级分列"))
story.append(bullet("批量录入：数值批量加减乘除、序列号快速填充、复选框批量插入（支持跨表）"))
story.append(bullet("合并/拆分：分类合并多表（关键词匹配+去重）、拆分工作簿、拆分数据"))
story.append(bullet("其他：聚光灯/指针（本加载项已有）、筛选组、分类汇总（字段求和/求唯一值/文本连接）"))
story.append(bullet("工程特征：<b>所有操作支持撤销</b>，突破 Excel 部分原生功能不可撤销的限制"))

story.append(Paragraph("2.3 ASAP Utilities（300+ 工具）", S["h2"]))
story.append(bullet("分组：<b>文本 / 数据 / 公式 / 选择 / 工作表 / 快捷方式</b>"))
story.append(bullet("特色工具：选择性粘贴组合（一步完成值+批注+格式）、公式转静态值、重置最后使用单元格（修复 Ctrl+End）、自定义快捷键（如 Ctrl+Alt+S 存盘并带时间戳备份）"))

story.append(Paragraph("2.4 Ablebits", S["h2"]))
story.append(bullet("Excel Suite：按字体颜色计算、单元格字符计数、跨多表聚合、随机数据生成、数字拼写（金额大写类）"))
story.append(bullet("Text Toolkit：文本前后追加、删除指定字符/数字/字符串、大小写转换、按指定位置插入文本"))

story.append(Paragraph("收敛结论", S["h2"]))
story.append(body("<b>离线工具的高频分组 = 文本 / 数据 / 单元格与选区 / 工作表 / 定位选择 / 视图</b>。共性特征：确定性（无概率性结果）、批量、可撤销、无需联网、秒级响应。业界 300+ 工具的产品最终都收敛到这几个分组，说明这是用户真实高频场景。"))

# ========== 三、推荐命令组 ==========
story.append(Paragraph("三、推荐命令组（离线 · 非 AI）与布局", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))

groups = [
    ("3.1 数据工具组", "Tier1 · 立即做", TIER1, [
        "查找/删除重复项（按列去重，保留首条）",
        "高级筛选增强（筛选选中区、一键清空筛选）",
        "数字 ↔ 文本互转（文本转数字、数字转文本保格式）",
        "分类汇总（按字段求和/求唯一值/文本连接）",
    ]),
    ("3.2 文本工具组", "Tier1 · 立即做", TIER1, [
        "大小写转换（全大写 / 全小写 / 首字母大写）",
        "添加文本（前后追加 / 指定位置插入）",
        "删除空格 / 删除指定字符",
        "提取数字 / 按位置截取文本",
        "智能分列（按分隔符 / 按位置自动拆分）",
        "合并多格文本（行/列方向合并）",
    ]),
    ("3.3 单元格工具组", "Tier2 · 短期做", TIER2, [
        "合并单元格并保留数据（不丢内容）",
        "批量填充（向下 / 向右 / 填充空白）",
        "选择性粘贴增强（值 + 格式组合一步完成）",
        "批量插入复选框（可选：链接到单元格）",
    ]),
    ("3.4 工作表工具组", "Tier2 · 短期做", TIER2, [
        "工作表快速导航（下拉/列表跳转，支持隐藏表）",
        "批量重命名 / 隐藏 / 取消隐藏",
        "跨表合并汇总（多工作表合并到一张，可去重）",
        "复制 / 移动工作表",
    ]),
    ("3.5 定位选择组", "Tier2 · 短期做", TIER2, [
        "定位公式 / 常量 / 空白单元格",
        "只选择可见单元格（规避筛选/隐藏行误操作）",
        "跳转到最后使用单元格",
    ]),
    ("3.6 视图工具组（扩展）", "Tier3 · 视反馈做", TIER3, [
        "显示公式（一键切换公式/结果视图）",
        "隐藏非选中区域 / 网格线切换 / 冻结切换",
    ]),
]

for title, tag, color, items in groups:
    t = styled_table(
        [[Paragraph(title, S["cellh"]), Paragraph(f"<b>{tag}</b>", S["cellh"])]] +
        [[Paragraph("<b>命令</b>", S["cellh"]), Paragraph(item, S["cell"])] for item in items],
        [1.7 * inch, 5.2 * inch],
        header_bg=color,
        font_size=9.5,
    )
    story.append(KeepTogether([t, Spacer(1, 10)]))

# 布局预览（功能区示意）
story.append(Paragraph("3.7 功能区布局示意（延续「JYYJ助手」Tab）", S["h2"]))
layout = [
    ["JYYJ助手", "数据工具", "文本工具", "单元格", "工作表", "定位选择", "视图"],
    ["打开面板(有)", "查重/去重", "大小写", "合并保留", "快速导航", "定位公式", "聚光灯(有)"],
    ["—", "高级筛选", "添加文本", "批量填充", "批量重命名", "定位空白", "指针(有)"],
    ["—", "数字互转", "删除空格", "选择性粘贴", "跨表合并", "选可见格", "显示公式*"],
    ["—", "分类汇总", "提取数字", "批量复选框", "复制移动", "跳转最后", "隐藏选区*"],
]
lt = styled_table([[Paragraph(c, S["cellh"]) for c in layout[0]]] +
                  [[Paragraph(c, S["cell"]) for c in row] for row in layout[1:]],
                  [0.95 * inch] * 7, header_bg=PRIMARY, font_size=8.5, align="CENTER")
lt.setStyle(TableStyle([
    ("ALIGN", (0, 1), (-1, -1), "CENTER"),
    ("FONTNAME", (0, 0), (-1, -1), "CJKFont"),
]))
story.append(KeepTogether([lt]))
story.append(Paragraph("* 为 Tier3 待定项；「有」= 已有命令。示意布局，非真实 Ribbon 渲染。", S["note"]))

# ========== 四、优先级与实现建议 ==========
story.append(Paragraph("四、优先级与实现建议", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))

prio = [
    ["命令组", "优先级", "复用现有", "实现要点"],
    ["数据工具组", "Tier1", "clean_excel_data / filter_excel_data", "去重/筛选/汇总复用现有工具内核，补「确认门控」；数字互转走批量写回"],
    ["文本工具组", "Tier1", "clean_excel_data（分列/修格式）", "纯本地字符处理，无需 AI；批量 + 可撤销，随改随验"],
    ["单元格工具组", "Tier2", "fill_blanks", "合并保留数据需防丢内容（关键操作加确认门控）"],
    ["工作表工具组", "Tier2", "跨表合并新工具", "多表合并注意 sheet 定位与去重，写回目标表显式指定"],
    ["定位选择组", "Tier2", "无", "纯选区操作，与 AI 工具链（选区→调用）天然兼容"],
    ["视图工具组", "Tier3", "聚光灯/指针成熟方案", "显示公式/冻结走原生；隐藏非选中区域注意恢复"],
]

def tier_mark(t):
    c = "#1f8a5f" if "Tier1" in t else ("#b45309" if "Tier2" in t else "#64748b")
    return Paragraph(f"<font color='{c}'><b>{t}</b></font>", S["cell"])

prio_rows = [[Paragraph(c, S["cellh"]) for c in prio[0]]]
for r in prio[1:]:
    prio_rows.append([
        Paragraph(r[0], S["cell"]),
        tier_mark(r[1]),
        Paragraph(r[2], S["cell"]),
        Paragraph(r[3], S["cell"]),
    ])
pt = styled_table(prio_rows, [1.15 * inch, 0.75 * inch, 1.9 * inch, 3.1 * inch],
                  header_bg=PRIMARY, font_size=9)
story.append(pt)

# ========== 五、与 AI 工具的分工原则 ==========
story.append(Paragraph("五、与 AI 工具的分工原则", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))
story.append(bullet("<b>离线工具</b>：确定性、即时、批量、可撤销 → 直接执行 + verify 自检闭环，无需联网与确认（破坏性操作除外）"))
story.append(bullet("<b>AI 工具</b>：智能、需理解语义 → 保留确认门控，处理模糊/语义级任务"))
story.append(bullet("<b>同一数据清洗分工</b>：离线做「格式化清洗」（类型/空格/大小写/重复），AI 做「语义清洗」（补全、归类、改写）"))
story.append(bullet("<b>优先级判断</b>：离线命令价值 = 使用频率 × 单次节省时间 × 确定性收益。文本与数据两类组是 300+ 工具产品的共同底座，应优先做"))

# ========== 六、落地建议 ==========
story.append(Paragraph("六、落地建议", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))
story.append(body("<b>第一批（Tier1）</b>：数据工具组 + 文本工具组。改动最小、价值最高——大量复用现有 clean/filter 工具内核，只需补「批量 + 可撤销 + 自检」通道。"))
story.append(body("<b>第二批（Tier2）</b>：单元格工具组 + 工作表工具组 + 定位选择组。关键操作（合并保留、去重删除）务必加确认门控，防数据丢失。"))
story.append(body("<b>分组位置</b>：延续现有「JYYJ助手」Tab，按「数据→文本→单元格→工作表→定位→视图」排列，与 AI 组（写公式/清洗/分析）并列，用户按任务类型即点即用。"))
story.append(body("<b>与自检体系打通</b>：离线命令同样落入「主步骤 + verify 自检」闭环，verify 未配置 Tool 必须 fail-closed 阻断，保持清单约束兜底。"))

# ========== 七、主要来源 ==========
story.append(Paragraph("七、主要来源", S["h1"]))
story.append(divider(CONTENT_WIDTH * 0.3, 2, ACCENT, 0, 10))
srcs = [
    "Kutools for Excel（300+ 高级功能）· https://zh-cn.extendoffice.com/product/kutools-for-excel.html",
    "Kutools 中文官网 · https://www.kutoolsforexcel.com/",
    "方方格子 Excel 工具箱帮助动画 · http://www.ffcell.com/home/ffcell_doc.aspx/vip.aspx",
    "方方格子 更新日志（筛选组/分类汇总/聚光灯）· http://www.ffcell.com/home/ffcell_updateinfo.aspx",
    "方方格子 分类合并多表 · http://www.ffcell.com/WenKu/article.aspx?id=616",
    "ASAP Utilities 工具说明（300+ 工具，中文）· https://www.asap-utilities.com/asap-utilities-excel-tools-description.php",
    "Ablebits Excel Suite · https://www.ablebits.com/excel-suite/index.php",
    "Ablebits Text Toolkit · https://www.ablebits.com/excel-text-toolkit/",
]
for s in srcs:
    story.append(bullet(s))

story.append(Spacer(1, 12))
story.append(divider(CONTENT_WIDTH, 1, HexColor("#e2e8f0"), 6, 8))
story.append(Paragraph("调研基于 2026-08-28 联网检索；优先级结合本加载项现有工具层（已具备能力优先复用）。", S["note"]))

doc = SimpleDocTemplate(OUT, pagesize=A4,
                        leftMargin=0.8 * inch, rightMargin=0.8 * inch,
                        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
                        title="JYYJ助手 离线命令组调研报告",
                        author="JYYJ助手")
doc.build(story)
print("OK", OUT, os.path.getsize(OUT))
