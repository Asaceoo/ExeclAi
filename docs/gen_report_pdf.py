# -*- coding: utf-8 -*-
"""将《竞品借鉴调研：JYYJ助手 × 方方格子》HTML 报告渲染为专业排版 PDF。"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, KeepTogether, Flowable, PageBreak)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------- CJK 字体 ----------
FONT = "YaHei"
fp = "C:/Windows/Fonts/msyh.ttc"
if not os.path.exists(fp):
    raise FileNotFoundError("未找到微软雅黑字体: " + fp)
pdfmetrics.registerFont(TTFont(FONT, fp))

# ---------- 配色 ----------
BRAND = HexColor("#2563eb")
BRAND2 = HexColor("#4f46e5")
ACCENT = HexColor("#2b6cb0")
INK = HexColor("#1f2733")
SUB = HexColor("#5a6472")
LINE = HexColor("#e3e8ef")
HEADBG = HexColor("#f0f4fb")
ROWBG = HexColor("#f7fafc")
GREEN = HexColor("#059669")
RED = HexColor("#dc2626")
AMBER = HexColor("#d97706")

# ---------- 样式 ----------
S = {
    "title": ParagraphStyle("t", fontName=FONT, fontSize=22, leading=29,
                            textColor=BRAND2, alignment=TA_CENTER, spaceAfter=6, wordWrap="CJK"),
    "subtitle": ParagraphStyle("st", fontName=FONT, fontSize=11.5, leading=17,
                               textColor=SUB, alignment=TA_CENTER, spaceAfter=4, wordWrap="CJK"),
    "meta": ParagraphStyle("m", fontName=FONT, fontSize=9, leading=13,
                           textColor=SUB, alignment=TA_CENTER, spaceAfter=2, wordWrap="CJK"),
    "h1": ParagraphStyle("h1", fontName=FONT, fontSize=15, leading=20,
                         textColor=BRAND, spaceBefore=16, spaceAfter=8, wordWrap="CJK"),
    "h2": ParagraphStyle("h2", fontName=FONT, fontSize=12, leading=17,
                         textColor=ACCENT, spaceBefore=10, spaceAfter=5, wordWrap="CJK"),
    "body": ParagraphStyle("b", fontName=FONT, fontSize=10.5, leading=16.5,
                           textColor=INK, spaceAfter=7, wordWrap="CJK"),
    "bullet": ParagraphStyle("bl", parent=None, fontName=FONT, fontSize=10.5, leading=16,
                             textColor=INK, leftIndent=14, bulletIndent=4, spaceAfter=5, wordWrap="CJK"),
    "caption": ParagraphStyle("c", fontName=FONT, fontSize=9, leading=13,
                              textColor=SUB, alignment=TA_CENTER, spaceBefore=4, spaceAfter=12, wordWrap="CJK"),
    "src": ParagraphStyle("src", fontName=FONT, fontSize=9.5, leading=15,
                          textColor=SUB, spaceAfter=3, leftIndent=0, wordWrap="CJK"),
}
# 用 Paragraph 子类渲染小标签
def _tag(text, color):
    return '<font color="#ffffff" size="8.5"><b> %s </b></font>' % text


# 彩色标签段落（简化为带“”的加粗色块文本）
def prio_lbl(text, color):
    return Paragraph(
        '<font color="%s"><b>■ %s</b></font>' % (color.hexval()[2:] and color.hexval(), text),
        S["h2"])


class Divider(Flowable):
    def __init__(self, width, height=2, color=BRAND, space_before=4, space_after=8):
        super().__init__()
        self.width, self.height, self.color = width, height, color
        self.spaceBefore, self.spaceAfter = space_before, space_after

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.rect(0, 0, self.width, self.height, fill=1, stroke=0)


def styled_table(data, widths=None, header_bg=BRAND, repeat=False):
    t = Table(data, colWidths=widths, repeatRows=1 if repeat else 0)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, 0), 9.5),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [ROWBG, colors.white]),
    ]))
    return t


doc = SimpleDocTemplate(
    "d:/execl1.0/docs/竞品借鉴-方方格子-调研报告.pdf",
    pagesize=A4,
    leftMargin=0.8 * inch, rightMargin=0.8 * inch,
    topMargin=0.7 * inch, bottomMargin=0.7 * inch)
story = []

# ===== 封面页 =====
story.append(Spacer(1, 1.5 * inch))
story.append(Paragraph("竞品借鉴调研", ParagraphStyle("tt", fontName=FONT, fontSize=30, leading=38,
                                                        textColor=BRAND2, alignment=TA_CENTER, wordWrap="CJK")))
story.append(Paragraph("JYYJ助手 × 方方格子（FFCell）", S["title"]))
story.append(Spacer(1, 0.15 * inch))
story.append(Divider(3.2 * inch, height=3, color=BRAND))
story.append(Spacer(1, 0.3 * inch))
story.append(Paragraph("从产品定位、能力重叠、交互形态三个维度拆解本方 Excel AI 加载项"
                       "（JYYJ助手）与国内头部效率工具箱「方方格子」的差异，"
                       "给出按优先级排序的可借鉴点与落地路线——同时明确哪些地方应当坚持 AI 原生优势、"
                       "不必向工具箱「降级」。", S["subtitle"]))
story.append(Spacer(1, 0.5 * inch))
story.append(Paragraph("调研日期：2026-08-28", S["meta"]))
story.append(Paragraph("来源：本地源码盘点 + 方方格子官网 / 更新日志 / 第三方评测", S["meta"]))
story.append(Paragraph("定位：内部决策参考", S["meta"]))
story.append(PageBreak())

# ===== 一、定位差异 =====
story.append(Paragraph("一、两家的定位差异：AI 原生 Agent vs 效率工具箱", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))
pos = [
    ["维度", "JYYJ助手（本项目）", "方方格子 FFCell"],
    ["产品形态", "C#/.NET COM 加载项 + WebView2 侧边栏 + 常驻 Python worker，AI 原生",
     "VSTO COM 加载项，纯本地效率工具箱（已逐步引入 AI）"],
    ["核心范式", "自然语言驱动：Agent 调用 200+ 工具、workflow 编排、技能（SKILL.md）",
     "点即用：把日常高频操作封装为 Ribbon 一键按钮与自定义函数"],
    ["能力入口", "侧边栏「聊天 / 工具 / 历史 / 技能 / 设置」为中心，能力绑在 AI 对话层",
     "功能区 Tab 分组 + 53 个 FF.* 工作表函数 + 宏收纳箱 / DIY 沉淀"],
    ["安全与治理", "风险分级 L0-L3、确认门、受控执行、撤销快照回滚、DPAPI",
     "常规 Office 操作，撤销依赖 Excel 本身"],
    ["程序化执行", "VBA + Python(xlwings) 双引擎，支持自定义代码沉淀", "仅 VBA 宏收纳箱 / 窗体"],
    ["商业化", "未涉及授权分层（内部 / 自用）", "免费基础 + 会员分层（月 8 元 / 年 45 元 / 终身 268 元）"],
]
story.append(styled_table(pos, widths=[0.95 * inch, 2.9 * inch, 2.9 * inch]))
story.append(Paragraph("结论：这是「AI 原生」与「效率工具」两种基因。真正可借鉴的不是再建一批批量工具，"
                       "而是把本项目已有 AI 能力的入口与交互打磨得像方方格子一样零门槛，"
                       "并补上它独有的几个高频交互。", S["body"]))

# ===== 二、能力重叠盘点 =====
story.append(Paragraph("二、能力重叠盘点：工具库其实已经「你有我也有」", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))
story.append(Paragraph("本项目工具注册中已覆盖方方格子绝大多数批量功能，差距主要体现在交互与形态，"
                       "而非「缺工具」。", S["body"]))
overlap = [
    ["方方格子模块", "本项目对应能力", "覆盖度"],
    ["文本提取 / 过滤 / 替换 / 拆分 / 删除",
     "<font size='8'>extract_text_characters、text_clean/split/merge、delete_text_matches、insert_text_into_cell、find_replace_range</font>",
     "已重叠"],
    ["数据清洗 / 排序 / 去重 / 筛选",
     "<font size='8'>clean_excel_data、sort_range、dedup_range、filter_excel_data、manage_excel_filter</font>",
     "已重叠"],
    ["高级排序 / 颜色排序 / 合并单元格排序",
     "<font size='8'>sort_range（含稳定性 tiebreaker）</font>", "大部分重叠"],
    ["汇总拆分 / 汇总大师 / 分类汇总",
     "<font size='8'>create_excel_pivot_summary、split_sheet_by_field、create_pivot_table、merge_sheets</font>",
     "已重叠且更强"],
    ["重复值工具 / 数据对比 / 两列匹配",
     "<font size='8'>fuzzy_match_columns、compare_two_sheets、cross_reference_columns</font>", "已重叠"],
    ["合并转换 / 拆分单元格 / 空单元格合并",
     "<font size='8'>merge_or_split_cells、merge_duplicate_cells、merge_or_split_cells</font>", "已重叠"],
    ["批量录入 / 序列 / 随机数",
     "<font size='8'>fill_number_sequence、fill_random_numbers、generate_permutation_list、smart_paste_delimiter</font>",
     "已重叠"],
    ["公式向导 / 公式转换 / 引用切换",
     "<font size='8'>explain_excel_formula、transform_excel_formulas、convert_formula_references、trace_formula_references、wrap_formulas_with_iferror</font>",
     "已重叠且更强"],
    ["宏收纳箱 / 自定义宏",
     "<font size='8'>save_vba_tool 沉淀 + CustomToolStore + VbaPresetCatalog</font>",
     "已有底层，缺可视化 UI"],
    ["身份证校验 / 拼音笔画 / 人民币大写 / 凑数",
     "<font size='8'>convert_text_to_pinyin、amount_to_uppercase、pick_numbers_to_sum、bankers_round_values</font>",
     "已重叠"],
    ["气泡式 / 一键取色 / 聚光灯 / 指针", "<font size='8'>—（无交互增强类）</font>", "完全缺失"],
    ["FF.* 工作表函数 / UDF 库",
     "<font size='8'>—（所有能力绑在 Agent 对话层，无独立函数注册）</font>", "形态缺失"],
]
story.append(styled_table(overlap, widths=[1.55 * inch, 3.55 * inch, 1.65 * inch]))
story.append(Paragraph("「完全缺失 / 形态缺失」两行，即本章的重点借鉴对象。", S["caption"]))

# ===== 三、值得借鉴点 =====
story.append(Paragraph("三、值得借鉴的重点（按优先级）", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))

def bullet(text):
    return Paragraph(text, S["bullet"], bulletText="•")

# P0-1
story.append(Paragraph(_tag("P0", "dc2626") + " 聚光灯 / 指针 / 阅读模式等「视图定位」交互增强",
                       ParagraphStyle("p0", fontName=FONT, fontSize=12.5, leading=18,
                                      textColor=RED, spaceBefore=8, spaceAfter=4, wordWrap="CJK")))
story.append(Paragraph("方方格子的高频杀手功能，帮用户在大量数据中快速定位目标。本项目为零。体量小、价值直观。", S["body"]))
for b in [
    "聚光灯：COM 事件驱动，高亮活动单元格所在整行 + 整列；支持自定义颜色、透明度、可选仅行 / 仅列；操作可撤销。与现有「事件监听 / 写入」基础设施贴合。",
    "指针工具：两个指向针标记重点关注单元格，方便演示 / 评审，成本极低。",
    "建议做成 Ribbon Toggle + 右键开关，天然可复用现有安全与命名管理设施。",
]:
    story.append(bullet(b))

# P0-2
story.append(Paragraph(_tag("P0", "dc2626") + " 工作表 UDF 函数库：把既有能力暴露为可复用的 FF.* 式函数",
                       ParagraphStyle("p0b", fontName=FONT, fontSize=12.5, leading=18,
                                      textColor=RED, spaceBefore=12, spaceAfter=4, wordWrap="CJK")))
story.append(Paragraph("方方格子注册 53 个 FF.* 自定义函数（FF.TextMatch、FF.RoundBank、FF.SumColX、FF.AIAsk…），"
                       "用户在单元格公式里直接调用。这是与「AI 对话式」能力互补的另一种形态。", S["body"]))
for b in [
    "面向纯 Excel 用户：不想「开侧边栏问 AI」，只想在公式里用 JYYJ.TextExtract(...)、JYYJ.RoundBank(...)。",
    "本项目的文本提取、金额大写、银行家舍入、编号补零、相似度匹配等纯函数型工具可批量沉淀为 UDF。",
    "已有 CustomToolStore 可扩展出「纯函数工具」类型，经 Ribbon / 收纳箱一键生成命名空间函数。",
    "注意：UDF 需满足确定性 + 无副作用约束，与受控执行 / 安全分级机制共存，需在技术方案中明确边界。",
]:
    story.append(bullet(b))

# P1-1
story.append(Paragraph(_tag("P1", "d97706") + " 高频一键工具的「免对话直达」入口",
                       ParagraphStyle("p1", fontName=FONT, fontSize=12.5, leading=18,
                                      textColor=AMBER, spaceBefore=12, spaceAfter=4, wordWrap="CJK")))
story.append(Paragraph("方方格子是「点按钮即执行」，本项目是「问 AI 再执行」。对高频零风险操作"
                       "（序号填充、删除空行、合并转换、格式清理等），免对话直达能显著降低使用门槛。", S["body"]))
for b in [
    "Ribbon 增加按场景分组的快捷按钮（文本 / 数据 / 公式 / 视图），一键调用对应工具（低风险直接执行，高风险仍走确认门）。",
    "右键菜单增强：把「快速清洗 / 提取数字 / 去空格 / 横向求和」等挂进选中区域右键菜单。",
    "面板增加「常用工具」缩略抽屉，别让高频能力深埋在对话里。",
]:
    story.append(bullet(b))

# P1-2
story.append(Paragraph(_tag("P1", "d97706") + " 宏收纳箱可视化 UI + 场景化导航 + 帮助动画",
                       ParagraphStyle("p1b", fontName=FONT, fontSize=12.5, leading=18,
                                      textColor=AMBER, spaceBefore=12, spaceAfter=4, wordWrap="CJK")))
story.append(Paragraph("本项目已有 CustomToolStore / save_vba_tool / VbaPresetCatalog 的底层，"
                       "缺的是方方格子「宏收纳箱 / DIY」那层用户可见的、可管理的「个人武器库」。", S["body"]))
for b in [
    "可视化收纳箱：按文本 / 数据 / 公式 / 视图场景分类浏览、收藏、增删已沉淀的宏与函数，替代目前只依赖 custom-tools 目录。",
    "场景化 UI 导航而非纯函数名组织，贴合用户操作习惯。",
    "每个工具配一句中文说明 + 示例动图（GIF），复用现有诊断 / 日志能力做「我做了什么」回显。",
]:
    story.append(bullet(b))

# P2
story.append(Paragraph(_tag("P2", "64748b") + " 垂直与生态功能（按需取舍）",
                       ParagraphStyle("p2", fontName=FONT, fontSize=12.5, leading=18,
                                      textColor=SUB, spaceBefore=12, spaceAfter=4, wordWrap="CJK")))
for b in [
    "字段排序：整理 ERP 导出报表字段、保存常用字段顺序 —— 与现有 add/insert/delete dimension 可组合，成本低。",
    "SQL 查询：对区域 / 表做类 SQL 查询。本项目有常驻 Python worker，可直接用内存数据解决方案，无需重写。",
    "图片 / 二维码专题：批量导入导出、二维码批生成 —— 独立赛道，除非有需求否则优先级低。",
    "授权分层（免费 / 会员）：目前为内部 / 自用，若无分发诉求暂不引入；若要对外，可借鉴「基础免费 + 高级工具分层」降低门槛的漏斗。",
]:
    story.append(bullet(b))

# ===== 四、反观 =====
story.append(Paragraph("四、反观：本项目已领先，不必向工具箱「降级」", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))
adv = [
    ["优势维度", "说明"],
    ["AI 原生理解与编排", "自然语言 → 多工具编排、workflow 流程、SKILL.md 技能体系，方方格子仅「AI 公式 / AI 工具箱」较浅，无法对标。"],
    ["安全治理", "风险分级 L0-L3、执行确认门、受控执行白名单、DPAPI 密钥加密、数据采样 / 额外发送提示 —— 工具箱普遍缺失。"],
    ["撤销与审计", "快照回滚、操作记录、会话 SQLite、Serilog 多文件日志 + DiagnosticsLog —— 工具箱只靠 Excel 原生撤销。"],
    ["VBA + Python 双引擎", "常驻 Python worker + xlwings，能跑复杂脚本 / 图表 / 跨表，工具箱只有宏。"],
    ["性能与稳健性积累", "项目记忆沉淀了大量 COM 僵尸进程清理、TLS、跨表图表、非 A1 原点偏移等实战约束，是独有护城河。"],
]
story.append(styled_table(adv, widths=[1.5 * inch, 5.25 * inch]))
story.append(Paragraph("设计基调：借鉴方方格子的是「入口与交互的零门槛」和「视图 / UDF / 收纳箱」这些形态，"
                       "而不是丢掉 AI Agent、安全治理或双引擎优势去堆砌普通批量按钮。"
                       "目标是「AI 能力 + 工具箱般的可达性」。", S["body"]))

# ===== 五、落地路线 =====
story.append(Paragraph("五、分阶段落地路线建议", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))
road = [
    ["阶段", "动作", "贴合现有架构点"],
    ["Phase 1", "聚光灯 + 指针交互增强；UDF 函数库（先沉淀 8-10 个纯函数型工具为 JYYJ.* 函数）",
     "Ribbon Toggle、COM 事件、CustomToolStore 扩展「纯函数工具」类型"],
    ["Phase 2", "高频工具免对话直达：Ribbon 场景分组 + 右键菜单增强 + 面板常用抽屉",
     "现有工具 Registry 复用 Schema 定义，直接映射到按钮"],
    ["Phase 3", "宏收纳箱可视化 UI + 场景化导航 + 工具帮助动画",
     "CustomToolStore 已有读写，补一层管理界面与示例库"],
    ["Phase 4", "字段排序、SQL 查询（走 Python worker）、可选授权分层",
     "PythonWorkerHost 已在；商业化按分发诉求再评估"],
]
story.append(styled_table(road, widths=[0.85 * inch, 3.1 * inch, 2.8 * inch]))

# ===== 来源 =====
story.append(Paragraph("附：主要信息来源", S["h1"]))
story.append(Divider(2.6 * inch, height=2, color=BRAND))
for src in [
    "方方格子官网-产品页：http://www.ffcell.com/home/ffcell.aspx",
    "方方格子官网-更新日志（V6.0.0.0）：http://www.ffcell.com/home/ffcell_updateinfo.aspx",
    "方方格子官网-会员 / 价格：http://www.ffcell.com/home/vip.aspx",
    "方方格子官网- Excel 扩展函数（53 个）：http://www.ffcell.com/WenKu/article.aspx?id=770",
    "方方格子官网- AI 公式 AI 拆分：http://www.ffcell.com/WenKu/article.aspx?id=789",
    "方方格子官网- 聚光灯：http://login.ffcell.com/WenKu/article.aspx?id=1003",
    "本项目能力清单：d:\\execl1.0 源码盘点（Ribbon.xml、各 *Registration.cs、CustomToolStore.cs、PythonWorkerHost.cs 等）",
]:
    story.append(Paragraph(src, S["src"]))

doc.build(story)
print("OK ->", "d:/execl1.0/docs/竞品借鉴-方方格子-调研报告.pdf")