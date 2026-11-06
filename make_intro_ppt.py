# -*- coding: utf-8 -*-
"""面试自我介绍 PPT（幸运咖 市场分析岗）"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

RED = RGBColor(0xD6, 0x40, 0x27)
DARK = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x77, 0x77, 0x77)
LIGHT = RGBColor(0xF7, 0xF3, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xE8, 0xA3, 0x3D)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
HERE = os.path.dirname(os.path.abspath(__file__))


def add_slide():
    return prs.slides.add_slide(BLANK)


def box(slide, l, t, w, h, fill=None, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1)
    return sp


def text(slide, l, t, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=6):
    """runs: list of paragraphs; each paragraph = list of (text, size, bold, color)"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        for (t, size, bold, color) in para:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = "微软雅黑"
    return tb


def header(slide, title, subtitle=None):
    box(slide, 0, 0, prs.slide_width, Inches(1.05), fill=WHITE)
    box(slide, 0, Inches(1.02), prs.slide_width, Pt(2.5), fill=RED)
    text(slide, Inches(0.6), Inches(0.18), Inches(10), Inches(0.7),
         [[(title, 26, True, DARK)]])
    if subtitle:
        text(slide, Inches(0.62), Inches(0.66), Inches(11), Inches(0.4),
             [[(subtitle, 12, False, GRAY)]])


def footer(slide, n):
    text(slide, Inches(12.3), Inches(7.05), Inches(0.9), Inches(0.35),
         [[(f"{n:02d}", 11, False, GRAY)]], align=PP_ALIGN.RIGHT)
    text(slide, Inches(0.6), Inches(7.05), Inches(6), Inches(0.35),
         [[("应聘：幸运咖 · 市场分析（郑州）", 10, False, GRAY)]])


# ===== S1 封面 =====
s = add_slide()
box(s, 0, 0, prs.slide_width, prs.slide_height, fill=RED)
box(s, Inches(0.9), Inches(2.0), Inches(0.12), Inches(2.6), fill=WHITE)
text(s, Inches(1.3), Inches(2.0), Inches(11), Inches(1.2),
     [[("自我介绍", 44, True, WHITE)]])
text(s, Inches(1.32), Inches(3.15), Inches(11), Inches(0.8),
     [[("应聘岗位：市场分析　|　幸运咖（蜜雪集团）　|　郑州", 20, False, WHITE)]])
text(s, Inches(1.32), Inches(4.15), Inches(11), Inches(1.2),
     [[("【姓名】　·　统计学本科　·　【电话】　·　【邮箱】", 14, False, WHITE)],
      [("随附作品：《平价现制饮品市场分析报告》（12 页，数据均标注来源）", 13, False, WHITE)]])

# ===== S2 我是谁 =====
s = add_slide()
header(s, "一页看懂我", "统计学出身的数据分析实践者")
cards = [
    ("科班方法论", "统计学本科：假设检验、回归分析、\n抽样调查的完整训练", RED),
    ("全链路动手", "Python / SQL / Excel /\n可视化出图 / 报告撰写", GOLD),
    ("产品化思维", "独立做过 2 个 AI 教育产品，\n习惯从业务需求倒推数据需求", RGBColor(0x3E, 0x7C, 0xB1)),
    ("行业热情", "长期跟踪现制饮品行业，\n独立完成本份市场分析报告", RGBColor(0x59, 0x8C, 0x5A)),
]
for i, (t, d, c) in enumerate(cards):
    l = Inches(0.6 + i * 3.12)
    box(s, l, Inches(1.6), Inches(2.9), Inches(2.5), fill=LIGHT)
    box(s, l, Inches(1.6), Inches(2.9), Inches(0.12), fill=c)
    text(s, l + Inches(0.2), Inches(1.95), Inches(2.5), Inches(0.5), [[(t, 17, True, DARK)]])
    text(s, l + Inches(0.2), Inches(2.55), Inches(2.55), Inches(1.4),
         [[(ln, 12, False, GRAY)] for ln in d.split("\n")], space_after=4)
text(s, Inches(0.6), Inches(4.5), Inches(12), Inches(2.2),
     [[("一句话版本：", 15, True, DARK),
       ("我能把『一堆公开数据』变成『一份老板能直接用的结论』——随附的 12 页市场分析报告就是证明。", 15, False, DARK)]],
     space_after=10)
footer(s, 2)

# ===== S3 技能 vs JD =====
s = add_slide()
header(s, "岗位要求 vs 我的匹配", "逐条对照 JD 任职要求")
rows = [
    ("JD 任职要求", "我的对应证据"),
    ("较强的数据敏感度与逻辑分析能力，\n能独立完成数据分析任务", "统计学本科 + 独立完成《平价现制饮品市场分析报告》：\n10+ 数据源、统一口径、5 维度分析、3 条结论"),
    ("熟练使用数据分析工具，良好的报告\n撰写与信息可视化能力", "Python(pandas/matplotlib)、SQL、Excel 透视表；\n报告含 5 张自绘图表，结论先行、图表佐证"),
    ("善于沟通协作，高效响应跨部门需求", "做产品项目时习惯于把结论翻译成业务语言：\n一句话结论 + 一张图 + 一条建议"),
    ("细致踏实、抗压、执行力和责任心", "报告中每个数字标注来源与统计时点，\n交叉验证后才写入结论"),
]
tbl = s.shapes.add_table(len(rows), 2, Inches(0.6), Inches(1.5), Inches(12.1), Inches(4.9)).table
tbl.columns[0].width = Inches(5.2)
tbl.columns[1].width = Inches(6.9)
for ri, row in enumerate(rows):
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        for i, ln in enumerate(val.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(13 if ri == 0 else 12)
            r.font.bold = (ri == 0) or (ci == 0)
            r.font.color.rgb = WHITE if ri == 0 else (DARK if ci == 0 else GRAY)
            r.font.name = "微软雅黑"
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED if ri == 0 else (LIGHT if ri % 2 == 0 else WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
footer(s, 3)

# ===== S4 作品：市场分析报告 =====
s = add_slide()
header(s, "我的作品：平价现制饮品市场分析报告", "完整报告随简历附送，此处展示 3 个核心结论")
s.shapes.add_picture(os.path.join(HERE, "charts", "chart_growth.png"),
                     Inches(0.5), Inches(1.45), width=Inches(6.2))
text(s, Inches(7.0), Inches(1.6), Inches(5.8), Inches(5.2),
     [[("结论 1　", 15, True, RED), ("平价带（≤10 元）是增速最快的价格带，2018-2023 CAGR 23.3%，2028E 规模 3719 亿元——幸运咖正站在最陡的坡上。", 13, False, DARK)]],
     space_after=14)
text(s, Inches(7.0), Inches(2.75), Inches(5.8), Inches(5.2),
     [[("结论 2　", 15, True, RED), ("下沉市场是基本盘：三线及以下占蜜雪内地门店 58%、占幸运咖门店 70%，该市场茶饮 CAGR 22.8%。", 13, False, DARK)]],
     space_after=14)
text(s, Inches(7.0), Inches(3.9), Inches(5.8), Inches(5.2),
     [[("结论 3　", 15, True, RED), ("现磨咖啡渗透率将从 9.4%（2023）升至 18.3%（2028E），幸运咖万店后『单店营收优先』的转向正当其时。", 13, False, DARK)]],
     space_after=14)
footer(s, 4)

# ===== S5 其他项目 =====
s = add_slide()
header(s, "其他项目经历", "数据分析能力在产品场景中的两次落地")
box(s, Inches(0.6), Inches(1.6), Inches(6.0), Inches(4.6), fill=LIGHT)
box(s, Inches(0.6), Inches(1.6), Inches(0.12), Inches(4.6), fill=RED)
text(s, Inches(0.95), Inches(1.85), Inches(5.4), Inches(4.2),
     [[("AI 考研数学错题系统", 17, True, DARK)],
      [("· 设计错题标签体系：知识点 × 错误类型 × 掌握度", 12.5, False, GRAY)],
      [("· 对错题做分类统计与薄弱点排序 = 一次完整的用户行为数据分析", 12.5, False, GRAY)],
      [("· 基于反馈持续修正统计口径", 12.5, False, GRAY)],
      [("锻炼能力：数据结构设计、口径管理", 12.5, True, RED)]], space_after=8)
box(s, Inches(6.9), Inches(1.6), Inches(6.0), Inches(4.6), fill=LIGHT)
box(s, Inches(6.9), Inches(1.6), Inches(0.12), Inches(4.6), fill=GOLD)
text(s, Inches(7.25), Inches(1.85), Inches(5.4), Inches(4.2),
     [[("背诵挖空 App", 17, True, DARK)],
      [("· 从 0 到 1：需求定义 → 内容结构化 → 上线验证", 12.5, False, GRAY)],
      [("· 把模糊需求拆解为可执行任务并逐项交付", 12.5, False, GRAY)],
      [("· 小范围用户验证后迭代 3 个版本", 12.5, False, GRAY)],
      [("锻炼能力：执行力、跨角色沟通", 12.5, True, RED)]], space_after=8)
footer(s, 5)

# ===== S6 岗位理解 =====
s = add_slide()
header(s, "我理解的这个岗位", "市场分析 ≠ 做表，而是『让业务少走弯路』")
items = [
    ("监测市场动态，输出定期报告", "→ 我会先搭指标框架：门店/同店/价格带/竞品动作，月报固定模板"),
    ("研究需求、用户行为与竞争格局", "→ 本份报告就是一次预演：价格带 × 下沉 × 咖啡三张地图"),
    ("支持策略制定，提供决策建议", "→ 每条结论配一条可执行建议，不做『只有数字没有态度』的报告"),
    ("专项调研 + 可视化成果", "→ Python 出图 + PPT 结构化表达，调研结束即成果交付"),
    ("维护统一分析标准与工具模板", "→ 口径文档化：来源、时点、计算公式写进模板，保证团队一致性"),
]
for i, (a, b) in enumerate(items):
    t = Inches(1.5 + i * 1.12)
    box(s, Inches(0.6), t, Inches(12.1), Inches(0.98), fill=LIGHT if i % 2 == 0 else WHITE)
    text(s, Inches(0.85), t + Inches(0.12), Inches(5.3), Inches(0.8), [[(a, 13.5, True, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(6.3), t + Inches(0.12), Inches(6.2), Inches(0.8), [[(b, 12.5, False, GRAY)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 6)

# ===== S7 为什么是幸运咖 =====
s = add_slide()
header(s, "为什么是幸运咖", "三个理由，都很实在")
text(s, Inches(0.7), Inches(1.7), Inches(12), Inches(4.8),
     [[("01　", 20, True, RED), ("赛道认同：", 17, True, DARK),
       ("6-8 元把咖啡卖进县城和大学城，是真正『把生意做在大多数人身边』的模式；我想参与它从万店到单店提质的下半场。", 14, False, DARK)]],
     space_after=18)
text(s, Inches(0.7), Inches(2.9), Inches(12), Inches(4.8),
     [[("02　", 20, True, RED), ("地域匹配：", 17, True, DARK),
       ("河南人，熟悉本地消费场景；幸运咖大本营就在郑州（河南门店 1500+），做本地市场调研有天然优势。", 14, False, DARK)]],
     space_after=18)
text(s, Inches(0.7), Inches(4.1), Inches(12), Inches(4.8),
     [[("03　", 20, True, RED), ("能力对口：", 17, True, DARK),
       ("岗位要的『独立分析 + 报告 + 可视化』，我已经用一份完整作品提前交了一次作业。", 14, False, DARK)]],
     space_after=18)
footer(s, 7)

# ===== S8 结尾 =====
s = add_slide()
box(s, 0, 0, prs.slide_width, prs.slide_height, fill=RED)
text(s, Inches(1.0), Inches(2.4), Inches(11.3), Inches(1.2),
     [[("谢谢，期待和雪王一起把咖啡卖进每个县城。", 30, True, WHITE)]])
text(s, Inches(1.02), Inches(3.7), Inches(11), Inches(1.6),
     [[("【姓名】　|　【电话】　|　【邮箱】", 16, False, WHITE)],
      [("附：《平价现制饮品市场分析报告》全文 12 页", 13, False, WHITE)]])

prs.save("面试自我介绍-幸运咖市场分析.pptx")
print("intro ppt saved:", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
