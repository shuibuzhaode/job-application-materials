# -*- coding: utf-8 -*-
"""生成针对幸运咖市场分析岗的简历 DOCX"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED = RGBColor(0xD6, 0x40, 0x27)
DARK = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x66, 0x66, 0x66)

doc = Document()
sec = doc.sections[0]
sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.6)
sec.left_margin, sec.right_margin = Cm(1.8), Cm(1.8)

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
style.font.color.rgb = DARK
style.paragraph_format.space_after = Pt(2)
style.paragraph_format.line_spacing = 1.25


def set_cn(run, size=None, bold=None, color=None):
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def heading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_cn(r, 12.5, True, RED)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:color"), "D64027")
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix)
        set_cn(r, 10.5, True, DARK)
    r = p.add_run(text)
    set_cn(r, 10.5, False, DARK)
    return p


def line(label, value):
    p = doc.add_paragraph()
    r = p.add_run(label)
    set_cn(r, 10.5, True, DARK)
    r = p.add_run(value)
    set_cn(r, 10.5, False, DARK)
    return p


# ===== 头部 =====
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run("【姓名】")
set_cn(r, 22, True, DARK)
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
r = p.add_run("求职意向：市场分析（幸运咖 · 郑州）　|　到岗时间：【填写】")
set_cn(r, 11, True, RED)

p = doc.add_paragraph()
r = p.add_run("电话：【手机号】　|　邮箱：【邮箱】　|　现居：【城市，如：郑州/河南】")
set_cn(r, 10, False, GRAY)

# ===== 教育背景 =====
heading("教育背景")
line("【学校名称】　统计学　本科　", "【20XX.09 - 20XX.06】")
bullet("主修课程：概率论与数理统计、回归分析、多元统计分析、时间序列分析、统计软件（Python/R）")
bullet("统计学科班训练：具备假设检验、抽样调查、回归建模的完整方法论基础，能独立完成从数据收集到结论输出的分析闭环")

# ===== 专业技能 =====
heading("专业技能（对应任职要求 1、2）")
bullet("Python 数据分析：", "熟练使用 pandas / numpy 完成数据清洗与统计，matplotlib 完成可视化出图（本次随附市场分析报告全部图表均由本人独立完成）")
bullet("SQL：", "掌握多表查询、聚合与窗口函数，可独立从数据库取数")
bullet("Excel：", "熟练运用透视表、VLOOKUP/INDEX-MATCH、条件格式搭建日常监控报表")
bullet("报告与可视化：", "具备『数据 → 图表 → 结论 → 建议』的完整报告撰写能力，熟悉 PPT 结构化表达")
bullet("数据源积累：", "熟悉上市公司财报（港交所披露）、极海/窄门门店监测、艾媒咨询、中国茶叶流通协会等公开行业数据渠道，注重口径标注与交叉验证")

# ===== 项目经历 =====
heading("项目经历")

p = doc.add_paragraph()
r = p.add_run("平价现制饮品（茶饮/咖啡）市场分析　|　独立研究项目　|　2026.10")
set_cn(r, 10.5, True, DARK)
bullet("背景：", "针对幸运咖所在的价格带与赛道，独立完成一份 12 页可视化市场分析报告（随简历附送）")
bullet("做法：", "收集蜜雪集团 2025 年报及 2026 中报、幸运咖官方门店公告、TOP6 茶饮品牌门店监测数据等 10+ 公开来源，统一口径后完成市场规模、竞争格局、价格带、下沉市场、咖啡赛道五个维度分析")
bullet("产出：", "5 张自绘图表 + 3 条核心结论（平价带 2018-2023 CAGR 23.3% 为最快价格带；三线以下占蜜雪内地门店 58%；现磨咖啡渗透率 2028E 翻倍至 18.3%），并针对幸运咖 2026『单店营收优先』战略提出 3 条可落地的市场监测建议")

p = doc.add_paragraph()
r = p.add_run("AI 考研数学错题系统　|　个人产品项目（AI 辅助开发）　|　2025 - 2026")
set_cn(r, 10.5, True, DARK)
bullet("独立定义需求与数据结构：", "设计错题标签体系（知识点 × 错误类型 × 掌握度），对错题数据做分类统计与薄弱点排序，本质是一次完整的『用户行为数据分析』实践")
bullet("迭代方式：", "基于使用反馈持续调整标签口径与统计逻辑，养成『先核对数据口径、再下结论』的习惯")

p = doc.add_paragraph()
r = p.add_run("背诵挖空 App　|　个人产品项目（AI 辅助开发）　|　2025")
set_cn(r, 10.5, True, DARK)
bullet("从 0 到 1 完成产品定义、内容结构化与上线验证，锻炼了把模糊需求拆解为可执行任务的能力（对应任职要求 3、4 的跨部门协作与执行力）")

# ===== 自我评价 =====
heading("自我评价（对应任职要求 3、4）")
bullet("数据敏感度：", "统计学出身 + 完整独立做过行业分析，看到数字会先问口径、来源和对比基期")
bullet("沟通协作：", "习惯把分析结论翻译成业务方听得懂的语言（一句话结论 + 一张图 + 一条建议），降低跨部门沟通成本")
bullet("细致踏实：", "报告中每个数字均标注来源与统计时点，交叉验证后才写入结论；抗压能力强，接受定期报告节奏与专项调研的并行安排")
bullet("动机：", "河南人，长期关注蜜雪/幸运咖从郑州走向全国的路径，希望以市场分析岗为起点深耕现制饮品行业")

doc.save("简历-市场分析-幸运咖.docx")
print("resume saved")
