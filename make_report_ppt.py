# -*- coding: utf-8 -*-
"""《平价现制饮品市场分析报告》PPT（12 页）"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

RED = RGBColor(0xD6, 0x40, 0x27)
DARK = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x77, 0x77, 0x77)
LIGHT = RGBColor(0xF7, 0xF3, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xE8, 0xA3, 0x3D)
BLUE = RGBColor(0x3E, 0x7C, 0xB1)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts")


def add_slide():
    return prs.slides.add_slide(BLANK)


def box(slide, l, t, w, h, fill=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    return sp


def text(slide, l, t, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=6):
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


def header(slide, kicker, title):
    box(slide, 0, 0, prs.slide_width, Inches(1.15), fill=WHITE)
    box(slide, 0, Inches(1.12), prs.slide_width, Pt(2.5), fill=RED)
    text(slide, Inches(0.6), Inches(0.14), Inches(11), Inches(0.4), [[(kicker, 11, True, RED)]])
    text(slide, Inches(0.6), Inches(0.48), Inches(12), Inches(0.6), [[(title, 24, True, DARK)]])


def footer(slide, n, src=None):
    text(slide, Inches(12.2), Inches(7.05), Inches(1.0), Inches(0.35),
         [[(f"{n:02d} / 12", 10, False, GRAY)]], align=PP_ALIGN.RIGHT)
    if src:
        text(slide, Inches(0.6), Inches(7.05), Inches(10.5), Inches(0.35), [[(src, 9, False, GRAY)]])


def pic(slide, name, l, t, w=None, h=None):
    kw = {"width": Inches(w)} if w else {"height": Inches(h)}
    slide.shapes.add_picture(os.path.join(HERE, name), Inches(l), Inches(t), **kw)


# ===== 01 封面 =====
s = add_slide()
box(s, 0, 0, prs.slide_width, prs.slide_height, fill=RED)
box(s, Inches(0.9), Inches(1.7), Inches(0.12), Inches(3.0), fill=WHITE)
text(s, Inches(1.3), Inches(1.7), Inches(11), Inches(0.6), [[("市场分析报告 · 独立研究", 15, True, WHITE)]])
text(s, Inches(1.3), Inches(2.3), Inches(11.2), Inches(1.6),
     [[("平价现制饮品市场分析", 42, True, WHITE)],
      [("茶饮格局 · 价格带 · 下沉市场 · 幸运咖的机会", 20, False, WHITE)]])
text(s, Inches(1.32), Inches(4.5), Inches(11), Inches(1.4),
     [[("作者：【姓名】　|　2026 年 10 月", 14, False, WHITE)],
      [("数据全部来自公开来源（财报 / 门店监测 / 行业报告），口径与出处见附录", 12, False, WHITE)]])

# ===== 02 摘要 =====
s = add_slide()
header(s, "EXECUTIVE SUMMARY", "三条核心结论")
concl = [
    ("平价带是最陡的坡", "单价 ≤10 元的平价现制饮品 2018-2023 CAGR 达 23.3%，2023 年规模 1366 亿元，预计 2028 年达 3719 亿元；增速高于大众带与高价带。", RED),
    ("下沉市场是基本盘", "三线及以下城市占蜜雪冰城内地门店 58%、占幸运咖门店约 70%；该层级茶饮市场 CAGR 22.8%，2023 年规模 960 亿元、占全行业 45.4%。", GOLD),
    ("咖啡是第二增长曲线", "现磨咖啡 2023-2028E CAGR 19.8% 领跑现制饮品；消费渗透率将从 9.4%（2023）升至 18.3%（2028E）。幸运咖 2025-11 破万店后转向『单店营收优先』，方向正确。", BLUE),
]
for i, (t, d, c) in enumerate(concl):
    tpos = Inches(1.5 + i * 1.75)
    box(s, Inches(0.6), tpos, Inches(12.1), Inches(1.55), fill=LIGHT)
    box(s, Inches(0.6), tpos, Inches(0.12), Inches(1.55), fill=c)
    text(s, Inches(0.95), tpos + Inches(0.15), Inches(11.5), Inches(0.5), [[(f"结论 {i+1}　{t}", 17, True, DARK)]])
    text(s, Inches(0.95), tpos + Inches(0.68), Inches(11.5), Inches(0.85), [[(d, 12.5, False, GRAY)]])
footer(s, 2)

# ===== 03 方法与口径 =====
s = add_slide()
header(s, "METHODOLOGY", "研究方法与数据口径")
text(s, Inches(0.6), Inches(1.5), Inches(6.0), Inches(5.2),
     [[("数据来源（四类）", 16, True, DARK)],
      [("① 公司财报：蜜雪集团 2025 年报、2026 中报（港交所披露）", 12.5, False, GRAY)],
      [("② 门店监测：极海/窄门等第三方门店数据、品牌官方公告", 12.5, False, GRAY)],
      [("③ 行业报告：艾媒咨询、中国茶叶流通协会、全球现制饮品行业报告", 12.5, False, GRAY)],
      [("④ 媒体核查：界面、36氪、新京报等对门店数与战略的交叉报道", 12.5, False, GRAY)],
      [("分析框架", 16, True, DARK)],
      [("市场规模 → 竞争格局 → 价格带 → 区域结构 → 单品赛道 → 公司深拆 → 机会/风险 → 建议", 12.5, False, GRAY)]],
     space_after=8)
box(s, Inches(7.0), Inches(1.5), Inches(5.7), Inches(5.0), fill=LIGHT)
text(s, Inches(7.3), Inches(1.75), Inches(5.1), Inches(4.5),
     [[("口径声明（诚实原则）", 15, True, RED)],
      [("· 不同机构对『市场规模』口径不一（终端零售额/企业收入），本报告统一采用终端零售额口径并标注出处", 12, False, GRAY)],
      [("· 门店数为统计时点数，各品牌时点不同（2026-08 为主），对比时已注明", 12, False, GRAY)],
      [("· 带 E 的数字为机构预测值，非事实值", 12, False, GRAY)],
      [("· 无一手调研数据；结论仅基于公开信息，供面试交流", 12, False, GRAY)]], space_after=8)
footer(s, 3)

# ===== 04 市场规模 =====
s = add_slide()
header(s, "MARKET SIZE", "市场盘子：整体放缓，结构分化")
pic(s, "chart_growth.png", 0.5, 1.5, w=7.4)
text(s, Inches(8.2), Inches(1.7), Inches(4.6), Inches(5.0),
     [[("整体：", 14, True, DARK), ("现制茶饮 2024 年约 3680 亿元，增速回落至个位数，行业进入存量提质期。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.2), Inches(2.8), Inches(4.6), Inches(5.0),
     [[("结构：", 14, True, DARK), ("增长不再平均——平价带、下沉市场、现磨咖啡三条线增速显著高于大盘。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.2), Inches(3.9), Inches(4.6), Inches(5.0),
     [[("含义：", 14, True, RED), ("对幸运咖而言，『赛道 β』仍在，但公司 α 要从开店速度切换到单店效率。", 12.5, True, DARK)]], space_after=12)
footer(s, 4, "来源：艾媒咨询 / 全球现制饮品行业报告 / 中国茶叶流通协会")

# ===== 05 门店格局 =====
s = add_slide()
header(s, "COMPETITIVE LANDSCAPE", "门店格局：一超 + 多强，连锁化率持续提升")
pic(s, "chart_stores_top6.png", 0.5, 1.45, w=7.6)
text(s, Inches(8.4), Inches(1.7), Inches(4.4), Inches(5.0),
     [[("一超：", 14, True, RED), ("蜜雪冰城 4.6 万店，约为第二名古茗的 3.2 倍，规模护城河=供应链议价权。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.4), Inches(2.9), Inches(4.4), Inches(5.0),
     [[("总量：", 14, True, DARK), ("全国茶饮门店约 45.2 万家（2026-06），增速仅 3.5%；个体单店持续退出，份额向头部集中。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.4), Inches(4.1), Inches(4.4), Inches(5.0),
     [[("风险信号：", 14, True, DARK), ("蜜雪 2025 年闭店 2527 家（+57%）、2026H1 闭店 1289 家——密度逼近天花板，同店质量比新店数量更重要。", 12.5, False, GRAY)]], space_after=12)
footer(s, 5, "来源：公开门店监测（2026-08-20）/ 蜜雪集团财报 / 红餐产业研究院")

# ===== 06 价格带 =====
s = add_slide()
header(s, "PRICE BANDS", "价格带地图：幸运咖卡在『无人区』")
pic(s, "chart_price_band.png", 0.4, 1.4, w=7.8)
text(s, Inches(8.5), Inches(1.7), Inches(4.3), Inches(5.2),
     [[("≤10 元：", 14, True, RED), ("茶饮侧蜜雪一家独大；咖啡侧幸运咖（6-8 元）暂无同量级对手。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.5), Inches(2.8), Inches(4.3), Inches(5.2),
     [[("10-20 元：", 14, True, DARK), ("最拥挤：古茗/沪上阿姨/茶百道/瑞幸混战，价格战频发（9.9 元常态化）。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.5), Inches(3.9), Inches(4.3), Inches(5.2),
     [[("≥20 元：", 14, True, DARK), ("高端收缩，喜茶/奈雪均已下放价格带；霸王茶姬以 20 元档突围。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.5), Inches(5.0), Inches(4.3), Inches(2.0),
     [[("→ 幸运咖的 6-8 元咖啡是目前唯一『低价 + 万店』的咖啡身位。", 12.5, True, DARK)]], space_after=8)
footer(s, 6, "来源：品牌公开菜单价 / 门店监测数据整理")

# ===== 07 下沉市场 =====
s = add_slide()
header(s, "LOWER-TIER MARKETS", "下沉市场：基本盘，也是下一步的增量")
pic(s, "chart_sinking.png", 0.6, 1.5, w=6.6)
text(s, Inches(7.6), Inches(1.7), Inches(5.2), Inches(5.0),
     [[("为什么下沉赢：", 15, True, DARK)],
      [("· 租金/人力低 → 单店模型在低客单价下依然成立", 12.5, False, GRAY)],
      [("· 竞争密度低：2023 年三线以下茶饮规模 960 亿、占 45.4%，CAGR 22.8% 高于高线", 12.5, False, GRAY)],
      [("· 蜜雪供应链（产地仓+物流）先覆盖下沉，成本优势自我强化", 12.5, False, GRAY)],
      [("但 2026 的故事变了：", 15, True, RED)],
      [("幸运咖下半年主战场转向长三角/珠三角一二线商业街，北京门店已破 70 家——下沉守成、高线进攻的双线作战，对市场监测的颗粒度要求更高（要分到城市层级）。", 12.5, False, GRAY)]], space_after=8)
footer(s, 7, "来源：蜜雪集团财报 / 极海品牌监测 / 行业报告")

# ===== 08 咖啡赛道 =====
s = add_slide()
header(s, "COFFEE TRACK", "咖啡赛道：渗透率翻倍窗口期")
rows = [
    ("品牌", "门店数（2026）", "主力价格带", "定位关键词"),
    ("瑞幸咖啡", "36,310（2026Q2）", "9.9-15 元", "写字楼/外卖/数字化"),
    ("库迪咖啡", "10,000+", "8.8-15 元", "价格战发起者"),
    ("幸运咖", "10,000+（2025-11）", "6-8 元", "县城/大学城/外带"),
    ("挪瓦咖啡", "10,000+", "10-15 元", "店中店/轻资产"),
]
tbl = s.shapes.add_table(len(rows), 4, Inches(0.6), Inches(1.5), Inches(7.4), Inches(3.4)).table
tbl.columns[0].width = Inches(1.8)
tbl.columns[1].width = Inches(2.2)
tbl.columns[2].width = Inches(1.6)
tbl.columns[3].width = Inches(1.8)
for ri, row in enumerate(rows):
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        cell.text = ""
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = val
        r.font.size = Pt(12)
        r.font.bold = (ri == 0)
        r.font.color.rgb = WHITE if ri == 0 else (RED if (ri == 3) else DARK)
        r.font.name = "微软雅黑"
        cell.fill.solid()
        cell.fill.fore_color.rgb = RED if ri == 0 else (RGBColor(0xFD, 0xF1, 0xEF) if ri == 3 else (LIGHT if ri % 2 == 0 else WHITE))
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
text(s, Inches(8.4), Inches(1.7), Inches(4.4), Inches(5.0),
     [[("窗口期：", 14, True, RED), ("现磨咖啡渗透率 9.4%（2023）→ 18.3%（2028E），CAGR 19.8% 为全行业最快。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.4), Inches(2.9), Inches(4.4), Inches(5.0),
     [[("幸运咖的牌：", 14, True, DARK), ("蜜雪供应链（海南烘焙线年产能 2 万吨+、巴西 40 亿采购计划）+ 6.6 元心智（活动单日 GMV +20%）。", 12.5, False, GRAY)]], space_after=12)
text(s, Inches(8.4), Inches(4.3), Inches(4.4), Inches(5.0),
     [[("幸运咖的短板：", 14, True, DARK), ("高线市场品牌认知弱于瑞幸；『专业咖啡』心智仍在建设（郑州旗舰店已试水 10-16 元手冲系列）。", 12.5, False, GRAY)]], space_after=12)
footer(s, 8, "来源：瑞幸财报 / 界面新闻 / 36氪 / 蜜雪集团公告")

# ===== 09 幸运咖深拆 =====
s = add_slide()
header(s, "DEEP DIVE", "幸运咖：从狂飙到提质的转折点")
pic(s, "chart_lucky_growth.png", 0.5, 1.45, w=7.6)
text(s, Inches(8.4), Inches(1.7), Inches(4.4), Inches(5.2),
     [[("三个阶段：", 15, True, DARK)],
      [("① 2017-2021 摸索期：定位不清，反响平平", 12.5, False, GRAY)],
      [("② 2022-2025 狂飙期：下沉+校园双轮，10 个月门店翻倍破万", 12.5, False, GRAY)],
      [("③ 2026 提质期：全年新增 ≤2000 家，『单店营收增长为主、规模为辅』", 12.5, False, GRAY)],
      [("对市场分析岗的含义：", 14, True, RED)],
      [("考核重心从『开了多少店』变为『单店卖了多少』——同店销售、复购、时段结构将成为核心监测指标。", 12.5, False, GRAY)]], space_after=8)
footer(s, 9, "来源：幸运咖官方公告 / 蜜雪集团财报 / 界面新闻")

# ===== 10 机会与风险 =====
s = add_slide()
header(s, "OPPORTUNITIES & RISKS", "机会与风险清单")
box(s, Inches(0.6), Inches(1.5), Inches(6.0), Inches(5.1), fill=RGBColor(0xEF, 0xF6, 0xEF))
text(s, Inches(0.9), Inches(1.7), Inches(5.4), Inches(4.7),
     [[("机会", 17, True, RGBColor(0x3F, 0x7A, 0x40))],
      [("· 咖啡渗透率翻倍窗口（9.4%→18.3%）", 12.5, False, GRAY)],
      [("· 6-8 元咖啡价格带暂无万店级对手", 12.5, False, GRAY)],
      [("· 蜜雪供应链复用：边际成本低于独立咖啡品牌", 12.5, False, GRAY)],
      [("· 校园渠道：约 1/4 门店在学校商圈，养成期用户终身价值高", 12.5, False, GRAY)],
      [("· 出海复制：马来西亚首店开业当日近 2000 杯", 12.5, False, GRAY)]], space_after=8)
box(s, Inches(6.9), Inches(1.5), Inches(6.0), Inches(5.1), fill=RGBColor(0xFB, 0xEE, 0xEC))
text(s, Inches(7.2), Inches(1.7), Inches(5.4), Inches(4.7),
     [[("风险", 17, True, RED)],
      [("· 集团利润承压：2026H1 净利 -14.7%，补贴退坡后高基数效应", 12.5, False, GRAY)],
      [("· 闭店率上行：2025 年闭店 2527 家，加盟商盈利是生命线", 12.5, False, GRAY)],
      [("· 高线进攻遇瑞幸密度网：写字楼/外卖场景非幸运咖传统强项", 12.5, False, GRAY)],
      [("· 价格战消耗：9.9/6.6 元活动拉新但压缩毛利", 12.5, False, GRAY)],
      [("· 心智天花板：『便宜』易得，『好喝』难建", 12.5, False, GRAY)]], space_after=8)
footer(s, 10)

# ===== 11 建议 =====
s = add_slide()
header(s, "RECOMMENDATIONS", "给幸运咖市场监测体系的三条建议（= 我入职后的工作清单）")
recs = [
    ("建一张『月度市场监测仪表盘』", "固定四组指标：自身（门店/同店/客单/时段）、竞品（瑞幸/库迪新店与价格动作）、价格带（6-8 元带份额）、区域（分城市层级）。模板化后每月 3 个工作日产出月报。"),
    ("画一张『县域空白地图』", "以县为单位叠加人口、GDP、学校数、现有门店与竞品距离，输出开店优先级评分——把『下沉还有空间』从口号变成排期表。"),
    ("立一套『口径与模板标准』", "所有报告统一标注：数据来源、统计时点、计算公式、预测/事实区分（对应 JD 第 5 条：维护统一分析标准与工具模板）。"),
]
for i, (t, d) in enumerate(recs):
    tpos = Inches(1.5 + i * 1.75)
    box(s, Inches(0.6), tpos, Inches(12.1), Inches(1.55), fill=LIGHT)
    box(s, Inches(0.6), tpos, Inches(0.9), Inches(1.55), fill=RED)
    text(s, Inches(0.6), tpos, Inches(0.9), Inches(1.55), [[(f"{i+1}", 30, True, WHITE)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(1.8), tpos + Inches(0.15), Inches(10.6), Inches(0.5), [[(t, 16, True, DARK)]])
    text(s, Inches(1.8), tpos + Inches(0.66), Inches(10.6), Inches(0.85), [[(d, 12, False, GRAY)]])
footer(s, 11)

# ===== 12 附录 =====
s = add_slide()
header(s, "APPENDIX", "数据来源与免责声明")
srcs = [
    "蜜雪集团 2025 年度业绩公告、2026 年中期报告（港交所披露，2026-03 / 2026-08）",
    "幸运咖官方公众号门店公告（2025-02 至 2025-11）；蜜雪集团对幸运咖规划的公开表述",
    "门店监测数据：极海品牌监测 / 窄门餐眼 / 公开整理（截至 2026-08-20）",
    "行业报告：艾媒咨询《2026-2027 中国新式茶饮消费趋势研究报告》、《全球现制饮品行业报告》（2026-10）、中国茶叶流通协会",
    "媒体核查：界面新闻、36氪、新京报、红餐网（2024-11 至 2026-08）",
]
text(s, Inches(0.6), Inches(1.5), Inches(12.1), Inches(4.0),
     [[(f"{i+1}. {x}", 12.5, False, DARK)] for i, x in enumerate(srcs)], space_after=8)
box(s, Inches(0.6), Inches(5.3), Inches(12.1), Inches(1.3), fill=LIGHT)
text(s, Inches(0.9), Inches(5.5), Inches(11.5), Inches(1.0),
     [[("免责声明：", 12.5, True, RED),
       ("本报告为个人独立研究作品，仅基于公开信息整理，不构成投资建议；带 E 数据为第三方机构预测值。部分指标统计时点不同，对比时已在文中标注。", 12, False, GRAY)]])
text(s, Inches(0.6), Inches(6.75), Inches(12), Inches(0.5),
     [[("作者：【姓名】　|　联系方式：【电话 / 邮箱】　|　完成时间：2026-10", 12, True, DARK)]])
footer(s, 12)

prs.save("平价现制饮品市场分析报告-幸运咖.pptx")
print("report saved")
