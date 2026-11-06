# -*- coding: utf-8 -*-
"""生成市场分析图表（数据来源见报告附录页）"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import os

for f in ["Microsoft YaHei", "SimHei"]:
    if any(f.lower() in x.name.lower() for x in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = f
        break
plt.rcParams["axes.unicode_minus"] = False

RED = "#D64027"
DARK = "#333333"
GRAY = "#9AA0A6"
GOLD = "#E8A33D"
BLUE = "#3E7CB1"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts")
os.makedirs(OUT, exist_ok=True)

def style(ax):
    for s in ["top", "right"]:
        ax.spines[s].set_visible(False)
    for s in ["left", "bottom"]:
        ax.spines[s].set_color("#CCCCCC")
    ax.tick_params(colors=DARK, labelsize=10)
    ax.grid(axis="y", color="#EEEEEE", zorder=0)

# 1. 茶饮门店 TOP6 品牌格局（2026-08 极海/百科口径）
brands = ["蜜雪冰城", "古茗", "沪上阿姨", "茶百道", "益禾堂", "霸王茶姬"]
stores = [46447, 14317, 11387, 8440, 7204, 7019]
fig, ax = plt.subplots(figsize=(9, 4.6), dpi=200)
colors = [RED] + ["#E9B8AE"] * 5
bars = ax.barh(brands[::-1], stores[::-1], color=colors[::-1], zorder=3)
for b, v in zip(bars, stores[::-1]):
    ax.text(v + 500, b.get_y() + b.get_height() / 2, f"{v:,}", va="center", fontsize=10, color=DARK)
ax.set_xlim(0, 52000)
ax.set_title("现制茶饮门店数 TOP6（2026年8月，单位：家）", fontsize=13, color=DARK, pad=12)
ax.set_xlabel("")
style(ax)
ax.grid(axis="x", color="#EEEEEE", zorder=0)
ax.grid(axis="y", visible=False)
fig.text(0.99, 0.01, "来源：公开门店监测数据（截至2026-08-20）", ha="right", fontsize=8, color=GRAY)
fig.tight_layout()
fig.savefig(f"{OUT}/chart_stores_top6.png", bbox_inches="tight")
plt.close(fig)

# 2. 价格带竞争地图：客单价 vs 门店数（气泡）
pts = [
    ("蜜雪冰城", 6, 46447, RED),
    ("幸运咖", 7, 10000, GOLD),
    ("益禾堂", 10, 7204, "#E9B8AE"),
    ("古茗", 15, 14317, "#E9B8AE"),
    ("沪上阿姨", 15, 11387, "#E9B8AE"),
    ("茶百道", 16, 8440, "#E9B8AE"),
    ("霸王茶姬", 20, 7019, "#E9B8AE"),
    ("瑞幸咖啡", 15, 36310, BLUE),
]
fig, ax = plt.subplots(figsize=(9, 5), dpi=200)
for name, price, st, c in pts:
    ax.scatter(price, st, s=st / 12, color=c, alpha=0.75, edgecolors="white", linewidths=1.2, zorder=3)
    ax.annotate(f"{name}\n{st/10000:.1f}万店", (price, st), textcoords="offset points",
                xytext=(0, 14 if st < 30000 else -26), ha="center", fontsize=9, color=DARK)
ax.axvspan(0, 10, color="#FDF1EF", zorder=0)
ax.axvspan(10, 20, color="#F7F7F7", zorder=0)
ax.text(5, 49000, "平价带 ≤10元", ha="center", fontsize=11, color=RED, weight="bold")
ax.text(15, 49000, "大众带 10-20元", ha="center", fontsize=11, color=GRAY, weight="bold")
ax.text(24, 49000, "高价带", ha="center", fontsize=11, color=GRAY, weight="bold")
ax.set_xlabel("主力客单价（元）", fontsize=10, color=DARK)
ax.set_ylabel("门店数（家）", fontsize=10, color=DARK)
ax.set_title("价格带竞争地图：气泡大小=门店规模（2026年）", fontsize=13, color=DARK, pad=12)
ax.set_xlim(0, 28)
ax.set_ylim(0, 53000)
style(ax)
fig.tight_layout()
fig.savefig(f"{OUT}/chart_price_band.png", bbox_inches="tight")
plt.close(fig)

# 3. 下沉市场：三线及以下占比
labels = ["蜜雪冰城\n内地门店", "幸运咖\n门店(2024)", "全行业\n市场规模(2023)"]
vals = [58, 70, 45.4]
fig, ax = plt.subplots(figsize=(8, 4.4), dpi=200)
bars = ax.bar(labels, vals, color=[RED, GOLD, "#B9C4CC"], width=0.55, zorder=3)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v}%", ha="center", fontsize=12, color=DARK, weight="bold")
ax.set_ylim(0, 85)
ax.set_title("三线及以下市场占比：下沉是基本盘", fontsize=13, color=DARK, pad=12)
style(ax)
ax.grid(axis="x", visible=False)
fig.text(0.99, 0.01, "来源：蜜雪集团财报 / 极海品牌监测 / 公开行业报告", ha="right", fontsize=8, color=GRAY)
fig.tight_layout()
fig.savefig(f"{OUT}/chart_sinking.png", bbox_inches="tight")
plt.close(fig)

# 4. 幸运咖门店增长曲线
x = ["2022", "2023", "2024-12", "2025-03", "2025-06", "2025-08", "2025-11", "2026(计划)"]
y = [2300, 3000, 4500, 5000, 6000, 7000, 10000, 12000]
fig, ax = plt.subplots(figsize=(9, 4.6), dpi=200)
ax.plot(x[:7], y[:7], marker="o", color=RED, linewidth=2.4, markersize=7, zorder=3)
ax.plot(x[6:], y[6:], marker="o", color=GRAY, linewidth=2, linestyle="--", markersize=7, zorder=3)
for i, v in enumerate(y):
    ax.annotate(f"{v:,}", (x[i], v), textcoords="offset points", xytext=(0, 10), ha="center", fontsize=9, color=DARK)
ax.annotate("2026年主动踩刹车：\n全年新增≤2000家\n转向单店营收优先", xy=(7, 12000), xytext=(4.6, 11500),
            fontsize=9, color=DARK, arrowprops=dict(arrowstyle="->", color=GRAY))
ax.set_title("幸运咖门店数：从狂飙到提质（家）", fontsize=13, color=DARK, pad=12)
style(ax)
fig.text(0.99, 0.01, "来源：幸运咖官方公告 / 蜜雪集团财报 / 媒体报道", ha="right", fontsize=8, color=GRAY)
fig.tight_layout()
fig.savefig(f"{OUT}/chart_lucky_growth.png", bbox_inches="tight")
plt.close(fig)

# 5. 平价带与咖啡赛道增速
fig, ax = plt.subplots(figsize=(9, 4.6), dpi=200)
cats = ["平价带(≤10元)\n2018-2023 CAGR", "三线以下茶饮\nCAGR", "现磨咖啡\n2023-2028E CAGR", "现制茶饮整体\n2024-2030E CAGR"]
vals = [23.3, 22.8, 19.8, 12.3]
bars = ax.bar(cats, vals, color=[RED, GOLD, BLUE, "#B9C4CC"], width=0.55, zorder=3)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.5, f"{v}%", ha="center", fontsize=12, color=DARK, weight="bold")
ax.set_ylim(0, 28)
ax.set_title("增速对比：平价 + 下沉 + 咖啡是三条最快赛道", fontsize=13, color=DARK, pad=12)
style(ax)
ax.grid(axis="x", visible=False)
ax.tick_params(axis="x", labelsize=9)
fig.text(0.99, 0.01, "来源：全球现制饮品行业报告 / 艾媒咨询 / 公开行业数据", ha="right", fontsize=8, color=GRAY)
fig.tight_layout()
fig.savefig(f"{OUT}/chart_growth.png", bbox_inches="tight")
plt.close(fig)

print("charts done:", os.listdir(OUT))
