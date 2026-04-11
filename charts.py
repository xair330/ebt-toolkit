"""
EBT分析工具包 - 图表渲染模块
所有可视化均使用玻璃座舱风格（浅色背景）
"""
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import matplotlib.colors as mcolors
import matplotlib.font_manager as fm
from matplotlib.font_manager import FontProperties
import seaborn as sns
import os

from config_mgr import cfg

def _setup_fonts():
    """注册鸿蒙字体，返回 (font_medium, font_black, font_regular)"""
    fonts = cfg["FONTS"]
    FONT_MEDIUM, FONT_BLACK, FONT_REGULAR = fonts["FONT_MEDIUM"], fonts["FONT_BLACK"], fonts["FONT_REGULAR"]
    for p in [FONT_MEDIUM, FONT_BLACK, FONT_REGULAR]:
        if os.path.exists(p):
            fm.fontManager.addfont(p)
    fmid = FontProperties(fname=FONT_MEDIUM)
    fblk = FontProperties(fname=FONT_BLACK)
    freg = FontProperties(fname=FONT_REGULAR)
    plt.rcParams["font.family"] = [fmid.get_name()]
    plt.rcParams["axes.unicode_minus"] = False
    return fmid, fblk, freg

FMID, FBLK, FREG = _setup_fonts()

T = cfg["THEME"]   # 颜色别名


def _set_glass_ax(ax, fig):
    fig.patch.set_facecolor(T["BG_DEEP"])
    ax.set_facecolor(T["GLASS_FACE"])
    for sp in ax.spines.values():
        sp.set_edgecolor(T["GLASS_EDGE"])
        sp.set_linewidth(1)
    ax.tick_params(colors=T["TEXT_DARK"], labelsize=11)
    ax.grid(color=T["GRID_GLASS"], linewidth=0.7, zorder=0)


def _title(ax, text):
    ax.set_title(text, fontproperties=FBLK, fontsize=15,
                 color=T["TEXT_DARK"], pad=18, fontweight="bold")


# ══════════════════════════════════════════════════════════════════
# 图1: 问题大类总览（环形图 + 条形图）
# ══════════════════════════════════════════════════════════════════
def plot_category_overview(cat_counts, output_path):
    palette = T["PALETTE"]
    fig = plt.figure(figsize=(18, 8), facecolor=T["BG_DEEP"])
    gs = gridspec.GridSpec(1, 2, width_ratios=[1, 1.3], figure=fig,
                           left=0.05, right=0.97, wspace=0.14, bottom=0.1, top=0.9)

    # 环形图
    ax1 = fig.add_subplot(gs[0])
    ax1.set_facecolor(T["GLASS_FACE"])
    wedges, lbl_texts, pct_texts = ax1.pie(
        cat_counts.values, labels=cat_counts.index, autopct="%1.1f%%",
        colors=palette, startangle=140, pctdistance=0.75,
        wedgeprops=dict(width=0.55, edgecolor=T["BG_DEEP"], linewidth=2.5),
        textprops={"fontsize": 13}
    )
    for t in lbl_texts:
        t.set_fontproperties(FMID); t.set_color(T["TEXT_DARK"]); t.set_fontsize(13)
    for pt in pct_texts:
        pt.set_fontsize(15); pt.set_fontproperties(FBLK); pt.set_color("#FFFFFF")
    ax1.text(0, 0, f"{sum(cat_counts)}\n记录总计", ha="center", va="center",
             fontproperties=FBLK, fontsize=15, color=T["TEXT_DARK"])
    _title(ax1, "问题大类分布总览")

    # 条形图
    ax2 = fig.add_subplot(gs[1])
    ax2.set_facecolor(T["GLASS_FACE"])
    for sp in ax2.spines.values(): sp.set_edgecolor(T["GLASS_EDGE"]); sp.set_linewidth(1)
    sorted_cats = cat_counts.sort_values()
    bar_colors = [palette[i % len(palette)] for i in range(len(sorted_cats))]
    bars = ax2.barh(sorted_cats.index, sorted_cats.values, color=bar_colors,
                    height=0.62, edgecolor=T["BG_DEEP"], linewidth=1)
    for bar, val in zip(bars, sorted_cats.values):
        ax2.text(val + 0.3, bar.get_y() + bar.get_height() / 2,
                 str(val), va="center", ha="left", fontproperties=FBLK,
                 fontsize=13, color=T["TEXT_DARK"])
    ax2.set_xlabel("出现频次", fontproperties=FMID, color=T["TEXT_MID"], fontsize=11)
    ax2.spines[["top", "right", "left"]].set_visible(False)
    ax2.spines["bottom"].set_color(T["GLASS_EDGE"])
    ax2.tick_params(colors=T["TEXT_DARK"], labelsize=11)
    ax2.grid(axis="x", color=T["GRID_GLASS"], linewidth=0.8)
    for lb in ax2.get_yticklabels(): lb.set_fontproperties(FMID)
    _title(ax2, "各问题大类频次排行")

    fig.suptitle(f"{cfg['REPORT_PERIOD']} EBT不足项 · 问题大类分析", fontproperties=FBLK,
                 fontsize=17, color=T["TEXT_DARK"], y=0.98)
    plt.savefig(output_path, dpi=150, facecolor=T["BG_DEEP"])
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 图2: 胜任力 × 问题大类 热力矩阵
# ══════════════════════════════════════════════════════════════════
def plot_heatmap(cross_df, output_path):
    fig, ax = plt.subplots(figsize=(17, 7), facecolor=T["BG_DEEP"])
    _set_glass_ax(ax, fig)
    ax.grid(False)
    cmap = sns.color_palette("flare", as_cmap=True)
    sns.heatmap(cross_df, annot=False, cmap=cmap, linewidths=1.8,
                linecolor=T["BG_DEEP"], ax=ax, cbar_kws={"shrink": 0.65})

    vmin, vmax = cross_df.values.min(), cross_df.values.max()
    for i, row_name in enumerate(cross_df.index):
        for j, col_name in enumerate(cross_df.columns):
            val = cross_df.loc[row_name, col_name]
            ratio = (val - vmin) / (vmax - vmin + 1e-9)
            txt_color = "#FFFFFF" if ratio > 0.45 else "#1C1C2E"
            ax.text(j + 0.5, i + 0.5, str(val), ha="center", va="center",
                    fontsize=13, fontproperties=FBLK, color=txt_color, zorder=5)

    cbar = ax.collections[0].colorbar
    cbar.ax.tick_params(colors=T["TEXT_DARK"], labelsize=10)
    cbar.outline.set_visible(False)

    ax.set_title(f"{cfg['REPORT_PERIOD']} EBT低分项胜任力分布", fontproperties=FBLK,
                 fontsize=15, color=T["TEXT_DARK"], pad=18)
    ax.set_xlabel("问题大类", fontproperties=FMID, color=T["TEXT_MID"], fontsize=12)
    ax.set_ylabel("胜任力维度", fontproperties=FMID, color=T["TEXT_MID"], fontsize=12)
    for lb in ax.get_xticklabels():
        lb.set_fontproperties(FMID); lb.set_fontsize(11); lb.set_color(T["TEXT_DARK"])
    for lb in ax.get_yticklabels():
        lb.set_fontproperties(FMID); lb.set_fontsize(11); lb.set_color(T["TEXT_DARK"])
    ax.tick_params(bottom=False, left=False)
    plt.tight_layout(pad=1.5)
    plt.savefig(output_path, dpi=150, facecolor=T["BG_DEEP"])
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 图3: 角色对比条形图
# ══════════════════════════════════════════════════════════════════
def plot_role_comparison(role_cross, output_path, role_order=None):
    if role_order is None:
        role_order = ["教员", "机长", "副驾驶"]
    role_colors = T["ROLE_COLORS"]
    palette = T["PALETTE"]

    role_cross = role_cross.reindex(columns=[r for r in role_order if r in role_cross.columns])
    sort_col = next((c for c in ["副驾驶", "机长", "教员"] if c in role_cross.columns),
                    role_cross.columns[0])
    role_cross = role_cross.sort_values(sort_col, ascending=False)

    x = np.arange(len(role_cross))
    n = len(role_cross.columns)
    offsets = np.linspace(-(n-1)/2, (n-1)/2, n) * 0.25

    fig, ax = plt.subplots(figsize=(17, 7.5), facecolor=T["BG_DEEP"])
    _set_glass_ax(ax, fig)

    for i, role in enumerate(role_cross.columns):
        vals = role_cross[role].values
        color = role_colors.get(role, palette[i % len(palette)])
        bars = ax.bar(x + offsets[i], vals, width=0.25, color=color, alpha=0.88,
                      label=role, edgecolor="white", linewidth=0.8, zorder=3)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.25,
                        str(val), ha="center", va="bottom",
                        fontproperties=FBLK, fontsize=12, color=T["TEXT_DARK"])

    ax.set_xticks(x)
    ax.set_xticklabels(role_cross.index, rotation=18, ha="right")
    for lb in ax.get_xticklabels(): lb.set_fontproperties(FMID); lb.set_fontsize(11)
    for lb in ax.get_yticklabels(): lb.set_fontproperties(FMID)
    ax.set_ylabel("出现频次", fontproperties=FMID, color=T["TEXT_MID"], fontsize=11)
    ax.spines[["top", "right"]].set_visible(False)
    ax.spines[["left", "bottom"]].set_color(T["GLASS_EDGE"])
    ax.grid(axis="y", color=T["GRID_GLASS"], linewidth=0.8, zorder=0)
    ax.legend(prop=FMID, frameon=False, fontsize=12, labelcolor=T["TEXT_DARK"], loc="upper right")
    _title(ax, "教员 / 机长 / 副驾驶  各问题大类失误频次对比")
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, facecolor=T["BG_DEEP"])
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 图4: 技术等级 × 胜任力 气泡图
# ══════════════════════════════════════════════════════════════════
def plot_level_bubble(bubble_data, level_order, comp_order, output_path):
    fig, ax = plt.subplots(figsize=(15, 7.5), facecolor=T["BG_DEEP"])
    _set_glass_ax(ax, fig)

    for gx in range(len(level_order)):
        ax.axvline(gx, color=T["GRID_GLASS"], linewidth=0.8, zorder=1)
    for gy in range(len(comp_order)):
        ax.axhline(gy, color=T["GRID_GLASS"], linewidth=0.8, zorder=1)

    sc = ax.scatter(bubble_data["lx"], bubble_data["cy"],
                    s=bubble_data["count"] * 120, c=bubble_data["count"],
                    cmap="Blues", alpha=0.85, edgecolors=T["GLASS_EDGE"], linewidths=0.8, zorder=3)
    for _, row in bubble_data.iterrows():
        ax.text(row["lx"], row["cy"], str(int(row["count"])),
                ha="center", va="center", fontproperties=FBLK, fontsize=11,
                color=T["TEXT_DARK"], zorder=4)

    ax.set_xticks(range(len(level_order)))
    ax.set_xticklabels(level_order, color=T["TEXT_DARK"], fontsize=12)
    ax.set_yticks(range(len(comp_order)))
    ax.set_yticklabels(comp_order, color=T["TEXT_DARK"], fontsize=12)
    for lb in ax.get_xticklabels(): lb.set_fontproperties(FMID)
    for lb in ax.get_yticklabels(): lb.set_fontproperties(FMID)
    ax.set_xlabel("技术等级（从低到高）", fontproperties=FMID, color=T["TEXT_MID"], fontsize=12)
    ax.set_ylabel("胜任力维度", fontproperties=FMID, color=T["TEXT_MID"], fontsize=12)
    ax.tick_params(colors=T["TEXT_DARK"], length=0)
    ax.set_xlim(-0.6, len(level_order) - 0.4)
    ax.set_ylim(-0.6, len(comp_order) - 0.4)
    ax.spines[:].set_visible(False)

    cbar = plt.colorbar(sc, ax=ax, shrink=0.6, pad=0.02)
    cbar.ax.tick_params(colors=T["TEXT_DARK"], labelsize=10)
    cbar.set_label("频次", color=T["TEXT_MID"], fontproperties=FMID)
    cbar.outline.set_visible(False)

    _title(ax, f"各技术等级  ×  胜任力维度  失误分布气泡图")
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, facecolor=T["BG_DEEP"])
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 图5: A330 vs A350 雷达图
# ══════════════════════════════════════════════════════════════════
def plot_aircraft_radar(radar_data, comp_labels, output_path):
    angles = np.linspace(0, 2*np.pi, len(comp_labels), endpoint=False).tolist()
    angles_plot = angles + angles[:1]

    fig, ax = plt.subplots(figsize=(10, 10), facecolor=T["BG_DEEP"],
                           subplot_kw=dict(polar=True))
    ax.set_facecolor(T["GLASS_FACE"])

    ac_styles = {
        "A330": {"color": "#2C6FAC", "marker": "o"},
        "A350": {"color": "#C47A12", "marker": "s"},
    }
    for ac, style in ac_styles.items():
        if ac not in radar_data: continue
        vals = radar_data[ac] + [radar_data[ac][0]]
        ax.plot(angles_plot, vals, "-", linewidth=2.5, label=ac,
                color=style["color"], marker=style["marker"], markersize=7, zorder=3)
        ax.fill(angles_plot, vals, alpha=0.12, color=style["color"])

    ax.set_thetagrids(np.degrees(angles), comp_labels, fontsize=12)
    for lb in ax.get_xticklabels():
        lb.set_fontproperties(FBLK); lb.set_color(T["TEXT_DARK"])
    for lb in ax.get_yticklabels():
        lb.set_fontproperties(FREG); lb.set_color(T["TEXT_MID"]); lb.set_fontsize(9)
    max_val = max(max(v) for v in radar_data.values() if v)
    ax.set_ylim(0, max_val * 1.25)
    ax.grid(color=T["GRID_GLASS"], linewidth=1)
    ax.spines["polar"].set_color(T["GLASS_EDGE"])
    ax.legend(loc="upper right", bbox_to_anchor=(1.35, 1.15),
              prop=FBLK, frameon=False, fontsize=13, labelcolor=T["TEXT_DARK"])
    ax.set_title("A330 vs A350  胜任力不足分布雷达图",
                 fontproperties=FBLK, fontsize=15, color=T["TEXT_DARK"], pad=28)
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, facecolor=T["BG_DEEP"])
    plt.close()


