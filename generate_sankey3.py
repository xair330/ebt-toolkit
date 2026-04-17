import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.path import Path
import matplotlib.patches as mpath_patches
import os
import sys

sys.path.insert(0, r'c:\工作文件\数据管理\IE\ebt_analysis_toolkit')
from theme_matrix_analysis import (
    load_data, expand_to_risk, COMP_ORDER, THEME_ORDER, RISK_ORDER, RISK_NAMES, fp
)

def plot_sankey_3_stage(df: pd.DataFrame, title: str, subtitle: str, out_path: str):
    df_risk = expand_to_risk(df)
    if df_risk.empty:
        print(f"  [跳过] {title} — 无数据")
        return

    # 两个阶段的流向
    flow1 = df_risk.groupby(["胜任力", "训练主题"]).size().reset_index(name="count")
    flow2 = df_risk.groupby(["训练主题", "核心风险"]).size().reset_index(name="count")

    # 节点提取与排序
    left_totals = flow1.groupby("胜任力")["count"].sum()
    left_nodes = [c for c in COMP_ORDER if c in left_totals.index]

    mid_totals = df_risk.groupby("训练主题").size().sort_values(ascending=False)
    mid_nodes = list(mid_totals.head(15).index)  # 取TOP 15个主题

    right_totals = flow2.groupby("核心风险")["count"].sum()
    right_nodes = [r for r in RISK_ORDER if r in right_totals.index]

    if not left_nodes or not mid_nodes or not right_nodes:
        print(f"  [跳过] {title} — 流量为空")
        return

    # 仅保留包含在TOP 15的主题的流向
    flow1 = flow1[flow1["训练主题"].isin(mid_nodes)]
    flow2 = flow2[flow2["训练主题"].isin(mid_nodes)]

    BG_MAIN = "#0A1020"
    BG_BODY = "#0D1526"

    max_nodes = max(len(left_nodes), len(mid_nodes), len(right_nodes))
    fig_w = 34
    fig_h = max(14, max_nodes * 0.8 + 2)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax.set_facecolor(BG_BODY)
    ax.set_xlim(-0.05, 1.05)
    ax.set_ylim(-0.5, max_nodes + 0.5)
    ax.axis("off")

    x_l, x_m, x_r = 0.05, 0.5, 0.95
    bar_w = 0.025

    def get_positions(nodes):
        n = len(nodes)
        offset = (max_nodes - n) / 2
        pos = {}
        for i, node in enumerate(nodes):
            pos[node] = max_nodes - 1 - offset - i
        return pos

    pos_l = get_positions(left_nodes)
    pos_m = get_positions(mid_nodes)
    pos_r = get_positions(right_nodes)

    comp_colors = {
        "FPA": "#3498DB", "FPM": "#2980B9", "COM": "#F1C40F", "LTW": "#E67E22",
        "SAW": "#9B59B6", "WLM": "#2ECC71", "PSD": "#E74C3C"
    }

    theme_colors = {}
    palette = ["#3498DB", "#2ECC71", "#E74C3C", "#F39C12", "#9B59B6",
               "#1ABC9C", "#E67E22", "#2980B9", "#27AE60", "#C0392B",
               "#D35400", "#8E44AD", "#16A085", "#F1C40F", "#7F8C8D"]
    for i, t in enumerate(mid_nodes):
        theme_colors[t] = palette[i % len(palette)]

    risk_colors = {
        "CFIT": "#E74C3C", "LOC": "#FF6B6B", "MAC": "#F39C12",
        "ARC": "#E67E22", "RE": "#D35400", "RI": "#C0392B",
        "GD": "#7F8C8D", "KSF": "#9B59B6", "IFD": "#3498DB",
        "HSE": "#2ECC71", "综合胜任力短板": "#1ABC9C", "综合能力": "#1ABC9C"
    }

    def draw_node(x, y, color, label, ha, text_offset=0.015, val=0):
        bar_h = 0.6
        rect = mpatches.FancyBboxPatch(
            (x - bar_w/2, y - bar_h/2), bar_w, bar_h,
            boxstyle="round,pad=0.01",
            facecolor=color, alpha=0.9, edgecolor="none")
        ax.add_patch(rect)
        if ha == "right":
            ax.text(x - bar_w/2 - text_offset, y, f"{label} ({int(val)})",
                    ha=ha, va="center", color="#D4E8FF", fontproperties=fp("bold", 16))
        elif ha == "left":
            ax.text(x + bar_w/2 + text_offset, y, f"{label} ({int(val)})",
                    ha=ha, va="center", color="#D4E8FF", fontproperties=fp("bold", 16))
        else: # center
            ax.text(x, y + bar_h/2 + 0.05, f"{label} ({int(val)})",
                    ha="center", va="bottom", color="#D4E8FF", fontproperties=fp("bold", 14))

    for node, y in pos_l.items():
        draw_node(x_l, y, comp_colors.get(node, "#888"), node, "right", val=left_totals.get(node, 0))
    for node, y in pos_m.items():
        draw_node(x_m, y, theme_colors.get(node, "#888"), node, "center", val=mid_totals.get(node, 0))
    for node, y in pos_r.items():
        cn = RISK_NAMES.get(node, node)
        draw_node(x_r, y, risk_colors.get(node, "#888"), f"{node} {cn}", "left", val=right_totals.get(node, 0))

    def draw_flow(flow_df, pos_start, pos_end, x_start, x_end, start_col, end_col, color_dict):
        total_flow = flow_df["count"].sum()
        max_left = flow_df.groupby(start_col)["count"].sum().max()
        if max_left <= 0: max_left = 1
        
        for _, row in flow_df.iterrows():
            st, en, cnt = row[start_col], row[end_col], row["count"]
            if st not in pos_start or en not in pos_end:
                continue
            y_s = pos_start[st]
            y_e = pos_end[en]
            color = color_dict.get(st, "#666")
            
            alpha = max(0.1, min(0.6, cnt / (total_flow * 0.05)))
            lw = max(0.8, min(20, cnt / max_left * 40))

            mid_x = (x_start + x_end) / 2
            verts = [(x_start + bar_w/2, y_s),
                     (mid_x, y_s), (mid_x, y_e),
                     (x_end - bar_w/2, y_e)]
            codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
            path = Path(verts, codes)
            patch = mpath_patches.PathPatch(
                path, facecolor="none", edgecolor=color,
                lw=lw, alpha=alpha, capstyle="round")
            ax.add_patch(patch)

            if cnt >= total_flow * 0.02:
                # Add text to highest curvature portion or near middle
                tx = mid_x + (0.01 if x_start == x_l else -0.01)
                ax.text(tx, (y_s + y_e)/2, str(cnt), ha="center", va="center",
                        color="#D4E8FF", fontsize=11, fontproperties=fp("bold", 12), alpha=0.9)

    draw_flow(flow1, pos_l, pos_m, x_l, x_m, "胜任力", "训练主题", comp_colors)
    draw_flow(flow2, pos_m, pos_r, x_m, x_r, "训练主题", "核心风险", theme_colors)

    total = int(df_risk.shape[0])
    fig.text(0.5, 0.96, title, ha="center", va="top", fontproperties=fp("black", 38), color="#D4E8FF")
    fig.text(0.5, 0.925, subtitle + f"   |   总流量：{total}", ha="center", va="top", fontproperties=fp("bold", 20), color="#6080A0")
    
    ax.text(x_l, max_nodes + 0.2, "胜任力 (Competency)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))
    ax.text(x_m, max_nodes + 0.2, "训练主题 (Training Theme)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))
    ax.text(x_r, max_nodes + 0.2, "核心风险 (Core Risk)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))

    plt.tight_layout(rect=[0, 0, 1, 0.90])
    plt.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()

if __name__ == "__main__":
    DATA_FILE = r"c:\工作文件\数据管理\IE\2026.1-3.xls"
    OUT_DIR = r"c:\工作文件\数据管理\IE\分析产出(2026)"
    
    df_all, df_weak = load_data(DATA_FILE)
    
    df_a330 = df_all[df_all["机型"].astype(str).str.contains("330", na=False)]
    df_a350 = df_all[df_all["机型"].astype(str).str.contains("350", na=False)]
    
    print(f"A330 数据量: {len(df_a330)}")
    print(f"A350 数据量: {len(df_a350)}")

    if len(df_a330) > 0:
        plot_sankey_3_stage(
            df_a330,
            title="胜任力 → 训练主题 → 核心风险 流向图 (A330)",
            subtitle="2026 Q1 A330 EBT 数据",
            out_path=os.path.join(OUT_DIR, "三级桑基图_A330.png")
        )
    
    if len(df_a350) > 0:
        plot_sankey_3_stage(
            df_a350,
            title="胜任力 → 训练主题 → 核心风险 流向图 (A350)",
            subtitle="2026 Q1 A350 EBT 数据",
            out_path=os.path.join(OUT_DIR, "三级桑基图_A350.png")
        )
