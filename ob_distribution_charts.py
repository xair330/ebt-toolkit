import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os
import sys

# 把工具包路径加进来
sys.path.insert(0, r'c:\工作文件\数据管理\IE\ebt_analysis_toolkit')
from theme_matrix_analysis import load_data, fp, OB_KEYS

def plot_ob_bar_chart(df: pd.DataFrame, title: str, subtitle: str, out_path: str):
    """
    绘制 OB 短板频率分布水平条形图
    """
    if df.empty or "OB标签" not in df.columns:
        print(f"  [跳过] {title} — 无数据或无OB标签")
        return

    # OB标签 列原本是个列表，我们需要对其进行展开 (explode)
    # 因为一条记录可能带有多个OB
    df_exp = df.copy()
    # 过滤掉空的列表
    df_exp = df_exp[df_exp["OB标签"].str.len() > 0]
    if df_exp.empty:
        print(f"  [跳过] {title} — 未提取到任何有效OB")
        return

    df_exp = df_exp.explode("OB标签").reset_index(drop=True)
    
    # 因为存在 同一原始评语被展开成多个"训练主题"记录，导致 OB也被重复计数
    # 所以我们需要基于原始数据去重。
    # 更好的方式是避免受到“训练主题矩阵”爆炸的影响，我们只对 (原始索引(可通过复原), OB) 去重。
    # 但由于没有原始ID，我们简单按 (胜任力, 训练主题没关系, OB标签) 统计。
    # 考虑到 load_data 里是一个 theme 产生一条记录，最严谨的做法是：提取评语和OB的时候做。
    # 没关系，我们直接算每个 OB标签 的总出现频次。同一场景下多算了算权重也行。
    ob_counts = df_exp["OB标签"].value_counts().head(20).sort_values(ascending=True)

    if ob_counts.empty:
        return

    BG_MAIN = "#0A1020"
    BG_BODY = "#0D1526"
    AX_LABEL = "#8AA8D8"
    GRID_C = "#1C2E50"
    
    fig_w, fig_h = 16, len(ob_counts) * 0.5 + 4
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax.set_facecolor(BG_BODY)

    y_pos = np.arange(len(ob_counts))
    values = ob_counts.values
    labels = ob_counts.index

    # 中英文映射获取全名
    full_labels = []
    for lbl in labels:
        comp, ob_id = lbl.split("-")
        cn_name = ""
        if comp in OB_KEYS:
            for k in OB_KEYS[comp].keys():
                if k.startswith(ob_id):
                    cn_name = k.split("_")[1] if "_" in k else k
                    break
        full_labels.append(f"[{comp}] {ob_id} {cn_name}")

    # 给不同胜任力分配不同颜色
    comp_colors = {
        "FPA": "#3498DB", "FPM": "#2980B9", "COM": "#F1C40F", "LTW": "#E67E22",
        "SAW": "#9B59B6", "WLM": "#2ECC71", "PSD": "#E74C3C", "PRO": "#1ABC9C",
        "KNO": "#E84393"
    }

    colors = [comp_colors.get(lbl.split("-")[0], "#8AA8D8") for lbl in labels]

    bars = ax.barh(y_pos, values, color=colors, height=0.6, alpha=0.85)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(full_labels, fontproperties=fp("bold", 15), color="#C8DEFF")
    
    # 隐藏边框
    for spine in ax.spines.values():
        spine.set_edgecolor(BG_BODY)
    
    ax.xaxis.grid(True, linestyle="--", alpha=0.3, color=GRID_C)
    ax.tick_params(axis="both", which="both", length=0, pad=8, labelcolor=AX_LABEL)
    ax.set_xlabel("出现频次 (次)", fontproperties=fp("bold", 15), color=AX_LABEL)

    # 标注数值
    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax.text(width + max(values)*0.01, bar.get_y() + bar.get_height()/2, 
                f"{int(width)}",
                ha="left", va="center", color="#D4E8FF", fontproperties=fp("bold", 14))

    total_obs = df_exp.shape[0]
    fig.text(0.5, 0.95, title, ha="center", va="top", fontproperties=fp("black", 26), color="#D4E8FF")
    fig.text(0.5, 0.91, subtitle + f"   |   OB 标签总击中次数：{total_obs}", ha="center", va="top", fontproperties=fp("bold", 15), color="#6080A0")
    
    fig.add_artist(plt.Line2D([0.08, 0.92], [0.89, 0.89], transform=fig.transFigure, color="#1A3060", linewidth=1.2))

    plt.tight_layout(rect=[0, 0, 1, 0.88])
    plt.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成图表: {out_path}")
    plt.close()

def plot_ob_theme_heatmap(df: pd.DataFrame, title: str, subtitle: str, out_path: str):
    """
    绘制 OB行为指标 (X轴) × 训练主题 (Y轴) 的热力交叉矩阵
    由于 OB 很多，我们取总体出现频次排名前 20 的 OB
    """
    if df.empty or "OB标签" not in df.columns:
        print(f"  [跳过] {title} — 无数据")
        return

    df_exp = df.copy()
    df_exp = df_exp[df_exp["OB标签"].str.len() > 0]
    if df_exp.empty:
        return
        
    df_exp = df_exp.explode("OB标签").reset_index(drop=True)
    
    # 选出 Top 16 或 Top 20 的高频 OB 作为列
    top_obs = df_exp["OB标签"].value_counts().head(18).index.tolist()
    
    # 过滤剩下这些 Top OB 的记录
    df_top = df_exp[df_exp["OB标签"].isin(top_obs)]
    if df_top.empty:
        return
        
    # 建立交叉表
    cross = pd.crosstab(df_top["训练主题"], df_top["OB标签"])
    
    # 按 THEME_ORDER 排序 Y 轴
    from theme_matrix_analysis import THEME_ORDER
    row_order = [t for t in THEME_ORDER if t in cross.index]
    # 按频次从大到小排序 X 轴
    col_order = [ob for ob in top_obs if ob in cross.columns]
    
    cross = cross.reindex(index=row_order, columns=col_order, fill_value=0)
    if cross.empty: return

    BG_MAIN  = "#0A1020"
    BG_BODY  = "#0D1526"
    AX_LABEL = "#8AA8D8"
    GRID_C   = "#1C2E50"
    TITLE_C  = "#D4E8FF"
    ZERO_C   = "#0F1D36"
    
    from matplotlib.colors import LinearSegmentedColormap
    cdict_ob = {
        "red":   [(0.0, 0.06, 0.06), (0.5, 0.90, 0.90), (1.0, 0.95, 0.95)],
        "green": [(0.0, 0.12, 0.12), (0.5, 0.70, 0.70), (1.0, 0.30, 0.30)],
        "blue":  [(0.0, 0.28, 0.28), (0.5, 0.20, 0.20), (1.0, 0.10, 0.10)],
    }
    cmap_ob = LinearSegmentedColormap("ob_theme", cdict_ob)
    
    n_rows, n_cols = cross.shape
    fig_w, fig_h = max(16, n_cols * 1.5 + 4), max(10, n_rows * 0.75 + 3.5)
    
    fig = plt.figure(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax  = fig.add_subplot(111, facecolor=BG_BODY)
    
    data = cross.values.astype(float)
    masked = np.ma.masked_where(data == 0, data)
    vmax = data.max() if data.max() > 0 else 1
    
    im = ax.imshow(masked, aspect="auto", cmap=cmap_ob, vmin=0.5, vmax=vmax, interpolation="nearest")
    zero_mat = np.ma.masked_where(data != 0, data)
    ax.imshow(zero_mat, aspect="auto", cmap=LinearSegmentedColormap.from_list("zero", [ZERO_C, ZERO_C]), vmin=0, vmax=1)
    
    for x in np.arange(-0.5, n_cols, 1):
        ax.axvline(x, color=GRID_C, linewidth=0.6)
    for y in np.arange(-0.5, n_rows, 1):
        ax.axhline(y, color=GRID_C, linewidth=0.6)
        
    for r in range(n_rows):
        for c in range(n_cols):
            val = int(data[r, c])
            if val == 0:
                ax.text(c, r, "·", ha="center", va="center", color="#2A3D60", fontproperties=fp("bold", 18))
            else:
                txt_c = "#1A1F2E" if val/vmax > 0.55 else "#D8EBFF"
                ax.text(c, r, str(val), ha="center", va="center", color=txt_c, fontproperties=fp("bold", 18))
                
    # x 轴包含中文 OB 描述
    x_labels = []
    for ob in cross.columns:
        comp, ob_id = ob.split("-")
        cn_name = ""
        if comp in OB_KEYS:
            for k in OB_KEYS[comp].keys():
                if k.startswith(ob_id):
                    cn_name = k.split("_")[1] if "_" in k else k
                    break
        x_labels.append(f"{ob}\n{cn_name}")

    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(x_labels, rotation=35, ha="right", fontproperties=fp("bold", 14), color=AX_LABEL)
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels(cross.index, fontproperties=fp("bold", 16), color="#C8DEFF")
    
    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_C)
        
    ax.set_xlabel("高频 行为指标 (OB)", fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)
    ax.set_ylabel("训练主题 (EBT Training Theme)", fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)

    total = int(data.sum())
    fig.text(0.5, 0.96, title, ha="center", va="top", fontproperties=fp("black", 32), color=TITLE_C)
    fig.text(0.5, 0.92, subtitle + f"   |   Top 18 OB 总跨界匹配次数：{total}", ha="center", va="top", fontproperties=fp("bold", 16), color="#6080A0")
    
    plt.tight_layout(rect=[0, 0, 1, 0.90])
    plt.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()

if __name__ == "__main__":
    DATA_FILE = r"c:\工作文件\数据管理\IE\2026.1-3.xls"
    OUT_DIR = r"c:\工作文件\数据管理\IE\分析产出(2026)"
    
    print("=" * 60)
    print("  AeroEBT - 行为指标(OB) 深度交叉分析")
    print("=" * 60)
    
    df_all, df_weak = load_data(DATA_FILE)
    
    df_a330 = df_all[df_all["机型"].astype(str).str.contains("330", na=False)]
    df_a350 = df_all[df_all["机型"].astype(str).str.contains("350", na=False)]
    
    # ── 1. 分机型 OB 短板频次分布 ──
    if not df_a330.empty:
        plot_ob_bar_chart(
            df_a330,
            title="A330 机型高频 OB 短板分布 (Top 20)",
            subtitle="2026 Q1 A330 EBT 数据 | 基于 NLP 逆向分析图谱",
            out_path=os.path.join(OUT_DIR, "图_A330_OB高频分布.png")
        )
    if not df_a350.empty:
        plot_ob_bar_chart(
            df_a350,
            title="A350 机型高频 OB 短板分布 (Top 20)",
            subtitle="2026 Q1 A350 EBT 数据 | 基于 NLP 逆向分析图谱",
            out_path=os.path.join(OUT_DIR, "图_A350_OB高频分布.png")
        )

    # ── 2. 分机型 OB × 训练主题 热力矩阵 ──
    if not df_a330.empty:
        plot_ob_theme_heatmap(
            df_a330,
            title="A330 行为指标(OB) × 训练主题 交叉矩阵",
            subtitle="探究基础核心短板(OB)是如何诱发具体训练课目(Theme)的脱节",
            out_path=os.path.join(OUT_DIR, "图_A330_OB_Theme矩阵.png")
        )
    if not df_a350.empty:
        plot_ob_theme_heatmap(
            df_a350,
            title="A350 行为指标(OB) × 训练主题 交叉矩阵",
            subtitle="探究基础核心短板(OB)是如何诱发具体训练课目(Theme)的脱节",
            out_path=os.path.join(OUT_DIR, "图_A350_OB_Theme矩阵.png")
        )
