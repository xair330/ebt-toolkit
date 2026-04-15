"""
B777 货机 EBT 分析入口
产出：
  1. 胜任力 × 训练主题  热力矩阵
  2. 训练主题 × 核心风险 热力矩阵
使用独立 cargo_config.json，不影响 A330/A350 的 ebt_config.json
"""
import os
import sys
import pandas as pd

# ── 必须在所有本地模块 import 之前设置 ──────────────────────────────
_here = os.path.dirname(os.path.abspath(__file__))
os.environ["EBT_CONFIG"] = os.path.join(_here, "cargo_config.json")
# ────────────────────────────────────────────────────────────────────

from config_mgr import cfg
from data_loader import load_and_clean
from theme_matrix_analysis import (
    THEME_KEYS, plot_heatmap, plot_theme_risk_heatmap
)


def pause_and_exit():
    input("\n分析完成，按回车键退出...")
    sys.exit(0)


def extract_all_with_desc(df: pd.DataFrame) -> pd.DataFrame:
    """
    提取所有有不足文字评语的记录（不论得分高低）。
    适用于单列模式——货机数据大量3分记录也有有价值的不足评语。
    """
    rows = []
    for _, row in df.iterrows():
        for _, code in cfg["COMPETENCIES"]:
            desc = str(row.get(f"{code}_不足", "")).strip()
            if not desc or desc in ("nan", "无", "-", ""):
                continue
            rows.append({
                "技术等级": str(row["技术等级"]),
                "机型":    row["机型"],
                "角色":    row["角色"],
                "等级":    row["等级"],
                "胜任力":  code,
                "得分":    row.get(f"{code}_得分"),
                "问题描述": desc.replace("\n", " "),
            })
    return pd.DataFrame(rows)


def build_themed_df(all_desc_df: pd.DataFrame) -> pd.DataFrame:
    """
    对低分项的问题描述进行训练主题关键词匹配，
    返回含 [胜任力, 训练主题, 得分, 角色, 机型, 等级] 列的长表。
    一条记录可命中多个训练主题（展开）。
    """
    rows = []
    for _, row in all_desc_df.iterrows():
        desc = str(row.get("问题描述", ""))
        matched = [t for t, kws in THEME_KEYS.items()
                   if any(kw in desc for kw in kws)]
        if not matched:
            continue
        for theme in matched:
            rows.append({
                "胜任力":   row["胜任力"],
                "训练主题": theme,
                "得分":     row["得分"],
                "角色":     row["角色"],
                "机型":     row["机型"],
                "等级":     row["等级"],
            })
    return pd.DataFrame(rows)


def main():
    # 支持拖拽 Excel 文件
    if len(sys.argv) > 1:
        dropped = sys.argv[1]
        if dropped.endswith(('.xls', '.xlsx')):
            cfg["DATA_FILE"] = dropped
            print(f"[提示] 检测到拖拽文件: {dropped}")

    DATA_FILE = cfg.get("DATA_FILE", "")
    if not os.path.exists(DATA_FILE):
        print(f"[错误] 找不到数据文件: {DATA_FILE}")
        print("请在 cargo_config.json 中配置正确的 DATA_FILE 路径。")
        pause_and_exit()

    OUTPUT_DIR = os.path.join(os.path.dirname(DATA_FILE), "分析产出(B777)")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    PERIOD = cfg.get("REPORT_PERIOD", "")

    print("=" * 56)
    print("      B777 货机 EBT 矩阵分析引擎")
    print("=" * 56)
    print(f"[*] 数据源: {DATA_FILE}")
    print(f"[*] 期次:   {PERIOD}")
    print(f"[*] 产出:   {OUTPUT_DIR}\n")

    # ── 数据加载 ──────────────────────────────────────────────────
    print("[1/3] 数据加载与清洗...")
    df = load_and_clean(DATA_FILE)
    all_desc_df = extract_all_with_desc(df)
    print(f"      共 {len(df)} 条记录，有评语记录 {len(all_desc_df)} 条")

    # ── 匹配训练主题 ──────────────────────────────────────────────
    print("[2/3] 匹配训练主题关键词...")
    df_themed = build_themed_df(all_desc_df)
    print(f"      命中训练主题记录: {len(df_themed)} 条")

    if df_themed.empty:
        print("[警告] 无记录命中任何训练主题，请检查 THEME_KEYS 关键词配置。")
        pause_and_exit()

    # 保存明细 CSV
    all_desc_df.to_csv(
        os.path.join(OUTPUT_DIR, "评语明细.csv"),
        index=False, encoding="utf-8-sig"
    )
    df_themed.to_csv(
        os.path.join(OUTPUT_DIR, "训练主题命中明细.csv"),
        index=False, encoding="utf-8-sig"
    )

    # ── 生成两个热力矩阵 ──────────────────────────────────────────
    print("[3/3] 生成热力矩阵...")

    # 矩阵1：胜任力 × 训练主题
    plot_heatmap(
        df_themed,
        title    = f"胜任力 × 训练主题  热力矩阵  ·  B777货机",
        subtitle = f"{PERIOD} | 全量有评语记录按关键词归类",
        out_path = os.path.join(OUTPUT_DIR, "矩阵1_胜任力×训练主题.png"),
    )

    # 矩阵2：训练主题 × 核心风险
    plot_theme_risk_heatmap(
        df_themed,
        title    = f"训练主题 × 核心风险  热力矩阵  ·  B777货机",
        subtitle = f"{PERIOD} | 训练主题通过 RISK_MAP 映射至对应核心风险",
        out_path = os.path.join(OUTPUT_DIR, "矩阵2_训练主题×核心风险.png"),
    )

    print("\n" + "=" * 56)
    print("全部分析完成！产出已保存至：")
    print(f" -> {OUTPUT_DIR}")
    print("=" * 56)
    pause_and_exit()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n[严重错误] {e}")
        traceback.print_exc()
        pause_and_exit()
