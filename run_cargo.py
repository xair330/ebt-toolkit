"""
B777 货机 EBT 分析入口
产出：
  图A  胜任力 × 训练主题  热力矩阵  · 低分(<3分)
  图B  胜任力 × 训练主题  热力矩阵  · 全员
  图C  训练主题 → 核心风险  桑基图   · 全员
  图D  胜任力 × 核心风险  热力矩阵  · 全员
  图E  胜任力 × 核心风险  热力矩阵  · 低分(<3分)
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
    THEME_KEYS,
    plot_heatmap,
    plot_risk_heatmap,
    plot_sankey,
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
    对评语记录进行训练主题关键词匹配，返回含 [胜任力, 训练主题, ...] 长表。
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

    print("=" * 58)
    print("      B777 货机 EBT 矩阵分析引擎")
    print("=" * 58)
    print(f"[*] 数据源: {DATA_FILE}")
    print(f"[*] 期次:   {PERIOD}")
    print(f"[*] 产出:   {OUTPUT_DIR}\n")

    # ── 数据加载 ──────────────────────────────────────────────────
    print("[准备] 数据加载与清洗...")
    df = load_and_clean(DATA_FILE)
    print(f"       共 {len(df)} 条记录")

    # 全员有评语记录
    all_desc  = extract_all_with_desc(df)
    # 低分（<3分）且有评语记录
    weak_desc = all_desc[all_desc["得分"] < 3.0].copy()
    print(f"       有评语记录: {len(all_desc)} 条  |  低分(<3分): {len(weak_desc)} 条")

    # 训练主题匹配
    df_themed_all  = build_themed_df(all_desc)
    df_themed_weak = build_themed_df(weak_desc)
    print(f"       训练主题命中: 全员 {len(df_themed_all)} 条  |  低分 {len(df_themed_weak)} 条\n")

    if df_themed_all.empty:
        print("[警告] 全员无记录命中任何训练主题，请检查 THEME_KEYS 配置。")
        pause_and_exit()

    # 保存明细 CSV
    all_desc.to_csv(os.path.join(OUTPUT_DIR, "评语明细_全员.csv"), index=False, encoding="utf-8-sig")
    weak_desc.to_csv(os.path.join(OUTPUT_DIR, "评语明细_低分.csv"), index=False, encoding="utf-8-sig")

    # ── 图A：胜任力 × 训练主题  ·  低分 ──────────────────────────
    print("[图A] 胜任力 × 训练主题  热力矩阵  · 低分(<3分)...")
    if not df_themed_weak.empty:
        plot_heatmap(
            df_themed_weak,
            title    = "胜任力 × 训练主题  热力矩阵  ·  低分预警（< 3分）",
            subtitle = f"{PERIOD} B777货机 | 得分低于3分的记录按评语关键词归类",
            out_path = os.path.join(OUTPUT_DIR, "图A_低分人员_训练主题矩阵.png"),
        )
    else:
        print("  [跳过] 低分记录为空")

    # ── 图B：胜任力 × 训练主题  ·  全员 ──────────────────────────
    print("[图B] 胜任力 × 训练主题  热力矩阵  · 全员...")
    plot_heatmap(
        df_themed_all,
        title    = "胜任力 × 训练主题  热力矩阵  ·  全员",
        subtitle = f"{PERIOD} B777货机 | 全量有评语记录按关键词归类",
        out_path = os.path.join(OUTPUT_DIR, "图B_训练主题矩阵.png"),
    )

    # ── 图C：训练主题 → 核心风险  桑基图  ·  全员 ─────────────────
    print("[图C] 训练主题 → 核心风险  桑基图  · 全员...")
    plot_sankey(
        df_themed_all,
        title    = "训练主题 → 核心风险  安全流向图  ·  全员",
        subtitle = f"{PERIOD} B777货机 | 训练主题通过 RISK_MAP 映射至对应核心风险",
        out_path = os.path.join(OUTPUT_DIR, "图C_全员_桑基图.png"),
    )

    # ── 图D：胜任力 × 核心风险  热力矩阵  ·  全员 ────────────────
    print("[图D] 胜任力 × 核心风险  热力矩阵  · 全员...")
    plot_risk_heatmap(
        df_themed_all,
        title    = "胜任力 × 核心风险  热力矩阵  ·  全员",
        subtitle = f"{PERIOD} B777货机 | 核心风险按胜任力维度聚合",
        out_path = os.path.join(OUTPUT_DIR, "图D_全员_胜任力×核心风险矩阵.png"),
    )

    # ── 图E：胜任力 × 核心风险  热力矩阵  ·  低分 ────────────────
    print("[图E] 胜任力 × 核心风险  热力矩阵  · 低分(<3分)...")
    if not df_themed_weak.empty:
        plot_risk_heatmap(
            df_themed_weak,
            title    = "胜任力 × 核心风险  热力矩阵  ·  低分预警（< 3分）",
            subtitle = f"{PERIOD} B777货机 | 仅统计得分低于3分的记录",
            out_path = os.path.join(OUTPUT_DIR, "图E_低分_胜任力×核心风险矩阵.png"),
        )
    else:
        print("  [跳过] 低分记录为空")

    print("\n" + "=" * 58)
    print("全部分析完成！产出已保存至：")
    print(f" -> {OUTPUT_DIR}")
    print("=" * 58)
    pause_and_exit()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n[严重错误] {e}")
        traceback.print_exc()
        pause_and_exit()
