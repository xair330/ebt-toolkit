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

    # ── 生成完整分析报告 ──────────────────────────────────────────
    print("[报告] 生成分析报告...")
    generate_report(
        df=df, all_desc=all_desc, weak_desc=weak_desc,
        df_themed_all=df_themed_all, df_themed_weak=df_themed_weak,
        period=PERIOD, output_dir=OUTPUT_DIR,
    )

    print("\n" + "=" * 58)
    print("全部分析完成！产出已保存至：")
    print(f" -> {OUTPUT_DIR}")
    print("=" * 58)
    pause_and_exit()




# ══════════════════════════════════════════════════════════════════
# 报告生成
# ══════════════════════════════════════════════════════════════════
def generate_report(df, all_desc, weak_desc,
                    df_themed_all, df_themed_weak,
                    period, output_dir):
    from theme_matrix_analysis import RISK_NAMES, expand_to_risk, CODE_TO_NAME

    total      = len(df)
    n_all_desc = len(all_desc)
    n_weak     = len(weak_desc)

    # 各胜任力低分技计
    comp_weak_cnt = weak_desc["胜任力"].value_counts()
    comp_all_cnt  = all_desc["胜任力"].value_counts()

    # 训练主题 TOP10（全员）
    theme_cnt_all  = df_themed_all["训练主题"].value_counts()
    theme_cnt_weak = df_themed_weak["训练主题"].value_counts() if not df_themed_weak.empty else theme_cnt_all * 0

    # 核心风险分布（全员）
    df_risk_all  = expand_to_risk(df_themed_all)
    df_risk_weak = expand_to_risk(df_themed_weak) if not df_themed_weak.empty else pd.DataFrame()
    risk_cnt_all  = df_risk_all["核心风险"].value_counts()  if not df_risk_all.empty  else pd.Series()
    risk_cnt_weak = df_risk_weak["核心风险"].value_counts() if not df_risk_weak.empty else pd.Series()

    # 角色分布（全员评语）
    role_cnt = all_desc.groupby(["角色","胜任力"]).size().unstack(fill_value=0)

    # ── 胜任力中文名映射（兼容cfg和CODE_TO_NAME）──
    comp_order = [c for _, c in cfg["COMPETENCIES"]]
    code_cn = {code: name for name, code in cfg["COMPETENCIES"]}

    def fmt_table(series, top=None, pct=True):
        """将 Series 转为 Markdown 表格行"""
        total_n = series.sum()
        rows = series.head(top) if top else series
        lines = []
        for idx, val in rows.items():
            p = f"{val/total_n*100:.1f}%" if pct and total_n else ""
            cn = RISK_NAMES.get(idx, idx)
            name = f"{idx}（{cn}）" if cn != idx else str(idx)
            lines.append(f"| {name} | {val} | {p} |")
        return "\n".join(lines)

    def fmt_comp_table(series):
        lines = []
        for code in comp_order:
            val = series.get(code, 0)
            cn  = code_cn.get(code, code)
            lines.append(f"| {code} | {cn} | {val} |")
        return "\n".join(lines)

    # ── 报告正文 ──────────────────────────────────────────────────
    report = f"""# {period} B777货机 EBT胜任力安全分析报告

> 数据来源：`cargo1-3.xlsx` · 分析工具：AeroEBT Toolkit  
> 生成时间：{pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}

---

## 一、数据概览

| 指标 | 数值 |
|------|------|
| 总检查记录数 | {total} 条 |
| 有效评语记录 | {n_all_desc} 条（占总记录 {n_all_desc/total*100:.1f}%）|
| 低分记录（< 3分） | {n_weak} 条（占有评语记录 {n_weak/n_all_desc*100:.1f}%）|
| 训练主题命中（全员） | {len(df_themed_all)} 条 |
| 训练主题命中（低分） | {len(df_themed_weak)} 条 |
| 涉及训练主题数（全员） | {df_themed_all['训练主题'].nunique()} 个 |
| 涉及核心风险类别（全员）| {df_risk_all['核心风险'].nunique() if not df_risk_all.empty else 0} 类 |

### 角色分布（有评语记录）

| 角色 | 记录数 | 占比 |
|------|--------|------|
{chr(10).join(f"| {r} | {v} | {v/n_all_desc*100:.1f}% |" for r,v in all_desc['角色'].value_counts().items())}

---

## 二、胜任力维度分析

### 2.1 各胜任力有评语记录数

| 代码 | 胜任力 | 全员有评语 |
|------|--------|----------|
{fmt_comp_table(comp_all_cnt)}

### 2.2 各胜任力低分(<3分)记录数

| 代码 | 胜任力 | 低分记录 |
|------|--------|---------|
{fmt_comp_table(comp_weak_cnt)}

> **关注重点**：低分记录最多的胜任力依次为 **{", ".join(f"{code_cn.get(c,c)}({c})" for c in comp_weak_cnt.head(3).index)}**

---

## 三、训练主题分析

### 3.1 训练主题频次 TOP10（全员有评语）

| 训练主题 | 命中次数 | 占比 |
|----------|---------|------|
{fmt_table(theme_cnt_all, top=10)}

### 3.2 训练主题频次 TOP10（低分 < 3分）

| 训练主题 | 命中次数 | 占比 |
|----------|---------|------|
{fmt_table(theme_cnt_weak, top=10) if not theme_cnt_weak.empty else "| — | — | — |"}

### 3.3 图表说明

**图A** — 低分人员训练主题矩阵

> 仅统计各胜任力得分低于3分的记录，直观呈现不达标区间的训练缺口分布。

![图A](图A_低分人员_训练主题矩阵.png)

**图B** — 全员训练主题矩阵

> 覆盖所有有不足评语的记录，反映整体训练重点分布。

![图B](图B_训练主题矩阵.png)

---

## 四、核心风险分析

### 4.1 核心风险频次（全员）

| 核心风险 | 命中次数 | 占比 |
|----------|---------|------|
{fmt_table(risk_cnt_all)}

### 4.2 核心风险频次（低分 < 3分）

| 核心风险 | 命中次数 | 占比 |
|----------|---------|------|
{fmt_table(risk_cnt_weak) if not risk_cnt_weak.empty else "| — | — | — |"}

> **高风险区域**：全员数据中命中最高的核心风险依次为  
> **{", ".join(f"{RISK_NAMES.get(r,r)}({r})" for r in risk_cnt_all.head(3).index) if not risk_cnt_all.empty else "—"}**

### 4.3 图表说明

**图C** — 训练主题→核心风险桑基图（全员）

> 直观展示评语中训练主题对应的安全风险流向，流线粗细表示频次。

![图C](图C_全员_桑基图.png)

**图D** — 胜任力×核心风险热力矩阵（全员）

![图D](图D_全员_胜任力×核心风险矩阵.png)

**图E** — 胜任力×核心风险热力矩阵（低分）

![图E](图E_低分_胜任力×核心风险矩阵.png)

---

## 五、分析结论与训练建议

### 5.1 主要发现

1. **低分集中领域**：低分记录最多的胜任力为 **{", ".join(f"{code_cn.get(c,c)}({c})" for c in comp_weak_cnt.head(3).index)}**，需重点关注。
2. **高频训练主题**：全员评语中出现最频繁的训练主题为 **{", ".join(f'"{t}"' for t in theme_cnt_all.head(3).index)}**。
3. **核心风险关联**：数据所反映的最主要核心风险为 **{", ".join(f"{RISK_NAMES.get(r,r)}({r})" for r in risk_cnt_all.head(3).index) if not risk_cnt_all.empty else "—"}**。

### 5.2 训练建议

| 优先级 | 训练主题 | 关联胜任力 | 关联核心风险 |
|--------|----------|----------|------------|
| 高 | {theme_cnt_all.index[0] if len(theme_cnt_all) > 0 else "—"} | — | — |
| 高 | {theme_cnt_all.index[1] if len(theme_cnt_all) > 1 else "—"} | — | — |
| 高 | {theme_cnt_all.index[2] if len(theme_cnt_all) > 2 else "—"} | — | — |

> ⚠️ 训练建议优先级列仅根据频次自动排序，最终排期需结合飞行运行实际风险评估。

---

*本报告由 AeroEBT Toolkit 自动生成 · {period}*
"""

    report_path = os.path.join(output_dir, f"分析报告_{period.replace(' ', '_')}.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"  [OK] 已生成: {report_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n[严重错误] {e}")
        traceback.print_exc()
        pause_and_exit()

