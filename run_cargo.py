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
    # 低分（<3分）且有评语记录——与A330/A350保持一致
    weak_threshold = cfg.get("WEAK_THRESHOLD", 3.0)
    weak_desc = all_desc[all_desc["得分"] < weak_threshold].copy()
    print(f"       有评语记录: {len(all_desc)} 条  |  低分(<{weak_threshold}分): {len(weak_desc)} 条")

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

    # ── 图A：胜任力 × 训练主题  ·  小于阈値（小于weak_threshold） ──────────
    print(f"[图A] 胜任力 × 训练主题  热力矩阵  · < {weak_threshold}分人员...")
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
    print(f"[图E] 胜任力 × 核心风险  热力矩阵  · < {weak_threshold}分...")
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
    from theme_matrix_analysis import RISK_NAMES, expand_to_risk

    total      = len(df)
    n_all_desc = len(all_desc)
    n_weak     = len(weak_desc)
    code_cn    = {code: name for name, code in cfg["COMPETENCIES"]}
    comp_order = [c for _, c in cfg["COMPETENCIES"]]

    # ── 基础统计 ─────────────────────────────────────────────────
    # 胜任力：全体命中 vs 低分命中
    comp_all  = df_themed_all["胜任力"].value_counts()
    comp_weak = df_themed_weak["胜任力"].value_counts() if not df_themed_weak.empty else pd.Series(dtype=int)

    # 训练主题：全体 vs 低分
    theme_all  = df_themed_all["训练主题"].value_counts()
    theme_weak = df_themed_weak["训练主题"].value_counts() if not df_themed_weak.empty else pd.Series(dtype=int)

    n_all_hit  = len(df_themed_all)
    n_weak_hit = len(df_themed_weak)

    # 核心风险
    df_risk_all  = expand_to_risk(df_themed_all)
    df_risk_weak = expand_to_risk(df_themed_weak) if not df_themed_weak.empty else pd.DataFrame()
    risk_all  = df_risk_all["核心风险"].value_counts()  if not df_risk_all.empty  else pd.Series(dtype=int)
    risk_weak = df_risk_weak["核心风险"].value_counts() if not df_risk_weak.empty else pd.Series(dtype=int)

    # ── 热点格（胜任力 × 训练主题 交叉次数）────────────────────────
    cross_all  = pd.crosstab(df_themed_all["胜任力"],  df_themed_all["训练主题"])
    cross_weak = pd.crosstab(df_themed_weak["胜任力"], df_themed_weak["训练主题"]) \
                 if not df_themed_weak.empty else pd.DataFrame()

    # 全体 TOP10 热点格
    all_hotspot = (cross_all.stack()
                   .reset_index()
                   .rename(columns={"胜任力": "code", "训练主题": "theme", 0: "cnt"})
                   .sort_values("cnt", ascending=False)
                   .head(10))

    # 低分 TOP10 高危格
    weak_hotspot = (cross_weak.stack()
                    .reset_index()
                    .rename(columns={"胜任力": "code", "训练主题": "theme", 0: "cnt"})
                    .sort_values("cnt", ascending=False)
                    .head(10)) if not cross_weak.empty else pd.DataFrame()

    # ── 薄弱场景低分率（至少出现过2次全体命中的主题）──────────────
    theme_rate = []
    for t in theme_all.index:
        a = int(theme_all.get(t, 0))
        w = int(theme_weak.get(t, 0)) if t in theme_weak else 0
        if a >= 2:
            theme_rate.append((t, a, w, w / a * 100))
    theme_rate.sort(key=lambda x: -x[3])

    # ── 低分率 TOP3 场景（用于训练建议自动生成）────────────────────
    top_rate     = [x for x in theme_rate if x[2] > 0][:3]
    top_weak_hot = weak_hotspot.head(3).to_dict("records") if not weak_hotspot.empty else []

    def _rate_str(a, w):
        return f"{w/a*100:.1f}%" if a > 0 else "—"

    def _comp_rows():
        lines = []
        for code in comp_order:
            cn  = code_cn.get(code, code)
            a   = int(comp_all.get(code, 0))
            w   = int(comp_weak.get(code, 0)) if code in comp_weak else 0
            r   = _rate_str(a, w) if a > 0 else "—"
            bold_a = f"**{a}**" if a == int(comp_all.max()) else str(a)
            bold_w = f"**{w}**" if w > 0 and w == int(comp_weak.max()) else str(w)
            lines.append(f"| **{code}** {cn} | {bold_a} | {bold_w} | {r} |")
        return "\n".join(lines)

    def _theme_top10():
        lines = []
        for i, (t, cnt) in enumerate(theme_all.head(10).items(), 1):
            w = int(theme_weak.get(t, 0)) if t in theme_weak else 0
            r = _rate_str(cnt, w)
            w_str = str(w) if w > 0 else "—"
            lines.append(f"| {i} | {t} | {cnt} | {w_str} | {r} |")
        return "\n".join(lines)

    def _all_hotspot_rows():
        lines = []
        for _, row in all_hotspot.iterrows():
            cn = code_cn.get(row["code"], row["code"])
            lines.append(f"| **{row['code']}** | {row['theme']} | **{int(row['cnt'])}** | — |")
        return "\n".join(lines)

    def _weak_hotspot_rows():
        if weak_hotspot.empty:
            return "| — | — | — |"
        lines = []
        for _, row in weak_hotspot.iterrows():
            lines.append(f"| **{row['code']}** | {row['theme']} | **{int(row['cnt'])}** |")
        return "\n".join(lines)

    def _rate_rank_rows():
        lines = []
        for i, (t, a, w, r) in enumerate(theme_rate[:12], 1):
            warn = " ⚠️" if r >= 15 else ""
            lines.append(f"| {i} | **{t}**{warn} | {r:.1f}% |")
        return "\n".join(lines)

    def _risk_rows(risk_ser):
        if risk_ser.empty:
            return "| — | — | — |"
        total_r = risk_ser.sum()
        lines = []
        for code, cnt in risk_ser.items():
            cn = RISK_NAMES.get(code, code)
            lines.append(f"| **{code}**（{cn}） | {cnt} | {cnt/total_r*100:.1f}% |")
        return "\n".join(lines)

    # ── 报告正文 ──────────────────────────────────────────────────
    report = f"""# {period} B777货机 EBT 胜任力训练分析报告

> **数据来源**：cargo1-3.xlsx | {period} B777货机 EBT训练记录
>
> **分析方法**：基于评语关键词归类，无匹配项不统计
>
> **数据记录**：{n_all_hit} 条全体命中 | **低分（< 3分）命中记录**：{n_weak_hit} 条
>
> 生成时间：{pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}

---

## 一、总体态势

### 1.1 胜任力负担分布（全体 vs 低分）

| 胜任力 | 全体命中 | 低分命中 | 低分占比 |
| --- | --- | --- | --- |
{_comp_rows()}

> **关键发现**：命中次数最多的两个胜任力为
> **{comp_order[comp_all.reindex(comp_order, fill_value=0).values.argmax()]}（{code_cn.get(comp_order[comp_all.reindex(comp_order, fill_value=0).values.argmax()],'—')}）**
> 和 **{comp_weak.index[0]}（{code_cn.get(comp_weak.index[0],'—')}）**（低分绝对值最高）。

---

## 二、高频训练场景分析（全体人员）

### 2.1 出现频次 TOP 10 训练主题

| 排名 | 训练主题 | 全体次数 | 低分次数 | 低分率 |
| --- | --- | --- | --- | --- |
{_theme_top10()}

> 低分率 = 低分（< 3分）次数 / 全体次数，反映该场景转化为真实能力瓶颈的概率。

---

## 三、高风险交叉点（热力矩阵解读）

### 3.1 全体 TOP 10 热点格

| 胜任力 | 训练主题 | 次数 | 解读 |
| --- | --- | --- | --- |
{_all_hotspot_rows()}

> 图B（全体训练主题矩阵）可视化上述热点分布。

![图B](图B_训练主题矩阵.png)

### 3.2 低分 TOP 10 高危格（最需强化训练的交叉区域）

| 胜任力 | 训练主题 | 低分次数 |
| --- | --- | --- |
{_weak_hotspot_rows()}

> 图A（低分预警矩阵）聚焦于得分低于3分的记录，揭示真实能力缺口所在。

![图A](图A_低分人员_训练主题矩阵.png)

---

## 四、薄弱场景低分率排名（高风险场景）

> 低分率 = 低分（< 3分）次数 / 全体次数，反映该场景转化为真实能力瓶颈的概率

| 排名 | 训练主题 | 低分率 |
| --- | --- | --- |
{_rate_rank_rows()}

---

## 五、核心风险分析

### 5.1 全体核心风险命中分布

| 核心风险 | 命中次数 | 占比 |
| --- | --- | --- |
{_risk_rows(risk_all)}

### 5.2 低分核心风险命中分布

| 核心风险 | 命中次数 | 占比 |
| --- | --- | --- |
{_risk_rows(risk_weak)}

> **{risk_all.index[0] if len(risk_all) > 0 else '—'}（{RISK_NAMES.get(risk_all.index[0],'—') if len(risk_all) > 0 else '—'}）**
> 是全员命中最多的核心风险，全员命中 **{int(risk_all.iloc[0]) if len(risk_all) > 0 else 0}次**。
>
> 低分矩阵中 **{f"{weak_hotspot.iloc[0]['code']}×{weak_hotspot.iloc[0]['theme']}={int(weak_hotspot.iloc[0]['cnt'])}" if not weak_hotspot.empty else '—'}** 最高，说明低分人员在该场景的能力缺口最为突出。

![图C](图C_全员_桑基图.png)

![图D](图D_全员_胜任力×核心风险矩阵.png)

![图E](图E_低分_胜任力×核心风险矩阵.png)

---

## 六、训练改进建议

### 🔴 第一优先级：立即强化（低分次数最多）

{"".join([f'''
### {i+1}. {r["code"]} × {r["theme"]}（低分{int(r["cnt"])}次）

- **问题**：该交叉维度低分次数居前，是最需深度干预的能力薄弱点
- **改进**：针对 {r["theme"]} 场景强化 {code_cn.get(r["code"],r["code"])} 维度专项训练
''' for i, r in enumerate(top_weak_hot)])}

---

### 🟡 第二优先级：重点关注（低分率高）

{"".join([f'''
### {i+1+len(top_weak_hot)}. {t}（低分率 {r:.1f}%）

- **改进**：该场景虽频次有限，但低分转化率高，建议在定期 EBT 训练中增设专项考核节点
''' for i, (t, a, w, r) in enumerate(top_rate)])}

---

## 七、总结

| 类型 | 核心问题 |
| --- | --- |
| **第一胜任力** | {comp_weak.index[0] if len(comp_weak) > 0 else '—'}（{code_cn.get(comp_weak.index[0],'—') if len(comp_weak) > 0 else '—'}）—低分绝对值最高 |
| **第一训练场景** | {theme_weak.index[0] if len(theme_weak) > 0 else '—'}（低分{int(theme_weak.iloc[0]) if len(theme_weak) > 0 else 0}次） |
| **最需关注的交叉点** | {f"{top_weak_hot[0]['code']}×{top_weak_hot[0]['theme']}" if top_weak_hot else '—'} |
| **隐性高风险** | {top_rate[0][0] if top_rate else '—'}（低分率{f"{top_rate[0][3]:.1f}%" if top_rate else '—'}，易被忽视） |

> 建议将上述分析结论纳入本季度训练讲评会核心议题，并在下一季度 EBT
> 科目编排中对第一、二优先级专项进行强化覆盖，季度末再次对比矩阵变化趋势。

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

