import os
import sys
import pandas as pd

# 将当前目录加入 path 以便引入
sys.path.insert(0, r'c:\工作文件\数据管理\IE\ebt_analysis_toolkit')

from theme_matrix_analysis import load_data, expand_to_risk, RISK_NAMES, OB_KEYS
from config_mgr import cfg

def get_ob_cn_name(ob_id_full):
    """ FPM-OB1 -> FPM-OB1 准确的官方中文描述 """
    if "-" not in ob_id_full:
        return ob_id_full
    comp, ob = ob_id_full.split("-")
    if comp in OB_KEYS:
        for k in OB_KEYS[comp].keys():
            if k.startswith(ob):
                # 如果有 "_"，取后半段，这就是正式库的文本
                return k.split("_")[1] if "_" in k else k
    return ob_id_full

def generate_fleet_report(df_fleet: pd.DataFrame, fleet_name: str, out_dir: str):
    if df_fleet.empty:
        print(f"  [跳过] {fleet_name} 无数据。")
        return

    total_records = len(df_fleet)
    
    # ── 1. 胜任力缺陷分布 ──
    comp_all = df_fleet["胜任力"].value_counts()
    
    # ── 2. 训练主题短板 ──
    theme_all = df_fleet["训练主题"].value_counts().head(10)
    
    # ── 3. 核心 OB 暴露分析 ──
    df_ob = df_fleet.copy()[df_fleet["OB标签"].str.len() > 0]
    ob_cnt = None
    if not df_ob.empty:
        df_ob = df_ob.explode("OB标签").reset_index(drop=True)
        ob_cnt = df_ob["OB标签"].value_counts().head(10)
        
    # ── 4. 底层流向风险 ──
    df_risk = expand_to_risk(df_fleet)
    risk_cnt = df_risk["核心风险"].value_counts().head(10) if not df_risk.empty else pd.Series()

    # 组装表格文本
    def _comp_rows():
        comps = ["FPA", "FPM", "COM", "LTW", "SAW", "WLM", "PSD", "PRO", "KNO"]
        code_cn = {code: name for name, code in cfg["COMPETENCIES"]}
        lines = []
        for code in comps:
            a = int(comp_all.get(code, 0))
            if a > 0:
                lines.append(f"| **{code}** ({code_cn.get(code, code)}) | {a} | {a/total_records*100:.1f}% |")
        return "\n".join(lines) if lines else "| — | — | — |"

    def _theme_rows():
        lines = []
        for i, (t, cnt) in enumerate(theme_all.items(), 1):
            lines.append(f"| {i} | {t} | {cnt} | {cnt/total_records*100:.1f}% |")
        return "\n".join(lines) if lines else "| — | — | — | — |"

    def _ob_rows():
        if ob_cnt is None or ob_cnt.empty:
            return "| — | — | — | — |"
        lines = []
        total_ob_hits = len(df_ob)
        for i, (ob, cnt) in enumerate(ob_cnt.items(), 1):
            cn = get_ob_cn_name(ob)
            lines.append(f"| {i} | **{ob}** | {cn} | {cnt} | {cnt/total_ob_hits*100:.1f}% |")
        return "\n".join(lines)

    def _risk_rows():
        if risk_cnt.empty:
            return "| — | — | — | — |"
        lines = []
        total_r = len(df_risk)
        for i, (idx, cnt) in enumerate(risk_cnt.items(), 1):
            cn = RISK_NAMES.get(idx, idx)
            lines.append(f"| {i} | **{idx}** | {cn} | {cnt} | {cnt/total_r*100:.1f}% |")
        return "\n".join(lines)
        
    # 抽取核心洞察用于建议
    top_theme = theme_all.index[0] if len(theme_all) > 0 else "N/A"
    top_ob = ob_cnt.index[0] if ob_cnt is not None and not ob_cnt.empty else "N/A"
    top_ob_cn = get_ob_cn_name(top_ob) if top_ob != "N/A" else "N/A"
    top_risk = risk_cnt.index[0] if not risk_cnt.empty else "N/A"
    top_risk_cn = RISK_NAMES.get(top_risk, "N/A")

    report = f"""# 2026 Q1 {fleet_name}机队 EBT运行数据分析报告

> **目标受众**：{fleet_name} 机队分管干部与飞行教员
> **数据范围**：{fleet_name} 机型全体 Q1 EBT 检查评分与评语记录
> **报告生成时间**：{pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}

---

## 本期执行摘要面板 (Executive Summary)

经过后端自然语言引擎的深度解析，本季度 {fleet_name} 机队的训练暴露呈现出以下核心焦点：

- 🔴 **最高发短板课目（宏观表象）**：`{top_theme}`
- 🟠 **最核心行为指标缺陷（底层诱因）**：`{top_ob} ({top_ob_cn})` (跨科目呈现)
- 🟡 **最终主导的运行安全隐患（定性红线）**：`{top_risk_cn}`

---

## 一、 顶层胜任力陷落分布

| 胜任力维 | 中文定义 | 缺陷提及频次 | 占比 |
| --- | --- | --- | --- |
{_comp_rows()}

---

## 二、 【行为解析提取】高频 OB 行为指标缺陷排行

通过对本季度 {fleet_name} 评语库与 ALTMS 标准 `obDictionary.ts` 的自动逆向解析，剔除无效描述后，我们捕捉到了机队飞行员最高发的行为模式偏差：

| 排名 | OB编码 | ALTMS 对应官方要求 | 跨域击中频次 | OB群内占比 |
| --- | --- | --- | --- | --- |
{_ob_rows()}

⚠️ **图示参考：机队高频 OB 短板直方图**
![{fleet_name}_OB高频分布](图_{fleet_name}_OB高频分布.png)

---

## 三、 显性化：具体训练课目突破口

基于胜任力与 OB 行为的缺陷，{fleet_name} 机队在以下实体训练科目中严重受挫：

| 排名 | EBT 训练科目 | 累积异常出现趟次 | 样本池发病率 |
| --- | --- | --- | --- |
{_theme_rows()}

针对上述科目不达标现象，我们使用“OB×Theme交叉热力矩阵”探究了到底是哪个行为坏毛病导致了这个科目挂科：

⚠️ **图示参考：探究核心短板(OB)如何诱发具体训练课目(Theme)的脱节**
![{fleet_name}_OB_Theme矩阵](图_{fleet_name}_OB_Theme矩阵.png)

---

## 四、 运行隐患流向指征 (Sankey 分析)

在确认了前端操作能力偏差后，我们通过 `RISK_MAP` 将这 {total_records} 项训练缺陷强制映射到了民航核心局方运行红线风险上。

| 排名 | 风险类 | 定性全称 | 该项风险汇聚流数 | 风险端占比 |
| --- | --- | --- | --- | --- |
{_risk_rows()}

⚠️ **图示参考：{fleet_name} 安全核心风险全要素桑基网**
![{fleet_name}_桑基图](三级桑基图_{fleet_name}.png)

---

## 五、 智能培训编排建议 (AI Diagnostics)

通过以上的层层解构，建议本季度 {fleet_name} 大队在排班特训中执行以下策略：

1. **针对性专项治理**：必须立刻下发关于 `{top_ob_cn}` 的安全通告，规范教员在带飞时严抓该细节。因为它是全场最高发的隐性杀手。
2. **模拟机排课倾斜**：在日常 SOP 检查中，加考 `{top_theme}` 场景的频率。
3. **安全宣贯靶向点**：重点向学员强调任何规避 `# {top_risk_cn} #` 的基本底线，这也是近期评估中最容易被触发的安全雷区。

> **备注**：所有判定基准已与新版 ALTMS 在线字典同步，数据分析全链路纯净隔离，由分析中台静态驱动。
"""
    
    report_path = os.path.join(out_dir, f"分析报告_{fleet_name}_2026_Q1.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"  [OK] 成功生成专属报告: {report_path}")

if __name__ == "__main__":
    DATA_FILE = r"c:\工作文件\数据管理\IE\2026.1-3.xls"
    OUT_DIR = r"c:\工作文件\数据管理\IE\分析产出(2026)"
    
    print("=" * 60)
    print("  AeroEBT - A330/A350 双机型 Markdown 自动报告生成器")
    print("=" * 60)
    
    df_all, _ = load_data(DATA_FILE)
    
    df_a330 = df_all[df_all["机型"].astype(str).str.contains("330", na=False)]
    df_a350 = df_all[df_all["机型"].astype(str).str.contains("350", na=False)]
    
    generate_fleet_report(df_a330, "A330", OUT_DIR)
    generate_fleet_report(df_a350, "A350", OUT_DIR)

