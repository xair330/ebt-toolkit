"""
EBT分析工具包 - 数据加载与预处理模块
"""
import pandas as pd
import re
from config_mgr import cfg


def load_and_clean(data_file: str) -> pd.DataFrame:
    """
    读取EBT检查记录 Excel 文件，重命名胜任力列，清洗数据类型。

    Returns:
        df: 已清洗的 DataFrame，胜任力列格式为 {CODE}_优点 / {CODE}_得分 / {CODE}_不足
    """
    df_raw = pd.read_excel(data_file)
    raw_cols = list(df_raw.columns)          # 保留原始列名用于后续评语检测
    new_cols = list(raw_cols)

    base_idx      = cfg.get("BASE_IDX", 19)        # 胜任力列块起始索引
    cols_per_comp = cfg.get("COLS_PER_COMP", 3)    # 每个胜任力占几列（3=优点/得分/不足，1=仅得分）

    idx = base_idx
    for ch_name, en_code in cfg["COMPETENCIES"]:
        if cols_per_comp >= 3:
            new_cols[idx]   = f"{en_code}_优点"
            new_cols[idx+1] = f"{en_code}_得分"
            new_cols[idx+2] = f"{en_code}_不足"
        else:  # cols_per_comp == 1
            new_cols[idx] = f"{en_code}_得分"
        idx += cols_per_comp

    df_raw.columns = new_cols
    df = df_raw
    first_code = cfg["COMPETENCIES"][0][1]
    df = df[pd.to_numeric(df[f"{first_code}_得分"], errors='coerce').notna()].reset_index(drop=True)

    for _, code in cfg["COMPETENCIES"]:
        df[f"{code}_得分"] = pd.to_numeric(df[f"{code}_得分"], errors="coerce")

    # 单列模式：自动检测文字评语列，加入 {CODE}_优点 / {CODE}_不足
    if cols_per_comp == 1:
        df = _attach_text_desc(df, raw_cols, cfg["COMPETENCIES"])

    def _get_aircraft(x):
        s = str(x)
        if "330" in s: return "A330"
        if "350" in s: return "A350"
        if "777" in s: return "B777"
        return "其他"
    df["机型"] = df["技术等级"].apply(_get_aircraft)
    df["角色"] = df["技术等级"].apply(_get_role)
    df["等级"] = df["技术等级"].apply(_get_level)

    return df


def extract_weak_records(df: pd.DataFrame, threshold: float = 3.0) -> pd.DataFrame:
    """
    提取所有得分低于阈值的胜任力记录，返回长表。

    Returns:
        issues_df: 含列 技术等级/机型/角色/等级/胜任力/得分/问题描述
    """
    cols_per_comp = cfg.get("COLS_PER_COMP", 3)
    rows = []
    for _, row in df.iterrows():
        for _, code in cfg["COMPETENCIES"]:
            score = row[f"{code}_得分"]
            desc  = row.get(f"{code}_不足", "")
            if pd.isna(desc): desc = ""
            desc = str(desc).strip()
            if pd.notna(score) and score < threshold:
                # 三列模式要求有不足描述；单列模式只要低分即记录
                if cols_per_comp >= 3 and not desc:
                    continue
                rows.append({
                    "技术等级": str(row["技术等级"]),
                    "机型":    row["机型"],
                    "角色":    row["角色"],
                    "等级":    row["等级"],
                    "胜任力":  code,
                    "得分":    score,
                    "问题描述": desc.replace("\n", " ") if desc else f"{code}得分偏低",
                })
    return pd.DataFrame(rows)


def classify_issues(issues_df: pd.DataFrame, category_keys: dict) -> pd.DataFrame:
    """
    根据关键词字典为每条不足描述打上一个或多个问题大类标签，返回爆炸展开的长表。
    """
    rows = []
    for _, row in issues_df.iterrows():
        matched = [cat for cat, kws in category_keys.items()
                   if any(kw in str(row["问题描述"]) for kw in kws)]
        if not matched:
            matched = ["其他"]
        for cat in matched:
            rows.append({**row.to_dict(), "问题大类": cat})
    return pd.DataFrame(rows)


def assess_comment_quality(text: str, cause_keys: list, sol_keys: list,
                           max_len=5, len_unit=100,
                           max_ob=10, ob_each=2,
                           max_cause=5, max_sol=10) -> dict:
    """
    对单条评语文本进行质量评分。
    Returns: dict with keys 字数/OB标签数/原因分析分/建议方案分/总分(满分100)
    """
    import re
    text = str(text)
    if text in ("nan", "") or text.strip() in ("无", "-"):
        return dict(字数=0, OB标签数=0, 原因分析分=0, 建议方案分=0, 总分=0)

    length   = len(text)
    ob_count = len(re.findall(r"OB", text, re.IGNORECASE))
    cause    = sum(1 for kw in cause_keys if kw in text)
    sol      = sum(2 for kw in sol_keys   if kw in text)

    raw = (min(max_len, length / len_unit)
           + min(max_ob,    ob_count * ob_each)
           + min(max_cause, cause)
           + min(max_sol,   sol))
    total_max = max_len + max_ob + max_cause + max_sol
    score = min(100, raw / total_max * 100)

    return dict(字数=length, OB标签数=ob_count, 原因分析分=cause,
                建议方案分=int(sol / 2), 总分=round(score, 2))


# ── 内部工具 ─────────────────────────────────────────────────────
def _attach_text_desc(df: pd.DataFrame, raw_cols: list, competencies: list) -> pd.DataFrame:
    """
    单列模式专用：扫描原始列名，找到每个胜任力对应的文字评语列（优点/不足），
    附加到 DataFrame 中（列名：{CODE}_优点 / {CODE}_不足）。
    规则：在索引>25的列中，找列名包含胜任力中文名的那一列，
          该列=优点，紧随其后一列=不足。
    """
    attached = 0
    for ch_name, en_code in competencies:
        pos = None
        for i, col in enumerate(raw_cols):
            if i > 25 and ch_name in str(col):
                pos = i
                break
        if pos is not None:
            df[f"{en_code}_优点"] = df.iloc[:, pos].astype(str).replace("nan", "")
            df[f"{en_code}_不足"] = df.iloc[:, pos + 1].astype(str).replace("nan", "")
            attached += 1
    return df


def _get_role(g) -> str:
    g = str(g)
    if "机长" in g: return "机长"
    if "教员" in g: return "教员"
    return "副驾驶"

def _get_level(g: str) -> str:
    m = re.search(r"([A-Z]\d?)类", str(g))
    return m.group(1) if m else ("学员" if "学员" in str(g) else "教员")
