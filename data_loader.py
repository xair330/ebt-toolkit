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
    df = pd.read_excel(data_file)
    new_cols = list(df.columns)

    base_idx = 19   # 胜任力列块起始索引（如有变化请调整此值）
    for ch_name, en_code in cfg["COMPETENCIES"]:
        new_cols[base_idx]   = f"{en_code}_优点"
        new_cols[base_idx+1] = f"{en_code}_得分"
        new_cols[base_idx+2] = f"{en_code}_不足"
        base_idx += 3

    df.columns = new_cols
    first_code = cfg["COMPETENCIES"][0][1]
    df = df[df[f"{first_code}_得分"] != "得分"].reset_index(drop=True)

    for _, code in cfg["COMPETENCIES"]:
        df[f"{code}_得分"] = pd.to_numeric(df[f"{code}_得分"], errors="coerce")

    df["机型"] = df["技术等级"].apply(
        lambda x: "A330" if "330" in str(x) else ("A350" if "350" in str(x) else "其他")
    )
    df["角色"] = df["技术等级"].apply(_get_role)
    df["等级"] = df["技术等级"].apply(_get_level)

    return df


def extract_weak_records(df: pd.DataFrame, threshold: float = 3.0) -> pd.DataFrame:
    """
    提取所有得分低于阈值的胜任力记录，返回长表。

    Returns:
        issues_df: 含列 技术等级/机型/角色/等级/胜任力/得分/问题描述
    """
    rows = []
    for _, row in df.iterrows():
        for _, code in cfg["COMPETENCIES"]:
            score = row[f"{code}_得分"]
            desc  = row[f"{code}_不足"]
            if pd.notna(score) and score < threshold and pd.notna(desc) and str(desc).strip():
                rows.append({
                    "技术等级": str(row["技术等级"]),
                    "机型":    row["机型"],
                    "角色":    row["角色"],
                    "等级":    row["等级"],
                    "胜任力":  code,
                    "得分":    score,
                    "问题描述": str(desc).replace("\n", " "),
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
def _get_role(g) -> str:
    g = str(g)
    if "机长" in g: return "机长"
    if "教员" in g: return "教员"
    return "副驾驶"

def _get_level(g: str) -> str:
    m = re.search(r"([A-Z]\d?)类", str(g))
    return m.group(1) if m else ("学员" if "学员" in str(g) else "教员")
