"""
AeroEBT 胜任力 × 训练主题 矩阵热力分析
要求: pip install pandas matplotlib seaborn openpyxl xlrd
"""
import os
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import seaborn as sns
import warnings
warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════
# 字体配置（HarmonyOS Sans SC）
# ══════════════════════════════════════════════════════════════════
FONT_PATHS = {
    "black":   r"C:\Users\xair3\AppData\Local\Microsoft\Windows\Fonts\HarmonyOS_Sans_SC_Black.ttf",
    "bold":    r"C:\Users\xair3\AppData\Local\Microsoft\Windows\Fonts\HarmonyOS_Sans_SC_Bold.ttf",
    "medium":  r"C:\Users\xair3\AppData\Local\Microsoft\Windows\Fonts\HarmonyOS_Sans_SC_Bold.ttf",
    "regular": r"C:\Users\xair3\AppData\Local\Microsoft\Windows\Fonts\HarmonyOS_Sans_SC_Bold.ttf",
}
_loaded_fonts = {}
for key, path in FONT_PATHS.items():
    if os.path.exists(path):
        _loaded_fonts[key] = fm.FontProperties(fname=path)
        fm.fontManager.addfont(path)

def fp(style="bold", size=12):
    """返回 FontProperties 对象，style='black'用于标题，其余统一用Bold"""
    path = FONT_PATHS.get(style, FONT_PATHS.get("bold", ""))
    if path and os.path.exists(path):
        return fm.FontProperties(fname=path, size=size)
    return fm.FontProperties(size=size)

# 设置 matplotlib 全局字体
if _loaded_fonts.get("bold"):
    matplotlib.rcParams["font.family"] = _loaded_fonts["bold"].get_name()
matplotlib.rcParams["axes.unicode_minus"] = False

# ══════════════════════════════════════════════════════════════════
# 胜任力定义（含中文全名，用于图表轴标签）
# ══════════════════════════════════════════════════════════════════
COMPETENCIES = [
    ("知识运用",          "KNO"),
    ("程序应用和遵守规章", "PRO"),
    ("自动航径管理",       "FPA"),
    ("人工航径管理",       "FPM"),
    ("沟  通",            "COM"),
    ("领导力和团队合作",   "LTW"),
    ("情景意识与信息管理", "SAW"),
    ("工作负荷管理",       "WLM"),
    ("问题解决与决策",     "PSD"),
]
# code → 中文简称（用于 Y 轴）
CODE_TO_NAME = {code: name for name, code in COMPETENCIES}
# Y轴顺序（按胜任力代码）
COMP_ORDER = [code for _, code in COMPETENCIES]

# ══════════════════════════════════════════════════════════════════
# 28个EBT训练主题关键词映射
# ══════════════════════════════════════════════════════════════════
THEME_KEYS = {
    # ─── 非技术类（紫色区域）6个 ───
    "非技术胜任力":         ["非技术", "人为因素", "NTS", "CRM", "机组资源"],
    "合规性":               ["合规", "SOP", "规章", "纪律", "标准程序", "随意",
                             "省略", "漏做", "不按程序", "跳步", "擅自"],
    "监控和交叉检查":       ["监控", "交叉检查", "发现晚", "未发现", "未觉察",
                             "注意力", "遗漏", "扫视", "缺少监控", "漏看", "未注意"],
    "意外性":               ["意外", "突发", "非预期", "猝不及防", "突然", "未预料"],
    "工作负荷、分心、压力": ["负荷", "分心", "压力", "急躁", "紧张", "忙乱",
                             "优先级", "主次", "手忙脚乱", "资源管理", "任务管理",
                             "分清主次", "节点"],
    "飞机系统管理":         ["系统管理", "ECAM", "QRH", "面板", "复位", "重置",
                             "做程序", "非正常程序", "记忆项目", "动作"],

    # ─── A列训练主题 6个 ───
    "恶劣天气":             ["天气", "雷雨", "降水", "积冰", "低能见",
                             "大雾", "结冰", "强对流", "冰雹", "雷暴"],
    "自动化管理":           ["自动化", "自动驾驶", " AP ", "AP断", "AP接", "自动油门",
                             "FMA", "FCU", "模式", "接通", "断开", "ATHR", "AFS",
                             "模式意识", "自动"],
    "复飞管理":             ["复飞", "GO AROUND", "GA ", "中断着陆", "复飞决断",
                             "决断复飞"],
    "人工航空器控制":       ["人工", "操纵", "手控", "杆量", "蹬舵", "偏出",
                             "修正", "打杆", "把控", "反应慢", "操控", "拉杆",
                             "推杆量", "操纵量", "人工操纵"],
    "差错管理，飞机状态管理不当": ["差错", "状态管理", "姿态", "能量", "速度管理",
                                  "下沉", "不当", "轨迹", "剖面", "高距比",
                                  "飞机状态", "坡度", "抬点", "五边"],
    "不稳定进近":           ["不稳定", "截获晚", "高进近", "进近不稳", "未截获",

 # ─── B列续 8个 ───
                             "不稳定进近"],
    "不利的风":             ["侧风", "正侧", "尾风", "阵风", "顶风", "颠簸",
                             "风分量", "侧风分量", "大侧风"],
    "风切变改出":           ["风切变", "WINDSHEAR", "微下击暴流", "逃逸机动",
                             "切变", "风切变改出"],
      "飞机系统故障":         ["故障", "失效", "卡阻", "告警", "MEL", "单通道",
                             "单发", "系统失效", "设备失效"],
    "进近能见度接近最低标准": ["能见度", "RVR", "最低标准", "决断高", " DA ", " MDA ",
                               "低标准进近", "能见接近", "低能见进近"],
    "着陆":                 ["着陆", "拉平", "平飘", "接地", "跳跃", "重着陆",
                             "仰角", "带杆", "下沉快", "浮地", "落地"],
    "跑道或滑行道道面状况": ["跑道", "滑行道", "湿滑", "污染", "摩擦", "道面",
                             "积水区", "打滑", "道面状况", "道面污染"],
    "地形":                 ["地形", "GPWS", "防撞", "拉起", "近地",
                             "标高", "地形起伏", "EGPWS", "近地告警"],
    "复杂状态的预防和改出": ["复杂状态", "大坡度", "异常姿态", "失速",
                             "改出", "预防复杂", "UAS", "不可靠空速"],

    # ─── C列训练主题 8个 ───
    "ATC":                  ["ATC", "管制", "陆空通话", "指令", "流控",
                             "听错", "雷达引导", "无线电", "管制员", "通波"],
    "发动机故障":           ["发动机", "单发", "熄火", "推力不对称", "发动机故障",
                             "FIRE", "ENG FAIL"],
    "火警和烟雾管理":       ["火警", "烟雾", "着火", "灭火", "排烟",
                             "烟雾管理", "火警管理", "FIRE WARNING"],
    "管理配载燃油性能差错": ["配载", "燃油", "性能", "油量", "重心",
                             "超重", "绿点速", "VLS", "V速度", "起飞性能"],
    "导航":                 ["导航", "偏航", "FMC", "MCDU", "进场程序",
                             "航路", "仪表进近", "程序设置", "RNAV", "RNP"],
    "飞行员失能":           ["失能", "生病", "接管", "生理", "身体不适",
                             "丧失能力", "飞行员失能", "机组失能"],
    "航空器冲突":           ["冲突", "TCAS", " RA ", " TA ", "避让",
                             "侵入", "间隔", "近似碰撞"],
    "特定运行或机型":        ["特定", "高高原", "延程", "RNP", "特殊机场",
                             "机型特点", "机型", "特定机场"],
}

# 训练主题显示顺序（与图片顺序一致）
THEME_ORDER = list(THEME_KEYS.keys())

# ══════════════════════════════════════════════════════════════════
# 核心风险映射：训练主题 → 关联的核心风险大类
# ══════════════════════════════════════════════════════════════════
RISK_MAP = {
    # ─── 非技术类 → 综合胜任力短板（单列）───
    "合规性":                           ["综合胜任力短板"],
    "监控和交叉检查":                   ["综合胜任力短板"],
    "意外性":                           ["综合胜任力短板"],
    "工作负荷、分心、压力":             ["综合胜任力短板"],
    "非技术胜任力":                     ["综合胜任力短板"],
    # ─── 技术类训练主题 → 核心风险大类 ───
    "着陆":                             ["ARC", "RE"],
    "复飞管理":                         ["CFIT", "LOC", "ARC"],
    "不稳定进近":                       ["CFIT", "ARC"],
    "飞机系统管理":                     ["KSF", "LOC"],
    "飞机系统故障":                     ["KSF", "LOC"],
    "发动机故障":                       ["KSF", "LOC"],
    "差错管理，飞机状态管理不当":       ["LOC", "CFIT"],
    "人工航空器控制":                   ["LOC", "ARC"],
    "自动化管理":                       ["LOC", "CFIT"],
    "恶劣天气":                         ["LOC", "IFD"],
    "不利的风":                         ["RE", "ARC", "LOC"],
    "风切变改出":                       ["CFIT", "LOC"],
    "地形":                             ["CFIT"],
    "复杂状态的预防和改出":             ["LOC"],
    "ATC":                              ["MAC"],
    "航空器冲突":                       ["MAC"],
    "导航":                             ["MAC", "CFIT"],
    "进近能见度接近最低标准":           ["CFIT", "RE"],
    "跑道或滑行道道面状况":             ["RE", "RI", "GD"],
    "火警和烟雾管理":                   ["IFD", "KSF"],
    "管理配载燃油性能差错":             ["LOC"],
    "飞行员失能":                       ["HSE"],
    "特定运行或机型":                   ["LOC", "CFIT"],
}

# 核心风险大类名称
RISK_NAMES = {
    "CFIT": "可控飞行撞地",
    "LOC":  "空中失控",
    "MAC":  "空中冲突",
    "ARC":  "非正常接触跑道",
    "RE":   "冲偏出跑道",
    "RI":   "跑道侵入",
    "GD":   "地面受损",
    "KSF":  "重要系统故障",
    "IFD":  "飞行中损伤",
    "HSE":  "人员伤病",
    "综合能力": "综合胜任力短板",
}

# 风险列显示顺序
RISK_ORDER = ["CFIT", "LOC", "MAC", "ARC", "RE", "RI", "GD",
              "KSF", "IFD", "HSE", "综合能力"]

# ══════════════════════════════════════════════════════════════════
# 数据加载与提取
# ══════════════════════════════════════════════════════════════════
def load_data(file_path: str):
    """
    读取 XLS，提取不足项记录并匹配训练主题。
    Returns:
        df_all  : 全体人员（有评语且命中训练主题的所有记录）
        df_weak : 低于3分人员（同上，仅得分 < 3.0）
    """
    df = pd.read_excel(file_path)
    cols = list(df.columns)
    base = 19
    for _, code in COMPETENCIES:
        cols[base]   = f"{code}_优点"
        cols[base+1] = f"{code}_得分"
        cols[base+2] = f"{code}_不足"
        base += 3
    df.columns = cols

    first_code = COMPETENCIES[0][1]
    df = df[df[f"{first_code}_得分"] != "得分"].reset_index(drop=True)
    for _, code in COMPETENCIES:
        df[f"{code}_得分"] = pd.to_numeric(df[f"{code}_得分"], errors="coerce")

    all_rows, weak_rows = [], []
    for _, row in df.iterrows():
        for _, code in COMPETENCIES:
            score   = row[f"{code}_得分"]
            comment = str(row[f"{code}_不足"]).strip()
            if not comment or comment in ("nan", "无", "-", ""):
                continue

            themes = [t for t, kws in THEME_KEYS.items()
                      if any(kw in comment for kw in kws)]
            if not themes:
                continue

            for t in themes:
                fleet = str(row.get("训练机型", "")).strip()
                rec = {"胜任力": code, "得分": score, "训练主题": t, "机型": fleet}
                all_rows.append(rec)
                if pd.notna(score) and score < 3.0:
                    weak_rows.append(rec)

    return pd.DataFrame(all_rows), pd.DataFrame(weak_rows)


# ══════════════════════════════════════════════════════════════════
# 通用热力矩阵绘图函数（航空 HUD 风格）
# ══════════════════════════════════════════════════════════════════
def _make_cross(df: pd.DataFrame) -> pd.DataFrame:
    """构建按预设顺序排列的交叉表，缺失填0"""
    cross = pd.crosstab(df["胜任力"], df["训练主题"])
    # 行：按 COMP_ORDER 排序，只保留实际出现的
    row_order = [c for c in COMP_ORDER if c in cross.index]
    # 列：按 THEME_ORDER 排序，只保留实际出现的
    col_order  = [t for t in THEME_ORDER if t in cross.columns]
    cross = cross.reindex(index=row_order, columns=col_order, fill_value=0)
    # Y轴直接使用英文缩写代码（保持 cross.index 不变，即已经是code)
    return cross


def plot_heatmap(df: pd.DataFrame, title: str, subtitle: str,
                 out_path: str, cmap_name: str = "hot_r"):
    """绘制精美航空风格热力矩阵图"""
    if df.empty:
        print(f"  [跳过] {title} — 无匹配数据")
        return

    cross = _make_cross(df)
    if cross.empty:
        print(f"  [跳过] {title} — 交叉表为空")
        return

    n_rows, n_cols = cross.shape

    # ── 颜色方案 ──────────────────────────────────
    # 深色航空底色                                  
    BG_BODY  = "#0D1526"   # 图表绘制区底色
    BG_MAIN  = "#0A1020"   # figure底色
    AX_LABEL = "#8AA8D8"   # 轴标签文字色
    TITLE_C  = "#D4E8FF"   # 标题色
    GRID_C   = "#1C2E50"   # 网格线色
    ZERO_C   = "#0F1D36"   # 数值为0的单元格色
    CBAR_TXT = "#A0C0E8"

    # 自定义 colormap: 深海蓝 → 琥珀 → 警告红
    cdict_hot = {
        "red":   [(0.0, 0.08, 0.08), (0.4, 0.82, 0.82), (1.0, 0.95, 0.95)],
        "green": [(0.0, 0.15, 0.15), (0.4, 0.48, 0.48), (1.0, 0.15, 0.15)],
        "blue":  [(0.0, 0.30, 0.30), (0.4, 0.07, 0.07), (1.0, 0.07, 0.07)],
    }
    cmap_custom = LinearSegmentedColormap("aero_hot", cdict_hot)

    # ── 动态画布尺寸 ────────────────────────────
    fig_w = max(20, n_cols * 0.85 + 4)
    fig_h = max(8,  n_rows * 1.0  + 3.5)

    fig = plt.figure(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax  = fig.add_subplot(111, facecolor=BG_BODY)

    # ── 绘制热力图 ───────────────────────────────
    # 为0的格子单独处理：用 masked array
    data   = cross.values.astype(float)
    masked = np.ma.masked_where(data == 0, data)
    vmax   = data.max() if data.max() > 0 else 1

    im = ax.imshow(masked, aspect="auto", cmap=cmap_custom,
                   vmin=0.5, vmax=vmax, interpolation="nearest")

    # 填充0值格子为深色
    zero_mat = np.ma.masked_where(data != 0, data)
    ax.imshow(zero_mat, aspect="auto",
              cmap=LinearSegmentedColormap.from_list("zero", [ZERO_C, ZERO_C]),
              vmin=0, vmax=1, interpolation="nearest")

    # ── 格子线 ───────────────────────────────────
    for x in np.arange(-0.5, n_cols, 1):
        ax.axvline(x, color=GRID_C, linewidth=0.6)
    for y in np.arange(-0.5, n_rows, 1):
        ax.axhline(y, color=GRID_C, linewidth=0.6)

    # ── 数值标注 ──────────────────────────────────
    for r in range(n_rows):
        for c in range(n_cols):
            val = int(data[r, c])
            if val == 0:
                ax.text(c, r, "·", ha="center", va="center",
                        color="#2A3D60", fontsize=18,
                        fontproperties=fp("bold", 18))
            else:
                # 根据热度决定标注颜色：浅背景用深色字，深背景用浅色字
                norm_val = val / vmax
                txt_c = "#1A1F2E" if norm_val > 0.55 else "#D8EBFF"
                ax.text(c, r, str(val), ha="center", va="center",
                        color=txt_c, fontsize=18, fontweight="bold",
                        fontproperties=fp("bold", 18))

    # ── 坐标轴 ────────────────────────────────────
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(cross.columns, rotation=42, ha="right",
                       fontproperties=fp("bold", 22), color=AX_LABEL)
    ax.set_yticks(range(n_rows))
    # Y轴使用英文缩写，与X轴统一字号
    ax.set_yticklabels(cross.index, rotation=0,
                       fontproperties=fp("bold", 22), color="#C8DEFF")
    ax.tick_params(axis="both", which="both", length=0, pad=8)

    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_C)

    # ── 轴标签 ────────────────────────────────────
    ax.set_xlabel("训练主题 (EBT Training Scenario Module)",
                  fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)
    ax.set_ylabel("胜任力 (Competency)",
                  fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)

    # ── 色条 (Colorbar) ───────────────────────────
    cbar = fig.colorbar(im, ax=ax, fraction=0.02, pad=0.015, aspect=30)
    cbar.ax.yaxis.label.set_color(CBAR_TXT)
    cbar.ax.tick_params(colors=CBAR_TXT, length=3)
    for label in cbar.ax.get_yticklabels():
        label.set_fontproperties(fp("bold", 17))
        label.set_color(CBAR_TXT)
    cbar.outline.set_edgecolor(GRID_C)
    cbar.set_label("出现次数", fontproperties=fp("bold", 17), color=CBAR_TXT)

    # ── 标题区 ────────────────────────────────────
    total = int(data.sum())
    fig.text(0.5, 0.97, title,
             ha="center", va="top", fontproperties=fp("black", 36),
             color=TITLE_C)
    fig.text(0.5, 0.935, subtitle + f"   |   总命中次数：{total}",
             ha="center", va="top", fontproperties=fp("bold", 18),
             color="#6080A0")

    # ── 装饰线 ────────────────────────────────────
    fig.add_artist(plt.Line2D([0.08, 0.92], [0.875, 0.875],
                              transform=fig.transFigure,
                              color="#1A3060", linewidth=1.2))

    plt.tight_layout(rect=[0, 0, 1, 0.87])
    plt.savefig(out_path, dpi=180, bbox_inches="tight",
                facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 训练主题 → 核心风险 展开辅助函数
# ══════════════════════════════════════════════════════════════════
def expand_to_risk(df: pd.DataFrame) -> pd.DataFrame:
    """将训练主题记录展开为核心风险记录（一对多展开）"""
    rows = []
    for _, r in df.iterrows():
        theme = r["训练主题"]
        risks = RISK_MAP.get(theme, [])
        for rk in risks:
            rows.append({**r.to_dict(), "核心风险": rk})
    return pd.DataFrame(rows)


def _make_risk_cross(df: pd.DataFrame) -> pd.DataFrame:
    """构建 胜任力 × 核心风险 交叉表"""
    df_risk = expand_to_risk(df)
    if df_risk.empty:
        return pd.DataFrame()
    cross = pd.crosstab(df_risk["胜任力"], df_risk["核心风险"])
    row_order = [c for c in COMP_ORDER if c in cross.index]
    col_order = [r for r in RISK_ORDER if r in cross.columns]
    return cross.reindex(index=row_order, columns=col_order, fill_value=0)


def plot_theme_risk_heatmap(df: pd.DataFrame, title: str, subtitle: str,
                             out_path: str):
    """
    绘制 训练主题 × 核心风险 热力矩阵。
    Y轴：训练主题（按THEME_ORDER排列）
    X轴：核心风险大类（代码+中文名）
    """
    df_risk = expand_to_risk(df)
    if df_risk.empty:
        print(f"  [跳过] {title} — 无匹配数据")
        return

    cross = pd.crosstab(df_risk["训练主题"], df_risk["核心风险"])
    row_order = [t for t in THEME_ORDER if t in cross.index]
    col_order  = [r for r in RISK_ORDER if r in cross.columns]
    if not row_order or not col_order:
        print(f"  [跳过] {title} — 交叉表为空")
        return
    cross = cross.reindex(index=row_order, columns=col_order, fill_value=0)

    n_rows, n_cols = cross.shape

    BG_BODY  = "#0D1526"
    BG_MAIN  = "#0A1020"
    AX_LABEL = "#8AA8D8"
    TITLE_C  = "#D4E8FF"
    GRID_C   = "#1C2E50"
    ZERO_C   = "#0F1D36"
    CBAR_TXT = "#A0C0E8"

    cdict_risk = {
        "red":   [(0.0, 0.06, 0.06), (0.35, 0.20, 0.20), (0.7, 0.90, 0.90), (1.0, 0.95, 0.95)],
        "green": [(0.0, 0.12, 0.12), (0.35, 0.55, 0.55), (0.7, 0.40, 0.40), (1.0, 0.12, 0.12)],
        "blue":  [(0.0, 0.28, 0.28), (0.35, 0.70, 0.70), (0.7, 0.10, 0.10), (1.0, 0.08, 0.08)],
    }
    cmap_risk = LinearSegmentedColormap("aero_risk2", cdict_risk)

    fig_w = max(14, n_cols * 1.8 + 5)
    fig_h = max(10, n_rows * 0.72 + 4)

    fig = plt.figure(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax  = fig.add_subplot(111, facecolor=BG_BODY)

    data   = cross.values.astype(float)
    masked = np.ma.masked_where(data == 0, data)
    vmax   = data.max() if data.max() > 0 else 1

    im = ax.imshow(masked, aspect="auto", cmap=cmap_risk,
                   vmin=0.5, vmax=vmax, interpolation="nearest")
    zero_mat = np.ma.masked_where(data != 0, data)
    ax.imshow(zero_mat, aspect="auto",
              cmap=LinearSegmentedColormap.from_list("zero", [ZERO_C, ZERO_C]),
              vmin=0, vmax=1, interpolation="nearest")

    for x in np.arange(-0.5, n_cols, 1):
        ax.axvline(x, color=GRID_C, linewidth=0.6)
    for y in np.arange(-0.5, n_rows, 1):
        ax.axhline(y, color=GRID_C, linewidth=0.6)

    for r in range(n_rows):
        for c in range(n_cols):
            val = int(data[r, c])
            if val == 0:
                ax.text(c, r, "·", ha="center", va="center",
                        color="#2A3D60", fontsize=18,
                        fontproperties=fp("bold", 18))
            else:
                norm_val = val / vmax
                txt_c = "#1A1F2E" if norm_val > 0.55 else "#D8EBFF"
                ax.text(c, r, str(val), ha="center", va="center",
                        color=txt_c, fontsize=16, fontweight="bold",
                        fontproperties=fp("bold", 16))

    # X轴：核心风险代码 + 中文名
    x_labels = [f"{code}\n{RISK_NAMES.get(code, code)}" for code in cross.columns]
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(x_labels, rotation=0, ha="center",
                       fontproperties=fp("bold", 15), color=AX_LABEL)
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels(cross.index, rotation=0,
                       fontproperties=fp("bold", 16), color="#C8DEFF")
    ax.tick_params(axis="both", which="both", length=0, pad=8)

    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_C)

    ax.set_xlabel("核心风险大类 (Core Risk Category)",
                  fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)
    ax.set_ylabel("训练主题 (EBT Training Theme)",
                  fontproperties=fp("bold", 18), color=AX_LABEL, labelpad=14)

    cbar = fig.colorbar(im, ax=ax, fraction=0.02, pad=0.015, aspect=30)
    cbar.ax.yaxis.label.set_color(CBAR_TXT)
    cbar.ax.tick_params(colors=CBAR_TXT, length=3)
    for label in cbar.ax.get_yticklabels():
        label.set_fontproperties(fp("bold", 15))
        label.set_color(CBAR_TXT)
    cbar.outline.set_edgecolor(GRID_C)
    cbar.set_label("出现次数", fontproperties=fp("bold", 16), color=CBAR_TXT)

    total = int(data.sum())
    fig.text(0.5, 0.97, title,
             ha="center", va="top", fontproperties=fp("black", 34),
             color=TITLE_C)
    fig.text(0.5, 0.935, subtitle + f"   |   总命中次数：{total}",
             ha="center", va="top", fontproperties=fp("bold", 17),
             color="#6080A0")
    fig.add_artist(plt.Line2D([0.06, 0.94], [0.928, 0.928],
                              transform=fig.transFigure,
                              color="#1A3060", linewidth=1.2))

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(out_path, dpi=180, bbox_inches="tight",
                facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()



# ══════════════════════════════════════════════════════════════════
# 胜任力 × 核心风险 热力矩阵图
# ══════════════════════════════════════════════════════════════════
def plot_risk_heatmap(df: pd.DataFrame, title: str, subtitle: str,
                      out_path: str):
    """绘制 胜任力 × 核心风险大类 热力矩阵"""
    cross = _make_risk_cross(df)
    if cross.empty:
        print(f"  [跳过] {title} — 无匹配数据")
        return

    n_rows, n_cols = cross.shape

    BG_BODY  = "#0D1526"
    BG_MAIN  = "#0A1020"
    AX_LABEL = "#8AA8D8"
    TITLE_C  = "#D4E8FF"
    GRID_C   = "#1C2E50"
    ZERO_C   = "#0F1D36"
    CBAR_TXT = "#A0C0E8"

    cdict_risk = {
        "red":   [(0.0, 0.06, 0.06), (0.35, 0.20, 0.20), (0.7, 0.90, 0.90), (1.0, 0.95, 0.95)],
        "green": [(0.0, 0.12, 0.12), (0.35, 0.55, 0.55), (0.7, 0.40, 0.40), (1.0, 0.12, 0.12)],
        "blue":  [(0.0, 0.28, 0.28), (0.35, 0.70, 0.70), (0.7, 0.10, 0.10), (1.0, 0.08, 0.08)],
    }
    cmap_risk = LinearSegmentedColormap("aero_risk", cdict_risk)

    fig_w = max(14, n_cols * 1.6 + 4)
    fig_h = max(8,  n_rows * 1.0  + 3.5)

    fig = plt.figure(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax  = fig.add_subplot(111, facecolor=BG_BODY)

    data   = cross.values.astype(float)
    masked = np.ma.masked_where(data == 0, data)
    vmax   = data.max() if data.max() > 0 else 1

    im = ax.imshow(masked, aspect="auto", cmap=cmap_risk,
                   vmin=0.5, vmax=vmax, interpolation="nearest")
    zero_mat = np.ma.masked_where(data != 0, data)
    ax.imshow(zero_mat, aspect="auto",
              cmap=LinearSegmentedColormap.from_list("zero", [ZERO_C, ZERO_C]),
              vmin=0, vmax=1, interpolation="nearest")

    for x in np.arange(-0.5, n_cols, 1):
        ax.axvline(x, color=GRID_C, linewidth=0.6)
    for y in np.arange(-0.5, n_rows, 1):
        ax.axhline(y, color=GRID_C, linewidth=0.6)

    for r in range(n_rows):
        for c in range(n_cols):
            val = int(data[r, c])
            if val == 0:
                ax.text(c, r, "·", ha="center", va="center",
                        color="#2A3D60", fontsize=18,
                        fontproperties=fp("bold", 18))
            else:
                norm_val = val / vmax
                txt_c = "#1A1F2E" if norm_val > 0.55 else "#D8EBFF"
                ax.text(c, r, str(val), ha="center", va="center",
                        color=txt_c, fontsize=18, fontweight="bold",
                        fontproperties=fp("bold", 18))

    # X轴：核心风险代码 + 中文名
    x_labels = [f"{code}\n{RISK_NAMES.get(code, code)}" for code in cross.columns]
    ax.set_xticks(range(n_cols))
    ax.set_xticklabels(x_labels, rotation=0, ha="center",
                       fontproperties=fp("bold", 16), color=AX_LABEL)
    ax.set_yticks(range(n_rows))
    ax.set_yticklabels(cross.index, rotation=0,
                       fontproperties=fp("bold", 22), color="#C8DEFF")
    ax.tick_params(axis="both", which="both", length=0, pad=8)

    for spine in ax.spines.values():
        spine.set_edgecolor(GRID_C)

    ax.set_xlabel("核心风险大类 (Core Risk Category)",
                  fontproperties=fp("bold", 20), color=AX_LABEL, labelpad=14)
    ax.set_ylabel("胜任力 (Competency)",
                  fontproperties=fp("bold", 20), color=AX_LABEL, labelpad=14)

    cbar = fig.colorbar(im, ax=ax, fraction=0.02, pad=0.015, aspect=30)
    cbar.ax.yaxis.label.set_color(CBAR_TXT)
    cbar.ax.tick_params(colors=CBAR_TXT, length=3)
    for label in cbar.ax.get_yticklabels():
        label.set_fontproperties(fp("bold", 17))
        label.set_color(CBAR_TXT)
    cbar.outline.set_edgecolor(GRID_C)
    cbar.set_label("出现次数", fontproperties=fp("bold", 17), color=CBAR_TXT)

    total = int(data.sum())
    fig.text(0.5, 0.97, title,
             ha="center", va="top", fontproperties=fp("black", 36),
             color=TITLE_C)
    fig.text(0.5, 0.935, subtitle + f"   |   总命中次数：{total}",
             ha="center", va="top", fontproperties=fp("bold", 19),
             color="#6080A0")

    fig.add_artist(plt.Line2D([0.08, 0.92], [0.932, 0.932],
                              transform=fig.transFigure,
                              color="#1A3060", linewidth=1.2))

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(out_path, dpi=180, bbox_inches="tight",
                facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 桑基图：训练主题 → 核心风险 流向分布
# ══════════════════════════════════════════════════════════════════
def plot_sankey(df: pd.DataFrame, title: str, subtitle: str,
                out_path: str):
    """
    绘制纯 matplotlib 桑基图，展示训练主题 → 核心风险的流向。
    左侧：训练主题（按频次降序排列TOP18）
    右侧：核心风险大类
    """
    df_risk = expand_to_risk(df)
    if df_risk.empty:
        print(f"  [跳过] {title} — 无数据")
        return

    # 按训练主题和核心风险分组统计
    flow = df_risk.groupby(["训练主题", "核心风险"]).size().reset_index(name="count")

    # 取左侧 TOP18 训练主题（按总量排序）
    theme_totals = flow.groupby("训练主题")["count"].sum().sort_values(ascending=False)
    top_themes = list(theme_totals.head(18).index)
    flow = flow[flow["训练主题"].isin(top_themes)]

    # 右侧：按 RISK_ORDER 排列
    risk_totals = flow.groupby("核心风险")["count"].sum()
    right_nodes = [r for r in RISK_ORDER if r in risk_totals.index]

    if not top_themes or not right_nodes:
        print(f"  [跳过] {title} — 流量为空")
        return

    BG_MAIN = "#0A1020"
    BG_BODY = "#0D1526"

    fig_w, fig_h = 26, max(14, len(top_themes) * 0.65 + 2)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax.set_facecolor(BG_BODY)
    ax.set_xlim(-0.1, 1.1)
    ax.set_ylim(-0.5, max(len(top_themes), len(right_nodes)) + 0.5)
    ax.axis("off")

    # ── 节点位置 ──────────────────────────────────
    left_x, right_x = 0.08, 0.92
    bar_w = 0.04

    # 左侧节点（训练主题）
    n_left = len(top_themes)
    left_positions = {}
    for i, theme in enumerate(top_themes):
        y = n_left - 1 - i
        left_positions[theme] = y

    # 右侧节点（核心风险）
    n_right = len(right_nodes)
    # 将右侧居中对齐到左侧范围
    right_offset = (n_left - n_right) / 2
    right_positions = {}
    for i, risk in enumerate(right_nodes):
        y = n_left - 1 - right_offset - i
        right_positions[risk] = y

    # ── 左侧颜色 ──────────────────────────────────
    theme_colors = {}
    # 柔和的航空色板
    palette = ["#3498DB", "#2ECC71", "#E74C3C", "#F39C12", "#9B59B6",
               "#1ABC9C", "#E67E22", "#2980B9", "#27AE60", "#C0392B",
               "#D35400", "#8E44AD", "#16A085", "#F1C40F", "#7F8C8D",
               "#2C3E50", "#E84393", "#00CEC9"]
    for i, theme in enumerate(top_themes):
        theme_colors[theme] = palette[i % len(palette)]

    # ── 右侧颜色（核心风险用固定色）──────────────
    risk_colors = {
        "CFIT": "#E74C3C", "LOC": "#FF6B6B", "MAC": "#F39C12",
        "ARC": "#E67E22", "RE": "#D35400", "RI": "#C0392B",
        "GD": "#7F8C8D", "KSF": "#9B59B6", "IFD": "#3498DB",
        "HSE": "#2ECC71", "综合胜任力短板": "#1ABC9C",
    }

    # ── 绘制左侧节点（训练主题条）──────────────
    max_left = max(theme_totals[t] for t in top_themes)
    for theme, y in left_positions.items():
        total = theme_totals[theme]
        bar_h = 0.7
        # 节点矩形
        rect = mpatches.FancyBboxPatch(
            (left_x - bar_w, y - bar_h/2), bar_w, bar_h,
            boxstyle="round,pad=0.01",
            facecolor=theme_colors[theme], alpha=0.85, edgecolor="none")
        ax.add_patch(rect)
        # 标签
        ax.text(left_x - bar_w - 0.01, y,
                f"{theme} ({total})",
                ha="right", va="center", color="#D4E8FF",
                fontproperties=fp("bold", 15))

    # ── 绘制右侧节点（风险条）──────────────────
    for risk, y in right_positions.items():
        total = int(risk_totals.get(risk, 0))
        bar_h = 0.7
        rect = mpatches.FancyBboxPatch(
            (right_x, y - bar_h/2), bar_w, bar_h,
            boxstyle="round,pad=0.01",
            facecolor=risk_colors.get(risk, "#888"), alpha=0.85, edgecolor="none")
        ax.add_patch(rect)
        cn_name = RISK_NAMES.get(risk, risk)
        ax.text(right_x + bar_w + 0.01, y,
                f"{risk} {cn_name} ({total})",
                ha="left", va="center", color="#D4E8FF",
                fontproperties=fp("bold", 16))

    # ── 绘制流线 ──────────────────────────────────
    from matplotlib.path import Path
    import matplotlib.patches as mpath_patches

    total_flow = flow["count"].sum()
    for _, row in flow.iterrows():
        theme, risk, cnt = row["训练主题"], row["核心风险"], row["count"]
        if theme not in left_positions or risk not in right_positions:
            continue
        y_l = left_positions[theme]
        y_r = right_positions[risk]

        alpha = max(0.08, min(0.55, cnt / (total_flow * 0.04)))
        lw = max(0.8, min(12, cnt / max_left * 25))

        # 贝塞尔曲线
        mid_x = (left_x + right_x) / 2
        verts = [(left_x, y_l),
                 (mid_x - 0.05, y_l),
                 (mid_x + 0.05, y_r),
                 (right_x, y_r)]
        codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
        path = Path(verts, codes)
        patch = mpath_patches.PathPatch(
            path, facecolor="none",
            edgecolor=theme_colors.get(theme, "#666"),
            lw=lw, alpha=alpha, capstyle="round")
        ax.add_patch(patch)

        # 流量大于阈值标注数字
        if cnt >= total_flow * 0.015:
            ax.text(mid_x, (y_l + y_r) / 2, str(cnt),
                    ha="center", va="center", color="#D4E8FF",
                    fontsize=11, fontproperties=fp("bold", 11), alpha=0.8)

    # ── 标题 ──────────────────────────────────────
    total = int(flow["count"].sum())
    fig.text(0.5, 0.97, title,
             ha="center", va="top", fontproperties=fp("black", 34),
             color="#D4E8FF")
    fig.text(0.5, 0.94, subtitle + f"   |   总流量：{total}",
             ha="center", va="top", fontproperties=fp("bold", 18),
             color="#6080A0")
    fig.add_artist(plt.Line2D([0.06, 0.94], [0.935, 0.935],
                              transform=fig.transFigure,
                              color="#1A3060", linewidth=1.2))

    # 左右列标题
    ax.text(left_x - bar_w - 0.01, n_left + 0.2,
            "训练主题 (Training Theme)",
            ha="right", va="bottom", color="#8AA8D8",
            fontproperties=fp("bold", 18))
    ax.text(right_x + bar_w + 0.01, n_left + 0.2,
            "核心风险 (Core Risk)",
            ha="left", va="bottom", color="#8AA8D8",
            fontproperties=fp("bold", 18))

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    plt.savefig(out_path, dpi=160, bbox_inches="tight",
                facecolor=BG_MAIN, edgecolor="none")
    print(f"  [OK] 已生成: {out_path}")
    plt.close()


# ══════════════════════════════════════════════════════════════════
# 主程序
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    DATA_FILE = r"c:\工作文件\数据管理\IE\2026.1-3.xls"
    if not os.path.exists(DATA_FILE):
        DATA_FILE = r"c:\工作文件\数据管理\检查员\2026.1-3.xls"

    OUT_DIR = r"c:\工作文件\数据管理\IE\分析产出(2026)"
    os.makedirs(OUT_DIR, exist_ok=True)

    print("=" * 60)
    print("  AeroEBT 胜任力 × 训练主题 矩阵分析")
    print(f"  数据源: {DATA_FILE}")
    print("=" * 60)

    print("\n[1/3] 读取并解析数据...")
    df_all, df_weak = load_data(DATA_FILE)
    print(f"  全体命中记录: {len(df_all)} 条")
    print(f"  低分(<3.0)命中记录: {len(df_weak)} 条")

    PERIOD = "2026 Q1"

    print("\n[2/3] 生成图表A（全体人员）...")
    plot_heatmap(
        df_all,
        title   = f"胜任力 × 训练主题  热力矩阵  ·  全体人员",
        subtitle= f"{PERIOD} EBT检查数据 | 按评语关键词归类，无匹配项不统计",
        out_path= os.path.join(OUT_DIR, "图A_全体人员_训练主题矩阵.png"),
    )

    print("\n[3/3] 生成图表B（低于3分人员）...")
    plot_heatmap(
        df_weak,
        title   = f"胜任力 × 训练主题  热力矩阵  ·  低分预警（< 3分）",
        subtitle= f"{PERIOD} EBT检查数据 | 仅统计各胜任力得分低于3分的记录",
        out_path= os.path.join(OUT_DIR, "图B_低分人员_训练主题矩阵.png"),
    )

    df_a330 = df_all[df_all["机型"].astype(str).str.contains("330", na=False)]
    df_a350 = df_all[df_all["机型"].astype(str).str.contains("350", na=False)]

    print(f"\n[4] 生成图表C（A330机型全体人员）... {len(df_a330)}条记录")
    if len(df_a330) > 0:
        plot_heatmap(
            df_a330,
            title   = f"胜任力 × 训练主题  热力矩阵  ·  A330 机型",
            subtitle= f"{PERIOD} A330检查数据 | 按评语关键词归类，无匹配项不统计",
            out_path= os.path.join(OUT_DIR, "图C_A330机型_训练主题矩阵.png"),
        )

    print(f"\n[5] 生成图表D（A350机型全体人员）... {len(df_a350)}条记录")
    if len(df_a350) > 0:
        plot_heatmap(
            df_a350,
            title   = f"胜任力 × 训练主题  热力矩阵  ·  A350 机型",
            subtitle= f"{PERIOD} A350检查数据 | 按评语关键词归类，无匹配项不统计",
            out_path= os.path.join(OUT_DIR, "图D_A350机型_训练主题矩阵.png"),
        )

    print("\n" + "=" * 60)
    print("  全部分析完成！输出目录:")
    print(f"  {OUT_DIR}")
    print("=" * 60)
