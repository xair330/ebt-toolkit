"""
EBT分析工具包 - 主运行入口
"""
import os
import sys
import pandas as pd
import numpy as np
import re

from config_mgr import cfg
from data_loader import load_and_clean, extract_weak_records, classify_issues, assess_comment_quality
from charts import (plot_category_overview, plot_heatmap, plot_role_comparison,
                    plot_level_bubble, plot_aircraft_radar)

def pause_and_exit():
    input("\n分析完成，按回车键退出...")
    sys.exit(0)

def main():
    # 支持拖拽：如果命令行传了参数且是Excel，覆盖配置里的DATA_FILE
    if len(sys.argv) > 1:
        dropped_file = sys.argv[1]
        if dropped_file.endswith(('.xls', '.xlsx')):
            cfg["DATA_FILE"] = dropped_file
            print(f"[提示] 检测到拖拽文件输入，使用文件: {dropped_file}")

    DATA_FILE = cfg.get("DATA_FILE", "")
    if not os.path.exists(DATA_FILE) or not DATA_FILE.endswith(('.xls', '.xlsx')):
        print(f"[错误] 无法找到有效的数据文件: {DATA_FILE}")
        print("请拖拽一个 Excel 文件到本程序上，或在 ebt_config.json 中配置正确的路径。")
        pause_and_exit()

    OUTPUT_DIR = os.path.join(os.path.dirname(DATA_FILE), "生成分析产出")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    print("==================================================")
    print("      EBT胜任力与评语大数据深度分析引擎")
    print("==================================================")
    print(f"[*] 数据源: {DATA_FILE}")
    print(f"[*] 期次: {cfg['REPORT_PERIOD']}")
    print(f"[*] 产出目录: {OUTPUT_DIR}\n")

    print("[1/7] 数据加载与清洗...")
    df = load_and_clean(DATA_FILE)
    issues_df = extract_weak_records(df, threshold=3.0)
    issues_df.to_csv(os.path.join(OUTPUT_DIR, "不足项明细.csv"), index=False, encoding="utf-8-sig")
    cat_df = classify_issues(issues_df, cfg["CATEGORY_KEYS"])
    cat_counts = cat_df["问题大类"].value_counts()
    comp_codes = [c for _, c in cfg["COMPETENCIES"]]
    print(f"      共 {len(df)} 条记录，其中低分项 {len(issues_df)} 条")

    print("[2/7] 图1: 问题大类分布...")
    plot_category_overview(cat_counts, os.path.join(OUTPUT_DIR, "图1_问题大类分布.png"))

    print("[3/7] 图2: 胜任力×问题大类热力矩阵...")
    cross = pd.crosstab(cat_df["胜任力"], cat_df["问题大类"])
    plot_heatmap(cross, os.path.join(OUTPUT_DIR, "图2_热力矩阵.png"))

    print("[4/7] 图3: 角色对比条形图...")
    role_cross = pd.crosstab(cat_df["问题大类"], cat_df["角色"])
    plot_role_comparison(role_cross, os.path.join(OUTPUT_DIR, "图3_角色对比.png"))

    print("[5/7] 图4: 等级×胜任力气泡图...")
    level_order = ["学员", "A1", "A2", "B", "C", "D", "Z", "教员"]
    bubble_data = cat_df.groupby(["等级", "胜任力"]).size().reset_index(name="count")
    bubble_data["lx"] = bubble_data["等级"].apply(lambda x: level_order.index(x) if x in level_order else -1)
    bubble_data["cy"] = bubble_data["胜任力"].apply(lambda x: comp_codes.index(x) if x in comp_codes else -1)
    bubble_data = bubble_data[bubble_data["lx"] >= 0]
    plot_level_bubble(bubble_data, level_order, comp_codes, os.path.join(OUTPUT_DIR, "图4_气泡图.png"))

    print("[6/7] 图5: A330 vs A350 雷达图...")
    radar_data = {}
    for ac in ["A330", "A350"]:
        sub = cat_df[cat_df["机型"] == ac]
        counts = [sub[sub["胜任力"] == c].shape[0] for c in comp_codes]
        total = sum(counts) or 1
        radar_data[ac] = [x/total for x in counts]
    plot_aircraft_radar(radar_data, comp_codes, os.path.join(OUTPUT_DIR, "图5_机型雷达图.png"))

    # PRO专项
    pro_df = issues_df[issues_df["胜任力"] == "PRO"].copy()
    def classify_pro(text):
        if any(kw in str(text) for kw in cfg["PRO_CALLOUT_KEYS"]):   return "标准喊话"
        if any(kw in str(text) for kw in cfg["PRO_CRITICAL_KEYS"]):  return "关键程序"
        return "一般性程序"
    pro_df["PRO子类"] = pro_df["问题描述"].apply(classify_pro)
    pro_summary = pro_df["PRO子类"].value_counts().reset_index()
    pro_summary.columns = ["分类区段", "提及总频次"]
    pro_summary.to_csv(os.path.join(OUTPUT_DIR, "PRO专项分析.csv"), index=False, encoding="utf-8-sig")

    REPORT_TITLE = f"{cfg['REPORT_PERIOD']} EBT胜任力与评语大数据深度分析报告"
    report = f"""# {REPORT_TITLE}

---

## 维度一：问题大类分布总览
![图1](图1_问题大类分布.png)

## 维度二：胜任力 × 问题大类 热力矩阵
![图2](图2_热力矩阵.png)

## 维度三：教员 / 机长 / 副驾驶 角色对比
![图3](图3_角色对比.png)

## 维度四：技术等级成长路径气泡图
![图4](图4_气泡图.png)

## 维度五：A330 vs A350 雷达图
![图5](图5_机型雷达图.png)

---

> 原始数据：`{DATA_FILE}`
> 分期：{cfg['REPORT_PERIOD']}
> 详细数据见同目录CSV文件
"""
    report_path = os.path.join(OUTPUT_DIR, "分析总报告.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
        
    print("\n" + "=" * 50)
    print("全部分析完成！所有产出已保存至：")
    print(f" -> {OUTPUT_DIR}")
    print("=" * 50)
    
    pause_and_exit()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n[严重错误] 程序运行过程中发生崩溃: {e}")
        traceback.print_exc()
        pause_and_exit()
