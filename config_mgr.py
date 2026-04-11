import json
import os
import sys

def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG = {
    "DATA_FILE": "请输入Excel文件路径.xls",
    "REPORT_PERIOD": "2026 Q1",
    "COMPETENCIES": [
        ["知识运用", "KNO"],
        ["程序应用和遵守规章", "PRO"],
        ["自动航径管理", "FPA"],
        ["人工航径管理", "FPM"],
        ["沟通", "COM"],
        ["领导力和团队合作", "LTW"],
        ["情景意识与信息管理", "SAW"],
        ["工作负荷管理", "WLM"],
        ["问题解决与决策", "PSD"]
    ],
    "CATEGORY_KEYS": {
        "操纵能力": ["拉平", "操纵量", "操纵偏", "杆量", "姿态", "侧风", "方向偏差", "蹬舵", "接地", "五边", "下沉", "平飘", "坡度", "抬点", "推杆", "带杆"],
        "程序执行": ["程序", "检查单", "遗漏", "漏项", "执行", "不完整", "错误", "复位", "ECAM", "QRH", "复飞程序", "流程", "读", "做"],
        "知识储备": ["不理解", "不清晰", "掌握不", "知识", "不了解", "不熟悉", "不知道", "理解偏", "不够熟"],
        "喊话沟通": ["喊话", "报叫", "证实", "报读", "漏报", "错报", "广播", "中英文", "报出"],
        "自动化管理": ["自动驾驶", "AP", "FMA", "AFS", "自动油门", "FCU", "自动化", "接通", "模式"],
        "能量管理": ["能量", "高距比", "下滑", "截获", "高进近", "不稳定进近", "速度管理", "下降率", "减速板"],
        "威胁管理": ["未发现", "发现晚", "未意识", "监控不", "缺少监控", "意识不", "未注意", "情景意识"],
        "决策质量": ["决策", "判断", "评估", "威胁识别", "信息收集", "感性决策", "备降", "复飞决断"],
        "机组配合": ["沟通", "团队", "配合", "提醒", "PM", "PF", "分工", "主动", "升级", "组员"],
        "负荷管理": ["负荷", "优先级", "分清主次", "任务管理", "资源管理", "规划", "节点", "急躁"]
    },
    "PRO_CALLOUT_KEYS": ["喊话", "报读", "证实", "读单", "错报", "漏报", "读"],
    "PRO_CRITICAL_KEYS": ["紧急", "火警", "释压", "单发", "复飞", "防冰", "风切变", "TCAS", "GPWS", "中断起飞", "记忆项目", "紧急下降", "失效", "ECAM", "严重", "关键", "下沉", "非正常"],
    "UAS_KEYS": ["UAS", "不可靠空速", "空速指示不可靠"],
    "QUALITY_CAUSE_KEYS": ["原因", "由于", "因为", "在于", "归因于", "导致", "造成", "未引起", "缺乏"],
    "QUALITY_SOL_KEYS": ["建议", "方法", "对策", "应该", "需", "注意", "加强", "提升", "多", "练习", "明确"],
    "THEME": {
        "BG_DEEP": "#EEF2FA",
        "BG_GLASS": "#F4F7FD",
        "GLASS_FACE": "#FFFFFFCC",
        "GLASS_EDGE": "#C8D4EE",
        "GRID_GLASS": "#D8E2F4",
        "TEXT_DARK": "#1E2A4A",
        "TEXT_MID": "#5A6A9A",
        "PALETTE": ["#2C6FAC", "#1A9480", "#D4820A", "#3D7DD4", "#1D6B5E", "#8A4E0B", "#205A9E", "#156B5B", "#A0600C", "#4B7FC9"],
        "ROLE_COLORS": {
            "教员": "#2C6FAC",
            "机长": "#C47A12",
            "副驾驶": "#1A9480"
        }
    },
    "FONTS": {
        "FONT_MEDIUM": "C:\\Users\\xair3\\AppData\\Local\\Microsoft\\Windows\\Fonts\\HarmonyOS_Sans_SC_Medium.ttf",
        "FONT_BLACK": "C:\\Users\\xair3\\AppData\\Local\\Microsoft\\Windows\\Fonts\\HarmonyOS_Sans_SC_Black.ttf",
        "FONT_REGULAR": "C:\\Users\\xair3\\AppData\\Local\\Microsoft\\Windows\\Fonts\\HarmonyOS_Sans_SC_Regular.ttf"
    }
}

cfg = {}

def load_or_create_config():
    global cfg
    config_path = os.path.join(get_exe_dir(), "ebt_config.json")
    if not os.path.exists(config_path):
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        cfg.update(DEFAULT_CONFIG)
        print(f"[配置初始化] 已生成默认配置文件: {config_path}")
    else:
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                cfg.update(loaded)
        except Exception as e:
            print(f"[配置错误] 读取 {config_path} 失败: {e}")
            print("将使用系统默认配置！")
            cfg.update(DEFAULT_CONFIG)

load_or_create_config()
