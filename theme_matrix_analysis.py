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
# OB行为指标解析词库 (NLP 映射映射)
# ══════════════════════════════════════════════════════════════════
OB_KEYS = {
    # ─── 包含四大核心非技术能力 (COM, PSD, LTW, WLM) 与关键技术能力 ───
    "COM": {
        "OB1_确定接收者已准备好并且能够接收信息": ["准备好", "接收者准备", "时机不对", "打断", "未注意状态"],
        "OB2_恰当选择沟通的内容、时机、方式和对象": ["时机", "对象", "方式和对象", "找谁", "无效沟通", "选错", "不仅"],
        "OB3_清晰、准确、简洁地传递信息": ["清晰", "准确", "简洁", "词不达意", "声音小", "音量小", "太小声", "冗长", "啰嗦", "重点"],
        "OB4_确认接收者展示出对重要信息的理解": ["确认接收", "是否理解", "未核实理解"],
        "OB5_接收信息时，积极倾听并展示理解": ["倾听", "展示理解", "不理解", "没听清", "听信", "接纳", "反馈"],
        "OB6_询问相关且有效的问题": ["询问", "质疑", "问清", "有效的问题", "不懂就问", "不问"],
        "OB7_适当升级沟通以解决已发现的偏差": ["升级沟通", "强硬", "干预无效时", "未大胆提出", "大胆交流", "再次提醒"],
        "OB8_以适合社会文化的方式使用和解读非语言沟通": ["手势", "肢体", "表情", "情绪识别", "非语言"],
        "OB9_遵守标准的无线电通话用语和程序": ["通话用语", "标准话", "陆空通话", "用语不标准", "ATC"],
        "OB10_使用英文准确应对数据链信息": ["英文", "数据链", "CPDLC"]
    },
    "PSD": {
        "OB1_及时识别、评估和管理威胁和差错": ["识别问题", "核心威胁", "问题的原因", "盲目", "未觉察", "未察觉", "发现不及时"],
        "OB2_从适当的来源寻求准确和充分的信息": ["收集", "获取", "适当的来源", "寻求", "充分的信息", "渠道", "看错", "flysmart", "全面"],
        "OB3_识别并核实出现的问题及原因": ["判断依据", "原因", "误判", "真因", "现象与本质"],
        "OB4_在保证安全的前提下，坚持不懈地解决问题": ["坚持不懈", "放弃", "迎难而上", "缺乏韧劲", "未彻底解决"],
        "OB5_确定并考虑适当的选项": ["选项", "备选", "单一", "方案", "计划单一", "兜底", "多重方案", "没有计划"],
        "OB6_应用适当和及时的决策技巧": ["急于下结论", "着急决策", "仓促", "武断", "拍脑袋", "犹豫", "主观臆断", "果断"],
        "OB7_根据需要监控、回顾以及调整决策": ["检查回顾", "调整策略", "持续关注", "验证", "修正", "未作回顾", "改变主意"],
        "OB8_在缺乏指导或程序的情况下随机应变": ["随机应变", "死板", "灵活处理", "变通"],
        "OB9_当遇到非预期事件时展现出复原力": ["复原力", "意外恢复", "突发应对"]
    },
    "LTW": {
        "OB1_鼓励团队参与和开放坦诚交流": ["鼓励", "意见", "想法", "建议", "氛围", "包容", "坦诚交流"],
        "OB2_需要时表现出主观能动性和提供指导": ["果断地领导", "主观能动性", "积极主动", "引导", "带头"],
        "OB3_使他人参与计划": ["参与计划", "单打独斗", "未让PM参与", "拉入团队"],
        "OB4_考虑他人的意见": ["倾听建议", "独断", "一意孤行", "参考意见", "采纳"],
        "OB5_适宜地给予和接收反馈意见": ["接受意见", "反馈", "听不进", "给予反馈"],
        "OB6_以有效的方式处理和解决冲突与分歧": ["冲突", "分歧", "调节矛盾", "各自为战"],
        "OB7_在需要时果断地领导": ["决断力", "机长担当", "领导力不足", "软弱"],
        "OB8_承担决策和行动的责任": ["承担", "推诿", "依赖", "缺乏担当", "推负责"],
        "OB9_遵照执行指令": ["执行指令", "抗拒", "不听口令", "擅自行动"],
        "OB10_应用有效的干预策略来解决已发现的偏差": ["偏差", "干预策略", "纠正", "兜底", "防患", "干预过晚"],
        "OB11_管理文化和语言方面的挑战": ["语言障碍", "文化差异", "外籍"]
    },
    "WLM": {
        "OB1_在各种情况下都有良好的自我管理（情绪、行为）": ["冷静", "慌乱", "急躁", "紧张", "情绪管理", "手忙脚乱"],
        "OB2_对任务进行有效的规划、优先级分配及节点安排": ["优先级", "主次", "重要紧急", "排序", "本末倒置", "规划"],
        "OB3_在执行任务时有效地管理时间": ["时间管理", "超时", "拖沓", "仓促执行"],
        "OB4_提供和给予协助": ["协助", "帮助", "支持", "接手"],
        "OB5_委派任务": ["委派", "分发", "一人包揽", "全干了"],
        "OB6_适当时寻求并接受协助": ["寻求协助", "死扛", "不懂求助", "拒绝帮忙"],
        "OB7_认真对动作进行监控、回顾和交叉检查": ["交叉检查", "漏看", "监控不到位", "监控负荷", "双重确认"],
        "OB8_核实任务是否已达到预期结果": ["核实预期", "结果偏差", "未验证效果"],
        "OB9_在干扰情形时能有效管理并恢复正常状态": ["分心", "干扰", "专注", "转移注意力", "抗干扰"]
    },
    "PRO": {
        "OB1_确定在哪里可以找到程序和法规": ["手册出处", "乱翻", "找不到", "不知道在哪"],
        "OB2_及时应用相关的运行规定、程序和技术": ["运行规定", "SOP", "偏离", "跳步", "漏做", "动作不规", "未能及时应用"],
        "OB3_遵循 SOP，除非更高的安全性指示需要适当偏离": ["未遵守SOP", "自创动作", "标新立异", "抄捷径"],
        "OB4_正确操作飞机系统和相关设备": ["面板", "输入错", "按错", "操作失误", "拨错", "CDU输入"],
        "OB5_监控飞机系统状态": ["系统状态", "未发觉故障", " ECAM 告警", "参数异常", "漏看状态"],
        "OB6_遵守适用法规": ["违反法规", "违法规限值", "突破红线"],
        "OB7_应用相关的程序知识": ["程序理解错误", "死记硬背", "不知所以然"]
    },
    "FPM": {
        "OB1_根据情况，以适宜的方式，精确、平稳地人工控制飞机": ["操纵量", "打杆", "蹬舵", "力矩", "带杆", "推杆", "平稳", "粗暴控制", "不精细"],
        "OB2_监控并识别与预计飞行航径的偏差，并采取适当措施": ["飞行轨迹", "偏离航径", "高了", "低了", "高度保持", "未修偏"],
        "OB3_使用关系式以及导航信号或目视信息来人工控制飞机": ["目视信息", "错看参考点", "错觉", "仪表交叉"],
        "OB4_管理飞行航径以实现最佳运行表现": ["最佳表现", "航迹粗糙", "非经济速度"],
        "OB5_在人工飞行期间保持预计航径，同时管理干扰": ["分心导致偏航", "注意力分配致偏"],
        "OB6_恰当的使用引导系统以匹配当时的情况": ["未使用引导", "错误跟随指引", "硬跟FD"],
        "OB7_有效监控飞行引导系统模式": ["未看FD模式", "无视指引工作"]
    },
    "FPA": {
        "OB1_根据已有的引导系统，恰当的使用以匹配当时的情况": ["该拔不拔", "晚接通", "盲目相信自动飞行"],
        "OB2_监控并识别与预计飞行航径的偏差，并采取适当措施": ["航迹偏离不知", "VDEV不看", "偏高不知道"],
        "OB3_管理飞行航径以实现最佳运行表现": ["剖面管理差", "高距比错", "过早减速"],
        "OB4_使用自动化功能保持预计飞行航径，同时管理其他任务和干扰": ["自动化负荷管理", "边按MCDU边偏离"],
        "OB5_根据飞行阶段和工作负荷，及时选择适当的自动化级别和模式": ["自动化级别错", "不需要的时候盲用模式", "降级处理不当"],
        "OB6_有效监控飞行引导系统，包括接通的状态和自动模式的转换": ["FMA", "模式", "降级", "未核实自动模式", "喊话错误", "接通状态"]
    },
    "SAW": {
        "OB1_监控并评估飞机及系统的状态": ["飞机状态", "系统状态", "未发觉", "警告灯亮不知"],
        "OB2_监控并评估飞机的能量状态及预计的飞行航径": ["能量状态", "高能量", "低能量", "减速慢", "偏高", "冲偏出倾向"],
        "OB3_监控和评估可能影响运行的整体环境": ["整体环境", "雷达", "气象", "风切变", "地标", "能见度"],
        "OB4_验证信息的准确性并检查过失误差": ["核对信息", "偏听偏信", "虚假高度"],
        "OB5_保持对相关人员表现能力的意识": ["同伴状态不佳", "关注PM", "关注管制员情况"],
        "OB6_制定有效的应对预案": ["应对预案", "没想后招", "未考虑特殊情况"],
        "OB7_对情况意识下降的迹象做出回应": ["懵懂", "意识丧失", "迷失", "反应迟缓", "管状视野"]
    },
    "KNO": {
        "OB1_展示有关限制和系统及相互作用的实用和适用知识": ["工作原理", "系统逻辑", "不知限值", "背不熟"],
        "OB2_展示所需的已公布的运行规定的知识": ["法规不明", "手册限制不清楚", "运行规定"],
        "OB3_展示有关物理学环境、空中交通环境的知识": ["高空物理", "颠簸层", "气流认识", "流量意识"],
        "OB4_展示有关适用法规的适当知识": ["适用要求", "民航规章", " CCAR "],
        "OB5_知道从哪里获得所需信息": ["瞎翻书", "不知道在哪看", "QRH找错"],
        "OB6_表现出对获取知识的积极兴趣": ["学习态度", "敷衍", "未提前准备"],
        "OB7_能够有效地应用知识": ["理论脱离实际", "书呆子", "刻板", "死搬硬套"]
    }
}

import re

def extract_obs_from_text(comment: str, comp_code: str) -> list:
    """
    核心OB提取算法：
    1. 优先提取手动显式标记 (如 OB2、OB-3)。
    2. 若没有，则走 NLP NLP_KEYS 字典进行倒装多模式匹配。
    """
    if not comment or comment in ("nan", "无", "-", ""):
        return []
        
    explicit_obs_list = []
    
    # 规则 1：匹配显式标签 (如 OB1, OB-2, OB 8)
    matches = re.findall(r'OB[-\s]?0?([1-9])', comment, re.IGNORECASE)
    for m in matches:
        explicit_obs_list.append(f"{comp_code}-OB{m}")
        
    # 如果找到了显式OB，优先使用教员打的标签（不再做多重猜测）
    if explicit_obs_list:
        return list(set(explicit_obs_list))
        
    # 规则 2：自然语义逆向匹配 (如果当前胜任力在词库中)
    nlp_obs_list = []
    if comp_code in OB_KEYS:
        # 遍历该胜任力下的所有OB特征进行多重匹配
        for ob_id, kws in OB_KEYS[comp_code].items():
            if any(kw in comment for kw in kws):
                # 提取数字，形如 "OB1"
                nlp_obs_list.append(f"{comp_code}-{ob_id.split('_')[0]}")
                
    return list(set(nlp_obs_list))

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

            obs = extract_obs_from_text(comment, code)
            
            for t in themes:
                fleet = str(row.get("训练机型", "")).strip()
                rec = {
                    "胜任力": code,
                    "得分": score,
                    "训练主题": t,
                    "机型": fleet,
                    "OB标签": obs
                }
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
# 桑基图：三级流向分布 (胜任力 → 训练主题 → 核心风险)
# ══════════════════════════════════════════════════════════════════
def plot_sankey_3_stage(df: pd.DataFrame, title: str, subtitle: str, out_path: str):
    df_risk = expand_to_risk(df)
    if df_risk.empty:
        print(f"  [跳过] {title} — 无数据")
        return

    # 两个阶段的流向
    flow1 = df_risk.groupby(["胜任力", "训练主题"]).size().reset_index(name="count")
    flow2 = df_risk.groupby(["训练主题", "核心风险"]).size().reset_index(name="count")

    # 节点提取与排序
    left_totals = flow1.groupby("胜任力")["count"].sum()
    left_nodes = [c for c in COMP_ORDER if c in left_totals.index]

    mid_totals = df_risk.groupby("训练主题").size().sort_values(ascending=False)
    mid_nodes = list(mid_totals.head(15).index)  # 取TOP 15个主题

    right_totals = flow2.groupby("核心风险")["count"].sum()
    right_nodes = [r for r in RISK_ORDER if r in right_totals.index]

    if not left_nodes or not mid_nodes or not right_nodes:
        print(f"  [跳过] {title} — 流量为空")
        return

    # 仅保留包含在TOP 15的主题的流向
    flow1 = flow1[flow1["训练主题"].isin(mid_nodes)]
    flow2 = flow2[flow2["训练主题"].isin(mid_nodes)]

    BG_MAIN = "#0A1020"
    BG_BODY = "#0D1526"

    max_nodes = max(len(left_nodes), len(mid_nodes), len(right_nodes))
    fig_w = 34
    fig_h = max(14, max_nodes * 0.8 + 2)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), facecolor=BG_MAIN)
    ax.set_facecolor(BG_BODY)
    ax.set_xlim(-0.05, 1.05)
    ax.set_ylim(-0.5, max_nodes + 0.5)
    ax.axis("off")

    x_l, x_m, x_r = 0.05, 0.5, 0.95
    bar_w = 0.025

    def get_positions(nodes):
        n = len(nodes)
        offset = (max_nodes - n) / 2
        pos = {}
        for i, node in enumerate(nodes):
            pos[node] = max_nodes - 1 - offset - i
        return pos

    pos_l = get_positions(left_nodes)
    pos_m = get_positions(mid_nodes)
    pos_r = get_positions(right_nodes)

    comp_colors = {
        "FPA": "#3498DB", "FPM": "#2980B9", "COM": "#F1C40F", "LTW": "#E67E22",
        "SAW": "#9B59B6", "WLM": "#2ECC71", "PSD": "#E74C3C"
    }

    theme_colors = {}
    palette = ["#3498DB", "#2ECC71", "#E74C3C", "#F39C12", "#9B59B6",
               "#1ABC9C", "#E67E22", "#2980B9", "#27AE60", "#C0392B",
               "#D35400", "#8E44AD", "#16A085", "#F1C40F", "#7F8C8D"]
    for i, t in enumerate(mid_nodes):
        theme_colors[t] = palette[i % len(palette)]

    risk_colors = {
        "CFIT": "#E74C3C", "LOC": "#FF6B6B", "MAC": "#F39C12",
        "ARC": "#E67E22", "RE": "#D35400", "RI": "#C0392B",
        "GD": "#7F8C8D", "KSF": "#9B59B6", "IFD": "#3498DB",
        "HSE": "#2ECC71", "综合胜任力短板": "#1ABC9C", "综合能力": "#1ABC9C"
    }

    def draw_node(x, y, color, label, ha, text_offset=0.015, val=0):
        bar_h = 0.6
        rect = mpatches.FancyBboxPatch(
            (x - bar_w/2, y - bar_h/2), bar_w, bar_h,
            boxstyle="round,pad=0.01",
            facecolor=color, alpha=0.9, edgecolor="none")
        ax.add_patch(rect)
        if ha == "right":
            ax.text(x - bar_w/2 - text_offset, y, f"{label} ({int(val)})",
                    ha=ha, va="center", color="#D4E8FF", fontproperties=fp("bold", 16))
        elif ha == "left":
            ax.text(x + bar_w/2 + text_offset, y, f"{label} ({int(val)})",
                    ha=ha, va="center", color="#D4E8FF", fontproperties=fp("bold", 16))
        else: # center
            ax.text(x, y + bar_h/2 + 0.05, f"{label} ({int(val)})",
                    ha="center", va="bottom", color="#D4E8FF", fontproperties=fp("bold", 14))

    for node, y in pos_l.items():
        draw_node(x_l, y, comp_colors.get(node, "#888"), node, "right", val=left_totals.get(node, 0))
    for node, y in pos_m.items():
        draw_node(x_m, y, theme_colors.get(node, "#888"), node, "center", val=mid_totals.get(node, 0))
    for node, y in pos_r.items():
        cn = RISK_NAMES.get(node, node)
        draw_node(x_r, y, risk_colors.get(node, "#888"), f"{node} {cn}", "left", val=right_totals.get(node, 0))

    def draw_flow(flow_df, pos_start, pos_end, x_start, x_end, start_col, end_col, color_dict):
        total_flow = flow_df["count"].sum()
        max_left = flow_df.groupby(start_col)["count"].sum().max()
        if max_left <= 0: max_left = 1
        
        for _, row in flow_df.iterrows():
            st, en, cnt = row[start_col], row[end_col], row["count"]
            if st not in pos_start or en not in pos_end:
                continue
            y_s = pos_start[st]
            y_e = pos_end[en]
            color = color_dict.get(st, "#666")
            
            alpha = max(0.1, min(0.6, cnt / (total_flow * 0.05)))
            lw = max(0.8, min(20, cnt / max_left * 40))

            mid_x = (x_start + x_end) / 2
            import matplotlib.patches as mpath_patches
            from matplotlib.path import Path
            verts = [(x_start + bar_w/2, y_s),
                     (mid_x, y_s), (mid_x, y_e),
                     (x_end - bar_w/2, y_e)]
            codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
            path = Path(verts, codes)
            patch = mpath_patches.PathPatch(
                path, facecolor="none", edgecolor=color,
                lw=lw, alpha=alpha, capstyle="round")
            ax.add_patch(patch)

            if cnt >= total_flow * 0.02:
                tx = mid_x + (0.01 if x_start == x_l else -0.01)
                ax.text(tx, (y_s + y_e)/2, str(cnt), ha="center", va="center",
                        color="#D4E8FF", fontsize=11, fontproperties=fp("bold", 12), alpha=0.9)

    draw_flow(flow1, pos_l, pos_m, x_l, x_m, "胜任力", "训练主题", comp_colors)
    draw_flow(flow2, pos_m, pos_r, x_m, x_r, "训练主题", "核心风险", theme_colors)

    total = int(df_risk.shape[0])
    fig.text(0.5, 0.96, title, ha="center", va="top", fontproperties=fp("black", 38), color="#D4E8FF")
    fig.text(0.5, 0.925, subtitle + f"   |   总流量：{total}", ha="center", va="top", fontproperties=fp("bold", 20), color="#6080A0")
    
    ax.text(x_l, max_nodes + 0.2, "胜任力 (Competency)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))
    ax.text(x_m, max_nodes + 0.2, "训练主题 (Training Theme)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))
    ax.text(x_r, max_nodes + 0.2, "核心风险 (Core Risk)", ha="center", va="bottom", color="#8AA8D8", fontproperties=fp("bold", 22))

    plt.tight_layout(rect=[0, 0, 1, 0.90])
    plt.savefig(out_path, dpi=160, bbox_inches="tight", facecolor=BG_MAIN, edgecolor="none")
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
