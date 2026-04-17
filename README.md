# AeroEBT Analyzer (v3.3 Glass Cockpit Edition)

AeroEBT Analyzer 是一款高度自动化、数据驱动的航空 EBT (Evidence-Based Training) 胜任力与飞行风险分析引擎。该版本已全面升级为本地离线的现代化 GUI 图形界面，并搭载了专门针对多机队（A330/A350）独立划分的高级深度数据分析流水线。

## 🌟 核心特性与升级内容

1. **多机队独立渲染沙盒**
   从底层自动物理隔离 A330 与 A350 数据，分别为两支机队独立运算并出具专门的针对性报告，数据不互串。
2. **三级核心风险桑基图（3-Level Sankey）**
   自动串联 `训练主题` → `OB（行为指标缺陷）` → `核心风险` 形成直观的三级流向监控图，深挖顶层事件的根本原因。
3. **OB 行为指标全维分析**
   基于官方 ALTMS 的标准字典（NLP解析），从评语长文本中精准捕捉机组的 OB 短板行为特征，生成 `OB高频频次分布卡` 以及 `OB × 训练主题热力矩阵`。
4. **Markdown 全自动拼装报告**
   无需人工梳理，系统自动将图表与数据结果无缝拼装为高质量的 `.md` 直出分析总报告。
5. **UI 图形与深色毛玻璃引擎**
   引入全自由度的配置面板界面，可一键控制全局背景透明度、主题色系、分析过滤参数与个性化数据筛查范围。
6. **绿色无捆绑 (纯净产出)**
   彻底清洗了老版本阶段一、二的繁杂过渡图表（不再强制生成冗余的图1~图K、过程CSV及旧版缝合报告），做到执行即拿干货。

---

## 🚀 快速开始

### 方法一：直接运行独立程序（推荐）
直接在项目 `dist/` 目录下找到打包好的**纯净版可执行文件**双击运行，无需任何环境配置：
```bash
双击运行 dist/AeroEBT_v3.3.exe
```
*(注意：请确保已关闭旧版本的程序界面，否则可能会碰到无法覆盖报错的问题。)*

### 方法二：基于 Python 源码运行
确保本地具备 Python 3 并在虚拟环境中安装了依赖项：
```bash
pip install pandas numpy matplotlib seaborn jieba customtkinter pillow
python gui_main.py
```

---

## 🗂 目录组织架构

```
ebt_analysis_toolkit/
├── dist/                          ← 编译发布的 .exe 产出仓
├── gui_main.py                    ← 工具主入口 (主窗口与GUI流程引擎)
├── theme_matrix_analysis.py       ← 核心 NLP 映射与行为指标 (OB) 分析模块
├── ob_distribution_charts.py      ← OB 等专项衍生图表绘制库
├── generate_sankey3.py            ← 第三代三级桑基图 (Sankey) 渲染引擎
├── fleet_report_generator.py      ← 双机型 Markdown 终极报告自动拼接器
├── data_loader.py                 ← 基础数据底盘交互服务
├── config_mgr.py / config.py      ← 色彩令牌、词云库和配置文件映射
├── ui_settings.json               ← 本地透明度与界面布局记忆存档
├── cargo_config.json              ← 货机特供扩展参数配置表
├── run_cargo.py                   ← 货机分支（B777等）专版后台流水线
└── README.md                      ← 您正在阅读的文件
```

## 📊 V3.3 自动化产出明细

每次执行完成，目标目录下均会整齐出具两套机队（A330、A350）的对应专版分析素材。

| 文件名示意 (以A330为例) | 内容说明 |
|:----|:----|
| `分析报告_A330_[时段].md` | 针对 A330 的 **最高维度综合文字叙述与数据图表** 完整交织报告。 |
| `三级桑基图_A330.png` | `训练主题` 流向 `违规操作模式 (OB)` 流向 `核心风险`。|
| `图_A330_OB高频分布.png` | 排名前 20 名最频发的危险机组操作习惯行为（OB）。|
| `图_A330_OB_Theme矩阵.png` | 将 OB 行为缺陷具体映射散播至各个 `训练阶段主题` 上的热力预警图。|

---

## 🛠️ 关于二次开发与构建

若您调整了底层逻辑图表代码并希望能分享带有您界面的更新执行档给其他人，可以使用 PyInstaller 快速打包重组：
```bash
python -m PyInstaller AeroEBT.spec --clean
```
编译时产生的 `build/` 环境及临时日志已被安全列为 `.gitignore` 不提交名录，只传编译完毕的成品。
