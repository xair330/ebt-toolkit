"""
AeroEBT Analyzer v3.0 — Glass Cockpit Edition
带完整设置面板：背景图、透明度、毛玻璃、字体、输出参数筛选
"""
import os
import sys
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import customtkinter as ctk
from PIL import Image, ImageFilter, ImageDraw
import pandas as pd
import numpy as np

from config_mgr import cfg
from data_loader import load_and_clean, extract_weak_records, classify_issues
from charts import (plot_category_overview, plot_heatmap, plot_role_comparison,
                    plot_level_bubble, plot_aircraft_radar)
from theme_matrix_analysis import (
    load_data as tma_load_data,
    plot_heatmap as tma_plot_heatmap,
    plot_risk_heatmap, plot_sankey,
)

# ══════════════════════════════════════════════════════════════════
# 全局主题色
# ══════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C_BG_DEEP   = "#0A1020"
C_PANEL     = "#141C30"
C_PANEL2    = "#1A2440"
C_TEXT      = "#D4E1F9"
C_TEXT_DIM  = "#6880A8"
C_TEAL      = "#1A9480"
C_AMBER     = "#C47A12"
C_HUD_GREEN = "#26E8A6"
C_BORDER    = "#1E3060"
C_ACCENT    = "#2860C0"

# 设置存储文件路径
_SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui_settings.json")

# 默认 UI 设置
DEFAULT_UI = {
    "bg_image_path": os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   r"..\AeroEBT\AEROEBT.png"),
    "panel_alpha":      0.88,          # 面板半透明度 0~1
    "blur_radius":      0,             # 毛玻璃模糊半径（像素）
    "font_family":      "HarmonyOS Sans SC",
    "font_size_base":   13,
    # 输出筛选参数
    "filter_jishu_dengji":  [],        # 技术等级多选（空=全选）
    "filter_jixing":        [],        # 机型（空=全选）
    "filter_yunyingjidi":   [],        # 运行基地（空=全选）
    "filter_xunlian_leixing": [],      # 训练类型（空=全选）
    # 输出图表选择
    "output_charts": {
        "图1_问题大类": True,
        "图2_热力矩阵": True,
        "图3_角色对比": True,
        "图4_气泡图":   True,
        "图5_机型雷达": True,
        "图A_全员主题矩阵": True,
        "图B_低分主题矩阵": True,
        "图C_A330主题矩阵": True,
        "图D_A350主题矩阵": True,
        "图E_全员桑基图":   True,
        "图F_A330桑基图":   True,
        "图G_A350桑基图":   True,
        "图H_全员风险矩阵": True,
        "图I_低分风险矩阵": True,
        "图J_A330风险矩阵": True,
        "图K_A350风险矩阵": True,
    },
}

def load_ui_settings():
    if os.path.exists(_SETTINGS_PATH):
        try:
            with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            # 合并 output_charts（防止字段缺失）
            merged_charts = dict(DEFAULT_UI["output_charts"])
            merged_charts.update(d.get("output_charts", {}))
            d["output_charts"] = merged_charts
            # 补全其他缺失键
            for k, v in DEFAULT_UI.items():
                d.setdefault(k, v)
            return d
        except Exception:
            pass
    return dict(DEFAULT_UI)

def save_ui_settings(ui):
    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(ui, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
# stdout 重定向
# ══════════════════════════════════════════════════════════════════
class StdoutRedirector:
    def __init__(self, widget):
        self.widget = widget
    def write(self, s):
        self.widget.configure(state="normal")
        self.widget.insert("end", s)
        self.widget.see("end")
        self.widget.configure(state="disabled")
    def flush(self):
        pass


# ══════════════════════════════════════════════════════════════════
# 设置面板弹窗
# ══════════════════════════════════════════════════════════════════
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, ui_settings: dict, on_apply):
        super().__init__(parent)
        self.title("⚙  AeroEBT 面板设置")
        self.geometry("780x680")
        self.configure(fg_color=C_BG_DEEP)
        self.resizable(True, True)
        self.grab_set()

        self.ui = dict(ui_settings)   # 工作副本
        self.on_apply = on_apply

        # ── Tab 容器 ──────────────────────────────────────────
        tabs = ctk.CTkTabview(self, fg_color=C_PANEL, segmented_button_fg_color=C_PANEL2,
                              segmented_button_selected_color=C_ACCENT,
                              segmented_button_unselected_hover_color=C_PANEL2,
                              text_color=C_TEXT, corner_radius=12)
        tabs.pack(fill="both", expand=True, padx=14, pady=(12, 0))

        tabs.add("🖼  外观")
        tabs.add("🔤  字体")
        tabs.add("🔽  筛选参数")
        tabs.add("📊  输出图表")

        self._build_tab_appearance(tabs.tab("🖼  外观"))
        self._build_tab_font(tabs.tab("🔤  字体"))
        self._build_tab_filter(tabs.tab("🔽  筛选参数"))
        self._build_tab_charts(tabs.tab("📊  输出图表"))

        # ── 底部按钮 ──────────────────────────────────────────
        btn_bar = ctk.CTkFrame(self, fg_color=C_PANEL, height=56, corner_radius=0)
        btn_bar.pack(fill="x", padx=0, pady=(0, 0), side="bottom")

        ctk.CTkButton(btn_bar, text="取 消", width=100,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      command=self.destroy).pack(side="right", padx=12, pady=10)

        ctk.CTkButton(btn_bar, text="应用并保存", width=140,
                      fg_color=C_ACCENT, hover_color="#3070D0",
                      text_color="white", font=ctk.CTkFont(weight="bold"),
                      command=self._apply).pack(side="right", padx=(0, 6), pady=10)

        ctk.CTkButton(btn_bar, text="恢复默认", width=110,
                      fg_color="transparent", border_width=1, border_color=C_AMBER,
                      text_color=C_AMBER, hover_color="#2A1A08",
                      command=self._reset_defaults).pack(side="left", padx=12, pady=10)

    # ── 外观 Tab ─────────────────────────────────────────────
    def _build_tab_appearance(self, tab):
        tab.configure(fg_color="transparent")
        sf = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        def section(text):
            ctk.CTkLabel(sf, text=text, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C_TEAL).pack(anchor="w", padx=4, pady=(14, 2))
            ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(0, 6))

        # 背景图片
        section("背景图片")
        bg_row = ctk.CTkFrame(sf, fg_color="transparent")
        bg_row.pack(fill="x", padx=4, pady=4)

        self._bg_var = ctk.StringVar(value=self.ui.get("bg_image_path", ""))
        bg_entry = ctk.CTkEntry(bg_row, textvariable=self._bg_var, fg_color=C_PANEL,
                                border_color=C_BORDER, text_color=C_TEXT, width=460)
        bg_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(bg_row, text="浏览…", width=80,
                      fg_color="transparent", border_width=1, border_color=C_TEAL,
                      text_color=C_TEAL, hover_color="#0E2020",
                      command=self._pick_bg).pack(side="left")

        ctk.CTkButton(bg_row, text="清除", width=60,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      command=lambda: self._bg_var.set("")).pack(side="left", padx=(6, 0))

        # 面板透明度
        section("面板半透明度")
        alpha_row = ctk.CTkFrame(sf, fg_color="transparent")
        alpha_row.pack(fill="x", padx=4, pady=4)

        self._alpha_val = ctk.CTkLabel(alpha_row, text=f"{self.ui.get('panel_alpha', 0.88):.0%}",
                                       width=52, text_color=C_HUD_GREEN,
                                       font=ctk.CTkFont(size=14, weight="bold"))
        self._alpha_val.pack(side="right", padx=(8, 0))

        self._alpha_slider = ctk.CTkSlider(
            alpha_row, from_=0.10, to=1.0,
            number_of_steps=90,
            button_color=C_TEAL, button_hover_color=C_HUD_GREEN,
            progress_color=C_TEAL, fg_color=C_PANEL2,
            command=self._alpha_changed)
        self._alpha_slider.set(self.ui.get("panel_alpha", 0.88))
        self._alpha_slider.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(sf, text="  数值越低面板越透明，背景图越显著",
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4)

        # 毛玻璃效果
        section("毛玻璃模糊强度")
        blur_row = ctk.CTkFrame(sf, fg_color="transparent")
        blur_row.pack(fill="x", padx=4, pady=4)

        self._blur_val = ctk.CTkLabel(blur_row, text=f"{self.ui.get('blur_radius', 0)}px",
                                      width=52, text_color=C_HUD_GREEN,
                                      font=ctk.CTkFont(size=14, weight="bold"))
        self._blur_val.pack(side="right", padx=(8, 0))

        self._blur_slider = ctk.CTkSlider(
            blur_row, from_=0, to=32,
            number_of_steps=32,
            button_color=C_TEAL, button_hover_color=C_HUD_GREEN,
            progress_color=C_TEAL, fg_color=C_PANEL2,
            command=self._blur_changed)
        self._blur_slider.set(self.ui.get("blur_radius", 0))
        self._blur_slider.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(sf, text="  0 = 无模糊，数值越大背景越模糊（毛玻璃感越强）",
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4)

    def _pick_bg(self):
        p = filedialog.askopenfilename(
            title="选择背景图片",
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.webp *.bmp"), ("所有文件", "*.*")])
        if p:
            self._bg_var.set(p)

    def _alpha_changed(self, v):
        self._alpha_val.configure(text=f"{float(v):.0%}")

    def _blur_changed(self, v):
        self._blur_val.configure(text=f"{int(float(v))}px")

    # ── 字体 Tab ─────────────────────────────────────────────
    def _build_tab_font(self, tab):
        tab.configure(fg_color="transparent")
        sf = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        def section(text):
            ctk.CTkLabel(sf, text=text, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C_TEAL).pack(anchor="w", padx=4, pady=(14, 2))
            ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(0, 6))

        FONT_PRESETS = [
            "HarmonyOS Sans SC",
            "Microsoft YaHei UI",
            "Source Han Sans SC",
            "Noto Sans CJK SC",
            "SimHei",
            "Arial",
            "Consolas",
        ]

        section("图表字体族")
        self._font_family_var = ctk.StringVar(value=self.ui.get("font_family", "HarmonyOS Sans SC"))
        font_menu = ctk.CTkOptionMenu(sf, values=FONT_PRESETS,
                                      variable=self._font_family_var,
                                      fg_color=C_PANEL, button_color=C_ACCENT,
                                      button_hover_color="#3070D0",
                                      text_color=C_TEXT, dropdown_fg_color=C_PANEL2)
        font_menu.pack(fill="x", padx=4, pady=4)
        ctk.CTkLabel(sf, text="  影响所有生成图表的中文字体",
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4)

        section("基础字号")
        size_row = ctk.CTkFrame(sf, fg_color="transparent")
        size_row.pack(fill="x", padx=4, pady=4)

        self._font_size_val = ctk.CTkLabel(size_row, text=f"{self.ui.get('font_size_base', 13)}pt",
                                           width=52, text_color=C_HUD_GREEN,
                                           font=ctk.CTkFont(size=14, weight="bold"))
        self._font_size_val.pack(side="right", padx=(8, 0))

        self._font_size_slider = ctk.CTkSlider(
            size_row, from_=9, to=22, number_of_steps=13,
            button_color=C_TEAL, button_hover_color=C_HUD_GREEN,
            progress_color=C_TEAL, fg_color=C_PANEL2,
            command=self._font_size_changed)
        self._font_size_slider.set(self.ui.get("font_size_base", 13))
        self._font_size_slider.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(sf, text="  图表坐标轴标签的基础字号（标题会自动按比例放大）",
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4)

        section("字体文件路径（TTF / OTF）")
        ctk.CTkLabel(sf, text="用于图表渲染，留空则使用上方字体族名称自动检索",
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4)

        font_paths = cfg.get("FONTS", {})
        self._font_path_vars = {}
        for label, key in [("Bold / 粗体", "FONT_BOLD"),
                            ("Black / 极黑", "FONT_BLACK"),
                            ("Medium / 常规", "FONT_MEDIUM")]:
            row = ctk.CTkFrame(sf, fg_color="transparent")
            row.pack(fill="x", padx=4, pady=3)
            ctk.CTkLabel(row, text=label, width=110, text_color=C_TEXT,
                         anchor="w").pack(side="left")
            var = ctk.StringVar(value=font_paths.get(key, ""))
            self._font_path_vars[key] = var
            e = ctk.CTkEntry(row, textvariable=var, fg_color=C_PANEL,
                             border_color=C_BORDER, text_color=C_TEXT)
            e.pack(side="left", fill="x", expand=True, padx=(4, 4))
            ctk.CTkButton(row, text="…", width=32,
                          fg_color="transparent", border_width=1, border_color=C_BORDER,
                          text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                          command=lambda v=var: v.set(
                              filedialog.askopenfilename(
                                  filetypes=[("字体文件", "*.ttf *.otf"), ("所有", "*.*")])
                              or v.get())).pack(side="left")

    def _font_size_changed(self, v):
        self._font_size_val.configure(text=f"{int(float(v))}pt")

    # ── 筛选参数 Tab ─────────────────────────────────────────
    def _build_tab_filter(self, tab):
        tab.configure(fg_color="transparent")
        sf = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        def section(text, hint=""):
            ctk.CTkLabel(sf, text=text, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C_TEAL).pack(anchor="w", padx=4, pady=(14, 0))
            if hint:
                ctk.CTkLabel(sf, text=hint, text_color=C_TEXT_DIM,
                             font=ctk.CTkFont(size=11)).pack(anchor="w", padx=4, pady=(0, 2))
            ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(2, 6))

        # ── 机型 ──────────────────────────────
        section("飞机机型", "勾选要分析的机型，留空=全部")
        JIXING_OPTS = ["330", "350", "329;330", "330;350", "320", "ARJ"]
        self._jixing_vars = {}
        row = ctk.CTkFrame(sf, fg_color="transparent")
        row.pack(fill="x", padx=4)
        selected_jx = self.ui.get("filter_jixing", [])
        for i, opt in enumerate(JIXING_OPTS):
            v = ctk.BooleanVar(value=(opt in selected_jx))
            self._jixing_vars[opt] = v
            ctk.CTkCheckBox(row, text=opt, variable=v,
                            text_color=C_TEXT,
                            checkmark_color="white",
                            fg_color=C_ACCENT, hover_color=C_PANEL2,
                            border_color=C_BORDER).grid(
                row=i//3, column=i%3, sticky="w", padx=10, pady=3)

        # ── 运行基地 ──────────────────────────
        section("运行基地", "")
        JIDI_OPTS = ["广州", "北京", "深圳"]
        self._jidi_vars = {}
        row2 = ctk.CTkFrame(sf, fg_color="transparent")
        row2.pack(fill="x", padx=4)
        selected_jd = self.ui.get("filter_yunyingjidi", [])
        for i, opt in enumerate(JIDI_OPTS):
            v = ctk.BooleanVar(value=(opt in selected_jd))
            self._jidi_vars[opt] = v
            ctk.CTkCheckBox(row2, text=opt, variable=v,
                            text_color=C_TEXT, checkmark_color="white",
                            fg_color=C_ACCENT, hover_color=C_PANEL2,
                            border_color=C_BORDER).grid(
                row=0, column=i, sticky="w", padx=10, pady=3)

        # ── 训练类型 ──────────────────────────
        section("训练类型", "")
        LEIXING_OPTS = ["检查课", "训练课"]
        self._leixing_vars = {}
        row3 = ctk.CTkFrame(sf, fg_color="transparent")
        row3.pack(fill="x", padx=4)
        selected_lt = self.ui.get("filter_xunlian_leixing", [])
        for i, opt in enumerate(LEIXING_OPTS):
            v = ctk.BooleanVar(value=(opt in selected_lt))
            self._leixing_vars[opt] = v
            ctk.CTkCheckBox(row3, text=opt, variable=v,
                            text_color=C_TEXT, checkmark_color="white",
                            fg_color=C_ACCENT, hover_color=C_PANEL2,
                            border_color=C_BORDER).grid(
                row=0, column=i, sticky="w", padx=10, pady=3)

        # ── 技术等级 ──────────────────────────
        section("技术等级", "勾选要纳入分析的等级，留空=全部")
        DENGJI_OPTS = [
            "330:A1类副驾驶", "330:A2类副驾驶", "330:B类副驾驶", "330:C类副驾驶",
            "330:D类副驾驶", "330:C类机长", "330:D类机长", "330:Z类机长",
            "330:资深副驾驶", "330:飞行教员A", "330:飞行教员B", "330:飞行教员C",
            "350:A1类副驾驶", "350:A2类副驾驶", "350:B类副驾驶", "350:C类副驾驶",
            "350:D类副驾驶", "350:B类机长", "350:C类机长", "350:D类机长",
            "350:飞行教员A", "350:飞行教员B",
        ]
        self._dengji_vars = {}
        dg_frame = ctk.CTkFrame(sf, fg_color="transparent")
        dg_frame.pack(fill="x", padx=4)
        selected_dg = self.ui.get("filter_jishu_dengji", [])
        for i, opt in enumerate(DENGJI_OPTS):
            v = ctk.BooleanVar(value=(opt in selected_dg))
            self._dengji_vars[opt] = v
            ctk.CTkCheckBox(dg_frame, text=opt, variable=v,
                            text_color=C_TEXT, checkmark_color="white",
                            fg_color=C_ACCENT, hover_color=C_PANEL2,
                            border_color=C_BORDER).grid(
                row=i//2, column=i%2, sticky="w", padx=10, pady=2)

        # ── 全选/清空 ──────────────────────────
        ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(12, 6))
        btn_row = ctk.CTkFrame(sf, fg_color="transparent")
        btn_row.pack(fill="x", padx=4, pady=4)
        ctk.CTkButton(btn_row, text="全选所有筛选", width=130,
                      fg_color="transparent", border_width=1, border_color=C_TEAL,
                      text_color=C_TEAL, hover_color="#0E2020",
                      command=self._select_all_filters).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="清空（全部纳入）", width=150,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      command=self._clear_all_filters).pack(side="left")

    def _select_all_filters(self):
        for v in self._jixing_vars.values(): v.set(True)
        for v in self._jidi_vars.values(): v.set(True)
        for v in self._leixing_vars.values(): v.set(True)
        for v in self._dengji_vars.values(): v.set(True)

    def _clear_all_filters(self):
        for v in self._jixing_vars.values(): v.set(False)
        for v in self._jidi_vars.values(): v.set(False)
        for v in self._leixing_vars.values(): v.set(False)
        for v in self._dengji_vars.values(): v.set(False)

    # ── 输出图表 Tab ─────────────────────────────────────────
    def _build_tab_charts(self, tab):
        tab.configure(fg_color="transparent")
        sf = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        CHART_GROUPS = [
            ("阶段一  问题大类分析", [
                ("图1_问题大类", "图1 · 问题大类分布环形图"),
                ("图2_热力矩阵", "图2 · 胜任力×问题大类热力矩阵"),
                ("图3_角色对比", "图3 · 教员/机长/副驾驶角色对比"),
                ("图4_气泡图",   "图4 · 技术等级成长路径气泡图"),
                ("图5_机型雷达", "图5 · A330 vs A350 胜任力雷达图"),
            ]),
            ("阶段二  训练主题分析", [
                ("图A_全员主题矩阵", "图A · 胜任力×训练主题  全体人员"),
                ("图B_低分主题矩阵", "图B · 胜任力×训练主题  低分预警（<3分）"),
                ("图C_A330主题矩阵", "图C · 胜任力×训练主题  A330机型"),
                ("图D_A350主题矩阵", "图D · 胜任力×训练主题  A350机型"),
            ]),
            ("阶段三  核心风险分析", [
                ("图E_全员桑基图",   "图E · 训练主题→核心风险桑基图  全员"),
                ("图F_A330桑基图",   "图F · 训练主题→核心风险桑基图  A330"),
                ("图G_A350桑基图",   "图G · 训练主题→核心风险桑基图  A350"),
                ("图H_全员风险矩阵", "图H · 胜任力×核心风险热力矩阵  全员"),
                ("图I_低分风险矩阵", "图I · 胜任力×核心风险热力矩阵  低分"),
                ("图J_A330风险矩阵", "图J · 胜任力×核心风险热力矩阵  A330"),
                ("图K_A350风险矩阵", "图K · 胜任力×核心风险热力矩阵  A350"),
            ]),
        ]

        saved_charts = self.ui.get("output_charts", DEFAULT_UI["output_charts"])
        self._chart_vars = {}

        for group_label, items in CHART_GROUPS:
            ctk.CTkLabel(sf, text=group_label,
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=C_TEAL).pack(anchor="w", padx=4, pady=(14, 2))
            ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(0, 6))

            grp_frame = ctk.CTkFrame(sf, fg_color="transparent")
            grp_frame.pack(fill="x", padx=4)
            for i, (key, label) in enumerate(items):
                v = ctk.BooleanVar(value=saved_charts.get(key, True))
                self._chart_vars[key] = v
                ctk.CTkCheckBox(grp_frame, text=label, variable=v,
                                text_color=C_TEXT, checkmark_color="white",
                                fg_color=C_ACCENT, hover_color=C_PANEL2,
                                border_color=C_BORDER).grid(
                    row=i//2, column=i%2, sticky="w", padx=10, pady=3)

        ctk.CTkFrame(sf, height=1, fg_color=C_BORDER).pack(fill="x", padx=4, pady=(12, 6))
        btn_row = ctk.CTkFrame(sf, fg_color="transparent")
        btn_row.pack(fill="x", padx=4, pady=4)
        ctk.CTkButton(btn_row, text="全选", width=80,
                      fg_color="transparent", border_width=1, border_color=C_TEAL,
                      text_color=C_TEAL, hover_color="#0E2020",
                      command=lambda: [v.set(True) for v in self._chart_vars.values()]
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="全不选", width=80,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      command=lambda: [v.set(False) for v in self._chart_vars.values()]
                      ).pack(side="left")

    # ── 应用 ─────────────────────────────────────────────────
    def _apply(self):
        self.ui["bg_image_path"]   = self._bg_var.get()
        self.ui["panel_alpha"]     = round(self._alpha_slider.get(), 2)
        self.ui["blur_radius"]     = int(self._blur_slider.get())
        self.ui["font_family"]     = self._font_family_var.get()
        self.ui["font_size_base"]  = int(self._font_size_slider.get())

        # 字体路径写回 cfg["FONTS"]
        for key, var in self._font_path_vars.items():
            cfg.setdefault("FONTS", {})[key] = var.get()

        # 筛选参数
        self.ui["filter_jishu_dengji"]    = [k for k, v in self._dengji_vars.items() if v.get()]
        self.ui["filter_jixing"]          = [k for k, v in self._jixing_vars.items() if v.get()]
        self.ui["filter_yunyingjidi"]     = [k for k, v in self._jidi_vars.items() if v.get()]
        self.ui["filter_xunlian_leixing"] = [k for k, v in self._leixing_vars.items() if v.get()]

        # 输出图表
        self.ui["output_charts"] = {k: v.get() for k, v in self._chart_vars.items()}

        save_ui_settings(self.ui)
        self.on_apply(self.ui)
        self.destroy()

    def _reset_defaults(self):
        if messagebox.askyesno("确认", "恢复所有设置为默认值？", parent=self):
            save_ui_settings(DEFAULT_UI)
            self.on_apply(dict(DEFAULT_UI))
            self.destroy()


# ══════════════════════════════════════════════════════════════════
# 主应用
# ══════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.ui = load_ui_settings()

        self.title("AeroEBT Analyzer  [Glass Cockpit Edition]")
        self.geometry("960x680")
        self.minsize(800, 580)
        self.configure(fg_color=C_BG_DEEP)

        # 背景图层
        self._bg_label = ctk.CTkLabel(self, text="")
        self._bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        self._apply_background()

        # 主布局
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_sidebar()
        self._build_console()

        # 状态条
        self._status_var = ctk.StringVar(value="SYSTEM READY")
        ctk.CTkLabel(self, textvariable=self._status_var,
                     text_color=C_TEXT_DIM, font=ctk.CTkFont(size=11),
                     fg_color="transparent").grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 6))

        sys.stdout = StdoutRedirector(self.console)
        print("AEROEBT v3.0 ── 系统初始化完成，等待数据接入...\n")

    # ── 背景图 ────────────────────────────────────────────────
    def _apply_background(self):
        path = self.ui.get("bg_image_path", "")
        blur = self.ui.get("blur_radius", 0)
        if path and os.path.exists(path):
            try:
                img = Image.open(path).convert("RGBA")
                if blur > 0:
                    img = img.filter(ImageFilter.GaussianBlur(radius=blur))
                # 窗口尺寸自适应（延迟读取）
                w = self.winfo_width() or 960
                h = self.winfo_height() or 680
                img = img.resize((max(w, 960), max(h, 680)), Image.LANCZOS)
                bg_ctk = ctk.CTkImage(light_image=img, dark_image=img,
                                      size=(max(w, 960), max(h, 680)))
                self._bg_label.configure(image=bg_ctk)
                self._bg_label._image = bg_ctk
                return
            except Exception:
                pass
        self._bg_label.configure(image=None)

    # ── Header ───────────────────────────────────────────────
    def _build_header(self):
        alpha = self.ui.get("panel_alpha", 0.88)
        hf = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        hf.grid(row=0, column=0, columnspan=2, sticky="nsew")

        inner = ctk.CTkFrame(hf, fg_color=C_PANEL, corner_radius=0, height=62)
        inner.pack(fill="both", expand=True, padx=0, pady=0)
        inner.pack_propagate(False)

        ctk.CTkLabel(inner,
                     text="✦  AEROEBT  COMPETENCY  ANALYSIS  ENGINE  v3.0",
                     font=ctk.CTkFont(family="HarmonyOS Sans SC Black", size=22, weight="bold"),
                     text_color=C_HUD_GREEN).pack(side="left", padx=22, pady=16)

        # 设置按钮
        ctk.CTkButton(inner, text="⚙  设 置", width=96, height=36,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      font=ctk.CTkFont(size=13),
                      command=self._open_settings).pack(side="right", padx=(0, 16), pady=13)

        # 当前筛选状态标签
        self._filter_badge = ctk.CTkLabel(
            inner, text=self._filter_summary(), text_color=C_AMBER,
            font=ctk.CTkFont(size=11), fg_color="transparent")
        self._filter_badge.pack(side="right", padx=(0, 8))

    def _filter_summary(self):
        parts = []
        jx = self.ui.get("filter_jixing", [])
        jd = self.ui.get("filter_yunyingjidi", [])
        lt = self.ui.get("filter_xunlian_leixing", [])
        dg = self.ui.get("filter_jishu_dengji", [])
        if jx:   parts.append(f"机型:{'+'.join(jx)}")
        if jd:   parts.append(f"基地:{'/'.join(jd)}")
        if lt:   parts.append(f"类型:{'/'.join(lt)}")
        if dg:   parts.append(f"等级:{len(dg)}项")
        if parts:
            return "筛选: " + "  |  ".join(parts)
        return "筛选: 全部数据（无过滤）"

    # ── 左侧面板 ─────────────────────────────────────────────
    def _build_sidebar(self):
        alpha = self.ui.get("panel_alpha", 0.88)
        self.sidebar = ctk.CTkFrame(self, fg_color=C_PANEL, width=256, corner_radius=12)
        self.sidebar.grid(row=1, column=0, sticky="nsew", padx=(14, 6), pady=(10, 6))
        self.sidebar.grid_propagate(False)

        # 滚动内容区
        sb_scroll = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent",
                                           scrollbar_button_color=C_BORDER,
                                           scrollbar_button_hover_color=C_ACCENT)
        sb_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        def sec_label(text):
            ctk.CTkLabel(sb_scroll, text=text,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=C_TEAL).pack(anchor="w", padx=14, pady=(12, 2))
            ctk.CTkFrame(sb_scroll, height=1, fg_color=C_BORDER).pack(
                fill="x", padx=10, pady=(0, 6))

        # ── 数据文件 ──────────────────────────────────────────
        sec_label("▸ FLIGHT DATA")
        self.file_path_var = ctk.StringVar(value=cfg.get("DATA_FILE", ""))
        ctk.CTkLabel(sb_scroll, textvariable=self.file_path_var,
                     text_color="#6880A8", wraplength=220, justify="left",
                     font=ctk.CTkFont(size=10)).pack(anchor="w", padx=14, pady=(0, 4))
        ctk.CTkButton(sb_scroll, text="LOAD FLIGHT DATA", height=32,
                      fg_color="transparent", border_width=1, border_color=C_TEAL,
                      text_color=C_TEAL, hover_color="#0E2020",
                      font=ctk.CTkFont(size=12),
                      command=self._browse_file).pack(fill="x", padx=12, pady=(0, 8))

        # ── 分析期次 ──────────────────────────────────────────
        sec_label("▸ PERIOD")
        self.entry_period = ctk.CTkEntry(sb_scroll, height=32,
                                         fg_color=C_BG_DEEP, border_color=C_BORDER,
                                         text_color=C_TEXT, font=ctk.CTkFont(size=12))
        self.entry_period.insert(0, cfg.get("REPORT_PERIOD", "2026 Q1"))
        self.entry_period.pack(fill="x", padx=12, pady=(0, 8))

        # ── 输出目录 ──────────────────────────────────────────
        sec_label("▸ OUTPUT DIR")
        default_out = os.path.join(os.path.dirname(
            cfg.get("DATA_FILE", os.path.expanduser("~"))), "分析产出(2026)")
        self.out_path_var = ctk.StringVar(value=default_out)
        ctk.CTkLabel(sb_scroll, textvariable=self.out_path_var,
                     text_color="#6880A8", wraplength=220, justify="left",
                     font=ctk.CTkFont(size=10)).pack(anchor="w", padx=14, pady=(0, 4))
        ctk.CTkButton(sb_scroll, text="SELECT FOLDER", height=32,
                      fg_color="transparent", border_width=1, border_color=C_BORDER,
                      text_color=C_TEXT_DIM, hover_color=C_PANEL2,
                      font=ctk.CTkFont(size=12),
                      command=self._browse_out).pack(fill="x", padx=12, pady=(0, 8))

        # ── 快速筛选预览 ──────────────────────────────────────
        sec_label("▸ ACTIVE FILTER")
        self._filter_preview = ctk.CTkLabel(sb_scroll,
                                             text=self._make_filter_preview(),
                                             text_color=C_TEXT_DIM,
                                             font=ctk.CTkFont(size=11),
                                             justify="left", wraplength=220)
        self._filter_preview.pack(anchor="w", padx=14, pady=(0, 8))

        # ── 执行按钮 ──────────────────────────────────────────
        self.btn_execute = ctk.CTkButton(
            self.sidebar, text="▶  EXECUTE  ANALYSIS",
            height=48, corner_radius=10,
            fg_color="transparent", border_width=2, border_color=C_AMBER,
            text_color=C_AMBER, hover_color="#2A1A08",
            font=ctk.CTkFont(family="HarmonyOS Sans SC", size=14, weight="bold"),
            command=self._run_analysis)
        self.btn_execute.pack(side="bottom", fill="x", padx=12, pady=12)

    def _make_filter_preview(self):
        jx  = self.ui.get("filter_jixing", [])
        jd  = self.ui.get("filter_yunyingjidi", [])
        lt  = self.ui.get("filter_xunlian_leixing", [])
        dg  = self.ui.get("filter_jishu_dengji", [])
        enabled = self.ui.get("output_charts", {})
        n_charts = sum(1 for v in enabled.values() if v)
        lines = [
            f"机型: {', '.join(jx) if jx else '全部'}",
            f"基地: {', '.join(jd) if jd else '全部'}",
            f"训练: {', '.join(lt) if lt else '全部'}",
            f"等级: {'已选 '+str(len(dg))+'项' if dg else '全部'}",
            f"图表: 已选 {n_charts}/16 张",
        ]
        return "\n".join(lines)

    # ── 控制台 ────────────────────────────────────────────────
    def _build_console(self):
        log_frame = ctk.CTkFrame(self, fg_color=C_PANEL, border_color=C_TEAL,
                                 border_width=1, corner_radius=12)
        log_frame.grid(row=1, column=1, sticky="nsew", padx=(0, 14), pady=(10, 6))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_frame, text="▸▸  FLIGHT SYSTEM CONSOLE",
                     font=ctk.CTkFont(family="Consolas", size=13, weight="bold"),
                     text_color=C_TEAL).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 0))

        self.console = ctk.CTkTextbox(
            log_frame, fg_color=C_BG_DEEP, text_color="#A9B7D1",
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=8)
        self.console.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.console.configure(state="disabled")

    # ── 操作方法 ─────────────────────────────────────────────
    def _browse_file(self):
        f = filedialog.askopenfilename(
            title="选择EBT原始数据",
            filetypes=[("Excel", "*.xls *.xlsx"), ("All", "*.*")])
        if f:
            self.file_path_var.set(f)
            cfg["DATA_FILE"] = f

    def _browse_out(self):
        d = filedialog.askdirectory(title="选择产出目录")
        if d:
            self.out_path_var.set(d)

    def _open_settings(self):
        SettingsDialog(self, self.ui, self._on_settings_applied)

    def _on_settings_applied(self, new_ui: dict):
        self.ui = new_ui
        self._apply_background()
        # 刷新筛选预览
        if hasattr(self, "_filter_preview"):
            self._filter_preview.configure(text=self._make_filter_preview())
        if hasattr(self, "_filter_badge"):
            self._filter_badge.configure(text=self._filter_summary())
        self._status_var.set("⚙ 设置已应用")

    def _run_analysis(self):
        data_path = self.file_path_var.get()
        if not os.path.exists(data_path):
            messagebox.showerror("无数据", "请先选择有效的Excel数据文件！")
            return
        cfg["REPORT_PERIOD"] = self.entry_period.get().strip()
        self.btn_execute.configure(state="disabled", text="⏳  ANALYZING...",
                                   border_color="#555", text_color="#555")
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")
        t = threading.Thread(target=self._analysis_task, daemon=True)
        t.start()

    def _analysis_task(self):
        try:
            self._do_analysis_logic()
            print("\n>>> ANALYSIS COMPLETE. ALL SYSTEMS NORMAL. <<<")
            self.after(0, lambda: messagebox.showinfo("完成",
                "全部分析完成！\n图表已保存至输出目录。"))
        except Exception:
            import traceback
            print(f"\n[CRITICAL ERROR]\n{traceback.format_exc()}")
        finally:
            self.after(0, lambda: self.btn_execute.configure(
                state="normal", text="▶  EXECUTE  ANALYSIS",
                border_color=C_AMBER, text_color=C_AMBER))

    # ── 数据筛选 ─────────────────────────────────────────────
    def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        """按 UI 设置对原始 DataFrame 进行筛选"""
        jx = self.ui.get("filter_jixing", [])
        jd = self.ui.get("filter_yunyingjidi", [])
        lt = self.ui.get("filter_xunlian_leixing", [])
        dg = self.ui.get("filter_jishu_dengji", [])

        if jx:
            df = df[df["机型"].astype(str).isin(jx)]
        if jd:
            df = df[df["运行基地"].astype(str).isin(jd)]
        if lt:
            df = df[df["训练类型"].astype(str).isin(lt)]
        if dg:
            df = df[df["技术等级"].astype(str).isin(dg)]
        return df

    def _chart_on(self, key: str) -> bool:
        return self.ui.get("output_charts", {}).get(key, True)

    # ── 分析逻辑 ─────────────────────────────────────────────
    def _do_analysis_logic(self):
        DATA_FILE  = cfg["DATA_FILE"]
        OUTPUT_DIR = self.out_path_var.get()
        PERIOD     = cfg["REPORT_PERIOD"]
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        print("══════════════════════════════════════════════════")
        print("   AeroEBT 大数据深度分析引擎 v3.0")
        print("   ┄ 问题大类 + 训练主题 + 核心风险 全维度 ┄")
        print("══════════════════════════════════════════════════")
        print(f"[*] 数据源:  {DATA_FILE}")
        print(f"[*] 期次:    {PERIOD}")
        print(f"[*] 产出:    {OUTPUT_DIR}")

        # 筛选条件摘要
        jx = self.ui.get("filter_jixing", [])
        jd = self.ui.get("filter_yunyingjidi", [])
        lt = self.ui.get("filter_xunlian_leixing", [])
        dg = self.ui.get("filter_jishu_dengji", [])
        if any([jx, jd, lt, dg]):
            print(f"[*] 筛选:    机型={jx or '全'}  基地={jd or '全'}"
                  f"  类型={lt or '全'}  等级={len(dg)+'项' if dg else '全'}")
        else:
            print("[*] 筛选:    无（全量数据）")
        print()

        # ── 阶段一：加载 & 问题大类 ─────────────────────────
        print("━" * 50)
        print("  阶段一  问题大类维度分析")
        print("━" * 50)

        print("\n[1/12] 数据加载与清洗...")
        df_raw = load_and_clean(DATA_FILE)
        df = self._apply_filters(df_raw)
        issues_df = extract_weak_records(df, threshold=3.0)
        issues_df.to_csv(os.path.join(OUTPUT_DIR, "不足项明细.csv"),
                         index=False, encoding="utf-8-sig")
        cat_df = classify_issues(issues_df, cfg["CATEGORY_KEYS"])
        cat_counts = cat_df["问题大类"].value_counts()
        comp_codes = [c for _, c in cfg["COMPETENCIES"]]
        print(f"      原始 {len(df_raw)} 条 → 筛选后 {len(df)} 条，"
              f"低分项 {len(issues_df)} 条\n")

        if self._chart_on("图1_问题大类"):
            print("[2/12] 渲染图1: 问题大类分布...")
            plot_category_overview(cat_counts,
                os.path.join(OUTPUT_DIR, "图1_问题大类分布.png"))

        if self._chart_on("图2_热力矩阵"):
            print("[3/12] 渲染图2: 胜任力×问题大类热力矩阵...")
            cross = pd.crosstab(cat_df["胜任力"], cat_df["问题大类"])
            plot_heatmap(cross, os.path.join(OUTPUT_DIR, "图2_热力矩阵.png"))

        if self._chart_on("图3_角色对比"):
            print("[4/12] 渲染图3: 角色对比...")
            role_cross = pd.crosstab(cat_df["问题大类"], cat_df["角色"])
            plot_role_comparison(role_cross,
                os.path.join(OUTPUT_DIR, "图3_角色对比.png"))

        if self._chart_on("图4_气泡图"):
            print("[5/12] 渲染图4: 等级气泡图...")
            level_order = ["学员", "A1", "A2", "B", "C", "D", "Z", "教员"]
            bubble_data = cat_df.groupby(["等级", "胜任力"]).size().reset_index(name="count")
            bubble_data["lx"] = bubble_data["等级"].apply(
                lambda x: level_order.index(x) if x in level_order else -1)
            bubble_data["cy"] = bubble_data["胜任力"].apply(
                lambda x: comp_codes.index(x) if x in comp_codes else -1)
            bubble_data = bubble_data[bubble_data["lx"] >= 0]
            plot_level_bubble(bubble_data, level_order, comp_codes,
                os.path.join(OUTPUT_DIR, "图4_气泡图.png"))

        if self._chart_on("图5_机型雷达"):
            print("[6/12] 渲染图5: 机型雷达图...")
            radar_data = {}
            for ac in ["A330", "A350"]:
                sub = cat_df[cat_df["机型"] == ac]
                counts = [sub[sub["胜任力"] == c].shape[0] for c in comp_codes]
                total = sum(counts) or 1
                radar_data[ac] = [x / total for x in counts]
            plot_aircraft_radar(radar_data, comp_codes,
                os.path.join(OUTPUT_DIR, "图5_机型雷达图.png"))

        # PRO 专项（不受图表开关控制，始终输出CSV）
        pro_df = issues_df[issues_df["胜任力"] == "PRO"].copy()
        def classify_pro(text):
            if any(kw in str(text) for kw in cfg["PRO_CALLOUT_KEYS"]): return "标准喊话"
            if any(kw in str(text) for kw in cfg["PRO_CRITICAL_KEYS"]): return "关键程序"
            return "一般性程序"
        pro_df["PRO子类"] = pro_df["问题描述"].apply(classify_pro)
        pro_summary = pro_df["PRO子类"].value_counts().reset_index()
        pro_summary.columns = ["分类区段", "提及总频次"]
        pro_summary.to_csv(os.path.join(OUTPUT_DIR, "PRO专项分析.csv"),
                           index=False, encoding="utf-8-sig")

        # ── 阶段二：训练主题 ─────────────────────────────────
        print("\n" + "━" * 50)
        print("  阶段二  训练主题维度分析")
        print("━" * 50)

        any_theme = any(self._chart_on(k) for k in
                        ["图A_全员主题矩阵", "图B_低分主题矩阵",
                         "图C_A330主题矩阵", "图D_A350主题矩阵"])
        any_risk  = any(self._chart_on(k) for k in
                        ["图E_全员桑基图", "图F_A330桑基图", "图G_A350桑基图",
                         "图H_全员风险矩阵", "图I_低分风险矩阵",
                         "图J_A330风险矩阵", "图K_A350风险矩阵"])

        tma_all = tma_weak = tma_a330 = tma_a350 = None

        if any_theme or any_risk:
            print("\n[7/12] 加载训练主题映射数据...")
            tma_all_raw, tma_weak_raw = tma_load_data(DATA_FILE)

            # 对训练主题数据也应用筛选（机型字段）
            jx_filter = [k for k in jx if k.replace(";", "").strip()]  # 取机型过滤
            tma_a330 = tma_all_raw[tma_all_raw["机型"].astype(str).str.contains("330", na=False)]
            tma_a350 = tma_all_raw[tma_all_raw["机型"].astype(str).str.contains("350", na=False)]

            # 基地/类型/等级筛选暂按全量（tma_load_data 只含训练主题命中记录）
            tma_all  = tma_all_raw
            tma_weak = tma_weak_raw
            print(f"      全员:{len(tma_all)}  低分:{len(tma_weak)}  "
                  f"A330:{len(tma_a330)}  A350:{len(tma_a350)}")
        else:
            print("\n[7/12] 训练主题图表均已跳过（设置中未启用）")

        print("\n[8/12] 渲染训练主题矩阵（图A-D）...")
        if tma_all is not None and self._chart_on("图A_全员主题矩阵"):
            tma_plot_heatmap(tma_all,
                title="胜任力 × 训练主题  热力矩阵  ·  全体人员",
                subtitle=f"{PERIOD} EBT检查数据 | 按评语关键词归类",
                out_path=os.path.join(OUTPUT_DIR, "图A_全体人员_训练主题矩阵.png"))

        if tma_weak is not None and self._chart_on("图B_低分主题矩阵"):
            tma_plot_heatmap(tma_weak,
                title="胜任力 × 训练主题  热力矩阵  ·  低分预警（< 3分）",
                subtitle=f"{PERIOD} EBT检查数据 | 仅统计各胜任力得分低于3分的记录",
                out_path=os.path.join(OUTPUT_DIR, "图B_低分人员_训练主题矩阵.png"))

        if tma_a330 is not None and len(tma_a330) > 0 and self._chart_on("图C_A330主题矩阵"):
            tma_plot_heatmap(tma_a330,
                title="胜任力 × 训练主题  热力矩阵  ·  A330 机型",
                subtitle=f"{PERIOD} A330检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图C_A330机型_训练主题矩阵.png"))

        if tma_a350 is not None and len(tma_a350) > 0 and self._chart_on("图D_A350主题矩阵"):
            tma_plot_heatmap(tma_a350,
                title="胜任力 × 训练主题  热力矩阵  ·  A350 机型",
                subtitle=f"{PERIOD} A350检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图D_A350机型_训练主题矩阵.png"))

        # ── 阶段三：核心风险 ─────────────────────────────────
        print("\n" + "━" * 50)
        print("  阶段三  核心风险关联分析")
        print("━" * 50)

        print("\n[9/12] 渲染桑基图（图E-G）...")
        if tma_all is not None and self._chart_on("图E_全员桑基图"):
            plot_sankey(tma_all,
                title="训练主题 → 核心风险  流向分析  ·  全体人员",
                subtitle=f"{PERIOD} EBT检查数据 | 训练主题经映射关联至核心风险大类",
                out_path=os.path.join(OUTPUT_DIR, "图E_全员_训练主题→核心风险_桑基图.png"))

        if tma_a330 is not None and len(tma_a330) > 0 and self._chart_on("图F_A330桑基图"):
            plot_sankey(tma_a330,
                title="训练主题 → 核心风险  流向分析  ·  A330 机型",
                subtitle=f"{PERIOD} A330检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图F_A330_训练主题→核心风险_桑基图.png"))

        if tma_a350 is not None and len(tma_a350) > 0 and self._chart_on("图G_A350桑基图"):
            plot_sankey(tma_a350,
                title="训练主题 → 核心风险  流向分析  ·  A350 机型",
                subtitle=f"{PERIOD} A350检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图G_A350_训练主题→核心风险_桑基图.png"))

        print("\n[10/12] 渲染胜任力×核心风险矩阵（图H-K）...")
        if tma_all is not None and self._chart_on("图H_全员风险矩阵"):
            plot_risk_heatmap(tma_all,
                title="胜任力 × 核心风险  热力矩阵  ·  全体人员",
                subtitle=f"{PERIOD} EBT检查数据 | 训练主题关联映射至核心风险大类",
                out_path=os.path.join(OUTPUT_DIR, "图H_全员_胜任力×核心风险矩阵.png"))

        if tma_weak is not None and self._chart_on("图I_低分风险矩阵"):
            plot_risk_heatmap(tma_weak,
                title="胜任力 × 核心风险  热力矩阵  ·  低分预警（< 3分）",
                subtitle=f"{PERIOD} EBT检查数据 | 仅统计得分低于3分的记录",
                out_path=os.path.join(OUTPUT_DIR, "图I_低分_胜任力×核心风险矩阵.png"))

        if tma_a330 is not None and len(tma_a330) > 0 and self._chart_on("图J_A330风险矩阵"):
            plot_risk_heatmap(tma_a330,
                title="胜任力 × 核心风险  热力矩阵  ·  A330 机型",
                subtitle=f"{PERIOD} A330检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图J_A330_胜任力×核心风险矩阵.png"))

        if tma_a350 is not None and len(tma_a350) > 0 and self._chart_on("图K_A350风险矩阵"):
            plot_risk_heatmap(tma_a350,
                title="胜任力 × 核心风险  热力矩阵  ·  A350 机型",
                subtitle=f"{PERIOD} A350检查数据",
                out_path=os.path.join(OUTPUT_DIR, "图K_A350_胜任力×核心风险矩阵.png"))

        # ── 阶段四：报告 ─────────────────────────────────────
        print("\n" + "━" * 50)
        print("  阶段四  报告汇总")
        print("━" * 50)

        print("\n[11/12] 生成 Markdown 总报告...")
        report_path = os.path.join(OUTPUT_DIR, "分析总报告.md")
        filter_note = self._make_filter_preview().replace("\n", " | ")
        report_content = f"""# {PERIOD} EBT胜任力深度分析报告

> 筛选条件：{filter_note}
> 原始数据：`{DATA_FILE}`
> 详细数据见同目录 CSV 文件

---

## 一、问题大类分布
![图1](图1_问题大类分布.png)

## 二、胜任力 × 问题大类 热力矩阵
![图2](图2_热力矩阵.png)

## 三、角色对比
![图3](图3_角色对比.png)

## 四、技术等级气泡图
![图4](图4_气泡图.png)

## 五、机型雷达图
![图5](图5_机型雷达图.png)

---

## 六、胜任力 × 训练主题 热力矩阵
### 6.1 全体人员
![图A](图A_全体人员_训练主题矩阵.png)
### 6.2 低分预警
![图B](图B_低分人员_训练主题矩阵.png)
### 6.3 A330
![图C](图C_A330机型_训练主题矩阵.png)
### 6.4 A350
![图D](图D_A350机型_训练主题矩阵.png)

---

## 七、训练主题 → 核心风险 桑基图
### 7.1 全员
![图E](图E_全员_训练主题→核心风险_桑基图.png)
### 7.2 A330
![图F](图F_A330_训练主题→核心风险_桑基图.png)
### 7.3 A350
![图G](图G_A350_训练主题→核心风险_桑基图.png)

---

## 八、胜任力 × 核心风险 热力矩阵
### 8.1 全员
![图H](图H_全员_胜任力×核心风险矩阵.png)
### 8.2 低分
![图I](图I_低分_胜任力×核心风险矩阵.png)
### 8.3 A330
![图J](图J_A330_胜任力×核心风险矩阵.png)
### 8.4 A350
![图K](图K_A350_胜任力×核心风险矩阵.png)
"""
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report_content)

        print("\n[12/12] 全部流程完毕！")
        print(f"\n[OK] 产出目录:\n{OUTPUT_DIR}")


# ══════════════════════════════════════════════════════════════════
# 入口
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import config_mgr
    app = App()
    app.mainloop()
