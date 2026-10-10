# -*- coding: utf-8 -*-
"""TankFiring 5星坦克属性查询工具

选择坦克 → 输入等级(1-13)、勋章等级(0-250) 与 五星坦克加成(0-65%) → 查询攻击/血量/维修。
五星坦克加成：在「等级基础值 + 勋章累加值」的基础上再乘以 (100+加成)%，四舍五入取整。
勾选属性卡片上的「已升级」：该属性基础值先 +4.5%，再累加熟练度、最后叠五星加成。
数据来源：TankFiring_5星坦克数据.xlsx（未找到时使用内置数据）。
"""

from __future__ import annotations

import math
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

# ---------------------------------------------------------------------------
# 数据（Excel 同步内置，保证脚本可独立运行）
# ---------------------------------------------------------------------------

LEVELS = list(range(1, 14))  # 1..13

# 逐级属性: {坦克: {"攻击": [lv1..lv13], "血量": [...], "维修": [...]}}
ATTRS = {
    "M1A2D Abrams": {
        "攻击": [2757, 2922, 3097, 3283, 3480, 3689, 3910, 4145, 4394, 4658, 4937, 5233, 5547],
        "血量": [13607, 14423, 15288, 16205, 17177, 18208, 19300, 20458, 21685, 22986, 24365, 25827, 27377],
        "维修": [2268, 2404, 2548, 2701, 2863, 3035, 3217, 3410, 3614, 3831, 4061, 4304, 4563],
    },
    "Type99A": {
        "攻击": [8230, 8724, 9247, 9802, 10390, 11013, 11674, 12374, 13116, 13903, 14737, 15621, 16558],
        "血量": [8157, 8646, 9165, 9715, 10298, 10916, 11571, 12265, 13001, 13781, 14608, 15484, 16413],
        "维修": [2719, 2882, 3055, 3238, 3433, 3639, 3857, 4088, 4334, 4594, 4869, 5161, 5471],
    },
    "T-14 Armata": {
        "攻击": [5471, 5799, 6147, 6516, 6907, 7321, 7760, 8226, 8720, 9243, 9798, 10386, 11009],
        "血量": [10866, 11518, 12209, 12942, 13719, 14542, 15415, 16340, 17320, 18359, 19461, 20629, 21867],
        "维修": [2415, 2560, 2713, 2876, 3049, 3232, 3426, 3631, 3849, 4080, 4325, 4584, 4859],
    },
    "K2": {
        "攻击": [9599, 10175, 10786, 11433, 12119, 12846, 13617, 14434, 15300, 16218, 17191, 18222, 19315],
        "血量": [6789, 7196, 7628, 8086, 8571, 9085, 9630, 10208, 10820, 11469, 12157, 12886, 13659],
        "维修": [2263, 2399, 2543, 2695, 2857, 3028, 3210, 3403, 3607, 3823, 4052, 4295, 4553],
    },
    "Merkava Mark IVM": {
        "攻击": [2287, 2424, 2569, 2723, 2886, 3059, 3243, 3438, 3644, 3863, 4095, 4341, 4601],
        "血量": [15841, 16791, 17798, 18866, 19998, 21198, 22470, 23818, 25247, 26762, 28368, 30070, 31874],
        "维修": [2640, 2798, 2966, 3144, 3333, 3533, 3745, 3970, 4208, 4460, 4728, 5012, 5312],
    },
    "Challenger2 TES": {
        "攻击": [6402, 6786, 7193, 7625, 8082, 8567, 9081, 9626, 10204, 10816, 11465, 12153, 12882],
        "血量": [9050, 9593, 10169, 10779, 11426, 12112, 12839, 13609, 14426, 15292, 16210, 17183, 18214],
        "维修": [2011, 2132, 2260, 2395, 2539, 2692, 2853, 3024, 3206, 3398, 3602, 3818, 4048],
    },
    "Leclerc T4": {
        "攻击": [3194, 3386, 3589, 3804, 4032, 4274, 4530, 4802, 5090, 5395, 5719, 6062, 6426],
        "血量": [11308, 11986, 12705, 13467, 14275, 15132, 16040, 17002, 18022, 19103, 20249, 21464, 22752],
        "维修": [1885, 1998, 2118, 2244, 2379, 2522, 2673, 2834, 3004, 3184, 3375, 3577, 3792],
    },
    "Type 10": {
        "攻击": [6869, 7281, 7718, 8181, 8672, 9192, 9744, 10329, 10949, 11606, 12302, 13040, 13822],
        "血量": [9506, 10076, 10681, 11322, 12001, 12721, 13484, 14293, 15151, 16060, 17024, 18045, 19128],
        "维修": [3169, 3359, 3560, 3774, 4000, 4240, 4495, 4764, 5050, 5353, 5675, 6015, 6376],
    },
    "Leopard 2AX": {
        "攻击": [4568, 4842, 5133, 5441, 5767, 6113, 6480, 6869, 7281, 7718, 8181, 8672, 9192],
        "血量": [12672, 13432, 14238, 15092, 15998, 16958, 17975, 19054, 20197, 21409, 22694, 24056, 25499],
        "维修": [2816, 2985, 3164, 3354, 3555, 3768, 3994, 4234, 4488, 4758, 5043, 5346, 5666],
    },
    "Oplot-M": {
        "攻击": [4568, 4842, 5133, 5441, 5767, 6113, 6480, 6869, 7281, 7718, 8181, 8672, 9192],
        "血量": [12672, 13432, 14238, 15092, 15998, 16958, 17975, 19054, 20197, 21409, 22694, 24056, 25499],
        "维修": [2816, 2985, 3164, 3354, 3555, 3768, 3994, 4234, 4488, 4758, 5043, 5346, 5666],
    },
    "Arjun MK2": {
        "攻击": [6869, 7281, 7718, 8181, 8672, 9192, 9744, 10329, 10949, 11606, 12302, 13040, 13822],
        "血量": [9506, 10076, 10681, 11322, 12001, 12721, 13484, 14293, 15151, 16060, 17024, 18045, 19128],
        "维修": [3169, 3359, 3560, 3774, 4000, 4240, 4495, 4764, 5050, 5353, 5675, 6015, 6376],
    },
    "Karrar": {
        "攻击": [2287, 2424, 2569, 2723, 2886, 3059, 3243, 3438, 3644, 3863, 4095, 4341, 4601],
        "血量": [15841, 16791, 17798, 18866, 19998, 21198, 22470, 23818, 25247, 26762, 28368, 30070, 31874],
        "维修": [2640, 2798, 2966, 3144, 3333, 3533, 3745, 3970, 4208, 4460, 4728, 5012, 5312],
    },
    "Ariete C2": {
        "攻击": [3194, 3386, 3589, 3804, 4032, 4274, 4530, 4802, 5090, 5395, 5719, 6062, 6426],
        "血量": [11308, 11986, 12705, 13467, 14275, 15132, 16040, 17002, 18022, 19103, 20249, 21464, 22752],
        "维修": [1885, 1998, 2118, 2244, 2379, 2522, 2673, 2834, 3004, 3184, 3375, 3577, 3792],
    },
    "Leclerc S1": {
        "攻击": [3194, 3386, 3589, 3804, 4032, 4274, 4530, 4802, 5090, 5395, 5719, 6062, 6426],
        "血量": [11308, 11986, 12705, 13467, 14275, 15132, 16040, 17002, 18022, 19103, 20249, 21464, 22752],
        "维修": [1885, 1998, 2118, 2244, 2379, 2522, 2673, 2834, 3004, 3184, 3375, 3577, 3792],
    },
    "M2020": {
        "攻击": [9599, 10175, 10786, 11433, 12119, 12846, 13617, 14434, 15300, 16218, 17191, 18222, 19315],
        "血量": [6789, 7196, 7628, 8086, 8571, 9085, 9630, 10208, 10820, 11469, 12157, 12886, 13659],
        "维修": [2263, 2399, 2543, 2695, 2857, 3028, 3210, 3403, 3607, 3823, 4052, 4295, 4553],
    },
}

# 熟练度加成解锁表（与 Excel「熟练度加成」Lv1-Lv250 宽表一致，源 proficency_bonus.dat）
# 每项为 [[等级, 加成值], ...]：0 级为 0，1 级解锁基础值，之后仅在特定等级 +1
# 加成值为整数百分比（血量% / 攻击% / 维修%）
BONUS_PTS = {
    "M1A2D Abrams": {
        "血量": [[0, 0], [1, 34], [246, 35]],
        "攻击": [[0, 0], [1, 6], [27, 7]],
        "维修": [[0, 0], [1, 5], [83, 6]],
    },
    "Type99A": {
        "血量": [[0, 0], [1, 20], [152, 21]],
        "攻击": [[0, 0], [1, 20], [107, 21]],
        "维修": [[0, 0], [1, 6], [51, 7]],
    },
    "T-14 Armata": {
        "血量": [[0, 0], [1, 27], [209, 28]],
        "攻击": [[0, 0], [1, 13], [81, 14]],
        "维修": [[0, 0], [1, 6], [241, 7]],
    },
    "K2": {
        "血量": [[0, 0], [1, 16], [7, 17]],
        "攻击": [[0, 0], [1, 24]],
        "维修": [[0, 0], [1, 5], [86, 6]],
    },
    "Merkava Mark IVM": {
        "血量": [[0, 0], [1, 39], [100, 40]],
        "攻击": [[0, 0], [1, 5], [71, 6]],
        "维修": [[0, 0], [1, 6], [101, 7]],
    },
    "Challenger2 TES": {
        "血量": [[0, 0], [1, 22], [94, 23]],
        "攻击": [[0, 0], [1, 16], [249, 17]],
        "维修": [[0, 0], [1, 5], [244, 6]],
    },
    "Leclerc T4": {
        "血量": [[0, 0], [1, 28], [183, 29]],
        "攻击": [[0, 0], [1, 7], [4, 8]],
        "维修": [[0, 0], [1, 4], [72, 5]],
    },
    "Type 10": {
        "血量": [[0, 0], [1, 23], [59, 24]],
        "攻击": [[0, 0], [1, 17], [207, 18]],
        "维修": [[0, 0], [1, 7], [20, 8]],
    },
    "Leopard 2AX": {
        "血量": [[0, 0], [1, 31], [81, 32]],
        "攻击": [[0, 0], [1, 11], [146, 12]],
        "维修": [[0, 0], [1, 7], [241, 8]],
    },
    "Oplot-M": {
        "血量": [[0, 0], [1, 31], [81, 32]],
        "攻击": [[0, 0], [1, 11], [146, 12]],
        "维修": [[0, 0], [1, 7], [241, 8]],
    },
    "Arjun MK2": {
        "血量": [[0, 0], [1, 23], [59, 24]],
        "攻击": [[0, 0], [1, 17], [207, 18]],
        "维修": [[0, 0], [1, 7], [20, 8]],
    },
    "Karrar": {
        "血量": [[0, 0], [1, 39], [100, 40]],
        "攻击": [[0, 0], [1, 5], [71, 6]],
        "维修": [[0, 0], [1, 6], [101, 7]],
    },
    "Ariete C2": {
        "血量": [[0, 0], [1, 28], [183, 29]],
        "攻击": [[0, 0], [1, 7], [4, 8]],
        "维修": [[0, 0], [1, 4], [72, 5]],
    },
    "Leclerc S1": {
        "血量": [[0, 0], [1, 28], [183, 29]],
        "攻击": [[0, 0], [1, 7], [4, 8]],
        "维修": [[0, 0], [1, 4], [72, 5]],
    },
    "M2020": {
        "血量": [[0, 0], [1, 16], [7, 17]],
        "攻击": [[0, 0], [1, 24]],
        "维修": [[0, 0], [1, 5], [86, 6]],
    },
}

# 基础信息（类型 / 装填 / 速度 / 精度）
META = {
    "M1A2D Abrams": ("遭遇型", 3, 100, 100),
    "Type99A": ("攻城型", 9, 100, 100),
    "T-14 Armata": ("突击型", 6, 100, 100),
    "K2": ("攻城型", 9, 100, 100),
    "Merkava Mark IVM": ("遭遇型", 3, 100, 100),
    "Challenger2 TES": ("突击型", 6, 100, 100),
    "Leclerc T4": ("遭遇型", 3, 100, 100),
    "Type 10": ("攻城型", 9, 100, 100),
    "Leopard 2AX": ("突击型", 6, 100, 100),
    "Oplot-M": ("突击型", 6, 100, 100),
    "Arjun MK2": ("攻城型", 9, 100, 100),
    "Karrar": ("遭遇型", 3, 100, 100),
    "Ariete C2": ("遭遇型", 3, 100, 100),
    "Leclerc S1": ("遭遇型", 3, 100, 100),
    "M2020": ("攻城型", 9, 100, 100),
}

TANK_ORDER = [
    "M1A2D Abrams", "Type99A", "T-14 Armata", "K2", "Merkava Mark IVM",
    "Challenger2 TES", "Leclerc T4", "Type 10", "Leopard 2AX", "Oplot-M",
    "Arjun MK2", "Karrar", "Ariete C2", "Leclerc S1", "M2020",
]

ATTR_KEYS = ("攻击", "血量", "维修")

# 配色
C_BG = "#1a1d23"
C_PANEL = "#242830"
C_CARD = "#2c313c"
C_CARD_ATK = "#3d2a2a"
C_CARD_HP = "#2a3d2e"
C_CARD_RPR = "#2a3340"
C_TEXT = "#e8eaed"
C_MUTED = "#9aa0a6"
C_ACCENT = "#f0b429"
C_INPUT = "#323842"
C_BTN_FG = "#1a1d23"


def load_from_excel(path: str) -> bool:
    """尝试从 Excel 覆盖内置数据。失败则返回 False。"""
    try:
        import openpyxl
    except ImportError:
        return False
    if not os.path.isfile(path):
        return False

    try:
        wb = openpyxl.load_workbook(path, data_only=True)

        # 逐级属性明细
        attrs: dict[str, dict[str, list[int]]] = {}
        if "逐级属性明细" in wb.sheetnames:
            for row in wb["逐级属性明细"].iter_rows(min_row=2, values_only=True):
                name_col = row[0]
                if not name_col or "—" not in str(name_col):
                    continue
                tank, attr = str(name_col).split("—", 1)
                tank, attr = tank.strip(), attr.strip()
                if attr not in ATTR_KEYS:
                    continue
                values = [int(v) for v in row[1:14] if v is not None]
                if len(values) == 13:
                    attrs.setdefault(tank, {})[attr] = values
        if attrs:
            ATTRS.clear()
            ATTRS.update(attrs)

        # 熟练度加成：读 Excel 宽表 Lv1-Lv250，重建解锁点
        if "熟练度加成" in wb.sheetnames:
            raw_curves: dict[str, dict[str, list[int]]] = {}
            for row in wb["熟练度加成"].iter_rows(min_row=2, values_only=True):
                tank = row[0]
                typ = str(row[1]) if row[1] else ""
                if not tank:
                    continue
                if "血量" in typ:
                    attr = "血量"
                elif "攻击" in typ:
                    attr = "攻击"
                elif "维修" in typ:
                    attr = "维修"
                else:
                    continue
                # C列起为 Lv1..Lv250
                vals = [int(v) for v in row[2:252] if v is not None]
                if len(vals) == 250:
                    raw_curves.setdefault(str(tank).strip(), {})[attr] = vals
            if raw_curves:
                new_pts: dict[str, dict[str, list[list[int]]]] = {}
                for tank, attrs in raw_curves.items():
                    for attr, vals in attrs.items():
                        pts = [[0, 0]]
                        for i, v in enumerate(vals, start=1):
                            if not pts or pts[-1][1] != v:
                                pts.append([i, int(v)])
                        new_pts.setdefault(tank, {})[attr] = pts
                BONUS_PTS.clear()
                BONUS_PTS.update(new_pts)

        # 总览 meta
        meta = {}
        if "5星坦克总览" in wb.sheetnames:
            for row in wb["5星坦克总览"].iter_rows(min_row=2, values_only=True):
                if not row[1]:
                    continue
                meta[str(row[1]).strip()] = (
                    str(row[2]) if row[2] else "",
                    int(row[3]) if row[3] is not None else 0,
                    int(row[4]) if row[4] is not None else 0,
                    int(row[5]) if row[5] is not None else 0,
                )
        if meta:
            META.clear()
            META.update(meta)
            # 保持与属性表一致的坦克顺序
            ordered = [t for t in TANK_ORDER if t in ATTRS]
            ordered += [t for t in ATTRS if t not in ordered]
            TANK_ORDER.clear()
            TANK_ORDER.extend(ordered)

        return True
    except Exception:
        return False


def level_bonus(tank: str, attr: str, medal: int) -> int:
    """该勋章等级的单级加成数值（绝对值）。"""
    pts = BONUS_PTS.get(tank, {}).get(attr) or [[0, 0]]
    m = max(0, min(250, int(medal)))
    val = pts[0][1]
    for lv, v in pts:
        if m >= lv:
            val = v
        else:
            break
    return int(val)


def bonus_sum(tank: str, attr: str, medal: int) -> int:
    """累加 Lv1..Lv当前勋章 的每级加成数值（绝对值，非百分比）。"""
    m = max(0, min(250, int(medal)))
    if m <= 0:
        return 0
    pts = BONUS_PTS.get(tank, {}).get(attr) or [[0, 0]]
    total = 0
    prev_lv, prev_v = 0, 0
    for lv, v in pts:
        if lv <= 0:
            prev_lv, prev_v = 0, int(v)
            continue
        start = max(prev_lv, 1)
        end = min(m, lv - 1)
        if end >= start:
            total += int(prev_v) * (end - start + 1)
        if m < lv:
            return total
        prev_lv, prev_v = int(lv), int(v)
    start = max(prev_lv, 1)
    if m >= start:
        total += int(prev_v) * (m - start + 1)
    return total


def base_stat(tank: str, attr: str, level: int) -> int:
    row = ATTRS.get(tank, {}).get(attr)
    if not row:
        return 0
    lv = max(1, min(13, int(level)))
    return row[lv - 1]


def final_stat(tank: str, attr: str, level: int, medal: int) -> tuple[int, int, int]:
    """返回 (基础值, 累加加成, 最终值)。

    最终值 = 基础值 + sum(Lv1..勋章等级的每级加成数值)。
    """
    base = base_stat(tank, attr, level)
    bonus = bonus_sum(tank, attr, medal)
    return base, bonus, base + bonus


class TankQueryApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("TankFiring · 5星坦克属性查询")
        self.root.geometry("960x640")
        self.root.minsize(900, 580)
        self.root.configure(bg=C_BG)

        self._setup_style()
        self._build_ui()
        self.query()

    def _setup_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background=C_BG)
        style.configure("Panel.TFrame", background=C_PANEL)
        style.configure("Card.TFrame", background=C_CARD)
        style.configure(
            "TLabel",
            background=C_BG,
            foreground=C_TEXT,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Title.TLabel",
            background=C_BG,
            foreground=C_ACCENT,
            font=("Microsoft YaHei UI", 16, "bold"),
        )
        style.configure(
            "Sub.TLabel",
            background=C_BG,
            foreground=C_MUTED,
            font=("Microsoft YaHei UI", 9),
        )
        style.configure(
            "Panel.TLabel",
            background=C_PANEL,
            foreground=C_TEXT,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "TCombobox",
            fieldbackground=C_INPUT,
            background=C_INPUT,
            foreground=C_TEXT,
            arrowcolor=C_TEXT,
            selectbackground=C_INPUT,
            selectforeground=C_TEXT,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", C_INPUT)],
            foreground=[("readonly", C_TEXT)],
            selectbackground=[("readonly", C_INPUT)],
            selectforeground=[("readonly", C_TEXT)],
        )
        style.configure(
            "TSpinbox",
            fieldbackground=C_INPUT,
            background=C_INPUT,
            foreground=C_TEXT,
            arrowcolor=C_TEXT,
        )
        style.configure(
            "Horizontal.TScale",
            background=C_BG,
            troughcolor=C_INPUT,
        )

    def _build_ui(self) -> None:
        # 顶部标题
        header = tk.Frame(self.root, bg=C_BG)
        header.pack(fill="x", padx=20, pady=(16, 8))
        tk.Label(
            header, text="5星坦克属性计算器", bg=C_BG, fg=C_ACCENT,
            font=("Microsoft YaHei UI", 18, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header, text="选择坦克 · 等级 1–13 · 勋章等级 0–250 · 五星坦克加成 0–65% · 查看攻击 / 血量 / 维修\n© 2026 hixanny",
            bg=C_BG, fg=C_MUTED, font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", pady=(2, 0))

        # 控制面板
        panel = tk.Frame(self.root, bg=C_PANEL, highlightbackground="#3a4050", highlightthickness=1)
        panel.pack(fill="x", padx=20, pady=8)
        inner = tk.Frame(panel, bg=C_PANEL)
        inner.pack(fill="x", padx=16, pady=14)

        # 坦克选择（第一行）
        row_tank = tk.Frame(inner, bg=C_PANEL)
        row_tank.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        inner.columnconfigure(0, weight=1)
        tk.Label(row_tank, text="坦克", bg=C_PANEL, fg=C_MUTED, width=22, anchor="w",
                 font=("Microsoft YaHei UI", 9)).pack(side="left")
        self.tank_var = tk.StringVar(value=TANK_ORDER[0])
        self.tank_cb = ttk.Combobox(
            row_tank, textvariable=self.tank_var, values=TANK_ORDER,
            state="readonly", width=28, font=("Microsoft YaHei UI", 10),
        )
        self.tank_cb.pack(side="left", padx=8)
        self.tank_cb.bind("<<ComboboxSelected>>", lambda _e: self.query())

        # 等级按钮（第二行）
        row_level = tk.Frame(inner, bg=C_PANEL)
        row_level.grid(row=1, column=0, sticky="ew", pady=4)
        tk.Label(row_level, text="等级 (1-13)", bg=C_PANEL, fg=C_MUTED, width=22, anchor="w",
                 font=("Microsoft YaHei UI", 9)).pack(side="left")
        self.level_var = tk.IntVar(value=1)
        self.level_btns: list[tk.Button] = []
        for lv in LEVELS:
            btn = tk.Button(
                row_level, text=str(lv), width=3, bd=0, padx=2, pady=2,
                cursor="hand2", font=("Microsoft YaHei UI", 10, "bold"),
                bg=C_INPUT, fg=C_TEXT, activebackground=C_ACCENT,
                activeforeground=C_BTN_FG, highlightthickness=0,
                command=lambda l=lv: self._set_level(l),
            )
            btn.pack(side="left", padx=2)
            self.level_btns.append(btn)

        # 勋章等级 + 滑条（第三行）
        row_medal = tk.Frame(inner, bg=C_PANEL)
        row_medal.grid(row=2, column=0, sticky="ew", pady=4)
        tk.Label(row_medal, text="勋章等级 (0-250)", bg=C_PANEL, fg=C_MUTED, width=22, anchor="w",
                 font=("Microsoft YaHei UI", 9)).pack(side="left")
        self.medal_var = tk.IntVar(value=0)
        self.medal_spin = ttk.Spinbox(
            row_medal, from_=0, to=250, textvariable=self.medal_var, width=6,
            font=("Microsoft YaHei UI", 10), command=self.query,
        )
        self.medal_spin.pack(side="left", padx=8)
        self.medal_spin.bind("<Return>", lambda _e: self.query())
        self.medal_spin.bind("<FocusOut>", lambda _e: self.query())
        self.medal_scale = tk.Scale(
            row_medal, from_=0, to=250, orient="horizontal", bg=C_PANEL, fg=C_TEXT,
            troughcolor=C_INPUT, highlightthickness=0, bd=0,
            activebackground=C_ACCENT, variable=self.medal_var,
            showvalue=False, command=lambda _v: self.query(),
        )
        self.medal_scale.pack(side="left", fill="x", expand=True, padx=(24, 0))

        # 五星坦克加成（第四行，在前两项数值基础上再乘 (100+pct)%）
        row_pct = tk.Frame(inner, bg=C_PANEL)
        row_pct.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        tk.Label(row_pct, text="五星坦克加成 (0-65%)", bg=C_PANEL, fg=C_MUTED, width=22,
                 anchor="w", font=("Microsoft YaHei UI", 9)).pack(side="left")
        self.pct_var = tk.IntVar(value=0)
        self.pct_spin = ttk.Spinbox(
            row_pct, from_=0, to=65, textvariable=self.pct_var, width=6,
            font=("Microsoft YaHei UI", 10), command=self.query,
        )
        self.pct_spin.pack(side="left", padx=8)
        self.pct_spin.bind("<Return>", lambda _e: self.query())
        self.pct_spin.bind("<FocusOut>", lambda _e: self.query())
        self.pct_scale = tk.Scale(
            row_pct, from_=0, to=65, orient="horizontal", bg=C_PANEL, fg=C_TEXT,
            troughcolor=C_INPUT, highlightthickness=0, bd=0,
            activebackground=C_ACCENT, variable=self.pct_var,
            showvalue=False, command=lambda _v: self.query(),
        )
        self.pct_scale.pack(side="left", fill="x", expand=True, padx=(24, 0))

        # 元信息
        self.meta_label = tk.Label(
            self.root, text="", bg=C_BG, fg=C_MUTED,
            font=("Microsoft YaHei UI", 9), anchor="w",
        )
        self.meta_label.pack(fill="x", padx=24, pady=(8, 4))

        # 三属性卡片
        cards = tk.Frame(self.root, bg=C_BG)
        cards.pack(fill="both", expand=True, padx=20, pady=(4, 12))
        cards.columnconfigure((0, 1, 2), weight=1, uniform="card")
        cards.rowconfigure(0, weight=1)

        self.card_frames = {}
        self.card_widgets = {}
        self.up_vars: dict[str, tk.IntVar] = {}
        self.up_checkboxes: dict[str, tk.Checkbutton] = {}
        card_cfg = [
            ("攻击", C_CARD_ATK, 0),
            ("血量", C_CARD_HP, 1),
            ("维修", C_CARD_RPR, 2),
        ]
        for name, color, col in card_cfg:
            frame = tk.Frame(cards, bg=color, highlightbackground="#3a4050", highlightthickness=1)
            frame.grid(row=0, column=col, sticky="nsew", padx=6, pady=4)
            self.card_frames[name] = frame

            title_row = tk.Frame(frame, bg=color)
            title_row.pack(pady=(14, 4), fill="x", padx=10)
            tk.Label(title_row, text=name, bg=color, fg=C_MUTED,
                     font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
            up_var = tk.IntVar(value=0)
            up_checkbox = tk.Checkbutton(
                title_row, text="已升级", variable=up_var,
                bg=color, activebackground=color, fg=C_TEXT, selectcolor=C_INPUT,
                font=("Microsoft YaHei UI", 9), command=self.query,
            )
            up_checkbox.pack(side="right")
            self.up_vars[name] = up_var
            self.up_checkboxes[name] = up_checkbox

            value = tk.Label(frame, text="—", bg=color, fg=C_TEXT,
                             font=("Microsoft YaHei UI", 28, "bold"))
            value.pack(pady=(4, 2))

            base = tk.Label(frame, text="基础 —", bg=color, fg=C_MUTED,
                            font=("Microsoft YaHei UI", 10))
            base.pack(pady=(10, 2))

            delta = tk.Label(frame, text="提升 —", bg=color, fg=C_MUTED,
                             font=("Microsoft YaHei UI", 9))
            delta.pack(pady=(6, 16))

            self.card_widgets[name] = {"value": value, "base": base, "delta": delta}

        # 底部 DPS / 综合
        footer = tk.Frame(self.root, bg=C_PANEL, highlightbackground="#3a4050", highlightthickness=1)
        footer.pack(fill="x", padx=20, pady=(0, 16))
        self.footer_label = tk.Label(
            footer, text="", bg=C_PANEL, fg=C_TEXT,
            font=("Microsoft YaHei UI", 10), anchor="w", justify="left",
        )
        self.footer_label.pack(fill="x", padx=16, pady=10)

    def _set_level(self, lv: int) -> None:
        self.level_var.set(lv)
        self.query()

    def _refresh_level_btns(self) -> None:
        """高亮当前选中等级对应的按钮。"""
        try:
            cur = int(self.level_var.get())
        except (tk.TclError, ValueError):
            cur = 1
        for i, btn in enumerate(self.level_btns, start=1):
            sel = (i == cur)
            btn.config(
                bg=C_ACCENT if sel else C_INPUT,
                fg=C_BTN_FG if sel else C_TEXT,
            )

    def _read_inputs(self) -> tuple[str, int, int, int] | None:
        tank = self.tank_var.get().strip()
        if tank not in ATTRS:
            messagebox.showwarning("提示", "请选择有效坦克")
            return None
        try:
            level = int(self.level_var.get())
        except (tk.TclError, ValueError):
            level = 13
        try:
            medal = int(self.medal_var.get())
        except (tk.TclError, ValueError):
            medal = 1
        try:
            pct = int(self.pct_var.get())
        except (tk.TclError, ValueError):
            pct = 0
        level = max(1, min(13, level))
        medal = max(0, min(250, medal))
        pct = max(0, min(65, pct))
        # 回写，避免越界显示
        self.level_var.set(level)
        self.medal_var.set(medal)
        self.pct_var.set(pct)
        return tank, level, medal, pct

    def query(self) -> None:
        parsed = self._read_inputs()
        if not parsed:
            return
        tank, level, medal, pct = parsed
        self._refresh_level_btns()

        # 13 级及以上：强制禁用并取消勾选所有已升级复选框
        if level >= 13:
            for attr in ATTR_KEYS:
                if attr in self.up_checkboxes:
                    self.up_checkboxes[attr].config(state='disabled')
                    self.up_vars[attr].set(0)
                    self.up_checkboxes[attr].deselect()
        else:
            # 1-12 级：正常启用复选框
            for attr in ATTR_KEYS:
                if attr in self.up_checkboxes:
                    self.up_checkboxes[attr].config(state='normal')

        ttype, reload_s, speed, acc = META.get(tank, ("", 0, 0, 0))
        self.meta_label.config(
            text=f"{tank}  ·  {ttype}  ·  装填 {reload_s}s  ·  速度 {speed}  ·  精度 {acc}"
        )

        results = {}
        for attr in ATTR_KEYS:
            base, bonus, mid = final_stat(tank, attr, level, medal)
            upgraded = bool(self.up_vars[attr].get())
            total = (base * 1.045 + bonus) if upgraded else float(mid)
            no_pct = math.floor(total + 0.5)
            final = math.floor(total * (100 + pct) / 100 + 0.5)
            up_add = round(base * 0.045) if upgraded else 0
            extra = final - no_pct
            results[attr] = (base, bonus, final, extra)
            w = self.card_widgets[attr]
            w["value"].config(text=f"{final:,}")
            w["base"].config(text=f"基础 {base:,}" + ("（已升级）" if upgraded else ""))
            parts = [f"熟练度共 {bonus:,}"]
            if upgraded:
                parts.append(f"升级 +{up_add:,}")
            if pct:
                parts.append(f"加成 +{extra:,}")
            w["delta"].config(text=" · ".join(parts))

        atk = results["攻击"][2]
        hp = results["血量"][2]
        rpr = results["维修"][2]
        dps = round(atk / reload_s, 1) if reload_s else 0
        self.footer_label.config(
            text=(
                f"综合：DPS≈{dps}  ·  "
                f"等级 {level}  ·  勋章 {medal}  ·  加成 {pct}%"
                f" (累加：血量+{results['血量'][1]:,} / 攻击+{results['攻击'][1]:,} / 维修+{results['维修'][1]:,})"
            )
        )


def main() -> None:
    # 优先读同目录 / 常见位置的 Excel
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "TankFiring_5星坦克数据.xlsx"),
        r"D:\AI\TankFiring_5星坦克数据.xlsx",
        os.path.join(os.getcwd(), "TankFiring_5星坦克数据.xlsx"),
    ]
    for path in candidates:
        if load_from_excel(path):
            break

    root = tk.Tk()
    # Windows 高 DPI
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    TankQueryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
