# -*- coding: utf-8 -*-
"""Рисунки к ПР2 (Жильникова, УЭБОиП): деревья событий и матрица рисков."""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

HERE = os.path.dirname(os.path.abspath(__file__))
plt.rcParams["font.family"] = "DejaVu Sans"


def sci(x):
    m, e = f"{x:.1e}".split("e")
    m = m.rstrip("0").rstrip(".").replace(".", ",")
    e = int(e)
    return f"{m}·10$^{{{e}}}$"


# ------------------------------------------------------------ деревья ---
def _leaves(n):
    return 1 if not n.get("ch") else sum(_leaves(c) for c in n["ch"])


def _layout(n, depth, y0, out):
    n["depth"] = depth
    if not n.get("ch"):
        n["y"] = y0 + 0.5
        out.append(n)
        return y0 + 1
    y = y0
    for c in n["ch"]:
        y = _layout(c, depth + 1, y, out)
    n["y"] = sum(c["y"] for c in n["ch"]) / len(n["ch"])
    return y


def _box(ax, x, y, w, h, text, fc="white", ec="black", fs=12, bold=False):
    ax.add_patch(FancyBboxPatch((x, y - h / 2), w, h,
                                boxstyle="square,pad=0", fc=fc, ec=ec, lw=0.9))
    ax.text(x + w / 2, y, text, ha="center", va="center", fontsize=fs,
            weight="bold" if bold else "normal", linespacing=1.15)


def _arrow(ax, x1, y1, x2, y2):
    xm = (x1 + x2) / 2
    ax.plot([x1, xm, xm], [y1, y1, y2], color="black", lw=0.8)
    ax.annotate("", xy=(x2, y2), xytext=(xm, y2),
                arrowprops=dict(arrowstyle="-|>", lw=0.8, color="black",
                                mutation_scale=9, shrinkA=0, shrinkB=0))


def draw_tree(root, f0, fname, *, col_w=2.7, gap=0.5, out_w=4.3, row_h=1.25,
              box_h=0.86, depth_max=None):
    leaves = []
    _layout(root, 0, 0, leaves)
    depth_max = max(l["depth"] for l in leaves)
    n = len(leaves)
    x_out = (depth_max + 1) * (col_w + gap)
    width = x_out + out_w + 1.6
    fig, ax = plt.subplots(figsize=(width * 0.95, n * row_h + 0.7))
    ax.set_xlim(-0.1, width)
    ax.set_ylim(n + 0.1, -0.45)
    ax.axis("off")

    def rec(nd, prob):
        x = nd["depth"] * (col_w + gap)
        fc = "#e8e8e8" if nd["depth"] == 0 else "white"
        _box(ax, x, nd["y"] , col_w, box_h, nd["t"], fc=fc)
        if nd.get("ch"):
            for c in nd["ch"]:
                _arrow(ax, x + col_w, nd["y"], c["depth"] * (col_w + gap), c["y"])
                rec(c, prob * c["p"])
        else:
            _arrow(ax, x + col_w, nd["y"], x_out, nd["y"])
            _box(ax, x_out, nd["y"], out_w, box_h, f'{nd["code"]}. {nd["out"]}',
                 fc=nd.get("fc", "white"))
            ax.text(x_out + out_w + 0.12, nd["y"], sci(prob), ha="left",
                    va="center", fontsize=12.5, weight="bold")

    rec(root, f0)
    ax.text(x_out + out_w / 2, -0.2, "Исход", ha="center", fontsize=12.5,
            weight="bold")
    ax.text(x_out + out_w + 0.12, -0.2, "Частота,\n1/год", ha="left",
            fontsize=12.5, weight="bold", va="center")
    fig.savefig(os.path.join(HERE, fname), dpi=200, bbox_inches="tight")
    plt.close(fig)


RED, ORANGE, YELLOW, GREEN = "#f4b6b6", "#f9d9a8", "#fbf1b0", "#cfe8c6"

TREE1 = {
    "t": "Разгерметизация\nрезервуара\nс толуолом 500 м³\nf = 1·10⁻³ 1/год",
    "ch": [
        {"t": "Пролив удержан\nобвалованием\nP = 0,9", "p": 0.9, "ch": [
            {"t": "Сбор разлива\nв течение 2 ч\nP = 0,8", "p": 0.8,
             "code": "И1", "out": "Локальный пролив в каре,\nнезначительное испарение",
             "fc": GREEN},
            {"t": "Задержка сбора\n(> 2 ч)\nP = 0,2", "p": 0.2,
             "code": "И2", "out": "Выброс паров толуола,\nпревышение ПДК\nна границе СЗЗ",
             "fc": YELLOW},
        ]},
        {"t": "Разрушение или\nпереполнение\nобвалования\nP = 0,1", "p": 0.1, "ch": [
            {"t": "Источник зажигания\nотсутствует\nP = 0,7", "p": 0.7,
             "code": "И3", "out": "Растекание за территорию,\nзагрязнение почв и\nгрунтовых вод",
             "fc": ORANGE},
            {"t": "Воспламенение\nпролива\nP = 0,3", "p": 0.3, "ch": [
                {"t": "Пожар потушен\nАУПТ\nP = 0,8", "p": 0.8,
                 "code": "И4", "out": "Локальный пожар пролива,\nвыброс продуктов горения",
                 "fc": ORANGE},
                {"t": "Отказ\nпожаротушения\nP = 0,2", "p": 0.2,
                 "code": "И5", "out": "Каскадный пожар\n(«эффект домино»),\nтоксичное облако, жертвы",
                 "fc": RED},
            ]},
        ]},
    ],
}

TREE2 = {
    "t": "Отказ локальных\nочистных\nсооружений\nf = 5·10⁻² 1/год",
    "ch": [
        {"t": "Автоанализатор\nобнаружил\nпревышение\nP = 0,85", "p": 0.85, "ch": [
            {"t": "Сток переключён\nв аварийную\nёмкость\nP = 0,9", "p": 0.9,
             "code": "И6", "out": "Сброс предотвращён,\nповторная очистка",
             "fc": GREEN},
            {"t": "Аварийная ёмкость\nпереполнена\nP = 0,1", "p": 0.1,
             "code": "И7", "out": "Кратковременный сброс\n(до 10 м³) в коллектор",
             "fc": YELLOW},
        ]},
        {"t": "Превышение\nне обнаружено\nавтоматикой\nP = 0,15", "p": 0.15, "ch": [
            {"t": "Выявлено\nлабораторией\nза сутки\nP = 0,7", "p": 0.7,
             "code": "И8", "out": "Сброс до 200 м³ с\nпревышением НДС",
             "fc": YELLOW},
            {"t": "Не выявлено\nв течение суток\nP = 0,3", "p": 0.3,
             "code": "И9", "out": "Длительный сброс,\nзагрязнение реки,\nгибель гидробионтов",
             "fc": ORANGE},
        ]},
    ],
}


# ------------------------------------------------------------ матрица ---
MATRIX = [  # строки: частота сверху вниз; столбцы: тяжесть
    ["А", "А", "А", "С"],
    ["А", "А", "В", "С"],
    ["А", "В", "В", "С"],
    ["А", "В", "С", "Д"],
    ["В", "С", "С", "Д"],
]
CAT_COLOR = {"А": RED, "В": ORANGE, "С": YELLOW, "Д": GREEN}


def draw_matrix(placed, fname):
    rows = ["Часто\n(> 1)", "Вероятно\n(1…10⁻²)", "Возможно\n(10⁻²…10⁻⁴)",
            "Редко\n(10⁻⁴…10⁻⁶)", "Невероятно\n(< 10⁻⁶)"]
    cols = ["Катастрофическая", "Критическая", "Некритическая", "Пренебрежимая"]
    fig, ax = plt.subplots(figsize=(9.6, 5.6))
    ax.set_xlim(-1.6, 4)
    ax.set_ylim(5, -0.7)
    ax.axis("off")
    for r in range(5):
        ax.text(-0.08, r + 0.5, rows[r], ha="right", va="center", fontsize=9.5)
        for c in range(4):
            cat = MATRIX[r][c]
            ax.add_patch(plt.Rectangle((c, r), 1, 1, fc=CAT_COLOR[cat], ec="black",
                                       lw=0.8))
            ax.text(c + 0.06, r + 0.1, cat, fontsize=12, weight="bold",
                    va="top", color="#555555")
    for c in range(4):
        ax.text(c + 0.5, -0.12, cols[c], ha="center", va="bottom", fontsize=9.5,
                weight="bold")
    ax.text(1.8, -0.55, "Тяжесть последствий", ha="center", fontsize=10.5,
            weight="bold")
    ax.text(-1.5, -0.3, "Частота, 1/год", fontsize=10, weight="bold")
    cells = {}
    for code, r, c in placed:
        cells.setdefault((r, c), []).append(code)
    for (r, c), codes in cells.items():
        lines = []
        for i in range(0, len(codes), 3):
            lines.append(", ".join(codes[i:i + 3]))
        ax.text(c + 0.55, r + 0.58, "\n".join(lines), ha="center", va="center",
                fontsize=9.5, weight="bold")
    fig.savefig(os.path.join(HERE, fname), dpi=200, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    draw_tree(TREE1, 1e-3, "pr2_tree1.png")
    draw_tree(TREE2, 5e-2, "pr2_tree2.png")
