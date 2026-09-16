# -*- coding: utf-8 -*-
"""Доска проекта внутреннего аудита СЭМ в цифровом планере (Яндекс Трекер).
Рисует изображение доски для вставки в отчёт как рис. 1."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

matplotlib.rcParams["font.family"] = "DejaVu Sans"

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "board_audit.png")

BG = "#f2f4f7"
HEAD = "#ffffff"
CARD = "#ffffff"
LINE = "#dfe3e9"
TXT = "#21242b"
MUTED = "#7b808a"
ACCENT = {"Подготовка к аудиту": "#ffcc00",
          "Проведение аудита": "#4b7bec",
          "Отчетность": "#26b47f"}

COLUMNS = [
    ("Подготовка к аудиту", [
        ("ECO-1", "Утверждение программы внутреннего аудита СЭМ", "Рослов Д. С.", "01.03.2027", "Выполнено"),
        ("ECO-2", "Анализ документации СЭМ цеха поверхностного монтажа", "Логинов В. А.", "05.03.2027", "Выполнено"),
        ("ECO-3", "Разработка контрольного листа (10 вопросов)", "Макаров З. И.", "10.03.2027", "Выполнено"),
        ("ECO-4", "Согласование сроков и состава группы с цехом", "Цветков А. Э.", "12.03.2027", "Выполнено"),
    ]),
    ("Проведение аудита", [
        ("ECO-5", "Вводное совещание с руководством цеха", "Рослов Д. С.", "15.03.2027", "Выполнено"),
        ("ECO-6", "Обход участка пайки и отмывки печатных плат", "Макаров З. И.", "16.03.2027", "Выполнено"),
        ("ECO-7", "Проверка обращения с отходами и химвеществами", "Цветков А. Э.", "17.03.2027", "Выполнено"),
        ("ECO-8", "Интервью с персоналом цеха", "Логинов В. А.", "18.03.2027", "Выполнено"),
        ("ECO-9", "Сбор и фиксация объективных свидетельств", "Рудяк", "18.03.2027", "Выполнено"),
    ]),
    ("Отчетность", [
        ("ECO-10", "Оформление листа несоответствий", "Рослов Д. С.", "19.03.2027", "Выполнено"),
        ("ECO-11", "Подготовка отчета по аудиту", "Рослов Д. С.", "22.03.2027", "Выполнено"),
        ("ECO-12", "Заключительное совещание, выдача отчета", "Рослов Д. С.", "23.03.2027", "Выполнено"),
        ("ECO-13", "Контроль выполнения корректирующих действий", "Макаров З. И.", "30.06.2027", "В работе"),
    ]),
]


def wrap(text, width):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if len(t) <= width:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def main():
    W, H = 15.6, 7.4
    fig = plt.figure(figsize=(W, H), dpi=160)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0), 100, 100, color=BG, zorder=0))

    # шапка планера
    ax.add_patch(Rectangle((0, 90.0), 100, 10.0, color=HEAD, zorder=1))
    ax.plot([0, 100], [90.0, 90.0], color=LINE, lw=1, zorder=2)
    ax.text(2.2, 96.4, "Яндекс Трекер", fontsize=11, color=MUTED, zorder=3)
    ax.text(2.2, 92.4, "ЭКОАУДИТ-2027 · Внутренний аудит СЭМ · ООО «ТехноПром»",
            fontsize=14, color=TXT, fontweight="bold", zorder=3)
    ax.text(97.8, 93.2, "Доска  ·  Спринт 1  ·  01.03.2027 – 30.06.2027",
            fontsize=10, color=MUTED, ha="right", zorder=3)

    col_w, gap, left = 31.0, 2.0, 1.8
    top = 87.0

    for ci, (title, cards) in enumerate(COLUMNS):
        x = left + ci * (col_w + gap)
        ax.add_patch(FancyBboxPatch((x, 2.0), col_w, top - 2.0,
                                    boxstyle="round,pad=0,rounding_size=1.2",
                                    facecolor="#e9ecf1", edgecolor=LINE,
                                    lw=1, zorder=1))
        ax.add_patch(Rectangle((x + 1.0, top - 3.0), 1.0, 2.0,
                               color=ACCENT[title], zorder=3))
        ax.text(x + 2.8, top - 2.6, title.upper(), fontsize=11.5,
                color=TXT, fontweight="bold", zorder=3)
        ax.text(x + col_w - 1.2, top - 2.5, str(len(cards)), fontsize=10.5,
                color=MUTED, ha="right", zorder=3)

        y = top - 5.2
        for key, name, who, date, status in cards:
            lines = wrap(name, 34)
            ch = 7.4 + 3.1 * len(lines)
            ax.add_patch(FancyBboxPatch((x + 1.0, y - ch), col_w - 2.0, ch,
                                        boxstyle="round,pad=0,rounding_size=0.9",
                                        facecolor=CARD, edgecolor=LINE,
                                        lw=1, zorder=3))
            ax.text(x + 2.0, y - 2.0, key, fontsize=9, color=ACCENT[title],
                    fontweight="bold", zorder=4)
            st_col = "#26b47f" if status == "Выполнено" else "#e08a1e"
            ax.text(x + col_w - 2.0, y - 2.0, status, fontsize=8.5,
                    color=st_col, ha="right", zorder=4)
            for li, ln in enumerate(lines):
                ax.text(x + 2.0, y - 5.8 - 3.1 * li, ln, fontsize=10,
                        color=TXT, zorder=4)
            ax.text(x + 2.0, y - ch + 1.8, who + "   ·   до " + date,
                    fontsize=8.5, color=MUTED, zorder=4)
            y -= ch + 1.8

    fig.savefig(OUT, dpi=160)
    print("saved", OUT)


if __name__ == "__main__":
    main()
