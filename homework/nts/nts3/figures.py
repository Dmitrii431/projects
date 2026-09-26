# -*- coding: utf-8 -*-
"""Графики и «скриншоты» кода и вывода для презентации НТС-3."""
import csv
import re

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from pygments import highlight
from pygments.formatters import ImageFormatter
from pygments.lexers import PythonLexer

import hypotheses as H  # noqa: E402  (выполняет расчёт, печатает результаты)

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 12})
BLUE, ORANGE, GREEN, RED, GREY = "#2E6FB7", "#E08A1E", "#3E9E52", "#B03A3A", "#9AA5B1"


def fig_phenomenon():
    years = sorted(H.it)
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 4.6),
                                 gridspec_kw={"width_ratios": [1.3, 1]})
    a1.plot(years, [H.it[y] for y in years], "-o", color=BLUE, lw=2.5,
            label="IT-сектор (J62–J63)")
    a1.plot(years, [H.al[y] for y in years], "-o", color=GREY, lw=2.5,
            label="Вся экономика")
    for y in years:
        a1.annotate(f"{H.it[y]:.1f}", (y, H.it[y]), textcoords="offset points",
                    xytext=(0, 8), ha="center", color=BLUE)
        a1.annotate(f"{H.al[y]:.1f}", (y, H.al[y]), textcoords="offset points",
                    xytext=(0, 8), ha="center", color="#555")
    a1.set_ylim(0, 80)
    a1.set_ylabel("% предприятий, использующих ИИ")
    a1.set_title("Использование ИИ, ЕС-27")
    a1.legend(loc="upper left", frameon=False)
    a1.spines[["top", "right"]].set_visible(False)
    ec = {r["time"]: float(r["value"]) for r in H.trend
          if r["nace_r2"] == H.IT and r["indic_is"] == "E_AI_EC"}
    ys = ["2023", "2024", "2025"]
    use = np.array([H.it[y] for y in ys])
    con = np.array([ec[y] for y in ys])
    never = 100 - use - con
    a2.bar(ys, use, color=BLUE, label="Используют ИИ")
    a2.bar(ys, con, bottom=use, color=ORANGE,
           label="Рассматривали, но отказались")
    a2.bar(ys, never, bottom=use + con, color="#D9DEE4",
           label="Не рассматривали")
    for i, y in enumerate(ys):
        a2.text(i, use[i] + con[i] / 2, f"{con[i]:.1f}", ha="center",
                va="center", color="white", weight="bold")
        a2.text(i, use[i] / 2, f"{use[i]:.1f}", ha="center", va="center",
                color="white", weight="bold")
    a2.set_title("IT-сектор ЕС: структура, %")
    a2.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=1,
              frameon=False, fontsize=10)
    a2.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig("fig_phenomenon.png", dpi=180)
    plt.close(fig)


def fig_barriers():
    ks = sorted(H.BARRIERS, key=lambda k: H.eu[k]["2025"])
    fig, ax = plt.subplots(figsize=(11, 5))
    y = np.arange(len(ks))
    ax.barh(y - 0.18, [H.eu[k]["2023"] for k in ks], 0.36, color="#C9D3DE",
            label="2023")
    ax.barh(y + 0.18, [H.eu[k]["2025"] for k in ks], 0.36,
            color=[ORANGE if k in H.TECH else BLUE for k in ks], label="2025")
    for i, k in enumerate(ks):
        ax.text(H.eu[k]["2025"] + 0.8, i + 0.18, f"{H.eu[k]['2025']:.1f}",
                va="center", fontsize=11)
    ax.set_yticks(y, [("Г1 " if k in H.TECH else "Г2 ") + H.BARRIERS[k]
                      for k in ks])
    ax.set_xlabel("% предприятий IT-сектора, рассматривавших ИИ и отказавшихся")
    ax.set_title("Причины отказа от ИИ, IT-сектор ЕС-27")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color="#C9D3DE", label="2023"),
                       Patch(color=ORANGE, label="2025, Г1 технологии"),
                       Patch(color=BLUE, label="2025, Г2 люди и среда")],
              frameon=False, loc="lower right")
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig("fig_barriers.png", dpi=180)
    plt.close(fig)


def fig_shift():
    ks = ["E_AI_BLEG", "E_AI_BCDP", "E_AI_BINC", "E_AI_BDDT"]
    fig, axs = plt.subplots(1, 4, figsize=(13, 4.4), sharey=True)
    for ax, k in zip(axs, ks):
        a = H.pick([r for r in H.ctry if r["time"] == "2023"
                    and r["geo"] not in H.AGG], indic_is=k, unit="PC_ENT_AI_EC")
        b = H.pick([r for r in H.ctry if r["time"] == "2025"
                    and r["geo"] not in H.AGG], indic_is=k, unit="PC_ENT_AI_EC")
        gg = sorted(set(a) & set(b))
        for g in gg:
            c = RED if b[g] > a[g] else GREEN
            ax.plot([0, 1], [a[g], b[g]], "-o", color=c, alpha=0.6, ms=4)
        ax.plot([0, 1], [np.median([a[g] for g in gg]),
                         np.median([b[g] for g in gg])], "-o", color="black",
                lw=3, ms=7)
        ax.set_xticks([0, 1], ["2023", "2025"])
        ax.set_title(H.BARRIERS[k], fontsize=12)
        ax.spines[["top", "right"]].set_visible(False)
    axs[0].set_ylabel("% отказавшихся по причине")
    fig.suptitle("Изменение барьеров по странам 2023→2025 (красный — рост, "
                 "зелёный — снижение, чёрный — медиана)", fontsize=12)
    fig.tight_layout()
    fig.savefig("fig_shift.png", dpi=180)
    plt.close(fig)


def fig_g4():
    R = np.load("corr_matrix.npy")
    names = [H.BARRIERS[k] for k in H.BARRIERS]
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(13, 5.2),
                                 gridspec_kw={"width_ratios": [1.15, 1]})
    im = a1.imshow(R, cmap="RdBu_r", vmin=-1, vmax=1)
    a1.set_xticks(range(len(names)), names, rotation=50, ha="right",
                  fontsize=10)
    a1.set_yticks(range(len(names)), names, fontsize=10)
    for i in range(len(names)):
        for j in range(len(names)):
            a1.text(j, i, f"{R[i, j]:.2f}", ha="center", va="center",
                    fontsize=8.5)
    a1.set_title("Связь барьеров между собой (Спирмен, n = 18)")
    fig.colorbar(im, ax=a1, fraction=0.046)
    ks = sorted(H.rs_use, key=lambda k: H.rs_use[k][0])
    vals = [H.rs_use[k][0] for k in ks]
    cols = [RED if H.rs_use[k][1] < 0.05 else GREY for k in ks]
    a2.barh(range(len(ks)), vals, color=cols)
    a2.set_yticks(range(len(ks)), [H.BARRIERS[k] for k in ks], fontsize=10)
    a2.axvline(0, color="black", lw=0.8)
    a2.set_xlim(-0.8, 0.8)
    a2.set_title("Барьер ↔ использование ИИ\n(красный — p < 0,05)")
    a2.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig("fig_g4.png", dpi=180)
    plt.close(fig)


def code_shot(src, start, end, out):
    """Скриншот фрагмента кода (строки start..end) с подсветкой."""
    lines = open(src).read().splitlines()[start - 1:end]
    fmt = ImageFormatter(font_name="DejaVu Sans Mono", font_size=15,
                         line_numbers=True, line_number_start=start,
                         style="friendly", image_pad=12)
    open(out, "wb").write(highlight("\n".join(lines), PythonLexer(), fmt))


def console_shot(text, out, title="python3 hypotheses.py"):
    """Скриншот вывода программы в стиле окна терминала."""
    font = ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf", 17)
    lines = [f"$ {title}"] + text.splitlines()
    w = max(font.getlength(l) for l in lines) + 40
    h = 22 * len(lines) + 60
    img = Image.new("RGB", (int(w), h), "#1E1E1E")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 30], fill="#3C3C3C")
    for i, c in enumerate(("#FF5F56", "#FFBD2E", "#27C93F")):
        d.ellipse([12 + i * 22, 9, 26 + i * 22, 23], fill=c)
    for i, l in enumerate(lines):
        col = "#9CDCFE" if l.startswith("===") else (
            "#6A9955" if l.startswith("$") else "#D4D4D4")
        if re.search(r"p = 0\.0[0-4]\d*\s*\(значимо\)|p = 0\.00", l):
            col = "#F48771"
        d.text((20, 42 + i * 22), l, font=font, fill=col)
    img.save(out)


def section(text, a, b):
    out, on = [], False
    for l in text.splitlines():
        if l.startswith("=== " + a):
            on = True
        elif b and l.startswith("=== " + b):
            on = False
        if on:
            out.append(l)
    return "\n".join(out).strip()


if __name__ == "__main__":
    fig_phenomenon()
    fig_barriers()
    fig_shift()
    fig_g4()
    code_shot("fetch_data.py", 16, 42, "code_fetch.png")
    src = open("hypotheses.py").read().splitlines()
    s1 = next(i for i, l in enumerate(src, 1) if "по странам, 2025" in l)
    s2 = next(i for i, l in enumerate(src, 1) if "Г3: смещение" in l)
    code_shot("hypotheses.py", s1, s2 - 2, "code_g12.png")
    s4 = next(i for i, l in enumerate(src, 1) if "Г4: «оцени" in l)
    code_shot("hypotheses.py", s2, len(src), "code_g34.png")
    res = open("results.txt").read()
    console_shot(section(res, "ФЕНОМЕН", "Г1"), "out_phen.png")
    console_shot(section(res, "Г1", "Г3"), "out_g12.png")
    console_shot(section(res, "Г3", None), "out_g34.png")
    print("figures ok")
