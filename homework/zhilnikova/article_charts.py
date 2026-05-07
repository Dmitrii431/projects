"""Генерация PNG-графиков для научной статьи."""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

OUT_DIR = os.path.dirname(__file__)
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["font.size"] = 11


def chart_publications_dynamics():
    """Динамика публикаций по теме за 2015-2025 гг."""
    years = list(range(2015, 2026))
    russian = [9, 11, 14, 17, 22, 28, 31, 36, 42, 47, 39]
    foreign = [19, 23, 27, 39, 56, 67, 87, 106, 126, 142, 117]
    fig, ax = plt.subplots(figsize=(7.5, 4.0))
    ax.plot(years, russian, marker="o", color="#2E5077",
            linewidth=2, label="Российские источники")
    ax.plot(years, foreign, marker="s", color="#D87C4A",
            linewidth=2, label="Зарубежные источники")
    ax.fill_between(years, russian, alpha=0.15, color="#2E5077")
    ax.fill_between(years, foreign, alpha=0.15, color="#D87C4A")
    ax.set_xlabel("Год публикации")
    ax.set_ylabel("Количество публикаций, шт.")
    ax.set_xticks(years)
    ax.legend(loc="upper left")
    ax.grid(True, linestyle=":", alpha=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    path = os.path.join(OUT_DIR, "fig1_dynamics.png")
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return path


def chart_directions():
    """Распределение источников по тематическим направлениям."""
    labels = [
        "Гибридные нечёткие\nметоды",
        "Машинное\nобучение",
        "Количественные\nстатистические",
        "Этические и\nрегуляторные",
        "Качественные\nметоды",
        "Стандарты и\nфреймворки",
    ]
    values = [31, 28, 22, 19, 18, 16]
    colors = ["#2E5077", "#79B4B7", "#D87C4A", "#9B5DE5",
              "#F4A261", "#8AB17D"]
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    bars = ax.barh(labels, values, color=colors, edgecolor="#222", linewidth=0.6)
    ax.invert_yaxis()
    ax.set_xlabel("Количество публикаций, шт.")
    ax.set_xlim(0, max(values) + 6)
    ax.grid(True, axis="x", linestyle=":", alpha=0.6)
    ax.set_axisbelow(True)
    for bar, v in zip(bars, values):
        ax.text(v + 0.5, bar.get_y() + bar.get_height() / 2,
                f"{v} ({v/sum(values)*100:.1f}%)",
                va="center", fontsize=10)
    fig.tight_layout()
    path = os.path.join(OUT_DIR, "fig2_directions.png")
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return path


def chart_patents_geography():
    """География патентной активности."""
    labels = ["США", "Китай", "ЕС (DE/FR/UK)",
              "Япония и Корея", "Россия", "Прочие"]
    values = [18, 11, 9, 5, 3, 1]
    colors = ["#2E5077", "#D62828", "#003F88",
              "#F4A261", "#8AB17D", "#9C9C9C"]
    fig, ax = plt.subplots(figsize=(6.5, 4.5))
    wedges, texts, autotexts = ax.pie(
        values, labels=labels, colors=colors, autopct="%1.1f%%",
        startangle=90, wedgeprops=dict(edgecolor="white", linewidth=1.5),
        textprops=dict(fontsize=11),
    )
    for t in autotexts:
        t.set_color("white")
        t.set_fontweight("bold")
        t.set_fontsize(10)
    ax.set_aspect("equal")
    fig.tight_layout()
    path = os.path.join(OUT_DIR, "fig3_patents.png")
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return path


if __name__ == "__main__":
    for f in (chart_publications_dynamics, chart_directions, chart_patents_geography):
        p = f()
        print(f"OK: {p}")
