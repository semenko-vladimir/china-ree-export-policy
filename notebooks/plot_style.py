from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
from cycler import cycler


ACADEMIC_COLORS = [
    "#1f4e79",
    "#8c5a2b",
    "#3f7f5f",
    "#6f5f90",
    "#7a7a7a",
    "#b05a4a",
    "#4f6f8f",
    "#9a7d34",
    "#5f7f7f",
    "#8f6f6f",
    "#2f5f5f",
    "#5f5f7f",
]

PARTNER_LABELS_RU = {
    "Austria": "Австрия",
    "Canada": "Канада",
    "France": "Франция",
    "Germany": "Германия",
    "Hong Kong, China": "Гонконг",
    "India": "Индия",
    "Italy": "Италия",
    "Japan": "Япония",
    "Korea, Rep.": "Республика Корея",
    "Malaysia": "Малайзия",
    "Netherlands": "Нидерланды",
    "Russian Federation": "Россия",
    "Singapore": "Сингапур",
    "Switzerland": "Швейцария",
    "Thailand": "Таиланд",
    "United Kingdom": "Великобритания",
    "United States": "США",
    "Vietnam": "Вьетнам",
}

EXPOSURE_GROUP_LABELS_RU = {
    "HighExposure top15": "Группа высокой исходной экспозиции (топ-15)",
    "Other countries": "Прочие страны",
}

PRODUCT_LABELS_RU = {
    "284690": "HS 284690: соединения редкоземельных металлов",
    "280530": "HS 280530: редкоземельные металлы, скандий и иттрий",
    "850511": "HS 850511: постоянные магниты",
    "HS 284690 — rare-earth compounds": "HS 284690: соединения редкоземельных металлов",
    "HS 280530 — rare-earth metals, scandium and yttrium": (
        "HS 280530: редкоземельные металлы, скандий и иттрий"
    ),
    "HS 850511 — permanent magnets": "HS 850511: постоянные магниты",
}


def set_plot_style() -> None:
    """Применить сдержанный академический стиль matplotlib с поддержкой кириллицы."""
    plt.style.use("default")
    plt.rcParams.update({
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "savefig.facecolor": "white",
        "font.family": "DejaVu Sans",
        "axes.unicode_minus": False,
        "axes.prop_cycle": cycler(color=ACADEMIC_COLORS),
        "axes.edgecolor": "#333333",
        "axes.linewidth": 0.8,
        "axes.titlesize": 13,
        "axes.titleweight": "normal",
        "axes.labelsize": 11,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "legend.fontsize": 9.5,
        "legend.title_fontsize": 10,
        "lines.linewidth": 1.8,
        "lines.markersize": 5,
        "grid.color": "#d9d9d9",
        "grid.linewidth": 0.7,
        "grid.linestyle": "-",
    })


def translate_label(label: str) -> str:
    """Перевести видимые динамические подписи в легендах графиков."""
    return PARTNER_LABELS_RU.get(
        label,
        EXPOSURE_GROUP_LABELS_RU.get(label, label),
    )


def product_label_ru(product_label: str) -> str:
    """Вернуть русскую подпись продукта для известного HS-кода или текстовой метки."""
    text = str(product_label)
    if text in PRODUCT_LABELS_RU:
        return PRODUCT_LABELS_RU[text]
    for hs_code, label in PRODUCT_LABELS_RU.items():
        if hs_code.isdigit() and hs_code in text:
            return label
    return text


def add_period_markers(ax) -> None:
    """Добавить общие вертикальные маркеры переходного периода."""
    ax.axvline(2015, linestyle="--", linewidth=1.0, color="#666666", alpha=0.85)
    ax.axvline(2016, linestyle=":", linewidth=1.2, color="#666666", alpha=0.9)


def format_axis(
    ax,
    title: str,
    xlabel: str = "Год",
    ylabel: Optional[str] = None,
    legend: bool = False,
    legend_outside: bool = False,
):
    """Единообразно оформить оси для графиков в академическом стиле."""
    ax.set_title(title, loc="left", pad=12)
    ax.set_xlabel(xlabel)
    if ylabel:
        ax.set_ylabel(ylabel)

    ax.grid(True, axis="y", alpha=0.8)
    ax.grid(False, axis="x")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="both", colors="#333333", length=3)

    if legend:
        if legend_outside:
            ax.legend(
                bbox_to_anchor=(1.02, 1),
                loc="upper left",
                frameon=False,
                borderaxespad=0,
            )
        else:
            ax.legend(frameon=False)

    return ax


def save_figure(path, fig=None) -> str:
    """Сохранить график в PNG с 300 dpi и добавить рядом копии SVG/PDF."""
    output = Path(path)
    fig = fig or plt.gcf()
    fig.tight_layout()
    fig.savefig(output, dpi=300, bbox_inches="tight")
    for suffix in (".svg", ".pdf"):
        fig.savefig(output.with_suffix(suffix), bbox_inches="tight")
    return str(output)
