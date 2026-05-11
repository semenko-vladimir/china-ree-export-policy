
# trading_partner_dynamics_analysis_all_codes.py
# Создает таблицы и графики динамики торговых партнеров для HS 284690, HS 280530 и HS 850511.
#
# Входные данные: data/
# Выходные данные: results/trading_partner_dynamics_all_hs_codes.xlsx
#                 figures/  (все графики)

import os
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, str(Path(__file__).parent.parent / "notebooks"))

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from plot_style import (
    add_period_markers,
    format_axis,
    product_label_ru,
    save_figure,
    set_plot_style,
    translate_label,
)


# -----------------------------
# Конфигурация
# -----------------------------

ROOT = Path(__file__).parent.parent
BASE_DIR   = ROOT / "data"
OUTPUT_DIR = ROOT / "figures"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = ROOT / "results" / "trading_partner_dynamics_all_hs_codes.xlsx"

SELECTED_YEARS = [2014, 2016, 2018, 2020, 2024]

set_plot_style()

SPECIAL_PARTNERS_TO_DROP = [
    "Other Asia, nes",
    "Unspecified",
    "Areas, nes",
    "World",
]

HS_CONFIGS = {
    "284690": {
        "input_file": "clean_china_ree_exports_284690_2010_2024.xlsx",
        "product_label": "HS 284690: соединения редкоземельных металлов",
        "sheet_candidates": ["final_panel_balanced", "panel_with_controls"],
    },
    "280530": {
        "input_file": "china_ree_exports_280530_with_controls.xlsx",
        "product_label": "HS 280530: редкоземельные металлы, скандий и иттрий",
        "sheet_candidates": ["final_panel_balanced", "panel_with_controls"],
    },
    "850511": {
        "input_file": "china_ree_exports_850511_with_controls.xlsx",
        "product_label": "HS 850511: постоянные магниты",
        "sheet_candidates": ["panel_with_controls", "final_panel_balanced"],
    },
}


# -----------------------------
# Загрузка и очистка
# -----------------------------

def choose_sheet(path: Path, candidates: list[str]) -> str:
    """Выбрать первый доступный лист из списка кандидатов."""
    sheets = pd.ExcelFile(path).sheet_names
    for sheet in candidates:
        if sheet in sheets:
            return sheet
    raise ValueError(
        f"В файле {path.name} нет подходящего листа. "
        f"Ожидался один из {candidates}; доступные листы: {sheets}"
    )


def load_panel(hs_code: str, config: dict) -> pd.DataFrame:
    """Загрузить и стандартизировать панель для одного HS-кода."""
    input_path = BASE_DIR / config["input_file"]
    if not input_path.exists():
        raise FileNotFoundError(
            f"Не найден входной файл для HS {hs_code}: {input_path}. "
            "Поместите файл в папку данных проекта."
        )

    sheet = choose_sheet(input_path, config["sheet_candidates"])
    df = pd.read_excel(input_path, sheet_name=sheet)

    required = {"partner", "year"}
    missing_required = required - set(df.columns)
    if missing_required:
        raise ValueError(f"В файле {input_path.name} нет колонок: {missing_required}")

    df["hs_code"] = hs_code
    df["product_label"] = config["product_label"]
    df["source_file"] = input_path.name
    df["source_sheet"] = sheet

    # Стандартизируем стоимость.
    if "export_value_usd" not in df.columns:
        if "export_value_1000usd" in df.columns:
            df["export_value_usd"] = pd.to_numeric(df["export_value_1000usd"], errors="coerce") * 1000
        elif "export_value_kusd" in df.columns:
            df["export_value_usd"] = pd.to_numeric(df["export_value_kusd"], errors="coerce") * 1000
        else:
            raise ValueError(f"В файле {input_path.name} нет export_value_usd или эквивалента в тыс. долл. США.")

    # Стандартизируем физический объем.
    if "quantity_kg" not in df.columns:
        if "quantity" in df.columns:
            df["quantity_kg"] = pd.to_numeric(df["quantity"], errors="coerce")
        elif "Quantity" in df.columns:
            df["quantity_kg"] = pd.to_numeric(df["Quantity"], errors="coerce")
        else:
            df["quantity_kg"] = np.nan

    df["partner"] = df["partner"].astype(str).str.strip()
    df["year"] = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["year"]).copy()
    df["year"] = df["year"].astype(int)

    df["export_value_usd"] = pd.to_numeric(df["export_value_usd"], errors="coerce").fillna(0)
    df["quantity_kg"] = pd.to_numeric(df["quantity_kg"], errors="coerce").fillna(0)

    df = df[~df["partner"].isin(SPECIAL_PARTNERS_TO_DROP)].copy()

    return df


def enrich_panel(df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитать доли, ранги, индикаторы положительных потоков и единицы стоимости/объема."""
    df = df.copy()

    df["positive_flow"] = (df["export_value_usd"] > 0).astype(int)

    df["total_value_clean"] = df.groupby("year")["export_value_usd"].transform("sum")
    df["total_quantity_clean"] = df.groupby("year")["quantity_kg"].transform("sum")

    df["share_clean"] = np.where(
        df["total_value_clean"] > 0,
        df["export_value_usd"] / df["total_value_clean"],
        np.nan,
    )

    df["quantity_share_clean"] = np.where(
        df["total_quantity_clean"] > 0,
        df["quantity_kg"] / df["total_quantity_clean"],
        np.nan,
    )

    df["rank_value"] = (
        df.groupby("year")["export_value_usd"]
        .rank(method="first", ascending=False)
        .astype(int)
    )

    df["export_value_musd"] = df["export_value_usd"] / 1_000_000
    df["share_pct"] = df["share_clean"] * 100
    df["quantity_mkg"] = df["quantity_kg"] / 1_000_000

    return df


# -----------------------------
# Таблицы
# -----------------------------

def make_top10_tables(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    top10_all = (
        df[df["rank_value"] <= 10]
        .sort_values(["year", "rank_value"])
        .copy()
    )

    top10_all = top10_all[
        [
            "hs_code", "product_label", "year", "rank_value", "partner",
            "export_value_musd", "share_pct", "quantity_mkg",
        ]
    ].rename(columns={"rank_value": "rank"})

    top10_all["partner"] = top10_all["partner"].map(translate_label)
    top10_selected = top10_all[top10_all["year"].isin(SELECTED_YEARS)].copy()
    return top10_all, top10_selected


def make_churn_table(df: pd.DataFrame, top10_all: pd.DataFrame) -> pd.DataFrame:
    sets_by_year = {
        year: set(top10_all[top10_all["year"] == year]["partner"])
        for year in sorted(df["year"].unique())
    }

    churn_rows = []
    previous = None

    for year, current in sets_by_year.items():
        if previous is None:
            entrants = current
            exits = set()
            retained = np.nan
            jaccard = np.nan
        else:
            entrants = current - previous
            exits = previous - current
            retained = len(current & previous)
            jaccard = 1 - retained / len(current | previous) if len(current | previous) else np.nan

        churn_rows.append({
            "hs_code": df["hs_code"].iloc[0],
            "product_label": df["product_label"].iloc[0],
            "year": year,
            "top10_partners": "; ".join(top10_all[top10_all["year"] == year]["partner"].tolist()),
            "entrants_vs_previous_year": "; ".join(sorted(entrants)),
            "exits_vs_previous_year": "; ".join(sorted(exits)),
            "num_entrants": len(entrants),
            "num_exits": len(exits),
            "num_retained": retained,
            "jaccard_churn_top10": jaccard,
        })

        previous = current

    return pd.DataFrame(churn_rows)


def make_yearly_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = df.groupby("year").agg(
        positive_partners=("positive_flow", "sum"),
        total_export_value_musd=("export_value_usd", lambda s: s.sum() / 1_000_000),
        total_quantity_mkg=("quantity_kg", lambda s: s.sum() / 1_000_000),
        top10_share_pct=("share_clean", lambda s: s.nlargest(10).sum() * 100),
        top5_share_pct=("share_clean", lambda s: s.nlargest(5).sum() * 100),
        hhi=("share_clean", lambda s: (s ** 2).sum()),
    ).reset_index()

    summary.insert(0, "hs_code", df["hs_code"].iloc[0])
    summary.insert(1, "product_label", df["product_label"].iloc[0])
    return summary


def choose_rank_partners(df: pd.DataFrame, max_partners: int = 12) -> list[str]:
    """
    Динамически выбрать партнеров для графиков рангов:
    - ведущие партнеры в выбранные годы;
    - ведущие докризисные партнеры по средней стоимости экспорта;
    - ограничение до max_partners по средней стоимости за весь период.
    """
    selected = set()

    for year in SELECTED_YEARS:
        sub = df[df["year"] == year].sort_values("export_value_usd", ascending=False).head(5)
        selected.update(sub["partner"].tolist())

    baseline = (
        df[df["year"].between(2010, 2014)]
        .groupby("partner")["export_value_usd"]
        .mean()
        .sort_values(ascending=False)
        .head(max_partners)
    )
    selected.update(baseline.index.tolist())

    ranking = (
        df[df["partner"].isin(selected)]
        .groupby("partner")["export_value_usd"]
        .mean()
        .sort_values(ascending=False)
    )

    return ranking.head(max_partners).index.tolist()


def make_rank_trajectory(df: pd.DataFrame) -> pd.DataFrame:
    partners = choose_rank_partners(df, max_partners=12)

    rank_traj = df[df["partner"].isin(partners)].copy()
    rank_traj["rank_display"] = rank_traj["rank_value"].where(rank_traj["export_value_usd"] > 0, np.nan)

    return rank_traj[
        [
            "hs_code", "product_label", "year", "partner", "rank_value", "rank_display",
            "export_value_musd", "share_pct", "quantity_mkg",
        ]
    ].sort_values(["partner", "year"])


# -----------------------------
# Графики
# -----------------------------

def save_rank_dynamics_plot(rank_traj: pd.DataFrame, hs_code: str, product_label: str) -> str:
    output = OUTPUT_DIR / f"trading_partners_rank_dynamics_hs{hs_code}.png"

    fig, ax = plt.subplots(figsize=(11, 7))
    for partner in rank_traj["partner"].drop_duplicates():
        sub = rank_traj[(rank_traj["partner"] == partner) & rank_traj["rank_display"].notna()]
        if len(sub):
            ax.plot(sub["year"], sub["rank_display"], marker="o", label=translate_label(partner))

    ax.invert_yaxis()
    max_rank = int(min(25, max(10, np.nanmax(rank_traj["rank_display"]) if rank_traj["rank_display"].notna().any() else 10)))
    ax.set_yticks(range(1, max_rank + 1))
    add_period_markers(ax)
    format_axis(
        ax,
        title=f"Динамика рангов выбранных импортёров, {product_label_ru(product_label)}",
        ylabel="Ранг по стоимости экспорта (1 = крупнейший)",
        legend=True,
        legend_outside=True,
    )
    save_figure(output, fig)
    plt.close(fig)

    return str(output)


def save_churn_plot(churn: pd.DataFrame, hs_code: str, product_label: str) -> str:
    output = OUTPUT_DIR / f"top10_churn_hs{hs_code}.png"

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(churn["year"], churn["num_entrants"], marker="o", label="Вход в топ-10")
    ax.plot(churn["year"], churn["num_exits"], marker="o", label="Выход из топ-10")
    add_period_markers(ax)
    format_axis(
        ax,
        title=f"Обновление состава топ-10 импортёров, {product_label_ru(product_label)}",
        ylabel="Количество партнёров",
        legend=True,
    )
    save_figure(output, fig)
    plt.close(fig)

    return str(output)


def save_positive_partners_plot(summary: pd.DataFrame, hs_code: str, product_label: str) -> str:
    output = OUTPUT_DIR / f"positive_partners_count_hs{hs_code}.png"

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(summary["year"], summary["positive_partners"], marker="o")
    add_period_markers(ax)
    format_axis(
        ax,
        title=f"Количество импортёров с положительным экспортом, {product_label_ru(product_label)}",
        ylabel="Количество импортёров",
    )
    save_figure(output, fig)
    plt.close(fig)

    return str(output)


def save_concentration_plot(summary: pd.DataFrame, hs_code: str, product_label: str) -> str:
    output = OUTPUT_DIR / f"concentration_hhi_top10_share_hs{hs_code}.png"

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(summary["year"], summary["hhi"], marker="o", label="HHI")
    add_period_markers(ax)
    format_axis(
        ax,
        title=f"Концентрация экспортных направлений, {product_label_ru(product_label)}",
        ylabel="Индекс Херфиндаля-Хиршмана (HHI)",
        legend=True,
    )
    save_figure(output, fig)
    plt.close(fig)

    return str(output)


def save_cross_code_plot(all_summary: pd.DataFrame, metric: str, ylabel: str, filename: str, title: str) -> str:
    output = OUTPUT_DIR / filename

    fig, ax = plt.subplots(figsize=(10, 6))
    for hs_code, sub in all_summary.groupby("hs_code"):
        label = f"HS {hs_code}"
        ax.plot(sub["year"], sub[metric], marker="o", label=label)

    add_period_markers(ax)
    format_axis(
        ax,
        title=title,
        ylabel=ylabel,
        legend=True,
    )
    save_figure(output, fig)
    plt.close(fig)

    return str(output)


# -----------------------------
# Основной запуск
# -----------------------------

def main():
    all_top10 = []
    all_top10_selected = []
    all_churn = []
    all_summary = []
    all_rank_traj = []
    chart_inventory = []

    for hs_code, config in HS_CONFIGS.items():
        print(f"Обработка HS {hs_code}...")

        df = load_panel(hs_code, config)
        df = enrich_panel(df)

        top10_all, top10_selected = make_top10_tables(df)
        churn = make_churn_table(df, top10_all)
        summary = make_yearly_summary(df)
        rank_traj = make_rank_trajectory(df)

        charts = [
            save_rank_dynamics_plot(rank_traj, hs_code, config["product_label"]),
            save_churn_plot(churn, hs_code, config["product_label"]),
            save_positive_partners_plot(summary, hs_code, config["product_label"]),
            save_concentration_plot(summary, hs_code, config["product_label"]),
        ]

        for chart in charts:
            chart_inventory.append({
                "hs_code": hs_code,
                "product_label": config["product_label"],
                "chart_file": chart,
            })

        all_top10.append(top10_all)
        all_top10_selected.append(top10_selected)
        all_churn.append(churn)
        all_summary.append(summary)
        all_rank_traj.append(rank_traj)

    top10_all_all = pd.concat(all_top10, ignore_index=True)
    top10_selected_all = pd.concat(all_top10_selected, ignore_index=True)
    churn_all = pd.concat(all_churn, ignore_index=True)
    summary_all = pd.concat(all_summary, ignore_index=True)
    rank_traj_all = pd.concat(all_rank_traj, ignore_index=True)
    rank_traj_all["partner"] = rank_traj_all["partner"].map(translate_label)

    # Сравнительные графики по HS-кодам.
    comparison_charts = [
        save_cross_code_plot(
            summary_all,
            metric="total_export_value_musd",
            ylabel="Стоимость экспорта, млн долл. США",
            filename="comparison_total_export_value_all_hs.png",
            title="Совокупная стоимость экспорта Китая по HS-кодам",
        ),
        save_cross_code_plot(
            summary_all,
            metric="positive_partners",
            ylabel="Количество импортёров",
            filename="comparison_positive_partners_all_hs.png",
            title="Количество импортёров с положительным экспортом по HS-кодам",
        ),
        save_cross_code_plot(
            summary_all,
            metric="hhi",
            ylabel="Индекс Херфиндаля-Хиршмана (HHI)",
            filename="comparison_hhi_all_hs.png",
            title="Концентрация экспортных направлений по HS-кодам",
        ),
        save_cross_code_plot(
            summary_all,
            metric="top10_share_pct",
            ylabel="Доля топ-10 направлений, %",
            filename="comparison_top10_share_all_hs.png",
            title="Доля топ-10 экспортных направлений по HS-кодам",
        ),
    ]

    for chart in comparison_charts:
        chart_inventory.append({
            "hs_code": "ALL",
            "product_label": "Сравнение по HS-кодам",
            "chart_file": chart,
        })

    chart_inventory = pd.DataFrame(chart_inventory)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        top10_all_all.to_excel(writer, sheet_name="top10_all_years", index=False)
        top10_selected_all.to_excel(writer, sheet_name="top10_selected_years", index=False)
        churn_all.to_excel(writer, sheet_name="top10_churn_yearly", index=False)
        summary_all.to_excel(writer, sheet_name="yearly_summary", index=False)
        rank_traj_all.to_excel(writer, sheet_name="rank_trajectories", index=False)
        chart_inventory.to_excel(writer, sheet_name="chart_inventory", index=False)

    print(f"Сохранена книга: {OUTPUT_FILE}")
    print(f"Графики сохранены в папке: {OUTPUT_DIR}")
    print("Графики:")
    for chart in chart_inventory["chart_file"].tolist():
        print(f" - {chart}")


if __name__ == "__main__":
    main()
