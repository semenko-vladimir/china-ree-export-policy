import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, str(Path(__file__).parent.parent / "notebooks"))

import pandas as pd
import matplotlib.pyplot as plt

from plot_style import add_period_markers, format_axis, save_figure, set_plot_style, translate_label

ROOT = Path(__file__).parent.parent
INPUT_FILE  = ROOT / "data"    / "clean_china_ree_exports_284690_2010_2024.xlsx"
OUTPUT_FILE = ROOT / "results" / "descriptive_export_structure_hs284690.xlsx"
FIGURES_DIR = ROOT / "figures"

SELECTED_YEARS = [2014, 2016, 2018, 2020, 2024]
SPECIAL_PARTNERS_TO_DROP = ["Other Asia, nes"]

set_plot_style()

df = pd.read_excel(INPUT_FILE, sheet_name="final_panel_balanced")

df = df[~df["partner"].isin(SPECIAL_PARTNERS_TO_DROP)].copy()
df["export_value_usd"] = pd.to_numeric(df["export_value_usd"], errors="coerce").fillna(0)
df["quantity_kg"] = pd.to_numeric(df["quantity_kg"], errors="coerce").fillna(0)

df["total_export_value_usd_clean"] = df.groupby("year")["export_value_usd"].transform("sum")
df["share_export_value_clean"] = df["export_value_usd"] / df["total_export_value_usd_clean"]

top10 = (
    df[df["year"].isin(SELECTED_YEARS)]
    .sort_values(["year", "export_value_usd"], ascending=[True, False])
    .groupby("year")
    .head(10)
    .copy()
)

top10["rank"] = top10.groupby("year")["export_value_usd"].rank(method="first", ascending=False).astype(int)
top10["export_value_musd"] = top10["export_value_usd"] / 1_000_000
top10["share_export_value_pct"] = top10["share_export_value_clean"] * 100
top10["quantity_mkg"] = top10["quantity_kg"] / 1_000_000

summary_rows = []
for year, g in df.groupby("year"):
    g = g.copy()
    g = g.sort_values("export_value_usd", ascending=False)
    shares = g.set_index("partner")["share_export_value_clean"]
    japan_share = shares.get("Japan", 0)
    us_share = shares.get("United States", 0)
    top5_share = g.head(5)["share_export_value_clean"].sum()
    hhi = (g["share_export_value_clean"] ** 2).sum()
    summary_rows.append({
        "year": year,
        "total_export_value_musd_clean": g["export_value_usd"].sum() / 1_000_000,
        "japan_share_pct": japan_share * 100,
        "us_share_pct": us_share * 100,
        "japan_us_share_pct": (japan_share + us_share) * 100,
        "top5_share_pct": top5_share * 100,
        "hhi_export_value_clean": hhi,
        "top1_partner": translate_label(g.iloc[0]["partner"]),
        "top1_share_pct": g.iloc[0]["share_export_value_clean"] * 100,
        "top5_partners": "; ".join(translate_label(partner) for partner in g.head(5)["partner"].tolist())
    })

shares_hhi = pd.DataFrame(summary_rows).sort_values("year")
selected_year_summary = shares_hhi[shares_hhi["year"].isin(SELECTED_YEARS)].copy()

findings = pd.DataFrame([
    ["Охват", "Описательный анализ структуры экспорта Китая по HS 284690 за 2010-2024 годы; 'Other Asia, nes' исключена из основных таблиц как специальная категория WITS."],
    ["Япония", f"Доля Японии: {selected_year_summary.iloc[0]['japan_share_pct']:.1f}% в 2014 году и {selected_year_summary.iloc[-1]['japan_share_pct']:.1f}% в 2024 году."],
    ["США", f"Доля США: {selected_year_summary.iloc[0]['us_share_pct']:.1f}% в 2014 году и {selected_year_summary.iloc[-1]['us_share_pct']:.1f}% в 2024 году."],
    ["Япония + США", f"Совокупная доля: {selected_year_summary.iloc[0]['japan_us_share_pct']:.1f}% в 2014 году и {selected_year_summary.iloc[-1]['japan_us_share_pct']:.1f}% в 2024 году."],
    ["Концентрация топ-5", f"Доля топ-5 направлений: {selected_year_summary.iloc[0]['top5_share_pct']:.1f}% в 2014 году и {selected_year_summary.iloc[-1]['top5_share_pct']:.1f}% в 2024 году."],
    ["HHI", f"HHI: {selected_year_summary.iloc[0]['hhi_export_value_clean']:.3f} в 2014 году и {selected_year_summary.iloc[-1]['hhi_export_value_clean']:.3f} в 2024 году."],
    ["Интерпретация", "Используйте этот блок как описательную статистику: регрессия ShareExportValue не дала статистически значимого результата, поэтому нельзя утверждать о значимом изменении долей только на основании регрессии."]
], columns=["вывод", "деталь"])

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    top10_report = top10.copy()
    top10_report["partner"] = top10_report["partner"].map(translate_label)

    findings.to_excel(writer, sheet_name="text_findings", index=False)
    top10_report[["year", "rank", "partner", "export_value_musd", "share_export_value_pct", "quantity_mkg"]].to_excel(writer, sheet_name="top10_selected_years", index=False)
    shares_hhi.to_excel(writer, sheet_name="shares_hhi", index=False)
    selected_year_summary.to_excel(writer, sheet_name="selected_year_summary", index=False)

fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(shares_hhi["year"], shares_hhi["japan_share_pct"], marker="o", label="Япония")
ax.plot(shares_hhi["year"], shares_hhi["us_share_pct"], marker="o", label="США")
ax.plot(shares_hhi["year"], shares_hhi["japan_us_share_pct"], marker="o", label="Япония и США")
ax.plot(shares_hhi["year"], shares_hhi["top5_share_pct"], marker="o", label="Топ-5 направлений")
add_period_markers(ax)
format_axis(
    ax,
    title="Доли в экспорте Китая, HS 284690",
    ylabel="Доля стоимости экспорта, %",
    legend=True,
)
save_figure(FIGURES_DIR / "export_structure_shares_hs284690.png", fig)
plt.close(fig)

fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(shares_hhi["year"], shares_hhi["hhi_export_value_clean"], marker="o")
add_period_markers(ax)
format_axis(
    ax,
    title="Концентрация экспорта по стоимости, HS 284690",
    ylabel="Индекс Херфиндаля-Хиршмана (HHI)",
)
save_figure(FIGURES_DIR / "export_structure_hhi_hs284690.png", fig)
plt.close(fig)

print(f"Сохранено: {OUTPUT_FILE}")
print(f"Сохранено: {FIGURES_DIR / 'export_structure_shares_hs284690.png'}")
print(f"Сохранено: {FIGURES_DIR / 'export_structure_hhi_hs284690.png'}")
