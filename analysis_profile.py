from pathlib import Path

import numpy as np
import pandas as pd


DATA = Path("systeme_electric")


def clean_code(series):
    return series.astype("string").str.strip()


def report_table(name, df, key=None):
    print(f"\n## {name}")
    print("shape", df.shape)
    print("full_duplicates", int(df.duplicated().sum()))
    print("missing_cells", int(df.isna().sum().sum()), "of", int(df.size), f"({df.isna().mean().mean():.2%})")
    if key:
        print("key_missing", int(df[key].isna().sum()))
        print("key_unique", int(df[key].nunique(dropna=True)))
        print("key_duplicates_rows", int(df[key].duplicated(keep=False).sum()))


moq = pd.read_excel(DATA / "MOQ SystemElectric.xlsx", sheet_name="Лист_1")
moq = moq.dropna(how="all").copy()
moq["Номенклатура.Код"] = clean_code(moq["Номенклатура.Код"])
report_table("MOQ", moq, "Номенклатура.Код")
print("missing_by_col", moq.isna().sum().to_dict())
print("multiplicity_stats", moq["Кратность"].describe().to_dict())
print("multiplicity_nonpositive", int((moq["Кратность"] <= 0).sum()))
print("multiplicity_top", moq["Кратность"].value_counts().head(12).to_dict())


qty = pd.read_excel(DATA / "Ежемесячные продажи в кол-м выражении SystemElectric 2024-2026.xlsx", sheet_name="Лист_1")
qty = qty[qty["Номенклатура.Код"].notna()].copy()
qty["Номенклатура.Код"] = clean_code(qty["Номенклатура.Код"])
month_cols = [c for c in qty.columns if any(str(c).startswith(x) for x in ["янв.", "февр.", "март", "апр.", "май", "июнь", "июль", "авг.", "сент.", "окт.", "нояб.", "дек."])]
qty[month_cols] = qty[month_cols].apply(pd.to_numeric, errors="coerce")
qty["Итого"] = pd.to_numeric(qty["Итого"], errors="coerce")
report_table("Monthly unit sales", qty, "Номенклатура.Код")
print("month_columns", len(month_cols), month_cols[0], month_cols[-1])
print("missing_core", qty[["Номенклатура", "Номенклатура.Код", "Артикул", "Кратность"]].isna().sum().to_dict())
print("monthly_missing", int(qty[month_cols].isna().sum().sum()), "monthly_zero", int((qty[month_cols].fillna(0) == 0).sum().sum()), "monthly_negative", int((qty[month_cols] < 0).sum().sum()))
print("skus_any_sales", int((qty[month_cols].fillna(0).sum(axis=1) > 0).sum()))
print("skus_no_sales", int((qty[month_cols].fillna(0).sum(axis=1) == 0).sum()))
print("skus_any_negative_month", int((qty[month_cols] < 0).any(axis=1).sum()))
calc_total = qty[month_cols].fillna(0).sum(axis=1)
print("total_mismatch_rows", int((calc_total - qty["Итого"].fillna(0)).abs().gt(1e-9).sum()))
print("annual_unit_totals", {y: float(qty[[c for c in month_cols if str(y) in c]].fillna(0).sum().sum()) for y in [2024, 2025, 2026]})
print("monthly_unit_totals")
print(qty[month_cols].sum().to_string())
print("top10_skus_units")
print(qty.assign(total_calc=calc_total).nlargest(10, "total_calc")[["Номенклатура.Код", "Артикул", "Номенклатура", "total_calc"]].to_string(index=False))
print("qty_multiplicity_unique", qty["Кратность"].value_counts(dropna=False).head(10).to_dict())


inv = pd.read_excel(DATA / "Ежемесячные остатки SystemElectric 2024-2026.xlsx", sheet_name="Лист_1")
inv = inv[inv["Номенклатура.Код"].notna()].copy()
inv["Номенклатура.Код"] = clean_code(inv["Номенклатура.Код"])
inv_months = [c for c in inv.columns if c not in ["№", "Номенклатура", "Номенклатура.Код", "Ед.изм"]]
report_table("Monthly inventory", inv, "Номенклатура.Код")
print("month_columns", len(inv_months), inv_months[0], inv_months[-1])
print("monthly_missing", int(inv[inv_months].isna().sum().sum()), "monthly_zero", int((inv[inv_months].fillna(0) == 0).sum().sum()), "monthly_negative", int((inv[inv_months] < 0).sum().sum()))
latest = inv[inv_months[-1]]
print("latest_inventory", {"missing": int(latest.isna().sum()), "zero": int(latest.eq(0).sum()), "negative": int(latest.lt(0).sum()), "positive": int(latest.gt(0).sum()), "sum": float(latest.sum())})
print("negative_inventory_skus_any", int((inv[inv_months] < 0).any(axis=1).sum()))


tx = pd.read_excel(DATA / "Динамика продаж_Syseme Electric_2025-2026.xlsx", sheet_name="Лист_1")
tx["Код"] = clean_code(tx["Код"])
tx["Дата_dt"] = pd.to_datetime(tx["Дата"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
report_table("Transactions", tx, "Код")
print("missing_by_col", tx.isna().sum().to_dict())
print("date_range", tx["Дата_dt"].min(), tx["Дата_dt"].max(), "bad_dates", int(tx["Дата_dt"].isna().sum()))
print("unique_documents", int(tx["Номер"].nunique()), "warehouses", tx["Склад"].value_counts(dropna=False).to_dict(), "units", tx["Ед."].value_counts(dropna=False).to_dict())
print("quantity_stats", tx["Количество"].describe(percentiles=[.01,.05,.25,.5,.75,.95,.99]).to_dict())
print("sign_rows", {"positive": int(tx["Количество"].gt(0).sum()), "zero": int(tx["Количество"].eq(0).sum()), "negative": int(tx["Количество"].lt(0).sum())})
print("sign_units", {"positive": float(tx.loc[tx["Количество"] > 0, "Количество"].sum()), "negative": float(tx.loc[tx["Количество"] < 0, "Количество"].sum()), "net": float(tx["Количество"].sum())})
print("transactions_by_year")
print(tx.groupby(tx["Дата_dt"].dt.year)["Количество"].agg(["count", "sum"]).to_string())
print("transactions_monthly_last18")
print(tx.groupby(tx["Дата_dt"].dt.to_period("M"))["Количество"].agg(["count", "sum"]).tail(18).to_string())


transit = pd.read_excel(DATA / "Товар в пути_SystemElectric на 22.09.2026.xlsx", sheet_name="TDSheet", header=1)
transit = transit.loc[:, transit.notna().any(axis=0)].copy()
transit["Код 1с"] = clean_code(transit["Код 1с"])
report_table("Current stock / in-transit", transit, "Код 1с")
print("columns", list(transit.columns))
print("missing_by_col", transit.isna().sum().to_dict())
for c in ["Сумма последние 12 мес", "   Ср мес за последние 12 мес", "Остаток", "Зарезервировано", "Свободный остаток", "Запас", "СЭ в пути 24.09", "Витрина", "Остаток ТЗ"]:
    s = pd.to_numeric(transit[c], errors="coerce")
    print(c, {"sum": float(s.sum()), "zero": int(s.eq(0).sum()), "negative": int(s.lt(0).sum()), "positive": int(s.gt(0).sum()), "median": float(s.median()), "p90": float(s.quantile(.9)), "max": float(s.max())})
sales12 = pd.to_numeric(transit["Сумма последние 12 мес"], errors="coerce")
free = pd.to_numeric(transit["Свободный остаток"], errors="coerce")
inroad = pd.to_numeric(transit["СЭ в пути 24.09"], errors="coerce")
avg12 = pd.to_numeric(transit["   Ср мес за последние 12 мес"], errors="coerce")
cover = (free + inroad) / avg12.replace(0, np.nan)
print("derived_health", {
    "active_skus": int(sales12.gt(0).sum()),
    "no_sales_12m": int(sales12.eq(0).sum()),
    "active_free_le_zero": int((sales12.gt(0) & free.le(0)).sum()),
    "active_free_plus_transit_le_zero": int((sales12.gt(0) & (free + inroad).le(0)).sum()),
    "active_cover_lt_1m": int((sales12.gt(0) & cover.lt(1)).sum()),
    "active_cover_gt_6m": int((sales12.gt(0) & cover.gt(6)).sum()),
    "active_cover_gt_12m": int((sales12.gt(0) & cover.gt(12)).sum()),
    "no_sales_positive_free": int((sales12.eq(0) & free.gt(0)).sum()),
})
print("cover_active_stats", cover[sales12.gt(0)].describe(percentiles=[.1,.25,.5,.75,.9,.95]).to_dict())
print("category_counts", transit["Категория 2026"].value_counts(dropna=False).sort_index().to_dict())


season_raw = pd.read_excel(DATA / "Сезонность SystemElectric 2024-2026.xlsx", sheet_name="Лист1", header=None)
money = season_raw.iloc[3:6, :14].copy()
money.columns = ["год", "янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", "итого"]
money = money.set_index("год").apply(pd.to_numeric, errors="coerce")
print("\n## Aggregate monetary sales / seasonality")
print(money.to_string())
print("2025_vs_2024", float(money.loc[2025, "итого"] / money.loc[2024, "итого"] - 1))
jan_sep = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен"]
print("2026_jan_sep_vs_2025", float(money.loc[2026, jan_sep].sum() / money.loc[2025, jan_sep].sum() - 1))
shares = money.loc[[2024, 2025], ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]].div(money.loc[[2024, 2025], "итого"], axis=0).mean()
print("average_month_share_2024_2025")
print(shares.sort_values(ascending=False).to_string())


sets = {
    "moq": set(moq["Номенклатура.Код"].dropna()),
    "qty": set(qty["Номенклатура.Код"].dropna()),
    "inv": set(inv["Номенклатура.Код"].dropna()),
    "tx": set(tx["Код"].dropna()),
    "transit": set(transit["Код 1с"].dropna()),
}
print("\n## Code coverage")
for a in sets:
    for b in sets:
        if a < b:
            print(a, b, "intersection", len(sets[a] & sets[b]), "only_a", len(sets[a] - sets[b]), "only_b", len(sets[b] - sets[a]))


# Exact reconciliation between transaction-level and monthly SKU totals over common dates.
tx2 = tx[tx["Дата_dt"].dt.year.between(2024, 2026)].copy()
month_map = {1:"янв.",2:"февр.",3:"март",4:"апр.",5:"май",6:"июнь",7:"июль",8:"авг.",9:"сент.",10:"окт.",11:"нояб.",12:"дек."}
tx2["month_col"] = tx2["Дата_dt"].dt.month.map(month_map) + " " + tx2["Дата_dt"].dt.year.astype(str)
long_qty = qty.melt(id_vars=["Номенклатура.Код"], value_vars=month_cols, var_name="month_col", value_name="summary_qty")
long_qty["summary_qty"] = long_qty["summary_qty"].fillna(0)
txm = tx2.groupby(["Код", "month_col"], as_index=False)["Количество"].sum().rename(columns={"Код":"Номенклатура.Код", "Количество":"tx_qty"})
cmp = long_qty.merge(txm, how="outer", on=["Номенклатура.Код", "month_col"]).fillna({"summary_qty":0,"tx_qty":0})
cmp["diff"] = cmp["summary_qty"] - cmp["tx_qty"]
print("\n## Transaction vs monthly-unit reconciliation")
print("pairs", len(cmp), "exact", int(cmp["diff"].abs().lt(1e-9).sum()), "mismatch", int(cmp["diff"].abs().ge(1e-9).sum()), "abs_diff", float(cmp["diff"].abs().sum()), "net_diff", float(cmp["diff"].sum()))
print("global_by_month")
print(cmp.groupby("month_col")[["summary_qty", "tx_qty", "diff"]].sum().to_string())
