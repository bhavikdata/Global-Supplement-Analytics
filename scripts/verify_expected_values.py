"""
Independently recomputes the workbook's key KPI and forecast figures directly
from data/clean/*.csv, using pandas rather than the Excel formulas. Use this
to verify the numbers on Executive_Dashboard / Sales_Analysis / Inventory_
Analysis / Forecast without needing to open Excel at all, or to double-check
them after opening the workbook and pressing Ctrl+Alt+F9 to recalculate.

Run:  python scripts/verify_expected_values.py
"""
import os
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(BASE, "data", "clean")

sales_rep = pd.read_csv(os.path.join(CLEAN, "Sales_Reporting.csv"))
inv_rep = pd.read_csv(os.path.join(CLEAN, "Inventory_Reporting.csv"))

print("=" * 70)
print("EXECUTIVE DASHBOARD — expected KPI values (filters = All / All)")
print("=" * 70)
total_rev = sales_rep["Revenue"].sum()
total_gp = sales_rep["GrossProfit"].sum()
print(f"Total Revenue:      ${total_rev:,.2f}")
print(f"Total Gross Profit: ${total_gp:,.2f}")
print(f"Gross Margin %:     {total_gp/total_rev*100:,.2f}%")
print(f"Units Sold:         {int(sales_rep['Units'].sum()):,}")

rev_2024 = sales_rep.loc[sales_rep.Year == 2024, "Revenue"].sum()
rev_2025 = sales_rep.loc[sales_rep.Year == 2025, "Revenue"].sum()
print(f"YoY Revenue Growth: {(rev_2025-rev_2024)/rev_2024*100:,.2f}%  (2024=${rev_2024:,.2f}, 2025=${rev_2025:,.2f})")
print(f"Overall Stockout Rate: {inv_rep['StockoutFlag'].mean()*100:,.2f}%")

print("\n" + "=" * 70)
print("SALES_ANALYSIS — Revenue by Category / Region / Channel")
print("=" * 70)
print(sales_rep.groupby("Category")["Revenue"].sum().round(2).to_string())
print()
print(sales_rep.groupby("Region")["Revenue"].sum().round(2).to_string())
print()
print(sales_rep.groupby("ChannelName")["Revenue"].sum().round(2).to_string())

print("\n" + "=" * 70)
print("INVENTORY_ANALYSIS — key figures")
print("=" * 70)
print(f"Average Days of Cover: {inv_rep['DaysOfCover'].mean(skipna=True):,.2f}")
latest = inv_rep["SnapshotMonth"].max()
alerts = inv_rep[(inv_rep.SnapshotMonth == latest) & (inv_rep.ReorderAlert == 1)]
print(f"Reorder alerts in latest month ({latest}): {len(alerts)}")
print(f"Slow-moving Product x Country combos (24 months): {int(inv_rep['SlowMoving'].sum())}")

print("\n" + "=" * 70)
print("FORECAST — seasonal-naive baseline, 3-month holdout evaluation")
print("=" * 70)
monthly = sales_rep.groupby("YearMonth")["Revenue"].sum().sort_index()
months = monthly.index.tolist()
actual = monthly.values
forecast = [None] * len(actual)
for i in range(len(actual)):
    if i >= 12:
        forecast[i] = actual[i - 12]
    elif i >= 1:
        forecast[i] = sum(actual[:i]) / i
holdout_idx = range(len(actual) - 3, len(actual))
abs_err = [abs(actual[i] - forecast[i]) for i in holdout_idx]
sq_err = [(actual[i] - forecast[i]) ** 2 for i in holdout_idx]
mae = sum(abs_err) / 3
rmse = (sum(sq_err) / 3) ** 0.5
wape = sum(abs_err) / sum(actual[i] for i in holdout_idx)
print(f"Holdout months: {[months[i] for i in holdout_idx]}")
print(f"MAE:  ${mae:,.2f}")
print(f"RMSE: ${rmse:,.2f}")
print(f"WAPE: {wape*100:,.2f}%")
print(f"Forecast Accuracy (1 - WAPE): {(1-wape)*100:,.2f}%")
print("\nDone. Compare these to the Excel workbook after pressing Ctrl+Alt+F9.")
