"""
Builds the denormalized 'reporting layer' tables (the equivalent of a Power
Query merge step) from the clean, normalized fact/dimension tables.
These are saved as CSVs and then loaded into the Excel workbook as static
Excel Tables (their values represent an ETL join, not row-level formulas --
documented in KPI_Definitions / README for performance reasons at 50k+ rows).
"""
import pandas as pd
import numpy as np
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(BASE, "data", "clean")

sales = pd.read_csv(os.path.join(CLEAN, "Sales.csv"))
products = pd.read_csv(os.path.join(CLEAN, "Products.csv"))
geography = pd.read_csv(os.path.join(CLEAN, "Geography.csv"))
channels = pd.read_csv(os.path.join(CLEAN, "Channels.csv"))
inventory = pd.read_csv(os.path.join(CLEAN, "Inventory.csv"))
po = pd.read_csv(os.path.join(CLEAN, "PurchaseOrders.csv"))

sales["OrderDate"] = pd.to_datetime(sales["OrderDate"])
sales["Year"] = sales["OrderDate"].dt.year
sales["YearMonth"] = sales["OrderDate"].dt.strftime("%Y-%m")
sales["GrossProfit"] = sales["Revenue"] - sales["Cost"]

sales_rep = sales.merge(products[["ProductID", "ProductName", "Category"]], on="ProductID", how="left")
sales_rep = sales_rep.merge(geography, on="CountryID", how="left")
sales_rep = sales_rep.merge(channels[["ChannelID", "ChannelName"]], on="ChannelID", how="left")
sales_rep = sales_rep[[
    "OrderLineID", "OrderID", "OrderDate", "Year", "YearMonth", "ProductID", "ProductName",
    "Category", "CountryID", "Country", "Region", "ChannelID", "ChannelName",
    "Units", "SellingPrice", "Discount", "Revenue", "Cost", "GrossProfit"
]]
sales_rep["OrderDate"] = sales_rep["OrderDate"].dt.strftime("%Y-%m-%d")
sales_rep.to_csv(os.path.join(CLEAN, "Sales_Reporting.csv"), index=False)

# --- Inventory reporting layer -----------------------------------------
inv_rep = inventory.merge(
    products[["ProductID", "ProductName", "Category", "ReorderPoint", "TargetStock"]],
    on="ProductID", how="left"
)
inv_rep = inv_rep.merge(geography, on="CountryID", how="left")
inv_rep["AvgDailySales"] = inv_rep["UnitsSold"] / 30.0
inv_rep["DaysOfCover"] = np.where(
    inv_rep["AvgDailySales"] > 0, inv_rep["ClosingStock"] / inv_rep["AvgDailySales"], np.nan
)
inv_rep["ReorderAlert"] = (inv_rep["ClosingStock"] < inv_rep["ReorderPoint"]).astype(int)
inv_rep["SlowMoving"] = ((inv_rep["UnitsSold"] < inv_rep["TargetStock"] * 0.03)).astype(int)
inv_rep = inv_rep[[
    "SnapshotMonth", "ProductID", "ProductName", "Category", "CountryID", "Country", "Region",
    "OpeningStock", "Receipts", "UnitsSold", "ClosingStock", "StockoutFlag",
    "ReorderPoint", "TargetStock", "DaysOfCover", "ReorderAlert", "SlowMoving"
]]
inv_rep.to_csv(os.path.join(CLEAN, "Inventory_Reporting.csv"), index=False)

# --- Purchase order reporting: on-time flag -----------------------------
po2 = po.copy()
po2["ExpectedReceiptDate"] = pd.to_datetime(po2["ExpectedReceiptDate"])
po2["ActualReceiptDate"] = pd.to_datetime(po2["ActualReceiptDate"], errors="coerce")
po2["Received"] = po2["ActualReceiptDate"].notna().astype(int)
po2["DelayDays"] = (po2["ActualReceiptDate"] - po2["ExpectedReceiptDate"]).dt.days
po2["OnTime"] = np.where(po2["Received"] == 1, (po2["DelayDays"] <= 0).astype(int), np.nan)
po2.to_csv(os.path.join(CLEAN, "PurchaseOrders_Reporting.csv"), index=False)

# --- Monthly company-level summary for forecast --------------------------
monthly = sales_rep.groupby("YearMonth", as_index=False).agg(Revenue=("Revenue", "sum"), Units=("Units", "sum"))
monthly = monthly.sort_values("YearMonth").reset_index(drop=True)
monthly.to_csv(os.path.join(CLEAN, "Monthly_Summary.csv"), index=False)

print("Sales_Reporting:", sales_rep.shape)
print("Inventory_Reporting:", inv_rep.shape)
print("PurchaseOrders_Reporting:", po2.shape)
print("Monthly_Summary:", monthly.shape)
print(monthly)
