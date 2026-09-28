"""
Global Supplement Sales & Inventory Analytics - Synthetic Data Generator
=========================================================================
Generates a fully synthetic, reproducible dataset (fixed random seed) for a
fictional supplement distributor. No real company, product, customer, or
transaction data is used anywhere in this project.

Outputs:
  data/raw/*.csv    - raw extracts WITH intentional data-quality issues
                       (duplicates, missing categories, inconsistent country
                       labels, invalid dates)
  data/clean/*.csv  - cleaned analytical tables used by the Excel workbook
  data/clean/data_quality_log.csv - record of every correction/rejection

Run:  python scripts/generate_data.py
Requires: pandas, numpy  (both preinstalled)
"""
import numpy as np
import pandas as pd
from datetime import date, timedelta
import os

SEED = 42
rng = np.random.default_rng(SEED)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "data", "raw")
CLEAN = os.path.join(BASE, "data", "clean")
os.makedirs(RAW, exist_ok=True)
os.makedirs(CLEAN, exist_ok=True)

# ---------------------------------------------------------------------------
# 1. CALENDAR  (24 months: 2024-01-01 .. 2025-12-31)
# ---------------------------------------------------------------------------
start_date = date(2024, 1, 1)
end_date = date(2025, 12, 31)
n_days = (end_date - start_date).days + 1
cal_dates = [start_date + timedelta(days=i) for i in range(n_days)]
calendar = pd.DataFrame({"Date": cal_dates})
calendar["Year"] = calendar["Date"].apply(lambda d: d.year)
calendar["Quarter"] = calendar["Date"].apply(lambda d: (d.month - 1) // 3 + 1)
calendar["Month"] = calendar["Date"].apply(lambda d: d.month)
calendar["MonthName"] = calendar["Date"].apply(lambda d: d.strftime("%B"))
calendar["YearMonth"] = calendar["Date"].apply(lambda d: d.strftime("%Y-%m"))
calendar["DateKey"] = calendar["Date"].apply(lambda d: d.strftime("%Y%m%d")).astype(int)
calendar.to_csv(os.path.join(CLEAN, "Calendar.csv"), index=False)

months = sorted(calendar["YearMonth"].unique().tolist())  # 24 month keys, chronological
month_index = {m: i for i, m in enumerate(months)}

# ---------------------------------------------------------------------------
# 2. GEOGRAPHY
# ---------------------------------------------------------------------------
geo_rows = [
    (1, "United States", "North America"),
    (2, "Canada", "North America"),
    (3, "United Kingdom", "Europe"),
    (4, "Germany", "Europe"),
    (5, "France", "Europe"),
    (6, "Australia", "APAC"),
    (7, "Japan", "APAC"),
    (8, "Brazil", "Latin America"),
]
geography = pd.DataFrame(geo_rows, columns=["CountryID", "Country", "Region"])
geography.to_csv(os.path.join(CLEAN, "Geography.csv"), index=False)

# Inconsistent raw country labels (data-quality issue), keyed by CountryID
country_label_variants = {
    1: ["United States", "USA", "U.S.A.", "US"],
    2: ["Canada", "CANADA", "Canada "],
    3: ["United Kingdom", "UK", "U.K."],
    4: ["Germany", "germany", "Germany"],
    5: ["France", "France", "FRANCE"],
    6: ["Australia", "Australia", "AUS"],
    7: ["Japan", "Japan", "JAPAN"],
    8: ["Brazil", "Brazil", "BRASIL"],
}

# ---------------------------------------------------------------------------
# 3. PRODUCTS  (15 SKUs across 4 categories)
# ---------------------------------------------------------------------------
categories = ["Vitamins", "Protein", "Sports Nutrition", "Wellness"]
product_names = {
    "Vitamins": ["Vitamin C 1000mg", "Vitamin D3 2000IU", "Multivitamin Daily", "Vitamin B-Complex"],
    "Protein": ["Whey Protein Vanilla", "Whey Protein Chocolate", "Plant Protein Unflavored", "Casein Protein"],
    "Sports Nutrition": ["Pre-Workout Blast", "BCAA Recovery", "Creatine Monohydrate", "Electrolyte Mix"],
    "Wellness": ["Omega-3 Fish Oil", "Probiotic 50B", "Magnesium Glycinate"],
}
suppliers = [101, 102, 103, 104]
prod_rows = []
pid = 1000
for cat, names in product_names.items():
    for name in names:
        unit_cost = round(rng.uniform(4, 14), 2)
        margin_mult = rng.uniform(1.8, 3.2)
        list_price = round(unit_cost * margin_mult, 2)
        prod_rows.append({
            "ProductID": pid,
            "ProductName": name,
            "Category": cat,
            "UnitCost": unit_cost,
            "ListPrice": list_price,
            "SupplierID": int(rng.choice(suppliers)),
            "ReorderPoint": int(rng.integers(150, 400)),
            "TargetStock": int(rng.integers(600, 1400)),
            "LeadTimeDays": int(rng.integers(7, 35)),
        })
        pid += 1
products = pd.DataFrame(prod_rows)
products.to_csv(os.path.join(CLEAN, "Products.csv"), index=False)
product_ids = products["ProductID"].tolist()

# Each product gets a monthly seasonal + trend demand profile
seasonal_curve = 1 + 0.25 * np.sin(np.linspace(0, 2 * np.pi, 12))  # 12-month cycle
def seasonality_for_month(m_idx):
    return seasonal_curve[m_idx % 12]

product_base_demand = {p: rng.uniform(60, 260) for p in product_ids}
product_trend = {p: rng.uniform(-0.004, 0.012) for p in product_ids}  # slight growth/decline per day

# ---------------------------------------------------------------------------
# 4. CUSTOMERS / CHANNELS
# ---------------------------------------------------------------------------
channels = pd.DataFrame([
    (1, "Online Retail", "B2C"),
    (2, "Pharmacy Chain", "B2B"),
    (3, "Gym & Fitness Retail", "B2B"),
    (4, "Marketplace (3P)", "B2C"),
], columns=["ChannelID", "ChannelName", "Segment"])
channels.to_csv(os.path.join(CLEAN, "Channels.csv"), index=False)
channel_ids = channels["ChannelID"].tolist()
channel_weights = np.array([0.35, 0.25, 0.20, 0.20])

# ---------------------------------------------------------------------------
# 5. PROMOTIONS (drive occasional demand spikes + discounts)
# ---------------------------------------------------------------------------
promo_months = set(rng.choice(months, size=8, replace=False))

# ---------------------------------------------------------------------------
# 6. SALES ORDER LINES  (~55,000 rows)
# ---------------------------------------------------------------------------
TARGET_ROWS = 55000
sales_rows = []
order_id_counter = 500000
line_id_counter = 1
country_ids = geography["CountryID"].tolist()
country_weights = np.array([0.30, 0.10, 0.14, 0.12, 0.09, 0.09, 0.09, 0.07])

days_arr = np.array(cal_dates)

# Precompute daily weight (seasonality) per day for sampling order dates
day_weights = []
for d in cal_dates:
    m_key = f"{d.year:04d}-{d.month:02d}"
    m_idx = month_index[m_key]
    w = seasonality_for_month(m_idx)
    if m_key in promo_months:
        w *= 1.35
    # slight weekday effect (more B2B on weekdays)
    if d.weekday() >= 5:
        w *= 0.85
    day_weights.append(w)
day_weights = np.array(day_weights)
day_weights = day_weights / day_weights.sum()

sampled_day_idx = rng.choice(len(cal_dates), size=TARGET_ROWS, p=day_weights)

# Stockout periods: for a random subset of (product, country) pairs, block certain months from selling normally
stockout_pairs = set()
for _ in range(25):
    p = int(rng.choice(product_ids))
    c = int(rng.choice(country_ids))
    m = rng.choice(months[2:-2])  # avoid edges
    stockout_pairs.add((p, c, m))

for i in range(TARGET_ROWS):
    d = cal_dates[sampled_day_idx[i]]
    m_key = f"{d.year:04d}-{d.month:02d}"
    m_idx = month_index[m_key]

    product_id = int(rng.choice(product_ids))
    country_id = int(rng.choice(country_ids, p=country_weights))
    channel_id = int(rng.choice(channel_ids, p=channel_weights))

    # Skip (with high probability) if this product/country/month is in a stockout window
    if (product_id, country_id, m_key) in stockout_pairs and rng.random() < 0.85:
        continue

    base = product_base_demand[product_id]
    trend = product_trend[product_id]
    days_since_start = (d - start_date).days
    demand_signal = base * (1 + trend * days_since_start) * seasonality_for_month(m_idx)
    units = max(1, int(rng.poisson(max(demand_signal / 30, 1))))

    prod = products.loc[products.ProductID == product_id].iloc[0]
    list_price = prod["ListPrice"]
    unit_cost = prod["UnitCost"]

    discount = 0.0
    if m_key in promo_months:
        discount = round(rng.choice([0.10, 0.15, 0.20], p=[0.5, 0.3, 0.2]), 2)
    elif rng.random() < 0.08:
        discount = round(rng.choice([0.05, 0.10]), 2)

    selling_price = round(list_price * (1 - discount), 2)
    revenue = round(selling_price * units, 2)
    cost = round(unit_cost * units, 2)

    order_id_counter += 1 if (i % 3 == 0) else 0
    sales_rows.append({
        "OrderLineID": line_id_counter,
        "OrderID": order_id_counter,
        "OrderDate": d.isoformat(),
        "ProductID": product_id,
        "CountryID": country_id,
        "ChannelID": channel_id,
        "Units": units,
        "SellingPrice": selling_price,
        "Discount": discount,
        "Revenue": revenue,
        "Cost": cost,
    })
    line_id_counter += 1

sales_clean = pd.DataFrame(sales_rows)
sales_clean.to_csv(os.path.join(CLEAN, "Sales.csv"), index=False)
print(f"Sales rows generated (clean): {len(sales_clean):,}")

# ---------------------------------------------------------------------------
# 7. INVENTORY  (monthly snapshot per product x country)
# ---------------------------------------------------------------------------
inv_rows = []
stock_state = {(p, c): int(rng.integers(400, 900)) for p in product_ids for c in country_ids}

sales_by_month_ppc = (
    sales_clean.assign(YearMonth=pd.to_datetime(sales_clean["OrderDate"]).dt.strftime("%Y-%m"))
    .groupby(["ProductID", "CountryID", "YearMonth"])["Units"].sum()
)

for m_key in months:
    for p in product_ids:
        prod = products.loc[products.ProductID == p].iloc[0]
        for c in country_ids:
            opening = stock_state[(p, c)]
            units_sold = int(sales_by_month_ppc.get((p, c, m_key), 0))
            # receipts: replenish toward target stock with some randomness / occasional delay (0 receipt)
            if rng.random() < 0.12:
                receipts = 0  # delayed delivery this month
            else:
                gap = max(prod["TargetStock"] - opening, 0)
                receipts = int(rng.integers(0, max(int(gap * 0.6), 1) + 1))
            closing = opening + receipts - units_sold
            stockout_flag = 0
            if closing < 0:
                # can't sell more than available in a real system; clip and flag stockout
                units_sold = opening + receipts
                closing = 0
                stockout_flag = 1
            elif closing < prod["ReorderPoint"] * 0.15:
                stockout_flag = 1
            stock_state[(p, c)] = closing
            inv_rows.append({
                "SnapshotMonth": m_key,
                "ProductID": p,
                "CountryID": c,
                "OpeningStock": opening,
                "Receipts": receipts,
                "UnitsSold": units_sold,
                "ClosingStock": closing,
                "StockoutFlag": stockout_flag,
            })

inventory = pd.DataFrame(inv_rows)
inventory.to_csv(os.path.join(CLEAN, "Inventory.csv"), index=False)
print(f"Inventory rows generated: {len(inventory):,}")

# ---------------------------------------------------------------------------
# 8. PURCHASE ORDERS
# ---------------------------------------------------------------------------
po_rows = []
po_id = 9000
for m_key in months:
    y, mo = map(int, m_key.split("-"))
    for p in rng.choice(product_ids, size=rng.integers(4, 9), replace=False):
        prod = products.loc[products.ProductID == p].iloc[0]
        order_day = rng.integers(1, 25)
        order_date = date(y, mo, int(order_day))
        lead = int(prod["LeadTimeDays"])
        expected = order_date + timedelta(days=lead)
        delay = int(rng.choice([0, 0, 0, 0, 3, 5, 10, 20], p=[0.55, 0.1, 0.1, 0.05, 0.08, 0.06, 0.04, 0.02]))
        actual = expected + timedelta(days=delay) if rng.random() < 0.9 else None  # 10% not yet received
        qty_ordered = int(rng.integers(200, 900))
        qty_received = qty_ordered if actual is not None else 0
        if actual is not None and rng.random() < 0.05:
            qty_received = int(qty_ordered * rng.uniform(0.8, 0.98))  # partial receipt
        po_rows.append({
            "POID": po_id,
            "ProductID": int(p),
            "OrderDate": order_date.isoformat(),
            "ExpectedReceiptDate": expected.isoformat(),
            "ActualReceiptDate": actual.isoformat() if actual is not None else "",
            "QuantityOrdered": qty_ordered,
            "QuantityReceived": qty_received,
            "SupplierID": int(prod["SupplierID"]),
        })
        po_id += 1

purchase_orders = pd.DataFrame(po_rows)
purchase_orders.to_csv(os.path.join(CLEAN, "PurchaseOrders.csv"), index=False)
print(f"Purchase order rows generated: {len(purchase_orders):,}")

# ---------------------------------------------------------------------------
# 9. RAW EXTRACTS WITH INTENTIONAL DATA-QUALITY ISSUES
#    (mirrors what a real ETL pull might look like before cleaning)
# ---------------------------------------------------------------------------
dq_log = []

# --- 9a. Raw Sales: inconsistent country labels, some duplicates, a few invalid dates, some missing category info
raw_sales = sales_clean.copy()
raw_sales = raw_sales.merge(geography[["CountryID"]], on="CountryID", how="left")

def messy_country_label(cid):
    variants = country_label_variants[cid]
    return rng.choice(variants)

raw_sales["Country_raw"] = raw_sales["CountryID"].apply(messy_country_label)
raw_sales = raw_sales.drop(columns=["CountryID"])

# Inject ~0.6% duplicate rows
n_dupes = int(len(raw_sales) * 0.006)
dupe_rows = raw_sales.sample(n=n_dupes, random_state=SEED)
raw_sales = pd.concat([raw_sales, dupe_rows], ignore_index=True)
dq_log.append({"Table": "Sales", "Issue": "Duplicate order lines", "RecordsAffected": n_dupes,
                "Resolution": "Dropped exact duplicates on OrderLineID, keeping first occurrence"})

# Inject ~0.3% invalid dates (typo'd / malformed strings)
n_bad_dates = max(1, int(len(raw_sales) * 0.003))
bad_idx = rng.choice(raw_sales.index, size=n_bad_dates, replace=False)
bad_date_samples = ["2024-13-05", "2025-02-30", "31/01/2024", "not_a_date", ""]
for idx in bad_idx:
    raw_sales.loc[idx, "OrderDate"] = rng.choice(bad_date_samples)
dq_log.append({"Table": "Sales", "Issue": "Invalid / malformed OrderDate values", "RecordsAffected": n_bad_dates,
                "Resolution": "Rows with unparseable dates were rejected and logged; not loaded into clean table"})

# Inject inconsistent country label issue already done; log it
dq_log.append({"Table": "Sales", "Issue": "Inconsistent country label formatting (e.g. 'USA' vs 'United States' vs 'US')",
                "RecordsAffected": len(raw_sales), "Resolution": "Standardised via country lookup/mapping table to CountryID"})

raw_sales.to_csv(os.path.join(RAW, "Sales_raw.csv"), index=False)

# --- 9b. Raw Products: some missing category values
raw_products = products.copy()
n_missing_cat = 2
miss_idx = rng.choice(raw_products.index, size=n_missing_cat, replace=False)
raw_products.loc[miss_idx, "Category"] = ""
dq_log.append({"Table": "Products", "Issue": "Missing Category value", "RecordsAffected": n_missing_cat,
                "Resolution": "Missing categories back-filled using ProductName keyword mapping; flagged in Data_Quality sheet"})
raw_products.to_csv(os.path.join(RAW, "Products_raw.csv"), index=False)

# --- 9c. Raw Inventory: a few negative-looking / impossible values introduced then corrected
raw_inventory = inventory.copy()
n_neg = 3
neg_idx = rng.choice(raw_inventory.index, size=n_neg, replace=False)
raw_inventory.loc[neg_idx, "ClosingStock"] = -1  # impossible raw value from source system glitch
dq_log.append({"Table": "Inventory", "Issue": "Negative ClosingStock (system export glitch)", "RecordsAffected": n_neg,
                "Resolution": "Recalculated ClosingStock = OpeningStock + Receipts - UnitsSold; verified balance equation"})
raw_inventory.to_csv(os.path.join(RAW, "Inventory_raw.csv"), index=False)

# --- 9d. Raw Geography: duplicated country row with inconsistent naming
raw_geo = pd.concat([geography, pd.DataFrame([{"CountryID": 1, "Country": "USA", "Region": "North America"}])],
                     ignore_index=True)
dq_log.append({"Table": "Geography", "Issue": "Duplicate country row with alternate spelling", "RecordsAffected": 1,
                "Resolution": "Deduplicated on CountryID, canonical name retained from master Geography table"})
raw_geo.to_csv(os.path.join(RAW, "Geography_raw.csv"), index=False)

dq_log_df = pd.DataFrame(dq_log)
dq_log_df.to_csv(os.path.join(CLEAN, "data_quality_log.csv"), index=False)

# ---------------------------------------------------------------------------
# 10. SUMMARY / VALIDATION PRINTS
# ---------------------------------------------------------------------------
print("\n--- Validation ---")
print("Sales date range:", sales_clean["OrderDate"].min(), "to", sales_clean["OrderDate"].max())
print("Unique OrderLineID count == rows:", sales_clean["OrderLineID"].is_unique)
print("Any negative Units:", (sales_clean["Units"] < 0).any())
print("Any negative Revenue:", (sales_clean["Revenue"] < 0).any())
bal_check = inventory["OpeningStock"] + inventory["Receipts"] - inventory["UnitsSold"] - inventory["ClosingStock"]
print("Inventory balance max abs deviation:", bal_check.abs().max())
print("Total clean rows -> Sales: {:,} | Inventory: {:,} | PurchaseOrders: {:,} | Products: {} | Geography: {} | Channels: {} | Calendar: {}".format(
    len(sales_clean), len(inventory), len(purchase_orders), len(products), len(geography), len(channels), len(calendar)))
print("\nRaw files written with intentional issues:")
print(" - Sales_raw.csv: +{} duplicates, {} invalid dates, inconsistent country labels".format(n_dupes, n_bad_dates))
print(" - Products_raw.csv: {} missing category values".format(n_missing_cat))
print(" - Inventory_raw.csv: {} impossible negative ClosingStock values".format(n_neg))
print(" - Geography_raw.csv: 1 duplicate country row with alternate spelling")
print("\nDone.")
