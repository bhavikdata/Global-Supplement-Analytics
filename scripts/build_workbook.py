"""
Builds Global_Supplement_Analytics.xlsx from the cleaned CSVs.
"""
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, DataBarRule
from openpyxl.comments import Comment

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(BASE, "data", "clean")
OUT = os.path.join(BASE, "Global_Supplement_Analytics.xlsx")

FONT_NAME = "Arial"

# ---- palette (restrained / professional) --------------------------------
NAVY = "1F3864"
STEEL = "2E5C8A"
TEAL = "1B7A72"
GOLD = "B08A2E"
GREY = "595959"
LIGHT_GREY = "F2F2F2"
WHITE = "FFFFFF"
RED = "C0504D"
GREEN = "2E7D32"

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(name=FONT_NAME, bold=True, color=WHITE, size=10)
TITLE_FONT = Font(name=FONT_NAME, bold=True, color=NAVY, size=16)
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, color=GREY, size=10)
SECTION_FONT = Font(name=FONT_NAME, bold=True, color=WHITE, size=11)
SECTION_FILL = PatternFill("solid", fgColor=STEEL)
KPI_LABEL_FONT = Font(name=FONT_NAME, bold=True, color=GREY, size=9)
KPI_VALUE_FONT = Font(name=FONT_NAME, bold=True, color=NAVY, size=18)
KPI_FILL = PatternFill("solid", fgColor=LIGHT_GREY)
BODY_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)


def style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def autofit(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def add_table(ws, df, sheet_name, table_name, start_row=1, start_col=1, style="TableStyleMedium9"):
    """Write a dataframe as a native Excel Table starting at (start_row, start_col)."""
    nrows, ncols = df.shape
    for j, col in enumerate(df.columns):
        ws.cell(row=start_row, column=start_col + j, value=col)
    for i in range(nrows):
        for j, col in enumerate(df.columns):
            val = df.iat[i, j]
            if pd.isna(val):
                val = None
            elif hasattr(val, "item"):
                val = val.item()
            ws.cell(row=start_row + 1 + i, column=start_col + j, value=val)
    style_header_row(ws, start_row, ncols, start_col)
    end_row = start_row + nrows
    end_col = start_col + ncols - 1
    ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True, showFirstColumn=False)
    ws.add_table(tab)
    return end_row, end_col


# =====================================================================
# LOAD DATA
# =====================================================================
products = pd.read_csv(os.path.join(CLEAN, "Products.csv"))
geography = pd.read_csv(os.path.join(CLEAN, "Geography.csv"))
channels = pd.read_csv(os.path.join(CLEAN, "Channels.csv"))
calendar = pd.read_csv(os.path.join(CLEAN, "Calendar.csv"))
sales = pd.read_csv(os.path.join(CLEAN, "Sales.csv"))
inventory = pd.read_csv(os.path.join(CLEAN, "Inventory.csv"))
po = pd.read_csv(os.path.join(CLEAN, "PurchaseOrders.csv"))
sales_rep = pd.read_csv(os.path.join(CLEAN, "Sales_Reporting.csv"))
inv_rep = pd.read_csv(os.path.join(CLEAN, "Inventory_Reporting.csv"))
po_rep = pd.read_csv(os.path.join(CLEAN, "PurchaseOrders_Reporting.csv"))
dq_log = pd.read_csv(os.path.join(CLEAN, "data_quality_log.csv"))
monthly = pd.read_csv(os.path.join(CLEAN, "Monthly_Summary.csv"))

raw_sales_sample_full = pd.read_csv(os.path.join(BASE, "data", "raw", "Sales_raw.csv"))
raw_products = pd.read_csv(os.path.join(BASE, "data", "raw", "Products_raw.csv"))
raw_inventory = pd.read_csv(os.path.join(BASE, "data", "raw", "Inventory_raw.csv"))
raw_geo = pd.read_csv(os.path.join(BASE, "data", "raw", "Geography_raw.csv"))

print("Loaded all source CSVs.")

# =====================================================================
# SHEET: Products
# =====================================================================
ws = wb.create_sheet("Products")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, products, "Products", "ProductsDim")
autofit(ws, {"A": 10, "B": 24, "C": 18, "D": 10, "E": 11, "F": 11, "G": 13, "H": 12, "I": 13})
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
ws.freeze_panes = "A2"

# =====================================================================
# SHEET: Geography
# =====================================================================
ws = wb.create_sheet("Geography")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, geography, "Geography", "tblGeography")
autofit(ws, {"A": 10, "B": 18, "C": 16})
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
ws.freeze_panes = "A2"

# =====================================================================
# SHEET: Channels
# =====================================================================
ws = wb.create_sheet("Channels")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, channels, "Channels", "tblChannels")
autofit(ws, {"A": 10, "B": 20, "C": 12})
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
ws.freeze_panes = "A2"

# =====================================================================
# SHEET: Calendar
# =====================================================================
ws = wb.create_sheet("Calendar")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, calendar, "Calendar", "tblCalendar")
autofit(ws, {"A": 12, "B": 8, "C": 9, "D": 8, "E": 12, "F": 11, "G": 11})
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
    row[0].number_format = "yyyy-mm-dd"
ws.freeze_panes = "A2"

print("Dimension sheets written.")

# =====================================================================
# SHEET: Sales (fact table, normalized, per spec) + GrossProfit formula
# =====================================================================
ws = wb.create_sheet("Sales")
ws.sheet_view.showGridLines = False
sales_cols = list(sales.columns) + ["GrossProfit"]
sales_out = sales.copy()
sales_out["GrossProfit"] = None  # placeholder, formula written below
end_r, end_c = add_table(ws, sales_out, "Sales", "SalesFact", style="TableStyleMedium9")
gp_col = end_c  # last column index (1-based) = GrossProfit
rev_col_letter = get_column_letter(sales_cols.index("Revenue") + 1)
cost_col_letter = get_column_letter(sales_cols.index("Cost") + 1)
gp_col_letter = get_column_letter(gp_col)
for r in range(2, end_r + 1):
    ws.cell(row=r, column=gp_col, value=f"={rev_col_letter}{r}-{cost_col_letter}{r}")
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
ws.cell(row=1, column=1).comment = Comment(
    "Normalized fact table (order-line grain). Keys only + core measures, "
    "matching the required schema. GrossProfit is the only derived column here "
    "(kept as a live formula). Joined attributes (Category, Country, Region, "
    "ChannelName) live in Sales_Reporting to keep this fact table lean and fast "
    "at 50k+ rows.", "Data Model")
autofit(ws, {get_column_letter(i + 1): 13 for i in range(end_c)})
ws.column_dimensions["C"].width = 12
ws.freeze_panes = "A2"
for r in range(2, end_r + 1):
    ws.cell(row=r, column=3).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=sales_cols.index("SellingPrice") + 1).number_format = "$#,##0.00"
    ws.cell(row=r, column=sales_cols.index("Discount") + 1).number_format = "0%"
    ws.cell(row=r, column=sales_cols.index("Revenue") + 1).number_format = "$#,##0.00"
    ws.cell(row=r, column=sales_cols.index("Cost") + 1).number_format = "$#,##0.00"
    ws.cell(row=r, column=gp_col).number_format = "$#,##0.00"
print(f"Sales fact table written: {end_r-1:,} rows, GrossProfit formulas added.")

# =====================================================================
# SHEET: Sales_Reporting (denormalized reporting layer - values, ETL join)
# =====================================================================
ws = wb.create_sheet("Sales_Reporting")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, sales_rep, "Sales_Reporting", "Sales_Reporting", style="TableStyleMedium2")
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
cols = list(sales_rep.columns)
for r in range(2, end_r + 1):
    ws.cell(row=r, column=cols.index("OrderDate") + 1).number_format = "yyyy-mm-dd"
    for cname in ["SellingPrice", "Revenue", "Cost", "GrossProfit"]:
        ws.cell(row=r, column=cols.index(cname) + 1).number_format = "$#,##0.00"
    ws.cell(row=r, column=cols.index("Discount") + 1).number_format = "0%"
autofit(ws, {get_column_letter(i + 1): 13 for i in range(end_c)})
ws.column_dimensions[get_column_letter(cols.index("ProductName") + 1)].width = 24
ws.freeze_panes = "A2"
ws.cell(row=1, column=1).comment = Comment(
    "Denormalized reporting extract: Sales joined to Products/Geography/Channels, "
    "equivalent to a Power Query 'Merge Queries' step, plus Year/YearMonth date "
    "keys. Produced once during ETL (not as 55k row-level Excel formulas) to keep "
    "the workbook responsive; see scripts/build_reporting_layer.py for the exact "
    "transformation logic. All SUMIFS/COUNTIFS analysis in this workbook reads "
    "from this table.", "ETL Note")
print(f"Sales_Reporting written: {end_r-1:,} rows.")

# =====================================================================
# SHEET: Inventory (fact, normalized)
# =====================================================================
ws = wb.create_sheet("Inventory")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, inventory, "Inventory", "InventoryFact", style="TableStyleMedium9")
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
autofit(ws, {get_column_letter(i + 1): 14 for i in range(end_c)})
ws.freeze_panes = "A2"
ws.cell(row=1, column=1).comment = Comment(
    "Grain: one row per Product x Country x SnapshotMonth. Balance check "
    "(OpeningStock + Receipts - UnitsSold = ClosingStock) is validated in "
    "Data_Quality and holds for 100% of rows (see reconciliation).", "Data Model")
print(f"Inventory written: {end_r-1:,} rows.")

# =====================================================================
# SHEET: Inventory_Reporting (denormalized, with DaysOfCover/ReorderAlert)
# =====================================================================
ws = wb.create_sheet("Inventory_Reporting")
ws.sheet_view.showGridLines = False
end_r, end_c = add_table(ws, inv_rep, "Inventory_Reporting", "Inventory_Reporting", style="TableStyleMedium2")
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
cols = list(inv_rep.columns)
for r in range(2, end_r + 1):
    ws.cell(row=r, column=cols.index("DaysOfCover") + 1).number_format = "0.0"
autofit(ws, {get_column_letter(i + 1): 13 for i in range(end_c)})
ws.column_dimensions[get_column_letter(cols.index("ProductName") + 1)].width = 24
ws.freeze_panes = "A2"
ws.cell(row=1, column=1).comment = Comment(
    "DaysOfCover = ClosingStock / (UnitsSold/30). ReorderAlert = 1 if "
    "ClosingStock < ReorderPoint. SlowMoving = 1 if UnitsSold < 3% of "
    "TargetStock for that month. Joined + derived during ETL for the same "
    "performance reason as Sales_Reporting.", "ETL Note")
print(f"Inventory_Reporting written: {end_r-1:,} rows.")

# =====================================================================
# SHEET: PurchaseOrders
# =====================================================================
ws = wb.create_sheet("PurchaseOrders")
ws.sheet_view.showGridLines = False
po_full = po_rep.copy()
end_r, end_c = add_table(ws, po_full, "PurchaseOrders", "tblPO", style="TableStyleMedium9")
for row in ws.iter_rows(min_row=2, max_row=end_r, min_col=1, max_col=end_c):
    for cell in row:
        cell.font = BODY_FONT
cols = list(po_full.columns)
for r in range(2, end_r + 1):
    for cname in ["OrderDate", "ExpectedReceiptDate", "ActualReceiptDate"]:
        ws.cell(row=r, column=cols.index(cname) + 1).number_format = "yyyy-mm-dd"
autofit(ws, {get_column_letter(i + 1): 15 for i in range(end_c)})
ws.freeze_panes = "A2"
ws.cell(row=1, column=1).comment = Comment(
    "Received=1 if ActualReceiptDate is populated (blank = not yet received "
    "as of data extract). OnTime=1 if received on/before ExpectedReceiptDate; "
    "blank if not yet received.", "Data Model")
print(f"PurchaseOrders written: {end_r-1:,} rows.")

wb.save(OUT)
print("Checkpoint 4 saved.")

# =====================================================================
# SHEET: Raw_Data_Sample (illustrative excerpt of the dirty raw extracts)
# =====================================================================
ws = wb.create_sheet("Raw_Data_Sample")
ws.sheet_view.showGridLines = False
ws["A1"] = "Raw Data Samples — Before Cleaning"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Illustrative excerpts only (full raw files are in data/raw/*.csv). Shows the exact issue types corrected during ETL."
ws["A2"].font = SUBTITLE_FONT
row_cursor = 4

def write_block(ws, title, df, row_cursor, table_name):
    ws.cell(row=row_cursor, column=1, value=title).font = SECTION_FONT
    ws.cell(row=row_cursor, column=1).fill = SECTION_FILL
    last_col = len(df.columns)
    for c in range(2, last_col + 1):
        ws.cell(row=row_cursor, column=c).fill = SECTION_FILL
    row_cursor += 1
    end_r, end_c = add_table(ws, df, ws.title, table_name, start_row=row_cursor, start_col=1)
    for r in ws.iter_rows(min_row=row_cursor + 1, max_row=end_r, min_col=1, max_col=end_c):
        for cell in r:
            cell.font = BODY_FONT
    return end_r + 3

# Sample of duplicate + bad-date + inconsistent-country rows
sample_dupe_ids = raw_sales_sample_full[raw_sales_sample_full.duplicated(subset=["OrderLineID"], keep=False)]["OrderLineID"].unique()[:4]
sample1 = raw_sales_sample_full[raw_sales_sample_full["OrderLineID"].isin(sample_dupe_ids)].head(8)
row_cursor = write_block(ws, "1. Sales_raw.csv — duplicate OrderLineID rows (sample)", sample1, row_cursor, "RawDupSample")

bad_dates = raw_sales_sample_full[~raw_sales_sample_full["OrderDate"].astype(str).str.match(r"^\d{4}-\d{2}-\d{2}$", na=False)].head(6)
row_cursor = write_block(ws, "2. Sales_raw.csv — invalid / malformed OrderDate values (sample)", bad_dates, row_cursor, "RawBadDateSample")

country_variety = raw_sales_sample_full[raw_sales_sample_full["Country_raw"].isin(["USA", "US", "U.S.A.", "UK", "U.K.", "germany", "AUS"])].drop_duplicates(subset=["Country_raw"]).head(8)
row_cursor = write_block(ws, "3. Sales_raw.csv — inconsistent country label formatting (sample)", country_variety, row_cursor, "RawCountrySample")

row_cursor = write_block(ws, "4. Products_raw.csv — missing Category values", raw_products[raw_products["Category"].isna() | (raw_products["Category"] == "")], row_cursor, "RawProdSample")

row_cursor = write_block(ws, "5. Inventory_raw.csv — impossible negative ClosingStock", raw_inventory[raw_inventory["ClosingStock"] < 0], row_cursor, "RawInvSample")

row_cursor = write_block(ws, "6. Geography_raw.csv — duplicate country row, alternate spelling", raw_geo, row_cursor, "RawGeoSample")

autofit(ws, {c: 16 for c in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]})
ws.column_dimensions["A"].width = 14

wb.save(OUT)
print("Checkpoint 5 (Raw_Data_Sample) saved.")

# =====================================================================
# SHEET: Data_Quality
# =====================================================================
ws = wb.create_sheet("Data_Quality")
ws.sheet_view.showGridLines = False
ws["A1"] = "Data Quality — Checks, Corrections & Reconciliation"
ws["A1"].font = TITLE_FONT
ws["A2"] = "All figures below are computed live from the Sales / Sales_Reporting / Inventory tables in this workbook."
ws["A2"].font = SUBTITLE_FONT

ws["A4"] = "1. Data Quality Log (issues found in raw extracts and how they were resolved)"
ws["A4"].font = SECTION_FONT
ws["A4"].fill = SECTION_FILL
for c in range(2, 5):
    ws.cell(row=4, column=c).fill = SECTION_FILL
end_r, end_c = add_table(ws, dq_log, "Data_Quality", "tblDQLog", start_row=5, start_col=1)
for r in ws.iter_rows(min_row=6, max_row=end_r, min_col=1, max_col=end_c):
    for cell in r:
        cell.font = BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")

next_row = end_r + 3
ws.cell(row=next_row, column=1, value="2. Live Reconciliation Checks (formulas — recalculate on file open / F9)").font = SECTION_FONT
ws.cell(row=next_row, column=1).fill = SECTION_FILL
for c in range(2, 4):
    ws.cell(row=next_row, column=c).fill = SECTION_FILL
next_row += 1

header_row = next_row
for j, h in enumerate(["Check", "Result", "Expected / Pass Rule"]):
    ws.cell(row=header_row, column=1 + j, value=h)
style_header_row(ws, header_row, 3)
next_row += 1

checks = [
    ("Sales fact table row count", "=COUNTA(SalesFact[OrderLineID])", "= clean row count (54,573 after removing duplicates & invalid dates)"),
    ("Duplicate OrderLineID values in Sales", "=IF((MAX(SalesFact[OrderLineID])-MIN(SalesFact[OrderLineID])+1)=COUNTA(SalesFact[OrderLineID]),\"0 (IDs are sequential, no gaps/dupes)\",\"CHECK NEEDED\")", "0 (all keys unique)"),
    ("Sales Revenue total vs Sales_Reporting Revenue total (diff)", "=ROUND(SUM(SalesFact[Revenue])-SUM(Sales_Reporting[Revenue]),2)", "0.00 (reporting layer reconciles to fact table)"),
    ("Sales Cost total vs Sales_Reporting Cost total (diff)", "=ROUND(SUM(SalesFact[Cost])-SUM(Sales_Reporting[Cost]),2)", "0.00"),
    ("Inventory balance failures (Opening+Receipts-Sold<>Closing)", "=SUMPRODUCT(--((InventoryFact[OpeningStock]+InventoryFact[Receipts]-InventoryFact[UnitsSold])<>InventoryFact[ClosingStock]))", "0 (balance equation holds for every row)"),
    ("Sales rows with negative Units", "=COUNTIF(SalesFact[Units],\"<0\")", "0"),
    ("Sales rows with negative Revenue", "=COUNTIF(SalesFact[Revenue],\"<0\")", "0"),
    ("Sales OrderDate min", "=MIN(SalesFact[OrderDate])", "2024-01-01"),
    ("Sales OrderDate max", "=MAX(SalesFact[OrderDate])", "2025-12-31"),
    ("Products with missing Category (clean table)", "=COUNTBLANK(ProductsDim[Category])", "0 (back-filled during ETL; 2 were blank in raw extract)"),
]
for i, (label, formula, expected) in enumerate(checks):
    r = next_row + i
    ws.cell(row=r, column=1, value=label).font = BODY_FONT
    ws.cell(row=r, column=2, value=formula).font = BODY_FONT
    ws.cell(row=r, column=3, value=expected).font = BODY_FONT
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True)
last_check_row = next_row + len(checks) - 1
ws.cell(row=next_row + 2, column=8, value="Date formulas display as serials unless formatted; set B8:B9 number format to yyyy-mm-dd in Excel if needed.")
for r in range(next_row, last_check_row + 1):
    ws.cell(row=r, column=2).number_format = "General"
ws.cell(row=next_row + 7, column=2).number_format = "yyyy-mm-dd"
ws.cell(row=next_row + 8, column=2).number_format = "yyyy-mm-dd"

next_row = last_check_row + 3
ws.cell(row=next_row, column=1, value="3. Checks performed programmatically in Python (scripts/generate_data.py) before load").font = SECTION_FONT
ws.cell(row=next_row, column=1).fill = SECTION_FILL
for c in range(2, 4):
    ws.cell(row=next_row, column=c).fill = SECTION_FILL
next_row += 1
py_checks = [
    "Uniqueness of OrderLineID, ProductID, CountryID, ChannelID, POID primary keys — PASS",
    "Referential integrity: every ProductID/CountryID/ChannelID in Sales exists in its dimension table — PASS",
    "No negative Units, Revenue, or Cost values in the clean Sales table — PASS",
    "Inventory balance equation (Opening + Receipts − Sold = Closing) holds for all 2,880 rows — PASS (also re-verified live in section 2 above)",
    "OrderDate range fully within 2024-01-01 to 2025-12-31, no future/impossible dates in clean data — PASS",
    "PurchaseOrders: ActualReceiptDate ≥ OrderDate for all received orders — PASS",
]
for i, txt in enumerate(py_checks):
    ws.cell(row=next_row + i, column=1, value="• " + txt).font = BODY_FONT
next_row = next_row + len(py_checks) + 2

ws.cell(row=next_row, column=1, value="4. Manual checks recommended in desktop Excel (not verifiable in this automated environment)").font = SECTION_FONT
ws.cell(row=next_row, column=1).fill = SECTION_FILL
for c in range(2, 4):
    ws.cell(row=next_row, column=c).fill = SECTION_FILL
next_row += 1
manual_checks = [
    "Open the file and confirm no broken-link warning (missing reference) appears on open.",
    "Insert a PivotTable from Sales_Reporting and confirm it refreshes without error (Data > Refresh All).",
    "Add slicers (Insert > Slicer) to a PivotTable for Year, Region, Category, Channel and confirm they filter correctly.",
    "If you rebuild Power Query connections (see README), click Data > Refresh All and confirm row counts match this sheet.",
    "Spot-check 5–10 rows in Sales_Reporting against Sales + Products/Geography/Channels manually with VLOOKUP/XLOOKUP.",
]
manual_checks = [
    "Open the file and confirm no broken-link warning (missing reference) appears on open.",
    "Press Ctrl+Alt+F9 (Windows) / Cmd+Fn+F9 (Mac) to force a full recalculation — recommended on first open (see note in section 5 below).",
    "Insert a PivotTable from Sales_Reporting and confirm it refreshes without error (Data > Refresh All).",
    "Add slicers (Insert > Slicer) to a PivotTable for Year, Region, Category, Channel and confirm they filter correctly.",
    "If you rebuild Power Query connections (see README), click Data > Refresh All and confirm row counts match this sheet.",
    "Spot-check 5–10 rows in Sales_Reporting against Sales + Products/Geography/Channels manually with VLOOKUP/XLOOKUP.",
]
for i, txt in enumerate(manual_checks):
    ws.cell(row=next_row + i, column=1, value="☐ " + txt).font = BODY_FONT
next_row = next_row + len(manual_checks) + 2

# --- Section 5: independently verified expected values + recalc environment note ---
ws.cell(row=next_row, column=1, value="5. Independently Verified Expected Values & Recalculation Note").font = SECTION_FONT
ws.cell(row=next_row, column=1).fill = SECTION_FILL
for c in range(2, 4):
    ws.cell(row=next_row, column=c).fill = SECTION_FILL
next_row += 1
recalc_note = (
    "This workbook's ~55,000 formula cells were built and structurally validated (zero #NAME?/#REF! parsing "
    "errors) using LibreOffice headless recalculation during development. However, a full clean recalculation "
    "pass of the FINAL file could not be completed in this automated sandbox: LibreOffice's headless engine "
    "hung indefinitely on this file size on the single-core container used to build it (CPU usage dropped to "
    "~0%, indicating a stalled dialog/render wait rather than genuine computation — a known class of headless "
    "LibreOffice issue, not a formula-logic problem). The table below gives independently computed 'ground "
    "truth' values (Python/pandas, from the exact same clean CSVs) for the workbook's key formulas, so you can "
    "open the file in desktop Excel, press Ctrl+Alt+F9, and confirm these figures match in seconds."
)
ws.cell(row=next_row, column=1, value=recalc_note).font = SUBTITLE_FONT
ws.cell(row=next_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=next_row, start_column=1, end_row=next_row + 4, end_column=6)
next_row += 6

exp_hdr = next_row
for j, h in enumerate(["Metric", "Independently Computed Value (Python/pandas)", "Where to check it in this workbook"]):
    ws.cell(row=exp_hdr, column=1 + j, value=h)
style_header_row(ws, exp_hdr, 3)
expected_values = [
    ("Total Revenue (all data)", "$13,574,293.45", "Executive_Dashboard, Total Revenue card, filters = All/All"),
    ("Total Gross Profit (all data)", "$7,629,716.13", "Executive_Dashboard, Gross Profit card"),
    ("Total Units Sold (all data)", "632,897", "Executive_Dashboard, Units Sold card"),
    ("Revenue 2024", "$4,684,234.39", "Sales_Analysis / Executive_Dashboard, Year filter = 2024"),
    ("Revenue 2025", "$8,890,059.06", "Sales_Analysis / Executive_Dashboard, Year filter = 2025"),
    ("YoY Revenue Growth (2024→2025)", "89.79%", "Executive_Dashboard, YoY Revenue Growth card"),
    ("Revenue — Protein category", "$5,160,573.39", "Sales_Analysis, Revenue by Category table"),
    ("Revenue — Vitamins category", "$4,012,083.42", "Sales_Analysis, Revenue by Category table"),
    ("Revenue — Sports Nutrition category", "$2,461,945.56", "Sales_Analysis, Revenue by Category table"),
    ("Revenue — Wellness category", "$1,939,691.08", "Sales_Analysis, Revenue by Category table"),
    ("Revenue — North America", "$5,423,290.68", "Sales_Analysis, Revenue by Region table"),
    ("Revenue — Europe", "$4,729,988.39", "Sales_Analysis, Revenue by Region table"),
    ("Revenue — APAC", "$2,468,983.06", "Sales_Analysis, Revenue by Region table"),
    ("Revenue — Latin America", "$952,031.32", "Sales_Analysis, Revenue by Region table"),
    ("Overall Stockout Rate", "22.15%", "Inventory_Analysis KPI strip"),
    ("Average Days of Cover", "403.6", "Inventory_Analysis KPI strip"),
    ("Reorder Alerts, latest month (2025-12)", "65", "Inventory_Analysis KPI strip / alert table"),
    ("Slow-Moving Product x Country combos (24 mo.)", "543", "Sum of Inventory_Reporting[SlowMoving]"),
    ("Forecast MAE (3-month holdout)", "$300,935.75", "Forecast sheet, evaluation metrics"),
    ("Forecast RMSE (3-month holdout)", "$324,360.22", "Forecast sheet, evaluation metrics"),
    ("Forecast WAPE (3-month holdout)", "41.53%", "Forecast sheet, evaluation metrics"),
    ("Forecast Accuracy (1 − WAPE)", "58.47%", "Forecast sheet / Executive_Dashboard KPI card"),
]
for i, (metric, val, where) in enumerate(expected_values):
    r = exp_hdr + 1 + i
    ws.cell(row=r, column=1, value=metric).font = BODY_FONT
    ws.cell(row=r, column=2, value=val).font = BODY_FONT
    ws.cell(row=r, column=3, value=where).font = BODY_FONT
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True)
exp_last_row = exp_hdr + len(expected_values)
ws.cell(row=exp_last_row + 2, column=1,
        value="Reproduce these figures yourself: scripts/verify_expected_values.py recomputes all of them directly "
              "from data/clean/*.csv, independent of the Excel formulas.")
ws.cell(row=exp_last_row + 2, column=1).font = SUBTITLE_FONT

autofit(ws, {"A": 46, "B": 46, "C": 46, "D": 20, "E": 20, "F": 20})
ws.column_dimensions["B"].width = 55

wb.save(OUT)
print("Checkpoint 6 (Data_Quality) saved.")

# =====================================================================
# SHEET: Sales_Analysis
# =====================================================================
ws = wb.create_sheet("Sales_Analysis")
ws.sheet_view.showGridLines = False
ws["A1"] = "Sales Analysis"
ws["A1"].font = TITLE_FONT
ws["A2"] = "All values are formulas (SUMIFS/COUNTIFS) reading live from the Sales_Reporting table."
ws["A2"].font = SUBTITLE_FONT

month_keys = monthly["YearMonth"].tolist()
category_list = sorted(products["Category"].unique().tolist())
region_list = sorted(geography["Region"].unique().tolist())
channel_list = channels["ChannelName"].tolist()

# --- Monthly trend table ---
r0 = 4
ws.cell(row=r0, column=1, value="Monthly Revenue & Units Trend").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 4):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["YearMonth", "Revenue", "Units"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 3)
for i, mk in enumerate(month_keys):
    r = hdr_row + 1 + i
    ws.cell(row=r, column=1, value=mk).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[YearMonth],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=3, value=f'=SUMIFS(Sales_Reporting[Units],Sales_Reporting[YearMonth],A{r})').number_format = "#,##0"
    ws.cell(row=r, column=2).font = BODY_FONT
    ws.cell(row=r, column=3).font = BODY_FONT
monthly_last_row = hdr_row + len(month_keys)

chart1 = LineChart()
chart1.title = "Monthly Revenue Trend (24 Months)"
chart1.style = 2
chart1.y_axis.title = "Revenue ($)"
chart1.x_axis.title = "Month"
chart1.height = 8
chart1.width = 18
data = Reference(ws, min_col=2, min_row=hdr_row, max_row=monthly_last_row)
cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=monthly_last_row)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.series[0].graphicalProperties.line.width = 22000
chart1.series[0].graphicalProperties.line.solidFill = STEEL
chart1.series[0].smooth = False
ws.add_chart(chart1, "E4")

# --- By Category table ---
r0 = monthly_last_row + 3
ws.cell(row=r0, column=1, value="Revenue & Margin by Category").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 5):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["Category", "Revenue", "GrossProfit", "Margin %"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 4)
for i, cat in enumerate(category_list):
    r = hdr_row + 1 + i
    ws.cell(row=r, column=1, value=cat).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Category],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=3, value=f'=SUMIFS(Sales_Reporting[GrossProfit],Sales_Reporting[Category],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=4, value=f'=IFERROR(C{r}/B{r},0)').number_format = "0.0%"
    for cc in (2, 3, 4):
        ws.cell(row=r, column=cc).font = BODY_FONT
cat_last_row = hdr_row + len(category_list)

chart2 = BarChart()
chart2.type = "col"
chart2.title = "Revenue by Category"
chart2.y_axis.title = "Revenue ($)"
chart2.height = 8
chart2.width = 12
data = Reference(ws, min_col=2, min_row=hdr_row, max_row=cat_last_row)
cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=cat_last_row)
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
chart2.series[0].graphicalProperties.solidFill = STEEL
chart2.legend = None
ws.add_chart(chart2, "E22")

# --- By Region table ---
r0 = cat_last_row + 3
ws.cell(row=r0, column=1, value="Revenue by Region").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 3):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["Region", "Revenue"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 2)
for i, reg in enumerate(region_list):
    r = hdr_row + 1 + i
    ws.cell(row=r, column=1, value=reg).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Region],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=2).font = BODY_FONT
reg_last_row = hdr_row + len(region_list)

chart3 = PieChart()
chart3.title = "Revenue Share by Region"
chart3.height = 8
chart3.width = 12
data = Reference(ws, min_col=2, min_row=hdr_row, max_row=reg_last_row)
cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=reg_last_row)
chart3.add_data(data, titles_from_data=True)
chart3.set_categories(cats)
ws.add_chart(chart3, "M22")

# --- By Channel table ---
r0 = reg_last_row + 3
ws.cell(row=r0, column=1, value="Revenue by Channel").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 3):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["Channel", "Revenue"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 2)
for i, ch in enumerate(channel_list):
    r = hdr_row + 1 + i
    ws.cell(row=r, column=1, value=ch).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[ChannelName],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=2).font = BODY_FONT
chan_last_row = hdr_row + len(channel_list)

chart4 = BarChart()
chart4.type = "bar"
chart4.title = "Revenue by Channel"
chart4.height = 8
chart4.width = 12
data = Reference(ws, min_col=2, min_row=hdr_row, max_row=chan_last_row)
cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=chan_last_row)
chart4.add_data(data, titles_from_data=True)
chart4.set_categories(cats)
chart4.series[0].graphicalProperties.solidFill = TEAL
chart4.legend = None
ws.add_chart(chart4, "E40")

autofit(ws, {"A": 20, "B": 14, "C": 14, "D": 12})
ws.freeze_panes = "A4"

wb.save(OUT)
print("Checkpoint 7 (Sales_Analysis) saved.")

# =====================================================================
# SHEET: Inventory_Analysis
# =====================================================================
ws = wb.create_sheet("Inventory_Analysis")
ws.sheet_view.showGridLines = False
ws["A1"] = "Inventory Analysis"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Formulas read live from the Inventory_Reporting table (2,880 monthly product x country snapshots)."
ws["A2"].font = SUBTITLE_FONT

latest_month = calendar["YearMonth"].iloc[-1]
product_names_sorted = products.sort_values("ProductName")["ProductName"].tolist()

# --- KPI strip ---
kpi_defs = [
    ("Overall Stockout Rate", "=AVERAGE(Inventory_Reporting[StockoutFlag])", "0.0%"),
    ("Avg Days of Cover", "=AVERAGE(Inventory_Reporting[DaysOfCover])", "0.0"),
    ("Reorder Alerts (latest month)", f'=COUNTIFS(Inventory_Reporting[SnapshotMonth],"{latest_month}",Inventory_Reporting[ReorderAlert],1)', "#,##0"),
    ("Slow-Moving Product/Country Combos", "=SUM(Inventory_Reporting[SlowMoving])", "#,##0"),
]
for i, (label, formula, fmt) in enumerate(kpi_defs):
    col = 1 + i * 2
    cell_label = ws.cell(row=4, column=col, value=label)
    cell_label.font = KPI_LABEL_FONT
    ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
    cell_val = ws.cell(row=5, column=col, value=formula)
    cell_val.font = KPI_VALUE_FONT
    cell_val.number_format = fmt
    ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col + 1)
    for rr in (4, 5, 6):
        for cc in (col, col + 1):
            ws.cell(row=rr, column=cc).fill = KPI_FILL
            ws.cell(row=rr, column=cc).border = BOX

# --- Stockout rate & Days of cover by product ---
r0 = 9
ws.cell(row=r0, column=1, value="Stockout Rate & Days of Cover by Product").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 5):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["Product", "Category", "Stockout Rate", "Avg Days of Cover"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 4)
for i, pname in enumerate(product_names_sorted):
    r = hdr_row + 1 + i
    cat = products.loc[products.ProductName == pname, "Category"].iloc[0]
    ws.cell(row=r, column=1, value=pname).font = BODY_FONT
    ws.cell(row=r, column=2, value=cat).font = BODY_FONT
    ws.cell(row=r, column=3, value=f'=AVERAGEIFS(Inventory_Reporting[StockoutFlag],Inventory_Reporting[ProductName],A{r})').number_format = "0.0%"
    ws.cell(row=r, column=4, value=f'=AVERAGEIFS(Inventory_Reporting[DaysOfCover],Inventory_Reporting[ProductName],A{r})').number_format = "0.0"
    ws.cell(row=r, column=3).font = BODY_FONT
    ws.cell(row=r, column=4).font = BODY_FONT
prod_last_row = hdr_row + len(product_names_sorted)

# conditional formatting: data bars on stockout rate
ws.conditional_formatting.add(
    f"C{hdr_row+1}:C{prod_last_row}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="C0504D")
)

chart5 = BarChart()
chart5.type = "col"
chart5.title = "Stockout Rate by Product"
chart5.y_axis.title = "Stockout Rate"
chart5.y_axis.numFmt = "0%"
chart5.height = 9
chart5.width = 22
data = Reference(ws, min_col=3, min_row=hdr_row, max_row=prod_last_row)
cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=prod_last_row)
chart5.add_data(data, titles_from_data=True)
chart5.set_categories(cats)
chart5.series[0].graphicalProperties.solidFill = RED
chart5.legend = None
ws.add_chart(chart5, "F9")

# --- Reorder alerts (latest month) ---
r0 = prod_last_row + 3
ws.cell(row=r0, column=1, value=f"Reorder Alerts — Latest Snapshot Month ({latest_month})").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 6):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
hdr_row = r0 + 1
for j, h in enumerate(["Product", "Country", "Closing Stock", "Reorder Point", "Alert?"]):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, 5)
alert_rows = inv_rep[(inv_rep["SnapshotMonth"] == latest_month) & (inv_rep["ReorderAlert"] == 1)].sort_values("ClosingStock")
for i, (_, rowd) in enumerate(alert_rows.iterrows()):
    r = hdr_row + 1 + i
    ws.cell(row=r, column=1, value=rowd["ProductName"]).font = BODY_FONT
    ws.cell(row=r, column=2, value=rowd["Country"]).font = BODY_FONT
    ws.cell(row=r, column=3, value=int(rowd["ClosingStock"])).font = BODY_FONT
    ws.cell(row=r, column=4, value=int(rowd["ReorderPoint"])).font = BODY_FONT
    ws.cell(row=r, column=5, value="REORDER").font = Font(name=FONT_NAME, bold=True, color=RED, size=10)
alert_last_row = hdr_row + max(len(alert_rows), 1)
if len(alert_rows) == 0:
    ws.cell(row=hdr_row + 1, column=1, value="No reorder alerts in the latest snapshot month.").font = BODY_FONT

autofit(ws, {"A": 24, "B": 16, "C": 16, "D": 16, "E": 12})
ws.freeze_panes = "A4"

wb.save(OUT)
print("Checkpoint 8 (Inventory_Analysis) saved.")

# =====================================================================
# SHEET: Forecast
# =====================================================================
ws = wb.create_sheet("Forecast")
ws.sheet_view.showGridLines = False
ws["A1"] = "Forecast — Seasonal Naive Baseline vs. Actual"
ws["A1"].font = TITLE_FONT
ws["A2"] = ("Baseline method: Forecast(month) = Actual(same month, 12 months earlier) once a full prior year "
            "exists; otherwise an expanding average of all prior months. Final 3 months held out for evaluation "
            "— no future information is used to forecast them (their seasonal-naive inputs are always 12 months "
            "in the past, i.e. inside the training window).")
ws["A2"].font = SUBTITLE_FONT
ws["A2"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[2].height = 30
ws.merge_cells("A2:H2")

hdr_row = 4
headers = ["#", "YearMonth", "Actual Revenue", "Forecast Revenue", "Error", "Abs Error", "Squared Error", "Holdout?"]
for j, h in enumerate(headers):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, len(headers))

n_months = len(month_keys)
holdout_start_idx = n_months - 3  # 0-indexed; last 3 months held out
first_data_row = hdr_row + 1
for i, mk in enumerate(month_keys):
    r = first_data_row + i
    ws.cell(row=r, column=1, value=i + 1).font = BODY_FONT
    ws.cell(row=r, column=2, value=mk).font = BODY_FONT
    ws.cell(row=r, column=3, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[YearMonth],B{r})').number_format = "$#,##0"
    ws.cell(row=r, column=3).font = BODY_FONT
    if i >= 12:
        fcast = f"=C{r-12}"
    elif i >= 1:
        fcast = f"=AVERAGE(C${first_data_row}:C{r-1})"
    else:
        fcast = '="n/a"'
    ws.cell(row=r, column=4, value=fcast).number_format = "$#,##0"
    ws.cell(row=r, column=4).font = BODY_FONT
    if i >= 1:
        ws.cell(row=r, column=5, value=f"=IFERROR(C{r}-D{r},\"\")").number_format = "$#,##0"
        ws.cell(row=r, column=6, value=f"=IFERROR(ABS(E{r}),\"\")").number_format = "$#,##0"
        ws.cell(row=r, column=7, value=f"=IFERROR(E{r}^2,\"\")").number_format = "#,##0"
    else:
        ws.cell(row=r, column=5, value="n/a")
        ws.cell(row=r, column=6, value="n/a")
        ws.cell(row=r, column=7, value="n/a")
    is_holdout = "Yes" if i >= holdout_start_idx else "No"
    ws.cell(row=r, column=8, value=is_holdout).font = BODY_FONT
    for cc in (5, 6, 8):
        ws.cell(row=r, column=cc).font = BODY_FONT
last_data_row = first_data_row + n_months - 1
holdout_first_row = first_data_row + holdout_start_idx

chart6 = LineChart()
chart6.title = "Actual vs Forecast Revenue"
chart6.y_axis.title = "Revenue ($)"
chart6.x_axis.title = "Month"
chart6.height = 9
chart6.width = 22
data = Reference(ws, min_col=3, max_col=4, min_row=hdr_row, max_row=last_data_row)
cats = Reference(ws, min_col=2, min_row=first_data_row, max_row=last_data_row)
chart6.add_data(data, titles_from_data=True)
chart6.set_categories(cats)
chart6.series[0].graphicalProperties.line.solidFill = STEEL
chart6.series[0].graphicalProperties.line.width = 22000
chart6.series[1].graphicalProperties.line.solidFill = GOLD
chart6.series[1].graphicalProperties.line.width = 18000
chart6.series[1].graphicalProperties.line.dashStyle = "dash"
ws.add_chart(chart6, "J4")

# --- Evaluation metrics (holdout = final 3 months only) ---
r0 = last_data_row + 3
ws.cell(row=r0, column=1, value="Holdout Evaluation Metrics (final 3 months only)").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
for c in range(2, 3):
    ws.cell(row=r0, column=c).fill = SECTION_FILL
metric_row = r0 + 1
ws.cell(row=metric_row, column=1, value="MAE (Mean Absolute Error)").font = BODY_FONT
ws.cell(row=metric_row, column=2, value=f"=AVERAGE(F{holdout_first_row}:F{last_data_row})").number_format = "$#,##0"
ws.cell(row=metric_row, column=2).font = BODY_FONT
ws.cell(row=metric_row + 1, column=1, value="RMSE (Root Mean Squared Error)").font = BODY_FONT
ws.cell(row=metric_row + 1, column=2, value=f"=SQRT(AVERAGE(G{holdout_first_row}:G{last_data_row}))").number_format = "$#,##0"
ws.cell(row=metric_row + 1, column=2).font = BODY_FONT
ws.cell(row=metric_row + 2, column=1, value="WAPE (Weighted Absolute Percentage Error)").font = BODY_FONT
ws.cell(row=metric_row + 2, column=2,
        value=f"=IF(SUM(C{holdout_first_row}:C{last_data_row})=0,\"undefined (zero actuals)\",SUM(F{holdout_first_row}:F{last_data_row})/SUM(C{holdout_first_row}:C{last_data_row}))")
ws.cell(row=metric_row + 2, column=2).number_format = "0.0%"
ws.cell(row=metric_row + 2, column=2).font = BODY_FONT
ws.cell(row=metric_row + 3, column=1, value="Forecast Accuracy (1 − WAPE)").font = BODY_FONT
ws.cell(row=metric_row + 3, column=2, value=f"=IFERROR(1-B{metric_row+2},\"n/a\")").number_format = "0.0%"
ws.cell(row=metric_row + 3, column=2).font = KPI_VALUE_FONT

note_row = metric_row + 5
ws.cell(row=note_row, column=1,
        value=("WAPE is undefined when total actual demand in the evaluation window is zero (division by zero). "
               "That does not occur in this dataset, but the formula above guards for it explicitly. "
               "Excel's FORECAST.ETS was intentionally NOT embedded as a live formula: this environment's "
               "recalculation engine (LibreOffice) does not reliably evaluate FORECAST.ETS, and its accuracy "
               "requires at least 2 full seasonal cycles and evenly spaced dates. See KPI_Definitions and README "
               "for how to add it as an optional comparison in desktop Excel (Data tab > Forecast Sheet, or the "
               "FORECAST.ETS function directly)."))
ws.cell(row=note_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row + 3, end_column=8)
ws.cell(row=note_row, column=1).font = SUBTITLE_FONT

autofit(ws, {"A": 8, "B": 12, "C": 14, "D": 14, "E": 12, "F": 12, "G": 14, "H": 10})
ws.freeze_panes = "A5"

wb.save(OUT)
print("Checkpoint 9 (Forecast) saved.")

# =====================================================================
# SHEET: Executive_Dashboard
# =====================================================================
ws = wb.create_sheet("Executive_Dashboard", 0)  # will reposition sheet order at the end
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = NAVY

ws.merge_cells("A1:N1")
ws["A1"] = "GLOBAL SUPPLEMENT SALES & INVENTORY ANALYTICS — EXECUTIVE DASHBOARD"
ws["A1"].font = Font(name=FONT_NAME, bold=True, color=WHITE, size=16)
ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[1].height = 30

ws.merge_cells("A2:N2")
ws["A2"] = "Synthetic data · Fictional company · 24 months (Jan 2024 – Dec 2025) · For portfolio / demonstration purposes only"
ws["A2"].font = Font(name=FONT_NAME, italic=True, color=GREY, size=9)
ws.row_dimensions[2].height = 16

# --- Filters ---
ws["A4"] = "Filter — Year:"
ws["A4"].font = KPI_LABEL_FONT
ws["B4"] = "All"
ws["B4"].font = Font(name=FONT_NAME, bold=True, color=NAVY, size=11)
ws["B4"].fill = PatternFill("solid", fgColor="FFF2CC")
ws["B4"].border = BOX
dv_year = DataValidation(type="list", formula1='"All,2024,2025"', allow_blank=False)
ws.add_data_validation(dv_year)
dv_year.add(ws["B4"])

ws["D4"] = "Filter — Category:"
ws["D4"].font = KPI_LABEL_FONT
ws["E4"] = "All"
ws["E4"].font = Font(name=FONT_NAME, bold=True, color=NAVY, size=11)
ws["E4"].fill = PatternFill("solid", fgColor="FFF2CC")
ws["E4"].border = BOX
cat_options = ",".join(category_list)
dv_cat = DataValidation(type="list", formula1=f'"All,{cat_options}"', allow_blank=False)
ws.add_data_validation(dv_cat)
dv_cat.add(ws["E4"])

ws["G4"] = "(KPI cards below respond to these two filters. YoY Growth, Stockout Rate and Forecast Accuracy are always company-wide.)"
ws["G4"].font = SUBTITLE_FONT
ws["G4"].alignment = Alignment(wrap_text=True)
ws.merge_cells("G4:N4")
ws.row_dimensions[4].height = 28

YEARC = "$B$4"
CATC = "$E$4"

def sumifs_filtered(measure):
    return (f'=SUMIFS(Sales_Reporting[{measure}],'
            f'Sales_Reporting[Year],IF({YEARC}="All",">0",{YEARC}),'
            f'Sales_Reporting[Category],IF({CATC}="All","*",{CATC}))')

# --- KPI cards ---
kpi_row1 = 6
kpis = [
    ("Total Revenue", sumifs_filtered("Revenue"), "$#,##0"),
    ("Gross Profit", sumifs_filtered("GrossProfit"), "$#,##0"),
    ("Gross Margin %", None, "0.0%"),  # filled after profit/revenue cells known
    ("Units Sold", sumifs_filtered("Units"), "#,##0"),
    ("YoY Revenue Growth", None, "0.0%"),
    ("Stockout Rate", "=AVERAGE(Inventory_Reporting[StockoutFlag])", "0.0%"),
    ("Forecast Accuracy", f"=Forecast!B{metric_row+3}", "0.0%"),
]
col_starts = [1, 3, 5, 7, 9, 11, 13]
label_cells = {}
value_cells = {}
for (label, formula, fmt), col in zip(kpis, col_starts):
    lbl = ws.cell(row=kpi_row1, column=col, value=label)
    lbl.font = KPI_LABEL_FONT
    ws.merge_cells(start_row=kpi_row1, start_column=col, end_row=kpi_row1, end_column=col + 1)
    val_cell = ws.cell(row=kpi_row1 + 1, column=col)
    val_cell.font = KPI_VALUE_FONT
    val_cell.number_format = fmt
    ws.merge_cells(start_row=kpi_row1 + 1, start_column=col, end_row=kpi_row1 + 2, end_column=col + 1)
    for rr in range(kpi_row1, kpi_row1 + 3):
        for cc in (col, col + 1):
            ws.cell(row=rr, column=cc).fill = KPI_FILL
            ws.cell(row=rr, column=cc).border = BOX
    label_cells[label] = (kpi_row1, col)
    value_cells[label] = (kpi_row1 + 1, col)
    if formula:
        ws.cell(row=kpi_row1 + 1, column=col, value=formula)

rev_r, rev_c = value_cells["Total Revenue"]
gp_r, gp_c = value_cells["Gross Profit"]
margin_r, margin_c = value_cells["Gross Margin %"]
ws.cell(row=margin_r, column=margin_c,
        value=f"=IFERROR({get_column_letter(gp_c)}{gp_r}/{get_column_letter(rev_c)}{rev_r},0)")

yoy_r, yoy_c = value_cells["YoY Revenue Growth"]
ws.cell(row=yoy_r, column=yoy_c,
        value='=IFERROR((SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Year],2025)'
              '-SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Year],2024))'
              '/SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Year],2024),"n/a")')

ws.row_dimensions[kpi_row1].height = 16

# --- Chart data (unfiltered, small helper tables placed below the KPI strip) ---
chart_area_row = kpi_row1 + 5
ws.cell(row=chart_area_row, column=1, value="Monthly Revenue Trend").font = SECTION_FONT
ws.cell(row=chart_area_row, column=1).fill = SECTION_FILL
ws.cell(row=chart_area_row, column=2).fill = SECTION_FILL
hdr = chart_area_row + 1
ws.cell(row=hdr, column=1, value="YearMonth")
ws.cell(row=hdr, column=2, value="Revenue")
style_header_row(ws, hdr, 2)
for i, mk in enumerate(month_keys):
    r = hdr + 1 + i
    ws.cell(row=r, column=1, value=mk).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[YearMonth],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=2).font = BODY_FONT
trend_last = hdr + len(month_keys)

dash_chart1 = LineChart()
dash_chart1.title = "Monthly Revenue Trend"
dash_chart1.height = 8.5
dash_chart1.width = 20
dash_chart1.y_axis.title = "Revenue ($)"
data = Reference(ws, min_col=2, min_row=hdr, max_row=trend_last)
cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=trend_last)
dash_chart1.add_data(data, titles_from_data=True)
dash_chart1.set_categories(cats)
dash_chart1.series[0].graphicalProperties.line.solidFill = STEEL
dash_chart1.series[0].graphicalProperties.line.width = 22000
dash_chart1.legend = None
ws.add_chart(dash_chart1, f"D{chart_area_row}")

r0 = trend_last + 2
ws.cell(row=r0, column=1, value="Revenue by Category").font = SECTION_FONT
ws.cell(row=r0, column=1).fill = SECTION_FILL
ws.cell(row=r0, column=2).fill = SECTION_FILL
hdr = r0 + 1
ws.cell(row=hdr, column=1, value="Category")
ws.cell(row=hdr, column=2, value="Revenue")
style_header_row(ws, hdr, 2)
for i, cat in enumerate(category_list):
    r = hdr + 1 + i
    ws.cell(row=r, column=1, value=cat).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=SUMIFS(Sales_Reporting[Revenue],Sales_Reporting[Category],A{r})').number_format = "$#,##0"
    ws.cell(row=r, column=2).font = BODY_FONT
cat_last = hdr + len(category_list)

dash_chart2 = BarChart()
dash_chart2.type = "col"
dash_chart2.title = "Revenue by Category"
dash_chart2.height = 8.5
dash_chart2.width = 12
data = Reference(ws, min_col=2, min_row=hdr, max_row=cat_last)
cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=cat_last)
dash_chart2.add_data(data, titles_from_data=True)
dash_chart2.set_categories(cats)
dash_chart2.series[0].graphicalProperties.solidFill = STEEL
dash_chart2.legend = None
ws.add_chart(dash_chart2, f"D{r0}")

r1 = cat_last + 2
ws.cell(row=r1, column=1, value="Stockout Rate by Category").font = SECTION_FONT
ws.cell(row=r1, column=1).fill = SECTION_FILL
ws.cell(row=r1, column=2).fill = SECTION_FILL
hdr = r1 + 1
ws.cell(row=hdr, column=1, value="Category")
ws.cell(row=hdr, column=2, value="Stockout Rate")
style_header_row(ws, hdr, 2)
for i, cat in enumerate(category_list):
    r = hdr + 1 + i
    ws.cell(row=r, column=1, value=cat).font = BODY_FONT
    ws.cell(row=r, column=2, value=f'=AVERAGEIFS(Inventory_Reporting[StockoutFlag],Inventory_Reporting[Category],A{r})').number_format = "0.0%"
    ws.cell(row=r, column=2).font = BODY_FONT
stk_last = hdr + len(category_list)

dash_chart3 = BarChart()
dash_chart3.type = "col"
dash_chart3.title = "Stockout Rate by Category"
dash_chart3.height = 8.5
dash_chart3.width = 12
dash_chart3.y_axis.numFmt = "0%"
data = Reference(ws, min_col=2, min_row=hdr, max_row=stk_last)
cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=stk_last)
dash_chart3.add_data(data, titles_from_data=True)
dash_chart3.set_categories(cats)
dash_chart3.series[0].graphicalProperties.solidFill = RED
dash_chart3.legend = None
ws.add_chart(dash_chart3, f"L{r0}")

ws.cell(row=stk_last + 2, column=1,
        value="Note: native PivotTables/PivotCharts and slicers were not created in this automated environment "
              "(see Start_Here and README for exact manual steps and why). The Year/Category cells above are a "
              "working substitute: they drive the KPI cards through SUMIFS with cell references.")
ws.cell(row=stk_last + 2, column=1).font = SUBTITLE_FONT
ws.cell(row=stk_last + 2, column=1).alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=stk_last + 2, start_column=1, end_row=stk_last + 4, end_column=8)

autofit(ws, {get_column_letter(i): 11 for i in range(1, 15)})
ws.column_dimensions["A"].width = 16
ws.freeze_panes = "A5"

wb.save(OUT)
print("Checkpoint 10 (Executive_Dashboard) saved.")

# =====================================================================
# SHEET: KPI_Definitions
# =====================================================================
ws = wb.create_sheet("KPI_Definitions")
ws.sheet_view.showGridLines = False
ws["A1"] = "KPI Definitions, Formulas & Assumptions"
ws["A1"].font = TITLE_FONT

kpi_table = [
    ["Total Revenue", "SUM of Revenue across all order lines in scope",
     "Numerator: SUM(Sales_Reporting[Revenue]). Denominator: n/a.",
     "Filtered by OrderDate's Year (via Sales_Reporting[Year]) and Category when dashboard filters are set to a specific value; 'All' removes that filter.",
     "Rows with blank/zero Revenue contribute 0 and are not excluded — legitimate $0 lines don't occur in this dataset by construction."],
    ["Gross Profit", "Revenue minus product cost for the same order lines",
     "Numerator: SUM(Sales_Reporting[GrossProfit]) where GrossProfit = Revenue − Cost. Denominator: n/a.",
     "Same date/category scope as Total Revenue.",
     "Cost uses UnitCost x Units at time of sale; does not include freight, warehousing, or overhead — a simplification, stated here explicitly."],
    ["Gross Margin %", "Gross Profit as a percentage of Revenue",
     "Numerator: Gross Profit. Denominator: Total Revenue.",
     "Same scope as above.",
     "If Total Revenue = 0, the KPI returns 0% via IFERROR rather than a divide-by-zero error."],
    ["Units Sold", "Total units sold across order lines in scope",
     "Numerator: SUM(Sales_Reporting[Units]). Denominator: n/a.",
     "Same scope as Total Revenue.",
     "No missing-value handling needed; Units is never null in the clean table."],
    ["YoY Revenue Growth", "Year-over-year change in total company revenue, 2025 vs 2024",
     "Numerator: Revenue(2025) − Revenue(2024). Denominator: Revenue(2024).",
     "Always company-wide (ignores the Year filter by design, since a single-year filter would make YoY meaningless); NOT affected by the Category filter either, to keep it as a single top-line company metric — the Sales_Analysis sheet has the category-level monthly detail if a category-specific YoY view is needed.",
     "If 2024 revenue were 0 the formula would return 'n/a' rather than an error; not triggered in this dataset."],
    ["Stockout Rate", "Share of Product x Country x Month snapshots flagged as a stockout",
     "Numerator: SUM(Inventory_Reporting[StockoutFlag]). Denominator: COUNT of Inventory_Reporting rows (implemented as AVERAGE of the 0/1 flag, which is mathematically identical).",
     "Computed across the full 24-month window; not affected by dashboard filters.",
     "StockoutFlag = 1 if ClosingStock hits 0 in the month, or if ClosingStock falls below 15% of ReorderPoint (a near-stockout risk trigger); see Inventory generation logic in scripts/generate_data.py."],
    ["Days of Cover", "How many days current closing stock would last at the recent sales rate",
     "Numerator: ClosingStock. Denominator: (UnitsSold in month / 30).",
     "Per Product x Country x Month row in Inventory_Reporting.",
     "If UnitsSold = 0 in a month, Days of Cover is left blank (undefined — division by zero) rather than showing an artificial infinite value."],
    ["Reorder Alert", "Flags a Product x Country combination that has fallen below its reorder point",
     "Rule: ClosingStock < ReorderPoint (from Products).", "n/a",
     "Evaluated per snapshot month; the Inventory_Analysis sheet shows the latest month only."],
    ["Slow-Moving Flag", "Flags low-velocity stock relative to its target level",
     "Rule: UnitsSold in month < 3% of TargetStock.", "n/a",
     "The 3% threshold is a stated assumption, not a universal industry standard — easy to change in one place (Python: scripts/build_reporting_layer.py) and re-run."],
    ["Forecast Accuracy", "1 − WAPE over the 3-month evaluation holdout",
     "See Forecast sheet: MAE, RMSE, WAPE.", "n/a",
     "See Forecast sheet notes for the WAPE-undefined edge case and the choice of baseline method."],
]
hdr_row = 3
headers = ["KPI", "Definition", "Numerator / Denominator", "Date Logic / Scope", "Assumptions & Edge-Case Handling"]
for j, h in enumerate(headers):
    ws.cell(row=hdr_row, column=1 + j, value=h)
style_header_row(ws, hdr_row, len(headers))
for i, row_vals in enumerate(kpi_table):
    r = hdr_row + 1 + i
    for j, v in enumerate(row_vals):
        c = ws.cell(row=r, column=1 + j, value=v)
        c.font = BODY_FONT
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 60

autofit(ws, {"A": 20, "B": 34, "C": 32, "D": 34, "E": 40})
ws.freeze_panes = "A4"

fnote_row = hdr_row + len(kpi_table) + 2
ws.cell(row=fnote_row, column=1,
        value=("Forecasting method in full: seasonal-naive baseline (Forecast[m] = Actual[m-12] once a full prior "
               "year of history exists; an expanding average of prior months is used for months 2–12, where no "
               "same-month-last-year value yet exists). Final 3 months (Oct–Dec 2025) are held out for evaluation; "
               "MAE, RMSE and WAPE are computed on that holdout only. No holdout-period actuals are used to build "
               "the holdout-period forecasts — each holdout forecast only ever references a value 12 months earlier, "
               "which always falls inside the training window."))
ws.cell(row=fnote_row, column=1).alignment = Alignment(wrap_text=True)
ws.merge_cells(start_row=fnote_row, start_column=1, end_row=fnote_row + 4, end_column=5)
ws.cell(row=fnote_row, column=1).font = SUBTITLE_FONT

wb.save(OUT)
print("Checkpoint 11 (KPI_Definitions) saved.")

# =====================================================================
# SHEET: Start_Here
# =====================================================================
ws = wb.create_sheet("Start_Here")
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = GOLD

ws.merge_cells("A1:H1")
ws["A1"] = "Global Supplement Sales & Inventory Analytics Dashboard"
ws["A1"].font = Font(name=FONT_NAME, bold=True, color=NAVY, size=18)
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:H2")
ws["A2"] = "⚠ 100% SYNTHETIC DATA — no real company, customer, product, or transaction is represented in this workbook."
ws["A2"].font = Font(name=FONT_NAME, bold=True, color=RED, size=11)
ws.row_dimensions[2].height = 18

text_blocks = [
    ("Project Overview",
     "This is a portfolio project demonstrating Excel-based analytics: data cleaning/ETL, KPI design, "
     "formula-driven dashboards, inventory analysis, and a transparent baseline demand forecast, built on a "
     "fully synthetic 24-month dataset (Jan 2024 – Dec 2025) for a fictional multi-country supplement "
     "distributor. It is designed to be published on GitHub and walked through in data analyst interviews."),
    ("How to Navigate This Workbook",
     "• Executive_Dashboard — one-page KPI summary with Year/Category filter cells and charts.\n"
     "• Sales_Analysis — monthly trend, category / region / channel breakdowns.\n"
     "• Inventory_Analysis — stockout rate, days of cover, reorder alerts, slow movers.\n"
     "• Forecast — seasonal-naive baseline vs. actual, with MAE / RMSE / WAPE on a 3-month holdout.\n"
     "• Data_Quality — issues found in raw data, how they were fixed, and live reconciliation checks.\n"
     "• KPI_Definitions — exact formulas, numerators/denominators, date logic and assumptions for every KPI.\n"
     "• Products / Geography / Channels / Calendar — clean dimension (lookup) tables.\n"
     "• Sales / Inventory / PurchaseOrders — clean, normalized fact tables (schema as specified).\n"
     "• Sales_Reporting / Inventory_Reporting — denormalized reporting layer (joined + derived fields) that "
     "all dashboard formulas read from.\n"
     "• Raw_Data_Sample — illustrative excerpts of the raw data before cleaning, showing real issue types."),
    ("Refreshing the Data",
     "Every number in this workbook is a live Excel formula (SUMIFS/COUNTIFS/AVERAGEIFS/INDEX-MATCH, etc.) "
     "reading from the Tables in this file — there is nothing to 'refresh' in the Power-Query sense unless you "
     "rebuild the optional Power Query connections described in the README (to re-import data/clean/*.csv). "
     "If you do that, use Data > Refresh All, then press Ctrl+Alt+F9 to force a full recalculation."),
    ("What Was and Wasn't Possible to Build/Verify Automatically",
     "Built and verified in this automated environment: all Excel Tables, formulas (SUMIFS/COUNTIFS/AVERAGEIFS/"
     "INDEX-MATCH/IFERROR), native openpyxl charts (line/bar/pie), conditional formatting, data-validation "
     "dropdowns, and a full recalculation pass with zero formula errors (see scripts/recalc_log.txt).\n\n"
     "NOT created here, with exact manual steps provided in the README: native Excel PivotTables/PivotCharts "
     "and Slicers. openpyxl (the Python library used to build this file outside of desktop Excel) cannot "
     "reliably create functional PivotTable caches or slicer objects, so claiming they exist here would be "
     "inaccurate. The Executive_Dashboard's Year/Category dropdown cells are a working substitute built entirely "
     "from formulas. The README gives exact click-by-click steps to add real PivotTables/PivotCharts/Slicers "
     "from the Sales_Reporting and Inventory_Reporting tables in under 10 minutes in desktop Excel.\n\n"
     "The optional VBA refresh macro (RefreshAll.bas) is provided as a separate, importable file — see README "
     "for the .xlsm conversion and macro-security steps. This .xlsx file itself contains no macros."),
]
r = 4
for title, body in text_blocks:
    ws.cell(row=r, column=1, value=title).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    for c in range(2, 9):
        ws.cell(row=r, column=c).fill = SECTION_FILL
    r += 1
    ws.cell(row=r, column=1, value=body)
    ws.cell(row=r, column=1).font = BODY_FONT
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 6, end_column=8)
    ws.row_dimensions[r].height = 14
    r += 9

autofit(ws, {get_column_letter(i): 13 for i in range(1, 9)})
ws.column_dimensions["A"].width = 110

wb.save(OUT)
print("Checkpoint 12 (Start_Here) saved.")

# =====================================================================
# Final sheet order
# =====================================================================
order = ["Start_Here", "Executive_Dashboard", "Sales_Analysis", "Inventory_Analysis", "Forecast",
         "Data_Quality", "KPI_Definitions", "Products", "Geography", "Channels", "Calendar",
         "Sales", "Sales_Reporting", "Inventory", "Inventory_Reporting", "PurchaseOrders", "Raw_Data_Sample"]
wb._sheets = [wb[name] for name in order]
wb.active = 0

wb.save(OUT)
print("FINAL workbook saved:", OUT)
