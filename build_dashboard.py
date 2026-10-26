#!/usr/bin/env python3
"""Build Sales_Dashboard.xlsx — KPI cards, 4 native Excel charts, summary tables.

Data: the SalesPulse sample dataset (1,507 sales transactions).
Charts are native openpyxl charts linked to the Summary tables, so they
redraw if the table values change. Summary tables are the pivot-equivalent
aggregations (openpyxl cannot create click-to-refresh PivotTables; the same
aggregations can be reproduced in Excel via Insert -> PivotTable on Data).
"""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.text import Font as ChartFont, CharacterProperties

SRC = Path("/home/hatch/workspace/your_files/codeorbit-internship/SalesPulse_AI/data/sample_sales.csv")
OUT = Path("/home/hatch/workspace/your_files/Excel_Sales_Dashboard")
OUT.mkdir(parents=True, exist_ok=True)

NAVY = "16294A"
NAVY2 = "1F3B63"
GOLD = "D4A017"
LIGHT = "E9EFF8"
WHITE = "FFFFFF"
GREY = "5F7391"

thin = Side(style="thin", color="B0BECF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

df = pd.read_csv(SRC)
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce", format="mixed")
df["Revenue"] = df["Quantity"] * df["Unit Price"] * (1 - df["Discount"])
df["Profit"] = df["Revenue"] - df["Quantity"] * df["Unit Cost"]
df["Month"] = df["Order Date"].dt.to_period("M").dt.to_timestamp()

total_revenue = df["Revenue"].sum()
total_profit = df["Profit"].sum()
total_orders = df["Order ID"].nunique()
total_units = int(df["Quantity"].sum())
aov = total_revenue / total_orders
margin = total_profit / total_revenue * 100
best_region = df.groupby("Region")["Revenue"].sum().idxmax()
best_product = df.groupby("Product Name")["Revenue"].sum().idxmax()
period = f"{df['Order Date'].min():%b %Y} – {df['Order Date'].max():%b %Y}"

monthly = df.groupby("Month").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"),
                                  Orders=("Order ID", "nunique")).reset_index()
by_region = df.groupby("Region").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"),
                                     Orders=("Order ID", "nunique")).reset_index()
by_region["Share %"] = by_region["Revenue"] / total_revenue * 100
by_region = by_region.sort_values("Revenue", ascending=False)
by_cat = df.groupby("Category").agg(Revenue=("Revenue", "sum"), Profit=("Profit", "sum"),
                                    Orders=("Order ID", "nunique")).reset_index().sort_values("Revenue", ascending=False)
top_prod = (df.groupby("Product Name").agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"),
                                           Orders=("Order ID", "nunique"))
            .reset_index().sort_values("Revenue", ascending=False).head(5))

print(f"Revenue {total_revenue:,.2f} | Profit {total_profit:,.2f} | Orders {total_orders} | "
      f"AOV {aov:,.2f} | Margin {margin:.1f}% | Best region {best_region} | Best product {best_product}")

wb = Workbook()

# ---------------------------------------------------------------- Data sheet
ws = wb.active
ws.title = "Data"
orig_cols = [c for c in df.columns if c not in ("Revenue", "Profit", "Month")]
headers = orig_cols + ["Revenue", "Profit"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=c, value=h)
    cell.font = Font(bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="center")
    cell.border = BORDER
for r, row in enumerate(df[orig_cols].itertuples(index=False), 2):
    vals = list(row)
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = BORDER
        if c in (8, 9, 10):
            cell.number_format = '#,##0.00'
        if c == 2:
            cell.number_format = 'YYYY-MM-DD'
    # Live Excel formulas for Revenue and Profit
    ws.cell(row=r, column=14, value=f"=H{r}*I{r}*(1-K{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=15, value=f"=N{r}-H{r}*J{r}").number_format = '#,##0.00'
widths = [11, 12, 11, 16, 11, 18, 13, 9, 11, 11, 9, 11, 16, 13, 13]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:O{len(df) + 1}"
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ------------------------------------------------------------- Summary sheet
sm = wb.create_sheet("Summary")


def write_table(ws, start_row, title, frame, money_cols=()):
    ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=13, color=NAVY)
    for c, h in enumerate(frame.columns, 1):
        cell = ws.cell(row=start_row + 1, column=c, value=h)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY2)
        cell.border = BORDER
    for r, row in enumerate(frame.itertuples(index=False), start_row + 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c in money_cols:
                cell.number_format = '#,##0.00'
            if hasattr(v, "strftime") and not isinstance(v, str):
                cell.number_format = 'YYYY-MM'
    return start_row + 2 + len(frame)  # first empty row after table


r = write_table(sm, 1, "Monthly trend", monthly, money_cols=(2, 3))
m_rows = len(monthly)
r = write_table(sm, r + 2, "Revenue by region", by_region, money_cols=(2, 3))
reg_rows = len(by_region)
reg_start = r - 1 - reg_rows  # header row of region table
r = write_table(sm, r + 2, "Revenue by category", by_cat, money_cols=(2, 3))
cat_rows = len(by_cat)
cat_start = r - 1 - cat_rows
r = write_table(sm, r + 2, "Top 5 products by revenue", top_prod, money_cols=(2,))
tp_rows = len(top_prod)
tp_start = r - 1 - tp_rows
for col, w in zip("ABCDE", [14, 16, 16, 10, 10]):
    sm.column_dimensions[col].width = w

# ------------------------------------------------------------ Dashboard sheet
db = wb.create_sheet("Dashboard", 0)
for col, w in zip("ABCDEFGHIJKL", [16] * 12):
    db.column_dimensions[col].width = w
db.sheet_properties.pageSetUpPr.fitToPage = True
db.page_setup.orientation = "landscape"
db.page_setup.fitToWidth = 1
db.page_setup.fitToHeight = 1

db.merge_cells("A1:L1")
t = db["A1"]
t.value = "SALES DASHBOARD"
t.font = Font(bold=True, size=22, color=NAVY)
db.merge_cells("A2:L2")
s = db["A2"]
s.value = f"Sample sales data  ·  {period}  ·  {len(df):,} transactions  ·  All figures computed from the Data sheet"
s.font = Font(size=11, color=GREY)

kpis = [
    ("Total Revenue", total_revenue, "₹ #,##0"),
    ("Total Profit", total_profit, "₹ #,##0"),
    ("Total Orders", total_orders, "#,##0"),
    ("Avg Order Value", aov, "₹ #,##0"),
    ("Profit Margin", margin / 100, "0.0%"),
    ("Best Region", best_region, None),
]
# KPI cards: rows 4-6, each 2 columns wide
for i, (label, value, fmt) in enumerate(kpis):
    c0 = 1 + i * 2
    c1 = c0 + 1
    db.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c1)
    lab = db.cell(row=4, column=c0, value=label)
    lab.font = Font(bold=True, size=11, color=WHITE)
    lab.fill = PatternFill("solid", fgColor=NAVY)
    lab.alignment = Alignment(horizontal="center")
    db.merge_cells(start_row=5, start_column=c0, end_row=6, end_column=c1)
    val = db.cell(row=5, column=c0, value=value)
    val.font = Font(bold=True, size=18, color=NAVY)
    val.fill = PatternFill("solid", fgColor="FFF3D6")
    val.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        val.number_format = fmt
    for rr in (4, 5, 6):
        for cc in (c0, c1):
            db.cell(row=rr, column=cc).border = BORDER
            if rr == 5:
                db.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor="FFF3D6")


def style_series(chart, fills):
    for ser, fill in zip(chart.series, fills):
        ser.graphicalProperties.solidFill = fill
        ser.graphicalProperties.line.solidFill = fill


# Monthly trend line chart (Summary rows 3..3+m_rows-1; header row 2)
lc = LineChart()
lc.title = "Monthly Revenue Trend"
lc.y_axis.title = "₹"
lc.x_axis.title = "Month"
lc.height = 10.5
lc.width = 26
data = Reference(sm, min_col=2, min_row=2, max_col=3, max_row=2 + m_rows)
cats = Reference(sm, min_col=1, min_row=3, max_row=2 + m_rows)
lc.add_data(data, titles_from_data=True)
lc.set_categories(cats)
style_series(lc, [NAVY, GOLD])
lc.legend.position = "b"
db.add_chart(lc, "A8")

# Region bar chart
bc = BarChart()
bc.type = "col"
bc.title = "Revenue by Region"
bc.y_axis.title = "₹"
bc.height = 9
bc.width = 14.5
data = Reference(sm, min_col=2, min_row=reg_start, max_row=reg_start + reg_rows)
cats = Reference(sm, min_col=1, min_row=reg_start + 1, max_row=reg_start + reg_rows)
bc.add_data(data, titles_from_data=True)
bc.set_categories(cats)
bc.shape = 4
style_series(bc, [NAVY2])
bc.legend = None
db.add_chart(bc, "A27")

# Category pie chart
pc = PieChart()
pc.title = "Revenue Share by Category"
pc.height = 9
pc.width = 14.5
data = Reference(sm, min_col=2, min_row=cat_start, max_row=cat_start + cat_rows)
cats = Reference(sm, min_col=1, min_row=cat_start + 1, max_row=cat_start + cat_rows)
pc.add_data(data, titles_from_data=True)
pc.set_categories(cats)
pc.dataLabels = DataLabelList()
pc.dataLabels.showPercent = True
db.add_chart(pc, "H27")

# Top products bar chart (horizontal)
bp = BarChart()
bp.type = "bar"
bp.title = "Top 5 Products by Revenue"
bp.height = 10
bp.width = 26
data = Reference(sm, min_col=2, min_row=tp_start, max_row=tp_start + tp_rows)
cats = Reference(sm, min_col=1, min_row=tp_start + 1, max_row=tp_start + tp_rows)
bp.add_data(data, titles_from_data=True)
bp.set_categories(cats)
style_series(bp, [GOLD])
bp.legend = None
db.add_chart(bp, "A45")

# Visible summary tables on the Dashboard (formulas -> Summary, single source of truth)
db["A62"] = "Revenue by Region"
db["A62"].font = Font(bold=True, size=13, color=NAVY)
for c, h in enumerate(["Region", "Revenue", "Share %"], 1):
    cell = db.cell(row=63, column=c, value=h)
    cell.font = Font(bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY2)
    cell.border = BORDER
for i in range(reg_rows):
    sr = reg_start + 1 + i
    db.cell(row=64 + i, column=1, value=f"=Summary!A{sr}").border = BORDER
    db.cell(row=64 + i, column=2, value=f"=Summary!B{sr}").number_format = '#,##0'
    db.cell(row=64 + i, column=2).border = BORDER
    db.cell(row=64 + i, column=3, value=f"=Summary!E{sr}").number_format = '0.0"%"'
    db.cell(row=64 + i, column=3).border = BORDER

db["E62"] = "Top 5 Products"
db["E62"].font = Font(bold=True, size=13, color=NAVY)
for c, h in enumerate(["Product", "Revenue", "Units"], 5):
    cell = db.cell(row=63, column=c, value=h)
    cell.font = Font(bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY2)
    cell.border = BORDER
for i in range(tp_rows):
    sr = tp_start + 1 + i
    db.cell(row=64 + i, column=5, value=f"=Summary!A{sr}").border = BORDER
    db.cell(row=64 + i, column=6, value=f"=Summary!B{sr}").number_format = '#,##0'
    db.cell(row=64 + i, column=6).border = BORDER
    db.cell(row=64 + i, column=7, value=f"=Summary!C{sr}").number_format = '#,##0'
    db.cell(row=64 + i, column=7).border = BORDER

db.sheet_properties.pageSetUpPr.fitToPage = True
db.print_title_rows = '1:2'

wb.save(OUT / "Sales_Dashboard.xlsx")
print("saved", OUT / "Sales_Dashboard.xlsx")
