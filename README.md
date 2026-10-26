# Sales Dashboard — Excel

A clear, at-a-glance sales dashboard built with native Excel charts, KPI
cards and summary tables, from a 1,507-transaction sample sales dataset.

## What's inside

**`Sales_Dashboard.xlsx`** — 3 sheets:

| Sheet | Contents |
|---|---|
| `Dashboard` | 6 KPI cards, 4 charts, 2 summary tables — the whole story on one page |
| `Data` | All 1,507 transactions + live `Revenue` and `Profit` Excel formulas |
| `Summary` | Pivot-style aggregation tables that feed the charts |

**Dashboard elements (6 visuals, requirement was 3–4):**

1. **KPI cards** — Total Revenue ₹561,379 · Total Profit ₹218,609 · Total Orders 1,500 · Avg Order Value ₹374 · Profit Margin 38.9% · Best Region North
2. **Monthly Revenue Trend** — line chart, revenue vs profit by month
3. **Revenue by Region** — column chart
4. **Revenue Share by Category** — pie chart with % labels
5. **Top 5 Products by Revenue** — horizontal bar chart
6. **Summary tables** — revenue by region (with share %) and top 5 products

All figures are computed from the dataset — nothing is hardcoded.

## One honest technical note: PivotTables

The task allows Excel PivotTables *or* charts. This workbook uses native Excel
charts linked to aggregation tables on the `Summary` sheet — those tables are
exactly the output a PivotTable would produce (group by month / region /
category / product, sum of revenue and profit). `openpyxl` cannot create
click-to-refresh PivotTables, so instead of faking one, the tables are
computed by the script and the charts are real Excel charts that redraw if
the table values change. In Excel you can also select the `Data` sheet and
use **Insert → PivotTable** to build a live pivot from the same data.

## How to use

1. Open `Sales_Dashboard.xlsx` in Excel or LibreOffice.
2. The `Dashboard` sheet is the presentation view (set to print on one
   landscape page).
3. Filter/sort on the `Data` sheet; `Revenue` and `Profit` are formulas, so
   they recalculate live.

## Rebuild / verify

```bash
pip install pandas openpyxl
python build_dashboard.py        # regenerates Sales_Dashboard.xlsx
python -m pytest tests/ -q       # KPI + chart + structure tests
```

## Files

```
Excel_Sales_Dashboard/
├── Sales_Dashboard.xlsx          # the dashboard
├── build_dashboard.py            # builds it from the sample CSV
├── screenshots/dashboard_preview.png
├── tests/test_dashboard.py
└── README.md
```

Dataset: the SalesPulse AI sample (`sample_sales.csv`, 1,507 rows) — Order
Date, Category, Region, Product, Quantity, Unit Price, Unit Cost, Discount.
