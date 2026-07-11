# End-to-End Analysis Case Study
### Financial Sample Dataset — Sales, Discounts & Profitability

Complete deliverable for the case study brief:

- ✅ Clean and explore the dataset in Python
- ✅ Analyze data using SQL and Pandas
- ✅ Visualize results using **Excel** and a **BI tool (Tableau)**
- ✅ Final report including charts, findings, and recommendations

## Contents

```
├── End_to_End_Analysis.ipynb           ← main notebook, executed with outputs (clean → SQL → charts)
├── 01_clean_data.py                    ← standalone script: cleaning & validation
├── 02_sql_analysis.py                  ← standalone script: SQL analysis (SQLite)
├── 03_visualize.py                     ← standalone script: Matplotlib/Seaborn charts
├── build_excel_dashboard.py            ← script that generates the Excel dashboard below
├── data/
│   ├── Sample_data.xlsx                ← raw source data (700 orders, 16 fields)
│   ├── cleaned_financial_data.csv      ← cleaned dataset (output of step 1)
│   └── financial_data.db               ← SQLite database (output of step 2)
├── charts/                             ← 6 PNG charts from the Python pipeline
│   ├── 01_profit_by_segment.png
│   ├── 02_margin_by_discount_band.png
│   ├── 03_heatmap_segment_discount.png
│   ├── 04_sales_profit_by_product.png
│   ├── 05_sales_by_country.png
│   └── 06_monthly_trend.png
├── excel_dashboard/
│   └── Financial_Analysis_Dashboard.xlsx   ← formula-driven pivot summaries + native Excel charts + Dashboard tab
├── bi_tableau/
│   ├── tableau_dashboard.png            ← screenshot of the Tableau dashboard ("Dashboard 2")
│   └── README.md                        ← sheet-by-sheet breakdown + rebuild steps
├── End_to_End_Analysis_Report.pdf       ← final written report (charts, findings, recommendations)
├── PDF_Explanation_Report.pdf           ← walkthrough of the whole pipeline: what each script does, how outputs were produced/verified, and how everything ties together (submission deliverable)
└── requirements.txt
```

## How to run

The notebook (`End_to_End_Analysis.ipynb`) and the Excel workbook already
contain executed/recalculated output — open and read them directly, no
re-running required. To reproduce from scratch:

```bash
pip install -r requirements.txt

python 01_clean_data.py            # -> data/cleaned_financial_data.csv
python 02_sql_analysis.py          # -> data/financial_data.db + printed SQL results
python 03_visualize.py             # -> charts/*.png
python build_excel_dashboard.py    # -> excel_dashboard/Financial_Analysis_Dashboard.xlsx
```

For Tableau, connect Tableau Desktop directly to
`data/cleaned_financial_data.csv` and follow `bi_tableau/README.md` to
rebuild the dashboard sheet-by-sheet.

## Pipeline overview

1. **Clean & validate** (`01_clean_data.py`) — load the raw 700-row Excel
   file, confirm there are no duplicates, fix the 53 blank `Discount Band`
   values (all had `Discounts == 0`, so relabeled `'None'`), and cross-check
   `Gross Sales`, `Sales`, and `Profit` against their formulas (0 mismatches).
2. **SQL analysis** (`02_sql_analysis.py`) — load the cleaned data into a
   SQLite database (`sales` table) and run the GROUP BY queries behind every
   figure in the report: profit by segment, margin by discount band, the
   segment × discount-band cross-tab, product performance, country
   performance, and the monthly trend.
3. **Python visualize** (`03_visualize.py`) — Matplotlib/Seaborn charts
   matching the report's six figures.
4. **Excel dashboard** (`build_excel_dashboard.py`) — the same six analyses
   rebuilt as live `SUMIF`/`SUMIFS` pivot-style tables (formulas, not
   hardcoded numbers) with native Excel charts, plus a one-page `Dashboard`
   tab with KPI tiles and copies of all six charts. Verified with zero
   formula errors on recalculation.
5. **Tableau dashboard** (`bi_tableau/`) — an 8-sheet interactive dashboard
   (seasonal trend, COGS trend, product bubbles, country map, etc.) with
   Country/Product/Sales/Profit/Month filters, screenshotted for this
   deliverable since Tableau workbooks need a licensed desktop app to open.

## Key findings (see full report for details)

- Enterprise is the only segment operating at a net loss (-3.13% margin),
  concentrated in Medium/High-discount orders.
- Channel Partners has the best margin in the business (73.1%).
- Margin erodes near-linearly with discount depth (21.9% → 9.1%).
- Paseo drives the most revenue; Amarilla is the most margin-efficient
  product.
- All five countries are solidly profitable — geography is not a risk
  factor.
- October and December are consistently the strongest revenue months.

## Cross-checks

Total Profit reconciles to **$16,893,702** across all three tools:
Python/SQL, the Excel `Dashboard` KPI tile, and the Tableau dashboard's
Profit tile — confirming the pipeline is internally consistent end to end.
