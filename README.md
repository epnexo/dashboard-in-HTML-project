# Telecom Transactions Report

A small, dependency-light pipeline that takes a messy telecom transactions CSV, cleans it, and
produces a self-contained interactive HTML dashboard, an Excel workbook and a data-cleaning
report — in about two seconds, with one double-click.

> **Demo data.** The file in `data/raw/` is synthetic demo data styled for Australia (AUD,
> Australian cities). It was derived from a public Nigerian sample by `scripts/make_demo_au.py`
> and does not represent real Australian figures.

## What you get

| Output | Description |
|---|---|
| `reports/telecom_report.html` | Single-file dashboard: six pages (Overview, Sales, Channels & regions, Operations, Customers, Data quality), filters, month/quarter chips, click-to-filter bars, tooltips, CSV export, print view, light/dark theme. |
| `reports/telecom_report.xlsx` | Summary (live formulas), Charts, Data and Cleaning sheets. |
| `cleaning/cleaning_report.html` | What was wrong in the source file and what was done about it. |
| `cleaning/findings_detail.csv` | Every correction, row by row, with the value before and after. |
| `data/clean/*_clean.csv` | The cleaned dataset. |

The Overview page opens with automatically generated **key findings** (for example, the region
whose failure rate is above the alert threshold and the revenue lost to failed transactions).
They are rule-based and recalculated with every filter; no AI runs at report time.

## Run it

Requirements: Python 3.10+ on Windows, plus two packages for the Excel file:

```
python -m pip install openpyxl pillow
```

Then either double-click **`run.bat`** and press *Clean and build report*, or:

```
python scripts/pipeline.py --open
```

To process your own file, drop the CSV into `data/raw/` (the newest file is used) or pass
`--input path/to/file.csv`. It must have these columns:

```
transaction_id, txn_date, customer_id, region, customer_segment, product_type,
bundle_name, channel, data_volume_mb, amount, status
```

## What the cleaning does

- Removes duplicate rows by `transaction_id`.
- Normalises dates written in five formats (`DD/MM/YYYY`, `MM-DD-YYYY`, `YYYY.MM.DD`,
  `DD Mon YYYY`, Excel serials) to `YYYY-MM-DD`; impossible or empty dates are left blank.
- Strips currency symbols, codes, commas and spaces from amounts; parentheses mean negative.
- Imputes empty amounts with the bundle's list price (its most frequent successful amount) and
  replaces outliers that do not match it. Both cases are flagged in `amount_flag`.
- Unifies region and channel aliases using `scripts/rules.json`.
- Fixes customer IDs with stray spaces or a missing `CUST` prefix.

## Customise without touching code

| File | What it controls |
|---|---|
| `scripts/settings.json` | Currency label, revenue and success-rate targets, failure alert threshold, inactivity window. |
| `scripts/rules.json` | Region and channel aliases. |
| `scripts/logo.png` | Logo shown in the dashboard header and the Excel workbook. |

## Project layout

```
run.bat                  double-click launcher
scripts/
  pipeline.py            cleaning + report generation (standard library only)
  excel.py               Excel workbook (openpyxl)
  app.py                 small desktop window (tkinter)
  report_template.html   dashboard template (HTML + CSS + vanilla JS, no libraries)
  make_demo_au.py        one-off generator of the Australian demo file
data/raw|clean|source    input, cleaned output, original sample
cleaning/                cleaning report and row-level findings
reports/                 HTML dashboard and Excel workbook
```

Built by EPNexo.
