<div align="center">

<img src="https://img.shields.io/badge/M--PESA-Analytics-14A800?style=for-the-badge&labelColor=121212&color=14A800" alt="M-PESA Analytics"/>

# M-PESA Financial Analytics System

**A complete financial intelligence system for Kenyan M-PESA statement data.**
Extract transactions from PDF statements, run comprehensive financial analysis,
and explore every result through an interactive browser dashboard — all in KES.

<br/>

[![Python](https://img.shields.io/badge/Python-3.10%2B-14A800?style=flat-square&logo=python&logoColor=white&labelColor=121212)](https://python.org)
[![pandas](https://img.shields.io/badge/pandas-2.0%2B-14A800?style=flat-square&logo=pandas&logoColor=white&labelColor=121212)](https://pandas.pydata.org)
[![Chart.js](https://img.shields.io/badge/Chart.js-4.4-14A800?style=flat-square&labelColor=121212)](https://chartjs.org)
[![pdf.js](https://img.shields.io/badge/pdf.js-3.11-14A800?style=flat-square&labelColor=121212)](https://mozilla.github.io/pdf.js/)
[![License](https://img.shields.io/badge/License-MIT-14A800?style=flat-square&labelColor=121212)](LICENSE)

<br/>

</div>

---

## What this system does
<img width="1898" height="898" alt="image" src="https://github.com/user-attachments/assets/6f13429c-ff82-416e-8776-6e92b87332e8" />

This system takes your Safaricom M-PESA PDF statements and turns them into a structured financial picture. It extracts every transaction exactly as printed, removes duplicates across overlapping statement periods, computes financial metrics at every time granularity, and presents everything through a self-contained dashboard that runs in any browser with no server required.

The pipeline has two parts that can be used independently or together:

- **Part 1** is a Python script that reads PDF files, extracts all transactions, and writes a clean CSV.
- **Part 2** is a Python analytics engine that receives the clean data and generates an Excel workbook plus an HTML dashboard.
- **The standalone dashboard** accepts the CSV or the raw PDFs directly in the browser, extracting and analysing without any Python required.
<img width="1881" height="881" alt="image" src="https://github.com/user-attachments/assets/4c223062-5755-4ccf-bf12-3f899a125042" />

---

## Repository contents

```
mpesa-analytics/
├── mpesa_part1_extractor.py     Python script: PDF to CSV extraction
├── mpesa_part2_analytics.py     Python script: analytics engine and dashboard generator
├── mpesa_dashboard.html         Self-contained browser dashboard (no server needed)
├── requirements.txt             Python dependencies
├── .gitignore
└── README.md
```

---

## System architecture

```
Your M-PESA PDFs
       |
       v
[ Part 1: Python extractor ]
  - Reads all PDFs in folder
  - Strips headers, footers, boilerplate
  - Assembles multi-line transaction rows
  - Parses receipt, date, time, details, amounts
  - Deduplicates across overlapping statement periods
  - Outputs: MPESA_verbatim_extract.csv
       |
       v
[ Part 2: Analytics engine ]          OR       [ HTML Dashboard ]
  - Feature engineering                           - Upload CSV or PDF directly
  - Financial metrics                             - pdf.js extracts PDF text
  - Entity extraction                             - Same Part 1 parser runs in browser
  - Anomaly detection (Z-score)                   - All analysis runs client-side
  - Excel export (6 sheets)                       - No data leaves your device
  - HTML dashboard output
```

---

## Quick start

### Option A: Python pipeline then browser dashboard

**Step 1: Install dependencies**

```bash
pip install -r requirements.txt
```

**Step 2: Place your PDFs**

Copy all your M-PESA PDF statements into a single folder. Overlapping periods are handled automatically. For example:

```
D:\Downloads\MPESA\
├── MPESA_Statement_Feb_2026.pdf
├── MPESA_Statement_Jan_Jun_2026.pdf
└── MPESA_Statement_Jan_Dec_2026.pdf
```

**Step 3: Configure the folder path**

Open `mpesa_part1_extractor.py` and set:

```python
PDF_FOLDER = Path(r"D:\Downloads\MPESA")
```

**Step 4: Run Part 1**

```bash
python mpesa_part1_extractor.py
```

This creates `MPESA_verbatim_extract.csv` in your PDF folder. The terminal prints a live progress report showing files found, pages processed, rows extracted, duplicates removed, and financial totals.

**Step 5: Open the dashboard**

Double-click `mpesa_dashboard.html`. Drop the CSV file onto the upload zone. Click Run Analytics.

### Option B: Browser-only (no Python required)

Open `mpesa_dashboard.html` directly in Chrome, Edge, or Firefox. Drop your PDF files or CSV onto the upload zone. The dashboard extracts and analyses everything inside the browser. An internet connection is required only to load Chart.js and pdf.js from CDN on first use.

---

## Part 1: PDF extractor in detail

The extractor mirrors the behaviour of `pdfplumber` at the text level:

| Step | What happens |
|---|---|
| Discovery | Scans `PDF_FOLDER` for all `.pdf` files alphabetically |
| Extraction | Opens each PDF page and extracts all text content |
| Cleaning | Removes page numbers, headers, footers, summary sections, and all Safaricom boilerplate using 28 noise patterns |
| Assembly | Joins continuation lines to their receipt anchor using the receipt number regex `^([A-Z0-9]{8,14})\s+` |
| Parsing | Extracts receipt number, ISO date, time, status, description, paid in, withdrawn, and balance from each assembled row |
| Deduplication | Removes exact duplicates across files using all 8 fields as the composite key |
| Sorting | Orders all rows chronologically oldest first |
| Export | Writes `MPESA_verbatim_extract.csv` with UTF-8 BOM encoding |

**Output columns**

| Column | Description |
|---|---|
| Row No. | Sequential row number after deduplication and sorting |
| Receipt No. | M-PESA receipt identifier |
| Completion Date | Transaction date in YYYY-MM-DD format |
| Completion Time | Transaction time in HH:MM:SS format |
| Details | Full transaction description exactly as printed |
| Transaction Status | Completed, Failed, Reversed, or Pending |
| Paid In | Amount received (positive, blank if none) |
| Withdrawn | Amount sent or paid (negative, blank if none) |
| Balance | Running M-PESA balance after the transaction |

**Handling overlapping statements**

M-PESA statements can cover overlapping periods. A January to June statement and a January to December statement both contain January and February. The deduplication step uses all 8 output fields as a composite key, so a transaction that appears in three files produces exactly one row in the output. The same-receipt-different-amount case (a transfer row and its charge row sharing the same receipt number) is preserved correctly because the amount fields differ.

---

## Part 2: Analytics engine in detail

Connect Part 2 to Part 1 by adding two lines to `mpesa_part1_extractor.py`:

At the top of the file:
```python
from mpesa_part2_analytics import run_analytics
```

At the end of `main()`, after `export_to_csv()`:
```python
run_analytics(df, PDF_FOLDER)
```

Part 2 then runs automatically after every Part 1 execution and writes all outputs to `PDF_FOLDER/mpesa_output/`.

**What Part 2 computes**

Income metrics including total, count, mean, median, standard deviation, daily average, weekly average, monthly average, quarterly average, and coefficient of variation.

Expense metrics across the same dimensions.

Cash flow metrics including net flow at daily, weekly, monthly, and overall granularity plus best and worst months by net surplus or deficit.

Savings metrics including savings rate, cumulative savings trajectory, and monthly savings rate.

Balance metrics including latest, highest, lowest, mean, and standard deviation.

Entity extraction that parses the Details column to identify the people, merchants, till numbers, and paybill numbers that appear most frequently as payment destinations or income sources.

Anomaly detection using Z-score analysis where any transaction with an absolute Z-score above 2.5 is flagged as statistically unusual.

Financial health score on a 0 to 100 scale computed from five weighted dimensions: savings rate, income stability, expense control, balance health, and fee ratio.

**Output files**

| File | Contents |
|---|---|
| `mpesa_clean.csv` | Full enriched dataset with all derived columns |
| `mpesa_clean.xlsx` | Excel workbook with six sheets: Transactions, Monthly Summary, By Category, Top Recipients, Top Senders, Anomalies |
| `mpesa_dashboard.html` | Self-contained HTML dashboard with all data embedded as JSON |

---

## Dashboard: features and navigation

The dashboard has nine pages accessible from the left sidebar. A time period slicer on every page (All, Hourly, Daily, Weekly, Monthly, Quarterly, Yearly) filters all charts and metrics to the selected window.

**Overview**
Five donut gauges showing income vs spend ratio, savings rate, transaction success rate, health score, and anomaly rate. Eight KPI cards each showing the headline figure with daily, weekly, and monthly averages beneath. Four charts: monthly income vs spending, spending by category, running balance, and transaction status breakdown.

**Cash Flow**
Best and worst month stat boxes with actual net figures and totals. Monthly net cash flow bars coloured green for surplus and lighter green for deficit. Cumulative savings trajectory. Monthly savings rate line with colour-coded data points.

**Income**
Seven KPI cards with full time-granularity breakdowns. Monthly income trend line. Income by source category donut.

**Spending**
Seven KPI cards with full time-granularity breakdowns. Horizontal bar chart of spending by category sized automatically to the number of categories. Monthly spending trend. Spending distribution histogram.

**Who Sends and Pays**
Two ranked lists: people and merchants you pay most, and people and sources who pay you most. Each shows entity name extracted from the Details column, total KES, transaction count, and average per transaction. A View all button opens a searchable, sortable popup with the complete list.

**Behavior**
Transaction count and KES amount by hour of day and by day of week. All four charts respond to the time period slicer and open transaction popups on click.

**Financial Health**
Composite health score with a colour-coded ring, radar chart across six dimensions, eight financial ratio rows with benchmark guidance, and automated narrative insights derived from the actual data.

**Anomalies**
All transactions with Z-score above 2.5 ranked by score. Clicking the flagged count opens a full sortable popup.

**All Transactions**
Full paginated ledger with 25 rows per page. Filter by month, direction, status, and free text search across details, receipt numbers, and amounts.

**Popup modals**
Every chart element is clickable and opens a popup showing all transactions that match that data point. Popups have a search field and a sort dropdown covering all columns with both ascending and descending options.

---

## Running environment

The Python scripts have been tested on:

- Windows 10 and 11 (Command Prompt, PowerShell, VS Code terminal)
- Ubuntu 22.04 and 24.04
- macOS 13 and 14
- Jupyter Notebook and JupyterLab
- Anaconda environments

The HTML dashboard runs in Chrome 100+, Edge 100+, and Firefox 100+. Safari 16+ works for CSV uploads. PDF extraction in the browser requires an internet connection on first load to fetch pdf.js from CDN (approximately 300 KB, cached permanently after the first load).

---

## Common questions

**My PDF extracts zero rows**

Safaricom PDFs must have a text layer. PDFs that are scanned images of statements have no extractable text. Download your statement fresh from the M-PESA app or MySafaricom portal rather than using a photographed or photocopied version.

**Some transaction details are blank**

This happens when the Details field in the PDF contains only amount tokens or status words, leaving nothing after the parser strips them. This is correct behaviour — the source PDF contains no additional description for those rows.

**The dashboard shows blank charts after upload**

If the CSV columns do not match the expected names exactly the parser will find no rows. The required column names are: `Receipt No.`, `Completion Date`, `Completion Time`, `Details`, `Transaction Status`, `Paid In`, `Withdrawn`, `Balance`. These are the exact names output by Part 1. If you are using a manually prepared CSV check that the column names match character for character including capitalisation and spacing.

**Re-uploading files does not work**

Click the Upload new files button in the sidebar footer to fully reset the dashboard before uploading a new set of files.

**The time slicer shows the same data regardless of selection**

The slicer filters by recency relative to today's date. If your statement data is entirely historical (for example all from 2026 and the slicer window is Daily meaning the last 24 hours), no rows fall in the window and the dashboard automatically falls back to showing all data. This prevents blank charts.

---

## Modifying the extraction logic

The core parsing constants in Part 1 can be adjusted without breaking the pipeline:

`NOISE_PATTERNS` in `mpesa_part1_extractor.py` — add any additional boilerplate lines you want removed from extracted text.

`TYPE_KEYWORDS` in `mpesa_part1_extractor.py` — add keywords to the transaction type classifier for custom categories.

`cat()` function in `mpesa_dashboard.html` — the browser-side categoriser uses the same keyword approach and can be extended in the same way.

---

## License

MIT License. You are free to use, modify, and distribute this code for personal or commercial purposes. Attribution is appreciated but not required.

---

<div align="center">

Built for Kenyan M-PESA users by [Kibet Philip](https://github.com/Apollop24)

</div>
