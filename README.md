# Automated Monthly Sales Report

> A single-file Python pipeline that generates realistic sales data, produces a fully-formatted Excel workbook with live formulas and native charts, renders a multi-page PDF report, and optionally publishes the result to Google Sheets.

![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)
![pandas](https://img.shields.io/badge/pandas-2.0%2B-150458.svg)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-217346.svg)
![ReportLab](https://img.shields.io/badge/reportlab-4.0%2B-red.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)

---

## Table of Contents

- [Overview](#overview)
- [Key Capabilities](#key-capabilities)
- [Architecture](#architecture)
- [Repository Layout](#repository-layout)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Generated Artifacts](#generated-artifacts)
- [Excel Report Specification](#excel-report-specification)
- [PDF Report Specification](#pdf-report-specification)
- [Google Sheets Integration](#google-sheets-integration)
- [Extending the Dataset](#extending-the-dataset)
- [Execution Flow](#execution-flow)
- [Running in Jupyter](#running-in-jupyter)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

**Automated Monthly Sales Report** is a self-contained reporting pipeline written in Python. It transforms raw sales transactions into presentation-ready deliverables — an Excel workbook, a PDF report, and optionally a Google Sheet — with no manual spreadsheet work and no dependency on external BI platforms.

The project is intended as both a practical automation tool and a reference implementation for developers who need to:

- Generate structured, realistic test data programmatically
- Produce styled Excel workbooks containing live formulas, conditional formatting, and native charts
- Build multi-page PDF reports with styled tables and embedded visualisations
- Publish the same content to Google Sheets for collaborative review

Everything is contained within a single file, `sales_report.py`, which runs unchanged as a script or inside a Jupyter notebook.

---

## Key Capabilities

| Domain | Capability |
|---|---|
| **Data generation** | 6,000 realistic sales orders with seasonality, discounts, cost/margin modelling, multi-region sales reps |
| **Reproducibility** | Seeded random number generator — identical inputs yield identical outputs |
| **Excel output** | Seven worksheets, eight KPI cards, live formulas, conditional formatting, four native Excel charts |
| **PDF output** | Multi-page landscape report with styled tables, embedded matplotlib charts, paginated footer |
| **Google Sheets** | Optional one-line upload that preserves formatting and charts |
| **Portability** | Runs as a CLI script or pasted directly into a Jupyter cell — no environment-specific code |
| **Minimal dependencies** | Core pipeline requires only `pandas` + `openpyxl`; PDF step adds `reportlab` + `matplotlib` |

---

## Architecture

```
┌────────────────────┐
│   Configuration    │  Paths, year, order count, seed
└─────────┬──────────┘
          ▼
┌────────────────────┐
│  Data generation   │  pandas + stdlib random
│  (6,000 × 13)      │  → sales_transactions.csv
└─────────┬──────────┘
          ▼
┌────────────────────┐        ┌────────────────────────┐
│  pandas summary    │        │  Excel builder         │
│  (console preview) │───────▶│  openpyxl → .xlsx      │
└────────────────────┘        └───────────┬────────────┘
                                          ▼
                              ┌────────────────────────┐
                              │  PDF builder           │
                              │  ReportLab → .pdf      │
                              └───────────┬────────────┘
                                          ▼
                              ┌────────────────────────┐
                              │  Google Sheets upload  │  (optional)
                              │  gspread + Drive API   │
                              └────────────────────────┘
```

Each stage is decoupled. The CSV is cached between runs, so regenerating the reports without regenerating data is a single re-execution.

---

## Repository Layout

```
sales_report/
├── sales_report.py                    # Complete project (single file)
├── README.md
└── output/                            # Created automatically on first run
    ├── sales_transactions.csv         # Raw data (6,000 rows)
    ├── Monthly_Sales_Report.xlsx      # Formatted Excel workbook
    ├── Monthly_Sales_Report.pdf       # Multi-page PDF report
    └── _pdf_charts/                   # PNG chart cache (regenerated each run)
        ├── chart_monthly.png
        ├── chart_region.png
        ├── chart_category.png
        └── chart_rep.png
```

Default output location on Windows: `C:\Users\Admin\output\`. This is configurable — see [Configuration](#configuration).

---

## Requirements

- **Python** 3.9 or later (3.10+ recommended; the codebase uses PEP 604 union syntax such as `str | Path`)
- **pip**
- **Operating system**: Windows, macOS, or Linux

### Runtime dependencies

```txt
pandas>=2.0
openpyxl>=3.1
reportlab>=4.0
matplotlib>=3.7
```

### Optional dependencies (Google Sheets)

```txt
gspread>=6.0
google-auth>=2.23
google-api-python-client>=2.100
```

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/yourname/automated-sales-report.git
cd automated-sales-report
```

**2. Install dependencies**

```bash
pip install pandas openpyxl reportlab matplotlib
```

**3. (Optional) Install Google Sheets dependencies**

```bash
pip install gspread google-auth google-api-python-client
```

A virtual environment is recommended:

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Usage

**Run as a script**

```bash
python sales_report.py
```

**Expected console output**

```
Working directory: C:\Users\Admin\output
CSV path         : C:\Users\Admin\output\sales_transactions.csv
Excel path       : C:\Users\Admin\output\Monthly_Sales_Report.xlsx
PDF path         : C:\Users\Admin\output\Monthly_Sales_Report.pdf

No CSV found — generated 6,000 dummy orders -> C:\Users\Admin\output\sales_transactions.csv

--- Monthly preview ---
           Revenue    Profit  Orders  Margin
Month
2025-01  $412,530   $143,220    418   34.7%
2025-02  $436,010   $152,180    441   34.9%
...
2025-12  $727,940   $256,470    748   35.2%

Excel report written -> C:\Users\Admin\output\Monthly_Sales_Report.xlsx
PDF report written   -> C:\Users\Admin\output\Monthly_Sales_Report.pdf
```

Open the resulting `.xlsx` file in Excel or LibreOffice once so that formulas are computed and cached (see [Formula caching](#formula-caching)).

---

## Configuration

All tunable parameters are declared in a single **CONFIG** block at the top of `sales_report.py`:

```python
OUTPUT_DIR = Path(r"C:\Users\Admin\output")   # Destination directory
YEAR       = 2025                             # Order year
N_ORDERS   = 6000                             # Number of transactions
SEED       = 42                               # RNG seed (reproducibility)
```

| Parameter | Description | Default |
|---|---|---|
| `OUTPUT_DIR` | Directory for CSV, XLSX, and PDF outputs | `C:\Users\Admin\output` |
| `YEAR` | Year assigned to the generated orders | `2025` |
| `N_ORDERS` | Total transactions to generate | `6000` |
| `SEED` | Seed controlling the random generator | `42` |

### Cross-platform path

For portability across Windows, macOS, and Linux, use a home-relative path:

```python
OUTPUT_DIR = Path.home() / "sales_report_output"
```

---

## Generated Artifacts

| Artifact | Description | Format |
|---|---|---|
| `sales_transactions.csv` | Raw source data — 13 columns × 6,000 rows | CSV |
| `Monthly_Sales_Report.xlsx` | Full formatted workbook with formulas and charts | Excel (OpenXML) |
| `Monthly_Sales_Report.pdf` | Presentation-ready PDF report, landscape A4 | PDF |
| `_pdf_charts/*.png` | Chart images embedded into the PDF | PNG @ 150 DPI |

---

## Excel Report Specification

### Worksheets

| # | Worksheet | Contents |
|---|---|---|
| 1 | **Dashboard** | Eight KPI cards and four native Excel charts (line, column, pie, bar) |
| 2 | **Monthly Summary** | Twelve months of `SUMIFS`/`COUNTIF` formulas plus month-over-month growth |
| 3 | **Region Summary** | Revenue, cost, and profit per region with data bars |
| 4 | **Product Summary** | Products ranked by revenue with colour-scale profit margins |
| 5 | **Category Summary** | Revenue share by category (source for the pie chart) |
| 6 | **Rep Performance** | `RANK()` scoring and quota attainment with red/amber/green scale |
| 7 | **Raw Data** | 6,000 rows with frozen header, autofilter, formatted dates and currency |

### Implementation highlights

```python
# Live formulas — all values computed by Excel itself
ws.cell(row=r, column=2, value=f"=SUMIFS({r_rev},{r_month},$A{r})")
ws.cell(row=r, column=9, value=f'=IFERROR(B{r}/B{r-1}-1,"")')

# Three-colour scale conditional formatting
ws.conditional_formatting.add(
    f"H{first_m}:H{last_m}",
    ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B",
    ),
)

# Data bars
ws.conditional_formatting.add(
    f"B{first_r}:B{last_r}",
    DataBarRule(
        start_type="num", start_value=0, end_type="max",
        color=BLUE, showValue=True,
    ),
)

# Native Excel chart — fully editable inside Excel
line = LineChart()
line.add_data(
    Reference(ws_m, min_col=2, min_row=hdr, max_row=last_m),
    titles_from_data=True,
)
line.set_categories(
    Reference(ws_m, min_col=1, min_row=first_m, max_row=last_m)
)
ws_d.add_chart(line, "A9")
```

### Dashboard KPI formulas

| KPI | Formula |
|---|---|
| Total Revenue | `='Monthly Summary'!B25` |
| Total Profit | `='Monthly Summary'!D25` |
| Profit Margin | `='Monthly Summary'!H25` |
| Total Orders | `='Monthly Summary'!F25` |
| Top Region | `=INDEX(...MATCH(MAX(...)...))` |
| Top Product | `=INDEX(...MATCH(MAX(...)...))` |

### Formula caching

`openpyxl` writes formulas **without cached values**. Excel and LibreOffice compute them on first open. If the file is parsed by `pandas` before being opened in a spreadsheet application, formula cells will appear as `NaN`.

**Resolution:** open the workbook once in Excel or LibreOffice. Excel will compute and cache all formula results; subsequent reads by `pandas` will return the cached values.

---

## PDF Report Specification

### Page layout

```
┌──────────────────────────────────────────────────────────────┐
│  Page 1   Cover · KPI grid · Monthly Summary · line chart    │
├──────────────────────────────────────────────────────────────┤
│  Page 2   Region table + bar chart                           │
│           Category table + pie chart                         │
├──────────────────────────────────────────────────────────────┤
│  Page 3   Product table · Rep table · rep bar chart          │
├──────────────────────────────────────────────────────────────┤
│  Page 4+  Raw data preview (first 200 rows)                  │
└──────────────────────────────────────────────────────────────┘
```

- **Page size:** landscape A4 to accommodate wide tables
- **Footer:** report title (left) and page number (right), drawn via a canvas callback
- **Table headers:** repeat across page breaks using `repeatRows=1`
- **Alternating row backgrounds:** via ReportLab's built-in `ROWBACKGROUNDS`

### Chart rendering

Charts are produced with matplotlib using a headless backend. The backend **must** be set before importing `pyplot`:

```python
import matplotlib
matplotlib.use("Agg")     # Headless — no GUI window required
import matplotlib.pyplot as plt
```

| Chart | Description |
|---|---|
| `chart_monthly.png` | Dual-axis line: revenue and profit by month |
| `chart_region.png` | Vertical bar: revenue by region |
| `chart_category.png` | Pie: revenue share by category |
| `chart_rep.png` | Horizontal bar: revenue per sales representative |

---

## Google Sheets Integration

The Google Sheets step is optional and disabled automatically when no credentials are present.

### 1. Create a service account

1. Navigate to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create or select a project
3. Under **APIs & Services → Library**, enable:
   - **Google Sheets API**
   - **Google Drive API**
4. Under **IAM & Admin → Service Accounts**, create a service account
5. Generate a **JSON key** for the account and download it

### 2. Place the key in `OUTPUT_DIR`

```
C:\Users\Admin\output\service_account.json
```

### 3. Re-run the pipeline

```bash
python sales_report.py
```

If `service_account.json` exists, the pipeline uploads the `.xlsx` as a **native Google Sheet**, preserving formatting and charts, and prints the shareable URL:

```
Google Sheet created -> https://docs.google.com/spreadsheets/d/1AbC.../edit
```

To grant a user access, extend the call in `main()`:

```python
link = upload_workbook(
    path,
    title="Monthly Sales Report",
    credentials=CREDENTIALS_PATH,
    share_with=["you@example.com"],
)
```

---

## Extending the Dataset

### Adding a product

```python
PRODUCTS["Tablet Pro 11"] = ("Hardware", 850.00, 520.00)
```

Then register a weight inside `generate_sales_pure_pandas()`:

```python
product_weights = {
    ...
    "Tablet Pro 11": 0.07,
}
```

### Adding a region

```python
REGIONS["Middle East"] = ["UAE", "Saudi Arabia", "Qatar"]
REPS["Middle East"]    = ["Ahmed Khan", "Layla Hassan", "Omar Farouk"]
```

Then update `region_weights` accordingly.

### Adjusting seasonality

```python
# Default — quiet summer, large Q4
month_weights = [0.85, 0.90, 1.00, 0.95, 1.05, 1.10,
                 0.95, 0.90, 1.15, 1.20, 1.35, 1.50]

# Flat — every month equal
month_weights = [1.0] * 12
```

### Adjusting discounts

```python
discount = round(random.uniform(0.0, 0.20), 2)   # Default: 0–20%
discount = round(random.uniform(0.0, 0.40), 2)   # Aggressive: 0–40%
```

---

## Execution Flow

1. **Configuration** — determines output directory, order year, order count, and RNG seed
2. **Data generation** — `generate_sales_pure_pandas()` produces a `(N_ORDERS × 13)` DataFrame using weighted random sampling for dates, regions, reps, products, and quantities
3. **CSV persistence** — an existing CSV is loaded if present; otherwise a new one is written
4. **pandas preview** — a `groupby("Month").agg(...)` summary is printed to the console
5. **Excel builder** — `build_report()` composes the seven worksheets, writes formulas via f-strings, applies styles and conditional formatting, and inserts native charts
6. **PDF builder** — `build_pdf_report()` renders four matplotlib PNGs and composes a ReportLab document with styled tables and embedded images
7. **Google Sheets upload** *(optional)* — `upload_workbook()` converts the XLSX to a native Google Sheet if credentials are present

---

## Running in Jupyter

The file is **notebook-safe** and can be pasted unchanged into a single Jupyter cell. Two guards make this possible:

```python
# Guard 1 — __file__ is undefined inside a notebook
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

# Guard 2 — __name__ == "__main__" is False inside a cell
try:
    __IPYTHON__          # Defined only within IPython / Jupyter
    main()
except NameError:
    if __name__ == "__main__":
        main()
```

---

## Troubleshooting

### `NameError: name '__file__' is not defined`

An older version of the script was pasted into a notebook. Update to the current version, which includes the `try/except __file__` guard.

### Excel cells appear empty when read programmatically

This is expected behaviour. See [Formula caching](#formula-caching). Open the workbook once in Excel or LibreOffice to populate the cached values.

### `ModuleNotFoundError: No module named 'matplotlib'`

```bash
pip install matplotlib
```

### `ValueError: Invalid character in path` on Windows

Use a raw string literal for the path:

```python
OUTPUT_DIR = Path(r"C:\Users\Admin\output")
```

### PDF chart images appear blurry

Increase the DPI inside `_build_pdf_charts()`:

```python
fig.savefig(p, dpi=200, bbox_inches="tight")   # default is 150
```

### Google Sheets upload fails with HTTP 403

The service-account email (visible inside the JSON key as `client_email`) must have **Editor** access to the target file. Either add the email as an editor manually or pass your own address via `share_with`.

### `reportlab.platypus.doctemplate.LayoutError: Flowable too large`

A table is wider than the available page width. Apply one of the following:

- Reduce the column widths in the affected table
- Switch that section to `portrait(A4)`
- Lower the font size in the table style

---

## Roadmap

- [ ] Command-line arguments (`--year`, `--orders`, `--seed`, `--output`)
- [ ] Pluggable data sources (SQL, REST API) instead of generated data
- [ ] Email delivery of the PDF (SMTP / SendGrid)
- [ ] Interactive HTML dashboard export via Plotly
- [ ] Docker image and scheduled GitHub Actions workflow
- [ ] Unit tests with `pytest` for the data generator and report builders
- [ ] Multi-year comparison worksheets

---

## Contributing

Contributions are welcome. To contribute:

1. Fork the repository
2. Create a feature branch

   ```bash
   git checkout -b feature/my-feature
   ```

3. Commit your changes

   ```bash
   git commit -m "Add feature: ..."
   ```

4. Push the branch

   ```bash
   git push origin feature/my-feature
   ```

5. Open a Pull Request

Please keep the single-file structure unless the project grows substantially. For significant changes, open an issue first to discuss the approach.

---

## License

Released under the **MIT License**.

```
MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Acknowledgements

This project is built on top of the following open-source libraries:

- [**pandas**](https://pandas.pydata.org/) — data manipulation and aggregation
- [**openpyxl**](https://openpyxl.readthedocs.io/) — Excel file generation
- [**ReportLab**](https://www.reportlab.com/) — PDF generation
- [**Matplotlib**](https://matplotlib.org/) — chart rendering
- [**gspread**](https://docs.gspread.org/) — Google Sheets automation

---

## Contact

**Maintainer:** Your Name
**Email:** you@example.com
**Repository:** 

If you find this project useful, consider starring the repository.# Automated Monthly Sales Report

> A single-file Python pipeline that generates realistic sales data, produces a fully-formatted Excel workbook with live formulas and native charts, renders a multi-page PDF report, and optionally publishes the result to Google Sheets.

![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)
![pandas](https://img.shields.io/badge/pandas-2.0%2B-150458.svg)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1%2B-217346.svg)
![ReportLab](https://img.shields.io/badge/reportlab-4.0%2B-red.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)

---

## Table of Contents

- [Overview](#overview)
- [Key Capabilities](#key-capabilities)
- [Architecture](#architecture)
- [Repository Layout](#repository-layout)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Generated Artifacts](#generated-artifacts)
- [Excel Report Specification](#excel-report-specification)
- [PDF Report Specification](#pdf-report-specification)
- [Google Sheets Integration](#google-sheets-integration)
- [Extending the Dataset](#extending-the-dataset)
- [Execution Flow](#execution-flow)
- [Running in Jupyter](#running-in-jupyter)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

**Automated Monthly Sales Report** is a self-contained reporting pipeline written in Python. It transforms raw sales transactions into presentation-ready deliverables — an Excel workbook, a PDF report, and optionally a Google Sheet — with no manual spreadsheet work and no dependency on external BI platforms.

The project is intended as both a practical automation tool and a reference implementation for developers who need to:

- Generate structured, realistic test data programmatically
- Produce styled Excel workbooks containing live formulas, conditional formatting, and native charts
- Build multi-page PDF reports with styled tables and embedded visualisations
- Publish the same content to Google Sheets for collaborative review

Everything is contained within a single file, `sales_report.py`, which runs unchanged as a script or inside a Jupyter notebook.

---

## Key Capabilities

| Domain | Capability |
|---|---|
| **Data generation** | 6,000 realistic sales orders with seasonality, discounts, cost/margin modelling, multi-region sales reps |
| **Reproducibility** | Seeded random number generator — identical inputs yield identical outputs |
| **Excel output** | Seven worksheets, eight KPI cards, live formulas, conditional formatting, four native Excel charts |
| **PDF output** | Multi-page landscape report with styled tables, embedded matplotlib charts, paginated footer |
| **Google Sheets** | Optional one-line upload that preserves formatting and charts |
| **Portability** | Runs as a CLI script or pasted directly into a Jupyter cell — no environment-specific code |
| **Minimal dependencies** | Core pipeline requires only `pandas` + `openpyxl`; PDF step adds `reportlab` + `matplotlib` |

---

## Architecture

```
┌────────────────────┐
│   Configuration    │  Paths, year, order count, seed
└─────────┬──────────┘
          ▼
┌────────────────────┐
│  Data generation   │  pandas + stdlib random
│  (6,000 × 13)      │  → sales_transactions.csv
└─────────┬──────────┘
          ▼
┌────────────────────┐        ┌────────────────────────┐
│  pandas summary    │        │  Excel builder         │
│  (console preview) │───────▶│  openpyxl → .xlsx      │
└────────────────────┘        └───────────┬────────────┘
                                          ▼
                              ┌────────────────────────┐
                              │  PDF builder           │
                              │  ReportLab → .pdf      │
                              └───────────┬────────────┘
                                          ▼
                              ┌────────────────────────┐
                              │  Google Sheets upload  │  (optional)
                              │  gspread + Drive API   │
                              └────────────────────────┘
```

Each stage is decoupled. The CSV is cached between runs, so regenerating the reports without regenerating data is a single re-execution.

---

## Repository Layout

```
sales_report/
├── sales_report.py                    # Complete project (single file)
├── README.md
└── output/                            # Created automatically on first run
    ├── sales_transactions.csv         # Raw data (6,000 rows)
    ├── Monthly_Sales_Report.xlsx      # Formatted Excel workbook
    ├── Monthly_Sales_Report.pdf       # Multi-page PDF report
    └── _pdf_charts/                   # PNG chart cache (regenerated each run)
        ├── chart_monthly.png
        ├── chart_region.png
        ├── chart_category.png
        └── chart_rep.png
```

Default output location on Windows: `C:\Users\Admin\output\`. This is configurable — see [Configuration](#configuration).

---

## Requirements

- **Python** 3.9 or later (3.10+ recommended; the codebase uses PEP 604 union syntax such as `str | Path`)
- **pip**
- **Operating system**: Windows, macOS, or Linux

### Runtime dependencies

```txt
pandas>=2.0
openpyxl>=3.1
reportlab>=4.0
matplotlib>=3.7
```

### Optional dependencies (Google Sheets)

```txt
gspread>=6.0
google-auth>=2.23
google-api-python-client>=2.100
```

---

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/yourname/automated-sales-report.git
cd automated-sales-report
```

**2. Install dependencies**

```bash
pip install pandas openpyxl reportlab matplotlib
```

**3. (Optional) Install Google Sheets dependencies**

```bash
pip install gspread google-auth google-api-python-client
```

A virtual environment is recommended:

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Usage

**Run as a script**

```bash
python sales_report.py
```

**Expected console output**

```
Working directory: C:\Users\Admin\output
CSV path         : C:\Users\Admin\output\sales_transactions.csv
Excel path       : C:\Users\Admin\output\Monthly_Sales_Report.xlsx
PDF path         : C:\Users\Admin\output\Monthly_Sales_Report.pdf

No CSV found — generated 6,000 dummy orders -> C:\Users\Admin\output\sales_transactions.csv

--- Monthly preview ---
           Revenue    Profit  Orders  Margin
Month
2025-01  $412,530   $143,220    418   34.7%
2025-02  $436,010   $152,180    441   34.9%
...
2025-12  $727,940   $256,470    748   35.2%

Excel report written -> C:\Users\Admin\output\Monthly_Sales_Report.xlsx
PDF report written   -> C:\Users\Admin\output\Monthly_Sales_Report.pdf
```

Open the resulting `.xlsx` file in Excel or LibreOffice once so that formulas are computed and cached (see [Formula caching](#formula-caching)).

---

## Configuration

All tunable parameters are declared in a single **CONFIG** block at the top of `sales_report.py`:

```python
OUTPUT_DIR = Path(r"C:\Users\Admin\output")   # Destination directory
YEAR       = 2025                             # Order year
N_ORDERS   = 6000                             # Number of transactions
SEED       = 42                               # RNG seed (reproducibility)
```

| Parameter | Description | Default |
|---|---|---|
| `OUTPUT_DIR` | Directory for CSV, XLSX, and PDF outputs | `C:\Users\Admin\output` |
| `YEAR` | Year assigned to the generated orders | `2025` |
| `N_ORDERS` | Total transactions to generate | `6000` |
| `SEED` | Seed controlling the random generator | `42` |

### Cross-platform path

For portability across Windows, macOS, and Linux, use a home-relative path:

```python
OUTPUT_DIR = Path.home() / "sales_report_output"
```

---

## Generated Artifacts

| Artifact | Description | Format |
|---|---|---|
| `sales_transactions.csv` | Raw source data — 13 columns × 6,000 rows | CSV |
| `Monthly_Sales_Report.xlsx` | Full formatted workbook with formulas and charts | Excel (OpenXML) |
| `Monthly_Sales_Report.pdf` | Presentation-ready PDF report, landscape A4 | PDF |
| `_pdf_charts/*.png` | Chart images embedded into the PDF | PNG @ 150 DPI |

---

## Excel Report Specification

### Worksheets

| # | Worksheet | Contents |
|---|---|---|
| 1 | **Dashboard** | Eight KPI cards and four native Excel charts (line, column, pie, bar) |
| 2 | **Monthly Summary** | Twelve months of `SUMIFS`/`COUNTIF` formulas plus month-over-month growth |
| 3 | **Region Summary** | Revenue, cost, and profit per region with data bars |
| 4 | **Product Summary** | Products ranked by revenue with colour-scale profit margins |
| 5 | **Category Summary** | Revenue share by category (source for the pie chart) |
| 6 | **Rep Performance** | `RANK()` scoring and quota attainment with red/amber/green scale |
| 7 | **Raw Data** | 6,000 rows with frozen header, autofilter, formatted dates and currency |

### Implementation highlights

```python
# Live formulas — all values computed by Excel itself
ws.cell(row=r, column=2, value=f"=SUMIFS({r_rev},{r_month},$A{r})")
ws.cell(row=r, column=9, value=f'=IFERROR(B{r}/B{r-1}-1,"")')

# Three-colour scale conditional formatting
ws.conditional_formatting.add(
    f"H{first_m}:H{last_m}",
    ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B",
    ),
)

# Data bars
ws.conditional_formatting.add(
    f"B{first_r}:B{last_r}",
    DataBarRule(
        start_type="num", start_value=0, end_type="max",
        color=BLUE, showValue=True,
    ),
)

# Native Excel chart — fully editable inside Excel
line = LineChart()
line.add_data(
    Reference(ws_m, min_col=2, min_row=hdr, max_row=last_m),
    titles_from_data=True,
)
line.set_categories(
    Reference(ws_m, min_col=1, min_row=first_m, max_row=last_m)
)
ws_d.add_chart(line, "A9")
```

### Dashboard KPI formulas

| KPI | Formula |
|---|---|
| Total Revenue | `='Monthly Summary'!B25` |
| Total Profit | `='Monthly Summary'!D25` |
| Profit Margin | `='Monthly Summary'!H25` |
| Total Orders | `='Monthly Summary'!F25` |
| Top Region | `=INDEX(...MATCH(MAX(...)...))` |
| Top Product | `=INDEX(...MATCH(MAX(...)...))` |

### Formula caching

`openpyxl` writes formulas **without cached values**. Excel and LibreOffice compute them on first open. If the file is parsed by `pandas` before being opened in a spreadsheet application, formula cells will appear as `NaN`.

**Resolution:** open the workbook once in Excel or LibreOffice. Excel will compute and cache all formula results; subsequent reads by `pandas` will return the cached values.

---

## PDF Report Specification

### Page layout

```
┌──────────────────────────────────────────────────────────────┐
│  Page 1   Cover · KPI grid · Monthly Summary · line chart    │
├──────────────────────────────────────────────────────────────┤
│  Page 2   Region table + bar chart                           │
│           Category table + pie chart                         │
├──────────────────────────────────────────────────────────────┤
│  Page 3   Product table · Rep table · rep bar chart          │
├──────────────────────────────────────────────────────────────┤
│  Page 4+  Raw data preview (first 200 rows)                  │
└──────────────────────────────────────────────────────────────┘
```

- **Page size:** landscape A4 to accommodate wide tables
- **Footer:** report title (left) and page number (right), drawn via a canvas callback
- **Table headers:** repeat across page breaks using `repeatRows=1`
- **Alternating row backgrounds:** via ReportLab's built-in `ROWBACKGROUNDS`

### Chart rendering

Charts are produced with matplotlib using a headless backend. The backend **must** be set before importing `pyplot`:

```python
import matplotlib
matplotlib.use("Agg")     # Headless — no GUI window required
import matplotlib.pyplot as plt
```

| Chart | Description |
|---|---|
| `chart_monthly.png` | Dual-axis line: revenue and profit by month |
| `chart_region.png` | Vertical bar: revenue by region |
| `chart_category.png` | Pie: revenue share by category |
| `chart_rep.png` | Horizontal bar: revenue per sales representative |

---

## Google Sheets Integration

The Google Sheets step is optional and disabled automatically when no credentials are present.

### 1. Create a service account

1. Navigate to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create or select a project
3. Under **APIs & Services → Library**, enable:
   - **Google Sheets API**
   - **Google Drive API**
4. Under **IAM & Admin → Service Accounts**, create a service account
5. Generate a **JSON key** for the account and download it

### 2. Place the key in `OUTPUT_DIR`

```
C:\Users\Admin\output\service_account.json
```

### 3. Re-run the pipeline

```bash
python sales_report.py
```

If `service_account.json` exists, the pipeline uploads the `.xlsx` as a **native Google Sheet**, preserving formatting and charts, and prints the shareable URL:

```
Google Sheet created -> https://docs.google.com/spreadsheets/d/1AbC.../edit
```

To grant a user access, extend the call in `main()`:

```python
link = upload_workbook(
    path,
    title="Monthly Sales Report",
    credentials=CREDENTIALS_PATH,
    share_with=["you@example.com"],
)
```

---

## Extending the Dataset

### Adding a product

```python
PRODUCTS["Tablet Pro 11"] = ("Hardware", 850.00, 520.00)
```

Then register a weight inside `generate_sales_pure_pandas()`:

```python
product_weights = {
    ...
    "Tablet Pro 11": 0.07,
}
```

### Adding a region

```python
REGIONS["Middle East"] = ["UAE", "Saudi Arabia", "Qatar"]
REPS["Middle East"]    = ["Ahmed Khan", "Layla Hassan", "Omar Farouk"]
```

Then update `region_weights` accordingly.

### Adjusting seasonality

```python
# Default — quiet summer, large Q4
month_weights = [0.85, 0.90, 1.00, 0.95, 1.05, 1.10,
                 0.95, 0.90, 1.15, 1.20, 1.35, 1.50]

# Flat — every month equal
month_weights = [1.0] * 12
```

### Adjusting discounts

```python
discount = round(random.uniform(0.0, 0.20), 2)   # Default: 0–20%
discount = round(random.uniform(0.0, 0.40), 2)   # Aggressive: 0–40%
```

---

## Execution Flow

1. **Configuration** — determines output directory, order year, order count, and RNG seed
2. **Data generation** — `generate_sales_pure_pandas()` produces a `(N_ORDERS × 13)` DataFrame using weighted random sampling for dates, regions, reps, products, and quantities
3. **CSV persistence** — an existing CSV is loaded if present; otherwise a new one is written
4. **pandas preview** — a `groupby("Month").agg(...)` summary is printed to the console
5. **Excel builder** — `build_report()` composes the seven worksheets, writes formulas via f-strings, applies styles and conditional formatting, and inserts native charts
6. **PDF builder** — `build_pdf_report()` renders four matplotlib PNGs and composes a ReportLab document with styled tables and embedded images
7. **Google Sheets upload** *(optional)* — `upload_workbook()` converts the XLSX to a native Google Sheet if credentials are present

---

## Running in Jupyter

The file is **notebook-safe** and can be pasted unchanged into a single Jupyter cell. Two guards make this possible:

```python
# Guard 1 — __file__ is undefined inside a notebook
try:
    BASE_DIR = Path(__file__).resolve().parent
except NameError:
    BASE_DIR = Path.cwd()

# Guard 2 — __name__ == "__main__" is False inside a cell
try:
    __IPYTHON__          # Defined only within IPython / Jupyter
    main()
except NameError:
    if __name__ == "__main__":
        main()
```

---

## Troubleshooting

### `NameError: name '__file__' is not defined`

An older version of the script was pasted into a notebook. Update to the current version, which includes the `try/except __file__` guard.

### Excel cells appear empty when read programmatically

This is expected behaviour. See [Formula caching](#formula-caching). Open the workbook once in Excel or LibreOffice to populate the cached values.

### `ModuleNotFoundError: No module named 'matplotlib'`

```bash
pip install matplotlib
```

### `ValueError: Invalid character in path` on Windows

Use a raw string literal for the path:

```python
OUTPUT_DIR = Path(r"C:\Users\Admin\output")
```

### PDF chart images appear blurry

Increase the DPI inside `_build_pdf_charts()`:

```python
fig.savefig(p, dpi=200, bbox_inches="tight")   # default is 150
```

### Google Sheets upload fails with HTTP 403

The service-account email (visible inside the JSON key as `client_email`) must have **Editor** access to the target file. Either add the email as an editor manually or pass your own address via `share_with`.

### `reportlab.platypus.doctemplate.LayoutError: Flowable too large`

A table is wider than the available page width. Apply one of the following:

- Reduce the column widths in the affected table
- Switch that section to `portrait(A4)`
- Lower the font size in the table style

---

## Roadmap

- [ ] Command-line arguments (`--year`, `--orders`, `--seed`, `--output`)
- [ ] Pluggable data sources (SQL, REST API) instead of generated data
- [ ] Email delivery of the PDF (SMTP / SendGrid)
- [ ] Interactive HTML dashboard export via Plotly
- [ ] Docker image and scheduled GitHub Actions workflow
- [ ] Unit tests with `pytest` for the data generator and report builders
- [ ] Multi-year comparison worksheets

---

## Contributing

Contributions are welcome. To contribute:

1. Fork the repository
2. Create a feature branch

   ```bash
   git checkout -b feature/my-feature
   ```

3. Commit your changes

   ```bash
   git commit -m "Add feature: ..."
   ```

4. Push the branch

   ```bash
   git push origin feature/my-feature
   ```

5. Open a Pull Request

Please keep the single-file structure unless the project grows substantially. For significant changes, open an issue first to discuss the approach.

---

## License

Released under the **MIT License**.

```
MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Acknowledgements

This project is built on top of the following open-source libraries:

- [**pandas**](https://pandas.pydata.org/) — data manipulation and aggregation
- [**openpyxl**](https://openpyxl.readthedocs.io/) — Excel file generation
- [**ReportLab**](https://www.reportlab.com/) — PDF generation
- [**Matplotlib**](https://matplotlib.org/) — chart rendering
- [**gspread**](https://docs.gspread.org/) — Google Sheets automation

---

## Contact

**Maintainer:** Asad khan
**Email:** asadkhanjadoon.ece@gmail.com


If you find this project useful, consider starring the repository.
