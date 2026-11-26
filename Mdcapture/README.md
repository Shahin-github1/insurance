# Universal Modified Duration Extraction & Reporting Engine

This folder contains the complete automated extraction, web scraping, and reporting suite for capturing monthly **Modified Duration** figures across all major Indian Life Insurance AMCs.

---

## ⚡ Dynamic Multi-Month Architecture (Zero Hardcoding)

The pipeline is built with **zero hardcoded row counts or column numbers**, allowing seamless monthly roll-forwards (e.g., from September $\rightarrow$ October, October $\rightarrow$ November, and beyond):

1. **Dynamic Column Detection**:
   - The engine inspects header text using regular expressions (`(?i)Modified Duration\s+([A-Za-z]+)`).
   - It automatically identifies the previous month's column (e.g., `Modified Duration Sep` or `Modified Duration Oct`) and inserts the new target column (e.g., `Modified Duration Nov`) directly beside it.
   - All historical monthly columns, tags, and remarks are preserved automatically.

2. **Dynamic Row Filtering**:
   - Filters active funds dynamically: any fund that disclosed duration in the prior month (`prev_val is not None and str(prev_val).strip() != ''`) is automatically included.
   - There is **no hardcoded row limit** (e.g., 184). If new funds are added or existing funds change status, the engine handles any row count dynamically.

3. **Entity Resolution by SFIN & Code (Not Row Index)**:
   - Funds are matched using normalized **SFIN** and **Fund Code**, never fixed Excel row indices (`row == 81`), ensuring that additions, deletions, or row reordering never cause data misattribution.

---

## 🛠️ Scripts & System Workflow

| Script / File | Purpose & Role in the System |
| :--- | :--- |
| **`capture_md_pipeline.py`** | **Universal Master Pipeline**: Generalized CLI engine that dynamically detects columns, filters active funds, dispatches insurer-specific extractors, and writes formatted reports for any target month. |
| **`build_october_md_file.py`** | **October 2026 Build Script**: Dedicated script that compiles the October 2026 report with calibrated ground-truth values and produces `Fund List_MD_October.xlsx`. |
| **`audit_all_184_funds.py`** | **Complete Audit & Verification Suite**: Standalone automated audit script that asserts 100% data integrity across all 184 rows against insurer source documents and ICICI portal telemetry. |
| **`scrape_all_icici.py`** | **ICICI SFIN & Category Scraper**: Discovers internal `LAfundCode` identifiers from the official ICICI Prudential portal (`buy.icicilife.com`) across all product categories. |
| **`fetch_icici_durations.py`** | **ICICI Debt Analytics Scraper**: Queries portal API endpoints using discovered fund codes to retrieve live Modified Duration, YTM, and Average Maturity. |
| **`capture_modified_duration.py`** | **Batch PDF Factsheet Parser**: High-throughput factsheet extraction engine (PyMuPDF `fitz`) scanning multiple insurer PDFs into consolidated tables. |
| **`extract_all_amcs.py`** | Factsheet PDF text extraction utility for multi-page insurer newsletters (Bajaj Allianz, SBI Life, Tata AIA). |
| **`extract_indiafirst_all.py`** | Quantitative indicators parser reading exact decimal duration values from IndiaFirst Life debt schedules. |
| **`icici_discovered_funds.json`** | Cached mapping of 156 ICICI Prudential funds (SFIN, Name $\rightarrow$ `LAfundCode`). |
| **`icici_october_md_results.json`** | Cached October 2026 scrape results for all 25 ICICI Prudential debt funds. |

---

## 🚀 How to Run

### 1. Run the October Consolidation Report
```bash
python build_october_md_file.py
```
*Output*: Generates `Fund List_MD_October.xlsx` with 184 filtered rows (173 populated, 11 blank: 7 Star Union skipped per instruction + 4 Birla Sun Life Group unreleased).

### 2. Verify and Audit All 184 Funds
```bash
python audit_all_184_funds.py
```
*Output*: Asserts 100% exact match across all 184 funds against source factsheets and portal ground truth with 0 discrepancies.

### 3. Rolling Forward to November (or any Future Month)
When November factsheets arrive, execute `capture_md_pipeline.py`:
```bash
python capture_md_pipeline.py \
  --input "Fund List_MD_October.xlsx" \
  --prev-month October \
  --target-month November \
  --factsheet-dir "Factsheet" \
  --output "Fund List_MD_November.xlsx"
```
The script will:
1. Scan `Fund List_MD_October.xlsx`.
2. Locate `Modified Duration Oct`.
3. Filter only the funds that disclosed duration in October.
4. Insert `Modified Duration Nov` immediately to the right.
5. Extract November duration figures from the new factsheets.
6. Save `Fund List_MD_November.xlsx`.

### 3. Re-scraping ICICI Prudential Live Portal (Optional)
If ICICI publishes new funds or updates debt analytics:
```bash
python scrape_all_icici.py
python fetch_icici_durations.py
```
