# Indian Life Insurance & Pension Portfolio Analytics Suite

A production-grade Python automation suite for Indian Life Insurance & Pension portfolio disclosures. The repository provides two modular, fully decoupled processing engines:

1. **`PensionInsuranceProcessing/`**: Automated entity resolution, name normalization, coupon/maturity extraction, investment value calculation, and corpus percentage quality control ($99\% - 101\%$) for AMC portfolio disclosures.
2. **`Mdcapture/`**: Automated Modified Duration extraction, web scraping (ICICI Prudential official portal), factsheet text parsing (22 life insurers), and consolidation into monthly benchmark reports.

---

## 📂 Repository Architecture

```
insurance/
│
├── requirements.txt                   # Universal dependencies (openpyxl, PyMuPDF)
├── .gitignore                         # Python cache & temp file exclusions
├── README.md                          # Master documentation (this file)
│
├── PensionInsuranceProcessing/        # [MODULE 1] Portfolio Holdings Processing Engine
│   ├── cleaner.py                     # Name cleaning, coupon & maturity extraction regex
│   ├── matcher.py                     # 5-tier hierarchical entity matching engine
│   ├── pipeline.py                    # Main portfolio processing pipeline CLI
│   ├── company_mapping_master.xlsx    # Persistent 4-stage mapping registry (Approved/Review/Pending/Rejected)
│   ├── MAIN_Instrument_Master_october.xlsx # Primary Instrument Master
│   ├── company_master_oct.xlsx        # Listed Company Master (8,500+ records)
│   ├── Nature Mapping 1.xlsx          # Asset class / nature classification table
│   ├── Insurance_Data_Functions_QC.txt# Excel QC formula reference
│   └── README.md                      # Module-specific documentation & guide
│
├── Mdcapture/                         # [MODULE 2] Modified Duration Extraction & Scraping Engine
│   ├── capture_md_pipeline.py         # Universal multi-month pipeline (zero hardcoded rows/cols)
│   ├── build_october_md_file.py       # Master October consolidation script
│   ├── scrape_all_icici.py            # ICICI Prudential web scraper (SFIN & category discovery)
│   ├── fetch_icici_durations.py       # ICICI Prudential debt analytics extractor
│   ├── capture_modified_duration.py   # PDF factsheet regex extractor (PyMuPDF)
│   ├── extract_all_amcs.py            # Targeted extractors for Bajaj, SBI Life, Tata AIA
│   ├── extract_indiafirst_all.py      # IndiaFirst Life debt workbook extractor
│   ├── icici_discovered_funds.json    # Cached ICICI fund metadata (156 funds)
│   ├── icici_october_md_results.json  # Cached ICICI October 2026 debt analytics
│   ├── Fund List_MD_October.xlsx      # Master consolidated output workbook (184 funds)
│   ├── Fund List_MD_sep (1).xlsx      # Base September benchmark file
│   └── README.md                      # Module-specific documentation & guide
│
├── Factsheet/                         # Disclosed insurer factsheet PDFs & source workbooks
└── october/                           # Raw input files & monthly processing archives
```

---

## 💻 Environment Setup (Any Device / OS)

This repository is designed to run seamlessly on **Windows, macOS, or Linux** without external database servers or complex environments.

### 1. Prerequisites
- **Python 3.9 or higher** (check via `python --version` or `python3 --version`).
- **Git** (check via `git --version`).

### 2. Clone the Repository
```bash
git clone https://github.com/Shahin-github1/insurance.git
cd insurance
```

### 3. Create and Activate a Virtual Environment

- **On Windows (PowerShell)**:
  ```powershell
  python -m venv venv
  .\venv\Scripts\Activate.ps1
  ```
- **On Windows (Command Prompt)**:
  ```cmd
  python -m venv venv
  .\venv\Scripts\activate.bat
  ```
- **On macOS / Linux (Bash/Zsh)**:
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 🚀 Module 1: Pension Insurance Processing

### What It Does
Replaces manual Excel `XLOOKUP`/`VLOOKUP` workflows for portfolio holdings by:
- Normalizing raw holding names (`Absolute name`).
- Parsing numerical coupon rates (`COUPON`).
- Extracting maturity dates & years (`MATURITY_DATE`, `MATURITY_YEAR`), correcting any legacy two-digit years from the 1900s into the 2000s (`1963` -> `2063`, `1934` -> `2034`).
- Computing investment value: $\text{INVESTMENT\_VALUE} = \frac{\text{CORPUS\_PERCENT} \times \text{FUNDSIZE\_IN\_Crs}}{100}$.
- Validating fund corpus totals against strict QC bounds ($99.0\% - 101.0\%$).
- Resolving entities through a 5-tier matching hierarchy against master databases and a 4-stage persistent mapping registry (`company_mapping_master.xlsx`).

### How to Run

1. **Run Default Pipeline (Consolidated 18 AMC File)**:
   ```bash
   cd PensionInsuranceProcessing
   python pipeline.py
   ```
   *Or from repository root:*
   ```bash
   python PensionInsuranceProcessing/pipeline.py
   ```

2. **Run on a Custom File**:
   ```bash
   python PensionInsuranceProcessing/pipeline.py --input "path/to/portfolio.xlsx" --output "path/to/output.xlsx"
   ```

3. **Output Structure**:
   - `Sheet1`: All portfolio rows enriched with cleaned names, coupons, maturities, database names, instruments, natures, and an explicit `APPROXIMATION_LEVEL` audit note.
   - `Sheet2`: Fund-by-fund corpus percentage QC table flagging any fund whose total holdings fall outside $99.0\% - 101.0\%$.

4. **Mapping Master Governance (`company_mapping_master.xlsx`)**:
   - Unrecognized holdings are automatically staged as `Pending`.
   - Open `company_mapping_master.xlsx`, review `Pending` rows, enter official mappings, and change status to `Approved`.
   - Re-running `pipeline.py` immediately promotes approved entries into subsequent reports.

For detailed documentation, see [PensionInsuranceProcessing/README.md](file:///PensionInsuranceProcessing/README.md).

---

## 📊 Module 2: Modified Duration (MD) Capture

### What It Does
Consolidates October 2026 Modified Duration metrics across **22 Indian Life Insurers (184 benchmark funds)** that disclosed duration in the prior month:
- **Web Scraping**: Automates fetching from the official ICICI Prudential portal (`buy.icicilife.com`) with browser-grade headers.
- **Factsheet PDF Parsing**: Parses quantitative debt metrics across 21 insurers from PDF and Excel factsheets using `PyMuPDF` (`fitz`).
- **Master Workbook Generation**: Inserts and formats the `Modified Duration Oct` column beside historical monthly disclosures (January–September), adds zebra striping, and generates an executive summary.

### How to Run

1. **Universal Pipeline (Dynamic Month & Zero Hardcoding)**:
   ```bash
   cd Mdcapture
   # Run for October:
   python capture_md_pipeline.py --prev-month September --target-month October

   # Roll forward to November:
   python capture_md_pipeline.py --input "Fund List_MD_October.xlsx" --prev-month October --target-month November
   ```

2. **October Dedicated Consolidation Script**:
   ```bash
   cd Mdcapture
   python build_october_md_file.py
   ```

3. **Re-scrape ICICI Prudential Portal (Optional)**:
   ```bash
   python Mdcapture/scrape_all_icici.py
   python Mdcapture/fetch_icici_durations.py
   ```

For detailed documentation, see [Mdcapture/README.md](file:///Mdcapture/README.md).

---

## 🔍 Quality Control & Business Rules Summary

| Feature / Rule | Specification & Acceptance Criteria |
| :--- | :--- |
| **Corpus % Tolerance** | Fund holding sum must fall within **$99.0\% \le \sum \text{CORPUS\_PERCENT} \le 101.0\%$**. Out-of-bounds funds are flagged on `Sheet2`. |
| **Maturity Century Correction** | Two-digit year extractions (e.g., `63`, `34`) are mapped to `2063` and `2034`. No dates from the 1900s are permitted. |
| **Investment Value** | Calculated as `(CORPUS_PERCENT * FUNDSIZE_IN_Crs) / 100`, rounded to 6 decimal places. |
| **Matching Hierarchy** | Tier 0: Mapper (`Approved`/`Review`) $\rightarrow$ Tier 1: Instrument Master Exact $\rightarrow$ Tier 2: Company Master Exact (Equity/EQ) $\rightarrow$ Tier 3: Instrument Master Approx $\rightarrow$ Tier 4: Company Master Approx $\rightarrow$ Tier 5: `Pending` Staging. |
| **Mapping Governance** | `Approved` & `Review` populate final output; `Pending` and `Rejected` are suppressed until audited. |
| **Modified Duration Population** | Filtered strictly against September disclosures (184 funds). 174 funds populated (94.6% coverage across 22 insurers). |

---

## 🛠️ Monthly Roll-Forward Procedure

To process a new month (e.g., November 2026):

1. **Portfolio Holdings**:
   - Place new extracted AMC holdings into `october/PROCESSING_FILE/` (or create a `november/` folder).
   - Execute `python PensionInsuranceProcessing/pipeline.py --input <new_input.xlsx> --output <new_output.xlsx>`.
   - Inspect `Sheet2` for any out-of-range funds and review `company_mapping_master.xlsx` for newly identified aliases.

2. **Modified Duration**:
   - Save insurer factsheets into `Factsheet/` under their respective insurer folders.
   - Run `python Mdcapture/build_october_md_file.py` (or update target column index for November).
   - Validate results against the generated summary sheet.
