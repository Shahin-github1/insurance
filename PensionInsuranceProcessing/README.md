# Pension & Insurance Portfolio Holdings Processing & Enrichment Pipeline

An automated, high-precision entity resolution and quantitative enrichment engine designed to process insurance and pension fund portfolio disclosures. This pipeline replaces manual Excel `XLOOKUP`/`VLOOKUP` workflows with an intelligent multi-tiered matching hierarchy, regex-driven data extraction, and quality control (QC) validation.

---

## 📌 Executive Summary & Key Capabilities

1. **Entity Normalization & Cleaning (`cleaner.py`)**:
   - Strips coupon rates (e.g., `7.80%`, `8.25%`), maturity dates (e.g., `28-Jun-27`, `04-08-2036`), and years (`2032`, `2027`) from raw company disclosure text.
   - Standardizes Sovereign and Treasury instruments (e.g., `364 DTB 14102027` -> `364 DAYS TREASURY BILL`).
   - Produces the normalized **`Absolute name`** column directly from `PDF_COMPANY_NAME`.

2. **Regex Feature Extraction**:
   - **Coupon Rate**: Extracts numerical coupon rates (e.g., `7.80%` -> `7.8`, `7.38%` -> `7.38`).
   - **Maturity Date & Year**: Extracts dates formatted as `DD-Mon-YY`, `DD-Mon-YYYY`, or `DD-MM-YYYY`.
   - **Century Correction**: Ensures no extracted dates fall into the 1900s. Years like `63` or `34` are automatically mapped to `2063` and `2034` (e.g. `28-Jun-1963` -> `28-Jun-2063`).

3. **Investment Value Calculation**:
   - Computes weighted investment value per holding:
     $$\text{INVESTMENT\_VALUE} = \frac{\text{CORPUS\_PERCENT} \times \text{FUNDSIZE\_IN\_Crs}}{100}$$
   - Rounded to 6 decimal places.

4. **Corpus QC Validation ($99.0\% - 101.0\%$)**:
   - Calculates the sum of `CORPUS_PERCENT` for every fund in the portfolio.
   - Applies strict acceptance bounds between **$99.0\%$ and $101.0\%$** (allowing for decimal rounding variations).
   - Generates an executive audit sheet (`Sheet2`) listing:
     - `PDF_FUND_NAME`
     - `HOLDINGS_COUNT`
     - `SUM_OF_CORPUS_PERCENT`
     - `SUM_OF_INVESTMENT_VALUE`
     - `QC_STATUS (99% - 101%)`: Automatically flagged as `PASS` or `FLAG / OUT OF RANGE`.

5. **5-Tier Hierarchical Matching Engine (`matcher.py`)**:
   - **Tier 0**: Checks persistent `company_mapping_master.xlsx` for known aliases and corporate name changes (`Approved` or `Review`).
   - **Tier 1**: Exact Match in **Instrument Master** (`MAIN_Instrument_Master_october.xlsx`).
   - **Tier 2**: Exact Match in **Company Master** (`company_master_oct.xlsx`) (Instrument = `"Equity"`, Nature = `"EQ"`).
   - **Tier 3**: Approximation Match in **Instrument Master**:
     - *Case-insensitive match*
     - *Normalized suffix / punctuation / OCR* (e.g., `Tata Stee| Limited` -> `Tata Steel Ltd.`)
     - *High-confidence Fuzzy match* ($\ge 88\%$)
   - **Tier 4**: Approximation Match in **Company Master**:
     - *Case-insensitive match*
     - *Normalized suffix / punctuation / OCR*
     - *High-confidence Fuzzy match* ($\ge 88\%$)
   - **Tier 5**: Fallback / Staging in `company_mapping_master.xlsx` as `Pending`.

6. **4-Stage Governance Lifecycle (`company_mapping_master.xlsx`)**:
   - `Approved`: Automatically used and populated in final output.
   - `Review`: High-confidence fuzzy match; populated in final output with an audit note.
   - `Pending`: Unmapped or low-confidence records held for manual verification; **not populated** in final output.
   - `Rejected`: Explicitly invalid or excluded holdings; **never populated**.

7. **Auditability (`APPROXIMATION_LEVEL`)**:
   - Every row is stamped with its exact match method (e.g., `Exact Match (Instrument Master)`, `Converted to lowercase`, `Normalized suffix/punctuation/OCR`, `Fuzzy token match (92%)`, or `[Pending] Unmatched`).

---

## 📁 Directory Structure

```
PensionInsuranceProcessing/
│
├── cleaner.py                     # Name normalization & regex extraction logic
├── matcher.py                     # 5-tier hierarchical matching engine
├── pipeline.py                    # End-to-end processing pipeline & CLI orchestrator
├── README.md                      # This documentation
│
├── company_mapping_master.xlsx    # Persistent 4-stage mapping master (Approved/Review/Pending/Rejected)
├── MAIN_Instrument_Master_october.xlsx # Primary Instrument Master (Debt + Equity)
├── company_master_oct.xlsx        # Listed Company Master (8,500+ listed equities)
├── Nature Mapping 1.xlsx          # Asset class / nature classification table
└── Insurance_Data_Functions_QC.txt# Formula and rule reference sheet
```

---

## 💻 Requirements & Installation

The pipeline requires **Python 3.9+** and standard Excel handling libraries:

```bash
pip install openpyxl
```

No external database servers or complex environments are required.

---

## 🚀 Usage Guide

### 1. Default Run (Consolidated 18 AMC File)
From within this directory:
```bash
python pipeline.py
```
Or from the project root:
```bash
python PensionInsuranceProcessing/pipeline.py
```

### 2. Processing a Custom File
Specify your input and desired output paths using `--input` and `--output`:
```bash
python pipeline.py --input "path/to/portfolio_input.xlsx" --output "path/to/portfolio_enriched.xlsx"
```
Or via positional arguments:
```bash
python pipeline.py "path/to/portfolio_input.xlsx" "path/to/portfolio_enriched.xlsx"
```

### 3. Custom Master Databases
You can override any master file path if using different monthly snapshots:
```bash
python pipeline.py \
  --input "input.xlsx" \
  --output "output.xlsx" \
  --mapper "company_mapping_master.xlsx" \
  --instrument-master "MAIN_Instrument_Master_october.xlsx" \
  --company-master "company_master_oct.xlsx"
```

---

## 🔄 Mapping Governance Workflow

When new unmapped company names appear in source disclosures:

1. **Automatic Staging**:
   - The pipeline logs unmatched records as `Pending` into `company_mapping_master.xlsx`.
   - The final output leaves `PDF_COMPANY_NAME_AS_PER_DATABASE` blank and sets `APPROXIMATION_LEVEL` to `[Pending] Unmatched - Requires Manual Review`.

2. **Manual Review**:
   - Open `company_mapping_master.xlsx`.
   - Filter `STATUS` column for `Pending`.
   - Review `RAW_NAME` and `CLEANED_NAME`. Fill in `MAPPED_DB_NAME`, `INSTRUMENT`, and `NATURE`.
   - Change `STATUS` from `Pending` to:
     - `Approved`: Confirmed official mapping.
     - `Review`: Provisional mapping needing secondary sign-off.
     - `Rejected`: Invalid holding / non-securities line item.

3. **Re-run Pipeline**:
   - Run `python pipeline.py` again.
   - All newly approved entries are automatically resolved and populated in the output.
