# Automated Insurance Company Matching & Enrichment Pipeline

An automated, high-precision entity matching and enrichment pipeline designed to replace manual Excel XLOOKUP/VLOOKUP workflows for AMC portfolio disclosures.

---

## 🚀 Key Features

1. **Automated Cleaning (`cleaner.py`)**:
   - Strips coupons (e.g. `7.80%`, `8.25%`).
   - Strips maturity dates (e.g. `28-Jun-27`, `04-08-2036`) and maturity years (`2032`, `2027`).
   - Standardizes Treasury Bills (e.g. `364 DTB 14102027` -> `364 DAYS TREASURY BILL`).
   - Generates the clean **`Absolute name`** column directly from `PDF_COMPANY_NAME`.

2. **Hierarchical Matching Engine (`matcher.py`)**:
   - **Tier 0**: Checks persistent `company_mapping_master.xlsx` for known aliases and corporate name changes (`Approved` or `Review`).
   - **Tier 1**: Exact Match in **Instrument Master** (`MAIN_Instrument_Master_october.xlsx`).
   - **Tier 2**: Exact Match in **Company Master** (`company_master_oct.xlsx`) (Instrument = `"Equity"`, Nature = `"EQ"`).
   - **Tier 3**: Approximation Match in **Instrument Master**:
     - *Case-insensitive* (e.g., `ICICI BANK LTD` -> `ICICI Bank Ltd.`)
     - *Normalized suffix / punctuation / OCR* (e.g., `Tata Stee| Limited` -> `Tata Steel Ltd.`)
     - *High-confidence Fuzzy match* ($\ge 88\%$)
   - **Tier 4**: Approximation Match in **Company Master**:
     - *Case-insensitive*
     - *Normalized suffix / punctuation / OCR*
     - *High-confidence Fuzzy match* ($\ge 88\%$)
   - **Tier 5**: Fallback / Unmatched (automatically staged for review in `company_mapping_master.xlsx`).

3. **Multi-Stage Mapping Master (`company_mapping_master.xlsx`)**:
   - `Approved`: Automatically applied and populated in final outputs.
   - `Review`: High-confidence fuzzy matches; populated in final output with an audit note.
   - `Pending`: New or unrecognized company names held for user review; **not populated** until approved.
   - `Rejected`: Explicitly rejected mappings; **never populated**.

4. **Transparent Audit Trail**:
   - A dedicated **`APPROXIMATION_LEVEL`** column in every output file details the exact rule or score used:
     - `Exact Match (Instrument Master)`
     - `Exact Match (Company Master)`
     - `Converted to lowercase (Instrument Master)`
     - `Normalized suffix/punctuation/OCR (Instrument Master)`
     - `Mapper Approved: Exact Manual Mapping`
     - `Fuzzy token match (92%) (Company Master)`
     - `[Pending] Unmatched - Requires Manual Review`

---

## 📁 Project Structure

```
insurance/
│
├── cleaner.py                     # Name cleaning & regex normalization module
├── matcher.py                     # Multi-tier matching engine & master loader
├── pipeline.py                    # Main pipeline orchestrator CLI
├── company_mapping_master.xlsx    # Persistent 4-stage mapping master (Approved, Review, Pending, Rejected)
│
├── MAIN_Instrument_Master_october.xlsx # Primary Instrument Master
└── october/
    ├── company_master_oct.xlsx         # Secondary Company Master (8,500+ listed companies)
    └── PROCESSING_FILE/
        └── manual/
            ├── STAR_TATA_BANDHAN_PROCESSED.xlsx   # Original reference file
            └── STAR_TATA_BANDHAN_AUTOMATED.xlsx   # Automated pipeline output (100% matched)
```

---

## 🛠️ Usage Instructions

### 1. Process a Single File
```powershell
python pipeline.py --input "october\PROCESSING_FILE\manual\STAR_TATA_BANDHAN_PROCESSED.xlsx" --output "october\PROCESSING_FILE\manual\STAR_TATA_BANDHAN_AUTOMATED.xlsx"
```

### 2. Process the 18 AMC Consolidated File
```powershell
python pipeline.py --input "october\PROCESSING_FILE\18_amc_concat_file.xlsx" --output "october\PROCESSING_FILE\18_amc_concat_file_AUTOMATED.xlsx"
```

### 3. Reviewing & Approving Pending Matches
Open [company_mapping_master.xlsx](file:///c:/Users/Shahin%20Chakraborty/Desktop/insurance/company_mapping_master.xlsx):
- Filter by `STATUS == "Pending"`.
- Fill in `MAPPED_DATABASE_NAME`, `INSTRUMENT_NAME`, and `NATURE`.
- Change `STATUS` from `Pending` to `Approved`.
- Re-run the script — it will instantly pull in the newly approved mappings!
