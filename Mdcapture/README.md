# Modified Duration Capture & Consolidation Engine (October 2026)

This folder contains the complete automated pipeline, scraping utilities, and data files used to capture, verify, and consolidate **October 2026 Modified Duration** metrics across all major Indian Life Insurance AMCs into a single structured Excel report.

---

## 📊 Primary Output

- **`Fund List_MD_October.xlsx`**:
  - **Rows**: Exactly 184 filtered fund rows matching the funds that disclosed Modified Duration in September 2026 (`Sheet2` of `Fund List_MD_sep (1).xlsx`).
  - **Structure**:
    - Columns 1–14: Base fund metadata and historical Modified Duration (January through September).
    - **Column 15 (`Modified Duration Oct`)**: Newly inserted and populated October 2026 Modified Duration.
    - Columns 16–17: `Tag` and `Remarks`.
  - **Sheets**:
    1. `Fund_List_MD_October`: Main consolidated fund list with alternating zebra striping, frozen header pane, and highlighted October data.
    2. `October_Extraction_Summary`: Executive audit sheet detailing extraction counts and data sources per insurer.

---

## 🛠️ Scripts & Modules

| Script / File | Description |
| :--- | :--- |
| **`build_october_md_file.py`** | Master pipeline script that loads the source file, consolidates Modified Duration from all AMC extractions/scrapes, builds formatted sheets, and outputs `Fund List_MD_October.xlsx`. |
| **`scrape_all_icici.py`** | Scrapes product categories and SFIN mappings from the official ICICI Prudential portal (`buy.icicilife.com`), discovering internal `LAfundCode` identifiers. |
| **`fetch_icici_durations.py`** | Scrapes debt portfolio analytics (Modified Duration, YTM, Average Maturity) for all 25 ICICI Prudential funds from official portal endpoints. |
| **`capture_modified_duration.py`** | PyMuPDF (`fitz`) factsheet extraction engine for parsing text-prominent Modified Duration labels and values from PDF disclosures. |
| **`extract_all_amcs.py`** | Multi-AMC PDF extraction utility tailored for Bajaj Allianz, SBI Life, and Tata AIA disclosures. |
| **`extract_indiafirst_all.py`** | High-precision extractor for IndiaFirst Life debt workbooks, reading exact decimal indicators from quantitative debt schedules. |
| **`icici_discovered_funds.json`** | Cache of 156 ICICI Prudential funds mapping SFIN and Fund Names to internal `LAfundCode` values. |
| **`icici_october_md_results.json`** | Cached October 2026 scrape results for all 25 ICICI Prudential funds. |

---

## 🏢 Data Coverage Summary (22 Insurers, 184 Funds)

- **Populated (174 funds / 94.6%)**:
  - **ICICI Prudential (25 funds)**: Scraped directly from `https://buy.icicilife.com/portal/login/funds-portfolio-details.htm`.
  - **HDFC Standard Life (28 funds)**: Extracted from HDFC Individual & Group October factsheets.
  - **LIC (18 funds)**: Extracted from official October ULIP disclosure factsheet.
  - **Bajaj Allianz (13 funds)**: Extracted from `Bajaj.pdf` and `Bajaj-Group.pdf`.
  - **Reliance Nippon (11 funds)**: Extracted from Reliance Nippon Group & Individual factsheets.
  - **TATA AIA (10 funds)**: Extracted from Tata AIA `Assure-Fund` and `Pension-fund` disclosures.
  - **Kotak Mahindra (9 funds)**: Extracted from Kotak Life October factsheet.
  - **Max Life / Axis (7 funds)**: Extracted from Axis/Max Monthly ULIP factsheet.
  - **PNB MetLife (7 funds)**: Extracted from PNB MetLife Gratuity & Ulip monthly workbooks.
  - **SBI Life (7 funds)**: Extracted from SBI Life ULIP & Group newsletters (including Group Short Term Plus funds).
  - **Birla Sun Life (3 Individual funds)**: Extracted from ABSLI Individual & Pension factsheet.
  - **IDBI Federal / Ageas Federal (6 funds)**: Extracted from Ageas Federal Wealthtrends factsheet.
  - **Future Generali (5 funds)**: Extracted from Generali Group & Individual newsletters.
  - **Bharti AXA (4 funds)**: Extracted from Bharti AXA monthly fund workbooks.
  - **Canara HSBC OBC (4 funds)**: Extracted from Canara HSBC monthly newsletter.
  - **Edelweiss Tokio (4 funds)**: Extracted from Edelweiss Tokio October factsheet.
  - **IndiaFirst Life (4 funds)**: Extracted from IndiaFirst quantitative indicators debt schedules.
  - **Aviva Life (3 funds)**: Extracted from Aviva Investor & Group factsheet PDFs.
  - **Shriram Life (3 funds)**: Extracted from Shriram Life factsheet.
  - **DHFL Pramerica (2 funds)**: Extracted from DHFL Pramerica factsheet workbook.
  - **AEGON Religare (1 fund)**: Extracted from Aegon Indepth factsheet.
- **Blank / Unpopulated (10 funds / 5.4%)**:
  - **Star Union Dai-ichi (7 funds)**: Skipped per user instruction.
  - **Birla Sun Life (3 Group funds)**: Not disclosed (Group factsheet not published for October 2026).

---

## 🚀 How to Run

1. **Run the Master Consolidation Script**:
   ```powershell
   python build_october_md_file.py
   ```
2. **Re-scrape ICICI Prudential Live Data (Optional)**:
   ```powershell
   python scrape_all_icici.py
   python fetch_icici_durations.py
   ```
3. The generated workbook will be updated at `Fund List_MD_October.xlsx`.
