import os
import sys
import openpyxl
from collections import defaultdict
from cleaner import clean_pdf_company_name, extract_coupon, extract_maturity
from matcher import CompanyMatcher

# =========================================================================
# CONFIGURATION - Set your default file paths here
# =========================================================================
INPUT_FILE = r"october\PROCESSING_FILE\18_amc_concat_file.xlsx"
OUTPUT_FILE = r"october\PROCESSING_FILE\18_amc_concat_file_AUTOMATED.xlsx"
MAPPER_FILE = r"company_mapping_master.xlsx"
INSTRUMENT_MASTER_FILE = r"MAIN_Instrument_Master_october.xlsx"
COMPANY_MASTER_FILE = r"october\company_master_oct.xlsx"
# =========================================================================

def run_pipeline(input_path: str = INPUT_FILE,
                 output_path: str = OUTPUT_FILE,
                 mapper_path: str = MAPPER_FILE,
                 im_path: str = INSTRUMENT_MASTER_FILE,
                 cm_path: str = COMPANY_MASTER_FILE):
    
    print("=" * 70)
    print(" INSURANCE COMPANY MATCHING & PORTFOLIO ENRICHMENT PIPELINE")
    print("=" * 70)
    print(f" Input File        : {input_path}")
    print(f" Output File       : {output_path}")
    print(f" Mapper Master     : {mapper_path}")
    print(f" Instrument Master : {im_path}")
    print(f" Company Master    : {cm_path}")
    print("=" * 70)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Initialize Matcher
    print("\n[1/4] Loading master databases and mapping registry...")
    matcher = CompanyMatcher(
        instrument_master_path=im_path,
        company_master_path=cm_path,
        mapper_file_path=mapper_path
    )

    # Load Workbook
    print(f"[2/4] Reading input workbook: {input_path}")
    wb = openpyxl.load_workbook(input_path)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active

    headers = [str(cell.value or "").strip() for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    def find_all_cols(target_names):
        """Finds all 1-based column indices matching target names."""
        indices = []
        for idx, h in enumerate(headers):
            for name in target_names:
                if h.strip().lower() == name.strip().lower():
                    indices.append(idx + 1)
                    break
        return indices

    def find_col(target_names):
        cols = find_all_cols(target_names)
        return cols[0] if cols else None

    # Locate key columns
    col_fund_name = find_col(["PDF_FUND_NAME", "FUND_NAME", "FUND NAME"])
    col_raw_name = find_col(["PDF_COMPANY_NAME", "PDF COMPANY NAME", "RAW_COMPANY_NAME"])
    col_abs_name = find_col(["Absolute name", "Absolute name ", "ACTUAL_NAME", "Actual Name To FINDOUT", "Cleaned Name"])
    col_db_name = find_col(["PDF_COMPANY_NAME_AS_PER_DATABASE", "COMPANY_NAME_AS_PER_DATABASE", "DB_NAME"])
    col_inst_db = find_col(["INSTRUMENT_NAME_AS_PER_DATABASE", "INSTRUMENT_AS_PER_DATABASE"])
    col_nat_db = find_col(["NATURE_AS_PER_DATABASE", "NATURE_DATABASE"])
    
    col_corpus_pct = find_col(["CORPUS_PERCENT", "CORPUS PERCENT", "CORPUS_%"])
    col_fund_size = find_col(["FUNDSIZE_IN_Crs", "FUNDSIZE_IN_CRS", "FUND SIZE", "FUND_SIZE"])
    col_inv_val = find_col(["INVESTMENT_VALUE", "INVESTMENT VALUE"])
    
    cols_coupon = find_all_cols(["COUPON", "COUPON_RATE"])
    cols_mat_date = find_all_cols(["MATURITY_DATE", "MATURITY DATE"])
    cols_mat_year = find_all_cols(["MATURITY_YEAR", "MATURITY YEAR"])
    col_approx = find_col(["APPROXIMATION_LEVEL", "MATCH_LEVEL", "MATCH_STATUS"])

    if not col_raw_name:
        raise ValueError("Column 'PDF_COMPANY_NAME' not found in input sheet!")

    # Append any missing essential columns
    max_c = len(headers)
    if not col_abs_name:
        max_c += 1; col_abs_name = max_c; ws.cell(1, col_abs_name, "Absolute name")
    if not col_db_name:
        max_c += 1; col_db_name = max_c; ws.cell(1, col_db_name, "PDF_COMPANY_NAME_AS_PER_DATABASE")
    if not col_inst_db:
        max_c += 1; col_inst_db = max_c; ws.cell(1, col_inst_db, "INSTRUMENT_NAME_AS_PER_DATABASE")
    if not col_nat_db:
        max_c += 1; col_nat_db = max_c; ws.cell(1, col_nat_db, "NATURE_AS_PER_DATABASE")
    if not col_inv_val:
        max_c += 1; col_inv_val = max_c; ws.cell(1, col_inv_val, "INVESTMENT_VALUE")
    if not col_approx:
        max_c += 1; col_approx = max_c; ws.cell(1, col_approx, "APPROXIMATION_LEVEL")

    total_rows = ws.max_row - 1
    print(f"\n[3/4] Processing and enriching {total_rows} rows...")

    match_stats = defaultdict(int)
    fund_corpus_map = defaultdict(float)
    fund_inv_map = defaultdict(float)
    fund_rows_map = defaultdict(int)

    for row_idx in range(2, ws.max_row + 1):
        raw_val = ws.cell(row_idx, col_raw_name).value
        raw_str = str(raw_val or "").strip()
        fund_str = str(ws.cell(row_idx, col_fund_name).value or "").strip() if col_fund_name else "Unknown Fund"

        # 1. Clean Company Name
        clean_name = clean_pdf_company_name(raw_str)
        ws.cell(row_idx, col_abs_name, clean_name)

        # 2. Extract Coupon
        c_rate = extract_coupon(raw_str)
        if cols_coupon:
            for c_col in cols_coupon:
                ws.cell(row_idx, c_col, c_rate if c_rate is not None else "")

        # 3. Extract Maturity Date & Year (1900s converted to 2000s)
        mat_date, mat_year = extract_maturity(raw_str)
        if cols_mat_date:
            for d_col in cols_mat_date:
                ws.cell(row_idx, d_col, mat_date if mat_date is not None else "")
        if cols_mat_year:
            for y_col in cols_mat_year:
                ws.cell(row_idx, y_col, mat_year if mat_year is not None else "")

        # 4. Calculate Investment Value: (CORPUS_PERCENT * FUNDSIZE_IN_Crs) / 100
        corpus_val = 0.0
        if col_corpus_pct:
            try:
                c_cell = ws.cell(row_idx, col_corpus_pct).value
                corpus_val = float(c_cell) if c_cell is not None else 0.0
            except (ValueError, TypeError):
                corpus_val = 0.0

        fund_size_val = 0.0
        if col_fund_size:
            try:
                f_cell = ws.cell(row_idx, col_fund_size).value
                fund_size_val = float(f_cell) if f_cell is not None else 0.0
            except (ValueError, TypeError):
                fund_size_val = 0.0

        inv_val = (corpus_val * fund_size_val) / 100.0 if (corpus_val and fund_size_val) else 0.0
        if col_inv_val:
            ws.cell(row_idx, col_inv_val, round(inv_val, 6) if inv_val else "")

        # Track Fund Totals for Corpus QC
        fund_corpus_map[fund_str] += corpus_val
        fund_inv_map[fund_str] += inv_val
        fund_rows_map[fund_str] += 1

        # 5. Hierarchical Company & Instrument Lookup
        res = matcher.match(raw_name=raw_str, clean_name=clean_name)
        status = res["status"]
        approx_level = res["approximation_level"]

        # Rule: Approved and Review populate final output. Pending/Rejected are held back.
        if status in ["Approved", "Review"]:
            ws.cell(row_idx, col_db_name, res["db_name"])
            ws.cell(row_idx, col_inst_db, res["instrument"])
            ws.cell(row_idx, col_nat_db, res["nature"])
            ws.cell(row_idx, col_approx, approx_level)
            match_stats[status] += 1
        else:
            ws.cell(row_idx, col_db_name, "")
            ws.cell(row_idx, col_inst_db, "")
            ws.cell(row_idx, col_nat_db, "")
            ws.cell(row_idx, col_approx, f"[{status}] {approx_level}")
            match_stats[status] += 1

    # [4/4] Generate / Update Sheet2: Fund Corpus Summary & QC Check
    print("\n[4/4] Generating Fund Corpus Summary & QC Check in Sheet2...")
    sheet2_name = "Sheet2"
    if sheet2_name in wb.sheetnames:
        ws_summary = wb[sheet2_name]
        ws_summary.delete_rows(1, ws_summary.max_row + 1)
    else:
        ws_summary = wb.create_sheet(title=sheet2_name)

    summary_headers = [
        "PDF_FUND_NAME",
        "HOLDINGS_COUNT",
        "SUM_OF_CORPUS_PERCENT",
        "SUM_OF_INVESTMENT_VALUE",
        "QC_STATUS (99% - 101%)"
    ]
    ws_summary.append(summary_headers)

    pass_funds = 0
    flag_funds = 0
    flagged_list = []

    for fund, total_pct in sorted(fund_corpus_map.items()):
        total_inv = fund_inv_map[fund]
        row_cnt = fund_rows_map[fund]
        
        # User Rule: Sum of CORPUS_PERCENT must be between 99.0% and 101.0%
        if 99.0 <= total_pct <= 101.0:
            qc_status = "PASS (Within 99% - 101%)"
            pass_funds += 1
        else:
            qc_status = f"FLAG / OUT OF RANGE ({total_pct:.2f}%)"
            flag_funds += 1
            flagged_list.append((fund, total_pct, row_cnt))

        ws_summary.append([
            fund,
            row_cnt,
            round(total_pct, 4),
            round(total_inv, 4),
            qc_status
        ])

    # Save output workbook
    wb.save(output_path)
    print(f"\nSuccessfully saved enriched workbook to:\n -> {output_path}")

    # Append any new pending records to mapper file
    matcher.save_pending_to_mapper()

    # Print Final Summary Report
    print("\n" + "=" * 70)
    print(" EXECUTION & QUALITY CONTROL SUMMARY")
    print("=" * 70)
    print(f"Total Rows Processed : {total_rows}")
    print(f"  Approved Mappings  : {match_stats['Approved']} ({match_stats['Approved']/total_rows*100:.2f}%)")
    print(f"  Review Mappings    : {match_stats['Review']} ({match_stats['Review']/total_rows*100:.2f}%)")
    print(f"  Pending Review     : {match_stats['Pending']}")
    print("-" * 70)
    print(f"Total Funds Checked  : {len(fund_corpus_map)}")
    print(f"  Corpus QC PASS     : {pass_funds} funds (Within 99.0% - 101.0%)")
    print(f"  Corpus QC FLAGGED  : {flag_funds} funds (Outside 99.0% - 101.0%)")
    print("=" * 70)

    if flagged_list:
        print("\nFlagged Funds for Manual Inspection (Outside 99% - 101% range):")
        for fund_name, tot_pct, r_cnt in flagged_list[:15]:
            print(f"  * {fund_name}: {tot_pct:.2f}% (Holdings: {r_cnt})")
        if len(flagged_list) > 15:
            print(f"  ... and {len(flagged_list) - 15} more (see Sheet2 in output file)")
    print("\nDone!\n")

if __name__ == "__main__":
    # Allows direct execution without CLI arguments, but also accepts arguments if provided
    in_file = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    out_file = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FILE
    run_pipeline(input_path=in_file, output_path=out_file)
