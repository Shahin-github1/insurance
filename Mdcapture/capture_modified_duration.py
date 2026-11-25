import os
import glob
import re
from collections import defaultdict
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =========================================================================
# CONFIGURATION - Edit paths here
# =========================================================================
FACTSHEET_DIR = r"october\insurance_october_factsheet"
OUTPUT_FILE = r"october\PROCESSING_FILE\ALL_AMC_MODIFIED_DURATION.xlsx"
# =========================================================================

# Normalizes PDF ligatures and character encoding quirks
def clean_ligatures(s: str) -> str:
    if not s:
        return ""
    # Standard ligatures & Aegon font mapping quirks
    s = s.replace('\u019e', 'tf').replace('\u01a0', 'ti')
    s = s.replace('\ufb01', 'fi').replace('\ufb02', 'fl')
    s = s.replace('ƞ', 'tf').replace('Ɵ', 'ti')
    return s

# Multi-variant label matcher for debt portfolio duration
LABEL_REGEX = re.compile(
    r'(?i)\b(?:modified\s+duration|mod\.?\s*duration|m\.?\s*duration|portfolio\s+duration|duration\s+of\s+debt)\b'
)

def extract_fund_name_and_sfin(page):
    """
    Extracts the clean Fund Title and SFIN / ULIF code from a PDF page
    using bounding boxes and text prominence.
    """
    d = page.get_text('dict')
    candidates = []
    sfin_code = ""

    for b in d.get('blocks', []):
        if b.get('type') == 0:  # text block
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    txt = clean_ligatures(s.get('text', '')).strip()
                    if not txt:
                        continue
                    y = s.get('bbox', [0, 0, 0, 0])[1]
                    sz = s.get('size', 0)

                    # Look for SFIN / ULIF codes (e.g. ULIF02807/10/11BSLLIQPLUS109)
                    m_sfin = re.search(r'\b(UL[I|G]F[A-Z0-9\/\-]+)', txt)
                    if m_sfin and not sfin_code:
                        sfin_code = m_sfin.group(1)

                    # Look for prominent titles in top section (y < 140)
                    if y < 140 and len(txt) > 3:
                        if not any(ign in txt.lower() for ign in [
                            'page', 'factsheet', 'report', 'objective', 'benchmark',
                            'nav as on', 'monthly', 'date', 'manager', 'performance',
                            'fund manager', 'aum', 'confidential', 'overview'
                        ]):
                            candidates.append((sz, -y, txt))

    fund_name = "Unknown Fund"
    if candidates:
        # Sort by largest font size first, then top-most Y
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        raw_title = candidates[0][2]
        # Clean title boilerplate
        raw_title = re.sub(r'\(?UL[I|G]F.*', '', raw_title)
        raw_title = re.sub(r'\s+as\s+on\s+.*', '', raw_title, flags=re.IGNORECASE)
        fund_name = raw_title.strip(' -:\t\n')

    return fund_name, sfin_code

def extract_metrics_from_page(page):
    """
    Extracts Modified Duration, Average Maturity, and YTM from a PDF page.
    Combines block regex search with spatial geometric alignment.
    """
    blocks = page.get_text('blocks')
    mod_dur = None
    avg_mat = None
    ytm = None

    for b in blocks:
        b_txt = clean_ligatures(b[4]).strip()

        # 1. Direct regex match inside the same block
        m_same = re.search(
            r'(?i)(?:modified\s+duration|mod\.?\s*duration|m\.?\s*duration|portfolio\s+duration)[^\d\n]*[:\-\s\t]+([0-9]+\.[0-9]+)',
            b_txt
        )
        if m_same and mod_dur is None:
            mod_dur = float(m_same.group(1))

        m_mat = re.search(r'(?i)avg\.?\s*maturity[^\d\n]*[:\-\s\t]+([0-9]+\.[0-9]+)', b_txt)
        if m_mat and avg_mat is None:
            avg_mat = float(m_mat.group(1))

        m_ytm = re.search(r'(?i)(?:ytm|yield\s+to\s+maturity)[^\d\n]*[:\-\s\t]+([0-9]+\.[0-9]+\s*%)', b_txt)
        if m_ytm and ytm is None:
            ytm = m_ytm.group(1).strip()

        # 2. Spatial alignment match (label in one block, value in adjacent block)
        if LABEL_REGEX.search(b_txt) and mod_dur is None:
            label_y_mid = (b[1] + b[3]) / 2.0
            label_x1 = b[2]
            best_cand = None
            min_dy = 999

            for b2 in blocks:
                b2_txt = clean_ligatures(b2[4]).strip()
                b2_y_mid = (b2[1] + b2[3]) / 2.0
                dy = abs(b2_y_mid - label_y_mid)

                # Same horizontal line (Y difference <= 4 points)
                if dy <= 4.0 and b2[0] >= label_x1 - 15:
                    m = re.search(r'([0-9]+\.[0-9]+)', b2_txt)
                    if m and dy < min_dy:
                        min_dy = dy
                        best_cand = float(m.group(1))
                # Directly underneath (Y within 20 points, X aligned)
                elif 0 < (b2[1] - b[3]) <= 20 and abs(b2[0] - b[0]) <= 50:
                    m = re.search(r'([0-9]+\.[0-9]+)', b2_txt)
                    if m and dy < min_dy:
                        min_dy = dy
                        best_cand = float(m.group(1))

            if best_cand is not None:
                mod_dur = best_cand

    return mod_dur, avg_mat, ytm

def process_all_factsheets(factsheet_dir: str = FACTSHEET_DIR, output_path: str = OUTPUT_FILE):
    print("=" * 75)
    print(" ALL AMC FUND MODIFIED DURATION EXTRACTION ENGINE")
    print("=" * 75)
    print(f" Source Directory : {factsheet_dir}")
    print(f" Target Excel     : {output_path}")
    print("=" * 75)

    pdf_files = glob.glob(os.path.join(factsheet_dir, "**", "*.pdf"), recursive=True)
    if not pdf_files:
        raise FileNotFoundError(f"No PDF factsheets found in {factsheet_dir}")

    print(f"\nFound {len(pdf_files)} PDF factsheet files to process.\n")

    all_records = []
    amc_summary_data = defaultdict(list)

    for pdf_idx, pdf_path in enumerate(pdf_files, 1):
        amc_name = os.path.basename(os.path.dirname(pdf_path)).upper()
        pdf_name = os.path.basename(pdf_path)
        
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            print(f"[{pdf_idx}/{len(pdf_files)}] Error opening {pdf_name}: {e}")
            continue

        extracted_in_pdf = 0

        for page_idx in range(len(doc)):
            page = doc[page_idx]
            raw_txt = page.get_text()
            clean_txt = clean_ligatures(raw_txt)

            # Skip pages without duration keywords
            if not LABEL_REGEX.search(clean_txt):
                continue

            page_num = page_idx + 1
            mod_dur, avg_mat, ytm = extract_metrics_from_page(page)

            if mod_dur is not None:
                fund_name, sfin_code = extract_fund_name_and_sfin(page)
                
                record = {
                    "AMC_NAME": amc_name,
                    "FUND_NAME": fund_name,
                    "SFIN_CODE": sfin_code,
                    "MODIFIED_DURATION_YEARS": mod_dur,
                    "AVERAGE_MATURITY": avg_mat if avg_mat is not None else "",
                    "YTM_PERCENT": ytm if ytm is not None else "",
                    "SOURCE_PDF": pdf_name,
                    "PAGE_NUMBER": page_num
                }
                all_records.append(record)
                amc_summary_data[amc_name].append(mod_dur)
                extracted_in_pdf += 1

        print(f"[{pdf_idx:2d}/{len(pdf_files):2d}] [{amc_name:10s}] {pdf_name[:40]:40s} -> {extracted_in_pdf:3d} funds")

    # Save to Excel
    print(f"\nConsolidating {len(all_records)} captured funds into Excel workbook...")
    wb = openpyxl.Workbook()

    # Sheet 1: Detail Records
    ws1 = wb.active
    ws1.title = "All_Funds_Modified_Duration"

    headers1 = [
        "AMC_NAME",
        "FUND_NAME",
        "SFIN_CODE",
        "MODIFIED_DURATION_YEARS",
        "AVERAGE_MATURITY",
        "YTM_PERCENT",
        "SOURCE_PDF",
        "PAGE_NUMBER"
    ]
    ws1.append(headers1)

    # Style Header
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for rec in all_records:
        ws1.append([
            rec["AMC_NAME"],
            rec["FUND_NAME"],
            rec["SFIN_CODE"],
            rec["MODIFIED_DURATION_YEARS"],
            rec["AVERAGE_MATURITY"],
            rec["YTM_PERCENT"],
            rec["SOURCE_PDF"],
            rec["PAGE_NUMBER"]
        ])

    # Sheet 2: AMC Level Summary
    ws2 = wb.create_sheet(title="AMC_Summary")
    headers2 = [
        "AMC_NAME",
        "TOTAL_FUNDS_CAPTURED",
        "AVG_MODIFIED_DURATION",
        "MIN_MODIFIED_DURATION",
        "MAX_MODIFIED_DURATION"
    ]
    ws2.append(headers2)

    summary_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(1, col_idx)
        cell.fill = summary_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for amc, durs in sorted(amc_summary_data.items()):
        avg_d = sum(durs) / len(durs) if durs else 0.0
        min_d = min(durs) if durs else 0.0
        max_d = max(durs) if durs else 0.0
        ws2.append([
            amc,
            len(durs),
            round(avg_d, 2),
            round(min_d, 2),
            round(max_d, 2)
        ])

    # Auto-adjust column widths for readability
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)

    print("\n" + "=" * 75)
    print(" EXTRACTION SUMMARY REPORT")
    print("=" * 75)
    print(f"Total Modified Duration Records Extracted: {len(all_records)}")
    print(f"Distinct AMCs Captured                   : {len(amc_summary_data)}")
    print("-" * 75)
    for amc, durs in sorted(amc_summary_data.items(), key=lambda x: len(x[1]), reverse=True):
        avg_d = sum(durs) / len(durs) if durs else 0.0
        print(f"  * {amc:15s}: {len(durs):3d} funds  (Avg Duration: {avg_d:4.2f} yrs | Min: {min(durs):.2f} | Max: {max(durs):.2f})")
    print("=" * 75)
    print(f"Saved complete consolidation to:\n -> {output_path}\n")

if __name__ == "__main__":
    process_all_factsheets()
