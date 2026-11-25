import os
import sys
import re
import json
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import fitz

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

def _find_file(filename, fallback_rel=None):
    if os.path.exists(filename):
        return filename
    in_script = os.path.join(SCRIPT_DIR, filename)
    if os.path.exists(in_script):
        return in_script
    in_root = os.path.join(PROJECT_ROOT, filename)
    if os.path.exists(in_root):
        return in_root
    if fallback_rel:
        in_fallback = os.path.join(PROJECT_ROOT, fallback_rel)
        if os.path.exists(in_fallback):
            return in_fallback
    return filename

def clean_sfin(sfin: str) -> str:
    if not sfin:
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(sfin)).upper()

def clean_ligatures(s: str) -> str:
    if not s:
        return ""
    s = s.replace('\u019e', 'tf').replace('\u01a0', 'ti')
    s = s.replace('\ufb01', 'fi').replace('\ufb02', 'fl')
    s = s.replace('ƞ', 'tf').replace('Ɵ', 'ti')
    return s

# =========================================================================
# INSURER-SPECIFIC EXTRACTORS
# =========================================================================

def extract_kotak_md(pdf_path: str, target_sfin: str) -> float:
    if not os.path.exists(pdf_path):
        return None
    c_target = clean_sfin(target_sfin)
    doc = fitz.open(pdf_path)
    for p in doc:
        txt = clean_sfin(p.get_text())
        if c_target in txt:
            words = [w for w in p.get_text('words') if w[0] < 160 and 320 < w[1] < 420]
            words.sort(key=lambda w: (w[1], w[0]))
            t = ' '.join(w[4] for w in words)
            m = re.search(r'Instruments\s*:\s*([0-9]+\.[0-9]+)', t)
            if not m:
                m = re.search(r':\s*([0-9]+\.[0-9]+)', t)
            if m:
                return float(m.group(1))
    return None

def extract_sbi_md(ulip_pdf: str, grp_pdf: str, target_sfin: str) -> float:
    c_target = clean_sfin(target_sfin)
    for path in [ulip_pdf, grp_pdf]:
        if not path or not os.path.exists(path):
            continue
        doc = fitz.open(path)
        for p in doc:
            txt = clean_sfin(p.get_text())
            if c_target in txt:
                words = [w for w in p.get_text('words') if 450 < w[0] < 550 and 150 < w[1] < 220]
                words.sort(key=lambda w: (w[1], w[0]))
                nums = [w[4] for w in words if re.match(r'^[0-9]+\.[0-9]+$', w[4])]
                if nums:
                    return float(nums[-1])
    return None

def extract_shriram_md(pdf_path: str, target_sfin: str) -> float:
    if not os.path.exists(pdf_path):
        return None
    c_target = clean_sfin(target_sfin)
    doc = fitz.open(pdf_path)
    for p in doc:
        txt = clean_sfin(p.get_text())
        if c_target in txt:
            blocks = p.get_text('blocks')
            for b in blocks:
                if 310 <= b[1] <= 350 and b[0] > 60:
                    nums = [l.strip() for l in b[4].split('\n') if re.match(r'^[0-9]+\.[0-9]+$', l.strip())]
                    if len(nums) >= 2:
                        return float(nums[-1] if len(nums) == 2 else nums[2])
                m = re.search(r'([0-9]+\.[0-9]+)\s*\n\s*[0-9]+\.[0-9]+%', b[4])
                if m and 'Grand Total' in b[4]:
                    return float(m.group(1))
    return None

def extract_tata_md(assure_pdf: str, pension_pdf: str, target_sfin: str) -> float:
    c_target = clean_sfin(target_sfin)
    for path in [assure_pdf, pension_pdf]:
        if not path or not os.path.exists(path):
            continue
        doc = fitz.open(path)
        for p in doc:
            txt = clean_sfin(p.get_text())
            if c_target in txt:
                blocks = p.get_text('blocks')
                for b in blocks:
                    if 'modified duration' in b[4].lower():
                        y0 = b[1]
                        for b2 in blocks:
                            if abs(b2[1] - y0) < 6 and 200 < b2[0] < 270:
                                val_str = b2[4].strip()
                                if re.match(r'^[0-9]+\.[0-9]+$', val_str):
                                    return float(val_str)
                        lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                        for l in lines:
                            m = re.search(r'([0-9]+\.[0-9]+)', l)
                            if m:
                                return float(m.group(1))
    return None

def extract_hdfc_md(indiv_pdf: str, group_pdf: str, target_sfin: str) -> float:
    c_target = clean_sfin(target_sfin)
    for path in [indiv_pdf, group_pdf]:
        if not path or not os.path.exists(path):
            continue
        doc = fitz.open(path)
        for pno in range(len(doc)):
            if pno < 3: # Skip Table of Contents / Overview pages
                continue
            p = doc[pno]
            txt = clean_sfin(p.get_text())
            if c_target in txt:
                blocks = p.get_text('blocks')
                for b in blocks:
                    if 'modified duration' in b[4].lower():
                        y0 = b[1]
                        for b2 in blocks:
                            if abs(b2[1] - y0) < 6 and 200 < b2[0] < 350:
                                val_str = b2[4].strip()
                                m = re.search(r'([0-9]+\.[0-9]+)', val_str)
                                if m:
                                    return float(m.group(1))
    return None

def extract_canara_md(pdf_path: str, target_sfin: str) -> float:
    if not os.path.exists(pdf_path):
        return None
    c_target = clean_sfin(target_sfin)
    doc = fitz.open(pdf_path)
    for p in doc:
        txt = clean_sfin(p.get_text())
        if c_target in txt:
            m = re.search(r'modified duration[^:]*:\s*([0-9]+\.[0-9]+)', p.get_text(), re.I)
            if m:
                return float(m.group(1))
    return None

def extract_edelweiss_md(pdf_path: str, target_sfin: str) -> float:
    if not os.path.exists(pdf_path):
        return None
    c_target = clean_sfin(target_sfin)
    doc = fitz.open(pdf_path)
    for pno in range(len(doc)):
        if pno < 2:
            continue
        p = doc[pno]
        txt = clean_sfin(p.get_text())
        if c_target in txt:
            for b in p.get_text('blocks'):
                if 'portfolio duration' in b[4].lower():
                    lines = [l.strip() for l in b[4].split('\n') if re.match(r'^[0-9]+\.[0-9]+$', l.strip())]
                    if lines:
                        return float(lines[0])
    return None

def extract_lic_md(pdf_path: str, target_sfin: str) -> float:
    if not os.path.exists(pdf_path):
        return None
    c_target = clean_sfin(target_sfin)
    doc = fitz.open(pdf_path)
    for p in doc:
        txt = clean_sfin(p.get_text())
        if c_target in txt:
            m = re.search(r'modified\s+duration\s*[:\s]*([0-9]+\.[0-9]+)', p.get_text(), re.I)
            if m:
                return float(m.group(1))
    return None

# =========================================================================
# MASTER DYNAMIC PIPELINE FUNCTION
# =========================================================================

def run_md_pipeline(
    input_file: str = None,
    output_file: str = None,
    target_month: str = "October",
    prev_month: str = None,
    factsheet_dir: str = None
):
    print("=" * 75)
    print(" UNIVERSAL MODIFIED DURATION EXTRACTION & REPORTING ENGINE")
    print("=" * 75)
    
    if not input_file:
        input_file = _find_file("Fund List_MD_sep (1).xlsx")
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Base input file not found: {input_file}")
        
    if not factsheet_dir:
        factsheet_dir = _find_file("Factsheet")
        
    print(f" Input File        : {input_file}")
    print(f" Factsheet Folder  : {factsheet_dir}")
    print(f" Target Month      : {target_month}")
    
    wb_src = openpyxl.load_workbook(input_file, data_only=True)
    ws_src = wb_src['Sheet2'] if 'Sheet2' in wb_src.sheetnames else wb_src.active
    
    # 1. DYNAMIC COLUMN DETECTION (No hardcoded column numbers)
    headers = [str(ws_src.cell(1, c).value or '').strip() for c in range(1, ws_src.max_column + 1)]
    
    prev_col = None
    if prev_month:
        for idx, h in enumerate(headers, start=1):
            if re.search(rf'modified\s+duration\s+{prev_month}', h, re.I):
                prev_col = idx
                break
    else:
        # Detect the rightmost "Modified Duration <Month>" column dynamically
        for idx in range(len(headers), 0, -1):
            h = headers[idx - 1]
            if re.search(r'modified\s+duration\s+([A-Za-z]+)', h, re.I):
                prev_col = idx
                prev_month = re.search(r'modified\s+duration\s+([A-Za-z]+)', h, re.I).group(1)
                break
                
    if not prev_col:
        raise ValueError(f"Could not locate previous month's Modified Duration column in headers: {headers}")
        
    print(f" Detected Prev Month Col : Column {prev_col} ('{headers[prev_col - 1]}')")
    
    # Identify key column indices dynamically
    def find_col_idx(candidates):
        for idx, h in enumerate(headers, start=1):
            if any(cand.lower() == h.lower() for cand in candidates):
                return idx
        return None

    col_insurer = find_col_idx(['Insurer Name', 'Insurer', 'Company']) or 1
    col_fund_name = find_col_idx(['Fund Name', 'Fund']) or 2
    col_sfin = find_col_idx(['SFIN', 'SFIN Code']) or 5
    col_code = find_col_idx(['Code']) or 4
    col_tag = find_col_idx(['Tag'])
    col_remarks = find_col_idx(['Remarks'])
    
    # 2. DYNAMIC ROW FILTERING (No hardcoded 184 count)
    target_rows = []
    for r in range(2, ws_src.max_row + 1):
        prev_val = ws_src.cell(r, prev_col).value
        # Keep any fund that disclosed duration in the prior month
        if prev_val is not None and str(prev_val).strip() != '':
            row_dict = {
                'src_row': r,
                'insurer': str(ws_src.cell(r, col_insurer).value or '').strip(),
                'fund_name': str(ws_src.cell(r, col_fund_name).value or '').strip(),
                'sfin': str(ws_src.cell(r, col_sfin).value or '').strip(),
                'code': str(ws_src.cell(r, col_code).value or '').strip(),
                'prev_md': prev_val,
                'cells': [ws_src.cell(r, c).value for c in range(1, ws_src.max_column + 1)]
            }
            row_dict['clean_sfin'] = clean_sfin(row_dict['sfin'])
            target_rows.append(row_dict)
            
    print(f" Dynamic Active Funds Filtered : {len(target_rows)} funds had data in {prev_month}")
    
    # Load ICICI Scraped Results if available
    icici_cache_path = _find_file("icici_october_md_results.json")
    icici_by_sfin = {}
    icici_by_code = {}
    if os.path.exists(icici_cache_path):
        with open(icici_cache_path, 'r', encoding='utf-8') as f:
            for item in json.load(f):
                if item.get('sfin'):
                    icici_by_sfin[clean_sfin(item['sfin'])] = item
                if item.get('code'):
                    icici_by_code[str(item['code']).strip()] = item
                    
    # Pre-load ALL_AMC_MODIFIED_DURATION if available
    all_amc_file = _find_file(r"october\PROCESSING_FILE\ALL_AMC_MODIFIED_DURATION.xlsx", r"october\PROCESSING_FILE\ALL_AMC_MODIFIED_DURATION.xlsx")
    all_amc_records = []
    if os.path.exists(all_amc_file):
        wb_md = openpyxl.load_workbook(all_amc_file, data_only=True)
        ws_md = wb_md['All_Funds_Modified_Duration'] if 'All_Funds_Modified_Duration' in wb_md.sheetnames else wb_md.active
        for r in range(2, ws_md.max_row + 1):
            all_amc_records.append({
                'amc': str(ws_md.cell(r, 1).value or '').strip().upper(),
                'fund_name': str(ws_md.cell(r, 2).value or '').strip(),
                'sfin': str(ws_md.cell(r, 3).value or '').strip(),
                'clean_sfin': clean_sfin(ws_md.cell(r, 3).value),
                'md': ws_md.cell(r, 4).value,
                'pdf': ws_md.cell(r, 7).value
            })

    # Factsheet File Paths
    factsheets = {
        'kotak': os.path.join(factsheet_dir, "kotak", "Kotak_Life_Individual_October_2026.pdf"),
        'sbi_ulip': os.path.join(factsheet_dir, "SBI", "SBI LIFE ULIP NEWS LETTER OCT 2026.pdf"),
        'sbi_grp': os.path.join(factsheet_dir, "SBI", "SBI LIFE GROUP NEWS LETTER OCT 2026.pdf"),
        'shriram': os.path.join(factsheet_dir, "shriram", "factsheet-for-the-month-of-oct-26.pdf"),
        'tata_assure': os.path.join(factsheet_dir, "1_mannual_Data", "tata", "Assure-Fund (4).pdf"),
        'tata_pension': os.path.join(factsheet_dir, "1_mannual_Data", "tata", "Pension-fund.pdf"),
        'hdfc_indiv': os.path.join(factsheet_dir, "HDFC", "fund-factsheets-individual.pdf"),
        'hdfc_grp': os.path.join(factsheet_dir, "HDFC", "Fund-Factsheets-Group.pdf"),
        'canara': os.path.join(factsheet_dir, "canara", "monthly-newsletter-october-26.pdf"),
        'edelweiss': os.path.join(factsheet_dir, "edelweiss", "October_2026.pdf"),
        'lic': os.path.join(factsheet_dir, "LIC", "ULIP Disclosure as on 31.10.2026.pdf")
    }

    # 3. EXTRACTION DISPATCHER (By SFIN & Insurer, NOT row index)
    populated_count = 0
    blank_count = 0
    
    print("\n[Processing] Extracting duration figures for active funds...")
    for row in target_rows:
        ins = row['insurer']
        sfin = row['sfin']
        c_sfin = row['clean_sfin']
        name = row['fund_name']
        code = row['code']
        
        val = None
        note = ""
        
        # Insurer 1: ICICI Prudential (Dynamic SFIN / Code lookup)
        if 'ICICI' in ins:
            rec = icici_by_sfin.get(c_sfin) or icici_by_code.get(code)
            if rec:
                val = rec.get('oct_md')
                note = "ICICI Official Web Portal"
                
        # Insurer 2: Star Union (Skip per user rule)
        elif 'Star Union' in ins:
            val = None
            note = "Star Union skipped per instruction"
            
        # Insurer 3: Birla Sun Life Group (Unreleased)
        elif 'Birla' in ins and ('ULGF' in sfin or 'GROUP' in name.upper()):
            val = None
            note = "Birla Sun Life Group factsheet not published"
            
        # Insurer 4: Kotak Mahindra Life (Direct PDF parsing)
        elif 'Kotak' in ins:
            val = extract_kotak_md(factsheets['kotak'], sfin)
            note = "Kotak Life Factsheet PDF"
            
        # Insurer 5: SBI Life (Direct PDF parsing)
        elif 'SBI' in ins:
            val = extract_sbi_md(factsheets['sbi_ulip'], factsheets['sbi_grp'], sfin)
            note = "SBI Life Factsheet PDF"
            
        # Insurer 6: Shriram Life (Direct PDF parsing)
        elif 'Shriram' in ins:
            val = extract_shriram_md(factsheets['shriram'], sfin)
            note = "Shriram Life Factsheet PDF"
            
        # Insurer 7: Tata AIA Life (Direct PDF parsing)
        elif 'TATA' in ins:
            val = extract_tata_md(factsheets['tata_assure'], factsheets['tata_pension'], sfin)
            note = "Tata AIA Factsheet PDF"
            
        # Insurer 8: HDFC Standard Life (Direct PDF parsing)
        elif 'HDFC' in ins:
            val = extract_hdfc_md(factsheets['hdfc_indiv'], factsheets['hdfc_grp'], sfin)
            note = "HDFC Factsheet PDF"
            
        # Insurer 9: Canara HSBC (Direct PDF parsing)
        elif 'Canara' in ins:
            val = extract_canara_md(factsheets['canara'], sfin)
            note = "Canara HSBC Factsheet PDF"
            
        # Insurer 10: Edelweiss Tokio (Direct PDF parsing)
        elif 'Edelweiss' in ins or 'Edelwiess' in ins:
            val = extract_edelweiss_md(factsheets['edelweiss'], sfin)
            note = "Edelweiss Tokio Factsheet PDF"
            
        # Insurer 11: LIC (Direct PDF parsing)
        elif 'Life Insurance Corporation' in ins or 'LIC' in ins:
            val = extract_lic_md(factsheets['lic'], sfin)
            note = "LIC ULIP Factsheet PDF"

        # Insurer 12-22: Fallback to ALL_AMC parsed records or specific files
        if val is None and not note:
            for rec in all_amc_records:
                if rec['clean_sfin'] and (rec['clean_sfin'] == c_sfin or rec['clean_sfin'] in c_sfin or c_sfin in rec['clean_sfin']):
                    val = rec['md']
                    note = f"ALL_AMC ({rec['pdf']})"
                    break

        if val is not None and str(val).strip() != '':
            try:
                row['target_val'] = float(str(val).replace('%', '').strip())
            except:
                row['target_val'] = val
            populated_count += 1
        else:
            row['target_val'] = None
            blank_count += 1
            
        row['match_note'] = note

    print(f" Summary: {populated_count} Populated, {blank_count} Blank")

    # 4. BUILD NEW WORKBOOK WITH DYNAMIC TARGET COLUMN INSERTION
    if not output_file:
        output_file = os.path.join(SCRIPT_DIR, f"Fund List_MD_{target_month}.xlsx")
        
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = f"Fund_List_MD_{target_month}"
    
    # Styles
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_new_col = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_new_col = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    font_body = Font(name="Calibri", size=11)
    font_body_bold = Font(name="Calibri", size=11, bold=True)
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Insert new month column immediately following prev_col
    new_headers = headers[:prev_col] + [f"Modified Duration {target_month[:3]}"] + headers[prev_col:]
    
    for c_idx, h in enumerate(new_headers, start=1):
        cell = ws_out.cell(1, c_idx, h)
        if c_idx == prev_col + 1:
            cell.font = font_new_col
            cell.fill = fill_new_col
        else:
            cell.font = font_header
            cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
    ws_out.row_dimensions[1].height = 28
    
    # Populate rows
    for r_idx, row in enumerate(target_rows, start=2):
        orig_cells = row['cells']
        new_row_vals = orig_cells[:prev_col] + [row['target_val']] + orig_cells[prev_col:]
        
        is_even = (r_idx % 2 == 0)
        for c_idx, val in enumerate(new_row_vals, start=1):
            cell = ws_out.cell(r_idx, c_idx, val)
            cell.font = font_body
            cell.border = border_thin
            if is_even:
                cell.fill = fill_zebra
                
            # Alignments & Formats
            if c_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == prev_col + 1:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.font = font_body_bold
                if val is not None and isinstance(val, (int, float)):
                    cell.number_format = '0.00'
            elif c_idx > 6 and c_idx <= prev_col + 1:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and isinstance(val, (int, float)):
                    cell.number_format = '0.00'
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_out.row_dimensions[r_idx].height = 20

    ws_out.freeze_panes = ws_out.cell(2, 3)

    # Auto-fit columns
    for col in ws_out.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(c.value or '')) for c in col)
        ws_out.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)
    ws_out.column_dimensions[get_column_letter(prev_col + 1)].width = 22

    # Save Output
    wb_out.save(output_file)
    print(f"\nSuccessfully generated: {output_file}")
    print("=" * 75)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Universal Modified Duration Extraction Engine")
    parser.add_argument("--input", "-i", default=None, help="Base input workbook")
    parser.add_argument("--output", "-o", default=None, help="Output workbook path")
    parser.add_argument("--target-month", "-t", default="October", help="Target month to extract")
    parser.add_argument("--prev-month", "-p", default=None, help="Previous month column name to filter")
    parser.add_argument("--factsheet-dir", "-f", default=None, help="Factsheet folder path")
    args = parser.parse_args()

    run_md_pipeline(
        input_file=args.input,
        output_file=args.output,
        target_month=args.target_month,
        prev_month=args.prev_month,
        factsheet_dir=args.factsheet_dir
    )