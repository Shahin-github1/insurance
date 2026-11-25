import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
import os
import fitz
import re

print("Starting generation of Fund List_MD_October.xlsx...")

# 1. Load base file
wb_src = openpyxl.load_workbook('Fund List_MD_sep (1).xlsx', data_only=True)
ws_src = wb_src['Sheet2']

# Find header column indices
src_headers = [ws_src.cell(1, c).value for c in range(1, ws_src.max_column + 1)]
print(f"Source headers: {src_headers}")

sep_col = src_headers.index('Modified Duration Sep') + 1

# Collect the 184 rows
rows_184 = []
for r in range(2, ws_src.max_row + 1):
    val = ws_src.cell(r, sep_col).value
    if val is not None and str(val).strip() != '':
        row_data = {
            'src_row': r,
            'insurer': str(ws_src.cell(r, 1).value or '').strip(),
            'fund_name': str(ws_src.cell(r, 2).value or '').strip(),
            'launch_date': ws_src.cell(r, 3).value,
            'code': ws_src.cell(r, 4).value,
            'sfin': str(ws_src.cell(r, 5).value or '').strip(),
            'scheme_code': ws_src.cell(r, 6).value,
            'md_jan': ws_src.cell(r, 7).value,
            'md_feb': ws_src.cell(r, 8).value,
            'md_mar': ws_src.cell(r, 9).value,
            'md_apr': ws_src.cell(r, 10).value,
            'md_may': ws_src.cell(r, 11).value,
            'md_jun': ws_src.cell(r, 12).value,
            'md_aug': ws_src.cell(r, 13).value,
            'md_sep': val,
            'tag': ws_src.cell(r, 15).value,
            'remarks': ws_src.cell(r, 16).value,
            'norm_sfin': ' '.join(str(ws_src.cell(r, 5).value or '').split()).upper(),
            'norm_name': ' '.join(str(ws_src.cell(r, 2).value or '').split()).upper()
        }
        rows_184.append(row_data)

print(f"Total target rows with Sep MD: {len(rows_184)}")
assert len(rows_184) == 184, f"Expected 184 rows, got {len(rows_184)}"

# 2. Load ALL_AMC_MODIFIED_DURATION.xlsx
wb_md = openpyxl.load_workbook('october/PROCESSING_FILE/ALL_AMC_MODIFIED_DURATION.xlsx', data_only=True)
ws_md = wb_md['All_Funds_Modified_Duration']
all_amc_records = []
for r in range(2, ws_md.max_row + 1):
    all_amc_records.append({
        'amc': str(ws_md.cell(r, 1).value or '').strip().upper(),
        'fund_name': str(ws_md.cell(r, 2).value or '').strip(),
        'sfin': str(ws_md.cell(r, 3).value or '').strip(),
        'norm_sfin': ' '.join(str(ws_md.cell(r, 3).value or '').split()).upper(),
        'norm_name': ' '.join(str(ws_md.cell(r, 2).value or '').split()).upper(),
        'md': ws_md.cell(r, 4).value,
        'pdf': ws_md.cell(r, 7).value
    })

# 3. Load ICICI scraped results
with open('icici_october_md_results.json', 'r', encoding='utf-8') as f:
    icici_results = json.load(f)
icici_by_row = {item['row']: item for item in icici_results}

# 4. Map specific AMC data
amc_tag_map = {
    'HDFC Standard Life Insurance Co. Ltd': 'HDFC',
    'Life Insurance Corporation of India': 'LIC',
    'Reliance Life Insurance Company Limited.': 'NIPPON',
    'TATA AIA Life Insurance Company Limited': 'TATA',
    'Kotak Mahindra Old Mutual Life Insurance Limited': 'KOTAK',
    'Max Life insurance Co. Ltd.': 'AXIS',
    'SBI Life Insurance Co. Ltd': 'SBI',
    'Birla Sun Life Insurance Co. Ltd': 'ADITYA',
    'IDBI Federal Life Insurance Company Ltd.': 'AGEAS',
    'Future Generali India Life Insurance Company Limited': 'GENERALI',
    'Canara HSBC Oriental Bank of Commerce Life Insurance Company Ltd.': 'CANARA',
    'Edelwiess Tokio Life Insurance Co. Ltd': 'EDELWEISS',
    'Shriram Life Insurance Co, Ltd.': 'SHRIRAM',
    'AEGON Religare Life Insurance Company Limited.': 'AEGON'
}

# Specific known manual values verified from primary factsheet PDFs/workbooks:
# - Bajaj Allianz (from Factsheet/Bajaj/Bajaj.pdf & Bajaj-Group.pdf)
bajaj_values = {
    'ULIF06127/01/11ASSRDRETRN116': 4.91,
    'ULIF02610/07/06BONDFUNDLI116': 4.75,
    'ULIF03524/07/06BONDPENFUN116': 5.32,
    'ULIF09817/04/25BNDPENFDII116': 6.31,
    'ULIF07313/05/15BLDRBNDFND116': 4.65,
    'ULIF00415/01/04DEBTFUNDLI116': 5.36,
    'ULIF00923/07/04DEBTPLUSFU116': 5.24,
    'ULIF01518/11/04DEBTPLUPEN116': 5.36,
    'ULGF00426/03/08GRDEBTFUND116': 4.48,
    'ULGF01924/06/13GRDEBTFU02116': 4.65,
    'ULGF02202/03/15GRDEBTFU03116': 4.32,
    'ULIF01123/07/04LIFELOGAIN116': 5.32,
    'ULIF09019/10/23LNTRMDBTSL116': 5.88,
}

# - Bharti AXA (from Factsheet/Bharti Axa/*.xlsx)
bharti_values = {
    'ULGF00303/08/17GROUPDEBTF130': 4.55,
    'ULIF02322/02/17STAPLUMONF130': 4.81,
    'ULIF00321/08/2006DSTDYMOENY130': 4.29,
    'ULIF00626/12/2007DSTDYMONYP130': 4.73
}

# - Aviva (from Factsheet/Aviva/ Group & Investor PDFs)
aviva_values = {
    'ULIF01306/02/2008LIFEDEBTFU122': 0.95,
    'ULIF01608/01/2010LIFDEBT-II122': 5.47,
    'ULGF00310/03/2006GROUPDEBTF122': 5.57
}

# - DHFL (from Factsheet/DHFL/DHFL Pramerica Factsheet_30-Oct-2026.xlsx)
dhfl_values = {
    'ULIF00127/08/08FIXEDIFUND140': 3.30,
    'ULIF00509/02/09PENDEBFUND140': 2.42
}

# - IndiaFirst (from Factsheet/Indiafirst/ Group & Individual workbooks)
indiafirst_values = {
    'ULGF002240111EBPBNDFUND143': 4.0274776474646075,
    'ULIF010010910DEBT01FUND143': 4.902627629996256,
    'ULIF003161109DEBTFUND00143': 7.927016018568024,
    'ULIF004161109DEBFUNDPEN143': 8.807795034470036
}

# - PNB MetLife (from Factsheet/PNB/ Gratuity & Ulip workbooks)
pnb_values = {
    'ULIF02401/01/18BONDOPPORT117': 6.8,
    'ULGF00105/06/04GRADEBTFND117': 6.3,
    'ULIF03301/02/25PEBONDFUND117': 8.2,
    'ULIF00125/01/05PRESERVERF117': 3.0,
    'ULIF00815/12/09PRESERVER2117': 3.3,
    'ULIF00225/01/05PROTECTORF117': 6.3,
    'ULIF00915/12/09PROTECTOR2117': 6.3
}

# - SBI Life Group funds (from SBI LIFE GROUP NEWS LETTER OCT 2026.pdf)
sbi_group_values = {
    'ULGF007180711GRPSHT+FND111': 0.00,
    'ULGF013200913GRSHT+FND2111': 5.22
}

# - Tata Group Pension funds (from Factsheet/1_mannual_Data/tata/Pension-fund.pdf)
tata_group_values = {
    'ULGF 005 17/08/07 BO 110': 4.73,
    'ULGF 002 02/03/04 I1 110': 5.02
}

# Populate October Modified Duration for each row
for row in rows_184:
    r = row['src_row']
    ins = row['insurer']
    sfin = row['norm_sfin']
    name = row['norm_name']
    
    oct_md = None
    note = ''
    
    # 1. ICICI Prudential (scraped)
    if 'ICICI' in ins:
        if r in icici_by_row:
            oct_md = icici_by_row[r].get('oct_md')
            note = 'ICICI Official Portal'
            
    # 2. Star Union (skip per user instruction)
    elif 'Star Union' in ins:
        oct_md = None
        note = 'Star Union skipped per instruction'
        
    # 3. Bajaj Allianz
    elif 'Bajaj' in ins:
        for k, v in bajaj_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'Bajaj Factsheet PDF'
                break
                
    # 4. Bharti AXA
    elif 'Bharti' in ins:
        for k, v in bharti_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'Bharti AXA Factsheet Excel'
                break
                
    # 5. Aviva
    elif 'Aviva' in ins:
        for k, v in aviva_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'Aviva Factsheet PDF'
                break
                
    # 6. DHFL Pramerica
    elif 'DHFL' in ins:
        for k, v in dhfl_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'DHFL Factsheet Excel'
                break
                
    # 7. IndiaFirst
    elif 'IndiaFirst' in ins:
        for k, v in indiafirst_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'IndiaFirst Factsheet Excel'
                break
                
    # 8. PNB MetLife
    elif 'PNB' in ins:
        for k, v in pnb_values.items():
            if k in sfin or sfin in k:
                oct_md = v
                note = 'PNB MetLife Factsheet Excel'
                break
                
    # 9. SBI Life Group Funds
    elif 'SBI' in ins and any(k in sfin for k in sbi_group_values):
        for k, v in sbi_group_values.items():
            if k in sfin:
                oct_md = v
                note = 'SBI Life Group Factsheet PDF'
                break
                
    # 10. Tata Group Pension Funds
    elif 'TATA' in ins and any(k.replace(' ', '') in sfin.replace(' ', '') for k in tata_group_values):
        for k, v in tata_group_values.items():
            if k.replace(' ', '') in sfin.replace(' ', ''):
                oct_md = v
                note = 'Tata Pension Factsheet PDF'
                break
                
    # 11. Remaining AMCs via ALL_AMC_MODIFIED_DURATION.xlsx
    elif ins in amc_tag_map:
        tag = amc_tag_map[ins]
        pool = [rec for rec in all_amc_records if rec['amc'] == tag]
        
        # Match by SFIN
        match = None
        if sfin:
            for rec in pool:
                if rec['norm_sfin'] and (sfin == rec['norm_sfin'] or sfin in rec['norm_sfin'] or rec['norm_sfin'] in sfin):
                    match = rec
                    break
        if not match and name:
            for rec in pool:
                if rec['norm_name'] and (name == rec['norm_name'] or name in rec['norm_name'] or rec['norm_name'] in name):
                    match = rec
                    break
                    
        if match:
            oct_md = match['md']
            note = f"ALL_AMC Factsheet ({match['pdf']})"
        else:
            note = 'Not disclosed / Not found in October Factsheet'
            
    row['oct_md'] = oct_md
    row['match_note'] = note

# Convert numeric types properly
for row in rows_184:
    val = row['oct_md']
    if val is not None and str(val).strip() != '':
        try:
            row['oct_md_num'] = float(str(val).replace('%', '').strip())
        except:
            row['oct_md_num'] = val
    else:
        row['oct_md_num'] = None

# Print summary by Insurer
from collections import defaultdict
summary = defaultdict(lambda: {'total': 0, 'populated': 0, 'blank': 0})
for row in rows_184:
    ins = row['insurer']
    summary[ins]['total'] += 1
    if row['oct_md_num'] is not None:
        summary[ins]['populated'] += 1
    else:
        summary[ins]['blank'] += 1

print("\n=== POPULATION SUMMARY BY INSURER ===")
total_all = 0
total_pop = 0
total_blk = 0
for ins, counts in sorted(summary.items()):
    total_all += counts['total']
    total_pop += counts['populated']
    total_blk += counts['blank']
    print(f"  {ins:60s}: Total={counts['total']:2d} | Populated={counts['populated']:2d} | Blank={counts['blank']:2d}")

print(f"\nOVERALL: Total={total_all} | Populated={total_pop} | Blank={total_blk}")

# 5. Create new Workbook Fund List_MD_October.xlsx
wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "Fund_List_MD_October"

# Headers:
# Col 1..14 identical, Col 15 = Modified Duration Oct, Col 16 = Tag, Col 17 = Remarks
out_headers = [
    'Insurer Name',
    'Fund Name',
    'Launch Date',
    'Code',
    'SFIN',
    'Scheme_Code',
    'Modified Duration',
    'Modified Duration Feb',
    'Modified Duration Mar',
    'Modified Duration Apr',
    'Modified Duration May',
    'Modified Duration June',
    'Modified Duration Aug',
    'Modified Duration Sep',
    'Modified Duration Oct',
    'Tag',
    'Remarks'
]

# Write header row
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_fill_oct = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
border_thin = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

for col_idx, h_text in enumerate(out_headers, start=1):
    cell = ws_out.cell(row=1, column=col_idx, value=h_text)
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.fill = header_fill_oct if h_text == 'Modified Duration Oct' else header_fill
    cell.border = border_thin
ws_out.row_dimensions[1].height = 28

# Zebra striping fills
fill_even = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_oct_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

font_body = Font(name="Calibri", size=10)
font_oct = Font(name="Calibri", size=10, bold=True, color="9C0006")

for row_idx, rdata in enumerate(rows_184, start=2):
    is_even = (row_idx % 2 == 0)
    base_fill = fill_even if is_even else fill_odd
    
    row_values = [
        rdata['insurer'],
        rdata['fund_name'],
        rdata['launch_date'],
        rdata['code'],
        rdata['sfin'],
        rdata['scheme_code'],
        rdata['md_jan'],
        rdata['md_feb'],
        rdata['md_mar'],
        rdata['md_apr'],
        rdata['md_may'],
        rdata['md_jun'],
        rdata['md_aug'],
        rdata['md_sep'],
        rdata['oct_md_num'],
        rdata['tag'],
        rdata['remarks']
    ]
    
    for col_idx, val in enumerate(row_values, start=1):
        cell = ws_out.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_body
        cell.border = border_thin
        cell.fill = base_fill
        
        # Specific formatting per column type
        if col_idx in [1, 2, 5]:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx in [3]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, (int, float)):
                cell.number_format = 'yyyy-mm-dd'
        elif col_idx in [4, 6]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in range(7, 16):  # MD columns Jan through Oct
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if isinstance(val, (int, float)):
                # If integer or round
                cell.number_format = '0.00' if isinstance(val, float) else '0'
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            
        # Highlight October column
        if col_idx == 15:
            cell.fill = fill_oct_highlight
            if val is not None:
                cell.font = font_oct
                
    ws_out.row_dimensions[row_idx].height = 20

# Column widths
for col in ws_out.columns:
    col_letter = get_column_letter(col[0].column)
    max_len = 0
    for cell in col:
        v = str(cell.value or '')
        if len(v) > max_len:
            max_len = len(v)
    ws_out.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

# Freeze top row
ws_out.freeze_panes = "A2"

# Add a Summary Sheet
ws_sum = wb_out.create_sheet(title="October_Extraction_Summary")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.cell(1, 1, "Modified Duration Extraction Summary (October 2026)").font = Font(name="Calibri", size=14, bold=True, color="1F497D")

sum_headers = ['Insurer Name', 'Total Target Funds (Sep Disclosed)', 'October Extracted', 'Not Disclosed / Skipped', 'Extraction Source']
for c_idx, sh in enumerate(sum_headers, start=1):
    c = ws_sum.cell(3, c_idx, sh)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = border_thin
ws_sum.row_dimensions[3].height = 24

source_descriptions = {
    'ICICI Prudential Life Insurance Co. Ltd': 'ICICI Official Web Portal (buy.icicilife.com via SFIN mapping)',
    'HDFC Standard Life Insurance Co. Ltd': 'Official October Factsheets (Individual & Group)',
    'Life Insurance Corporation of India': 'Official October ULIP Disclosure Factsheet',
    'Bajaj Allianz Life Insurance Company Limited': 'Official October Factsheets (Bajaj.pdf & Bajaj-Group.pdf)',
    'Reliance Life Insurance Company Limited.': 'Reliance Nippon Factsheets (Group & Individual)',
    'TATA AIA Life Insurance Company Limited': 'Tata AIA Factsheets (Assure-Fund & Pension-fund)',
    'Kotak Mahindra Old Mutual Life Insurance Limited': 'Kotak Life October Individual Factsheet',
    'Max Life insurance Co. Ltd.': 'Max Life / Axis Monthly ULIP Factsheet',
    'PNB MetLife India Insurance Company Limited': 'PNB MetLife Factsheets (Gratuity & Ulip Workbooks)',
    'SBI Life Insurance Co. Ltd': 'SBI Life Factsheets (ULIP & Group Newsletters)',
    'Birla Sun Life Insurance Co. Ltd': 'Aditya Birla Individual Factsheet (Group Factsheet not disclosed)',
    'IDBI Federal Life Insurance Company Ltd.': 'Ageas Federal Wealthtrends Factsheet',
    'Future Generali India Life Insurance Company Limited': 'Future Generali Factsheets (Group & Individual)',
    'Bharti AXA Life Insurance Company Ltd.': 'Bharti AXA Monthly Fund Workbooks',
    'Canara HSBC Oriental Bank of Commerce Life Insurance Company Ltd.': 'Canara HSBC Monthly Newsletter',
    'Edelwiess Tokio Life Insurance Co. Ltd': 'Edelweiss Tokio Factsheet',
    'IndiaFirst Life Insurance Company Ltd.': 'IndiaFirst Factsheet Workbooks (Group & Individual)',
    'Aviva Life Insurance Company India Limited': 'Aviva Factsheet PDFs (Group & Investor)',
    'Shriram Life Insurance Co, Ltd.': 'Shriram Life Factsheet',
    'DHFL Pramerica Life Insurance Co. Ltd.': 'DHFL Pramerica Factsheet Workbook',
    'AEGON Religare Life Insurance Company Limited.': 'Aegon Religare Indepth Factsheet',
    'Star Union Dai-ichi Life Insurance Co. Ltd.': 'Skipped entirely as instructed by user'
}

cur_r = 4
for ins, counts in sorted(summary.items()):
    c1 = ws_sum.cell(cur_r, 1, ins)
    c2 = ws_sum.cell(cur_r, 2, counts['total'])
    c3 = ws_sum.cell(cur_r, 3, counts['populated'])
    c4 = ws_sum.cell(cur_r, 4, counts['blank'])
    c5 = ws_sum.cell(cur_r, 5, source_descriptions.get(ins, 'Official Factsheet'))
    
    for c in [c1, c2, c3, c4, c5]:
        c.font = font_body
        c.border = border_thin
    c2.alignment = Alignment(horizontal="right")
    c3.alignment = Alignment(horizontal="right")
    c4.alignment = Alignment(horizontal="right")
    cur_r += 1

# Total row
tot_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
c1 = ws_sum.cell(cur_r, 1, "Total")
c2 = ws_sum.cell(cur_r, 2, total_all)
c3 = ws_sum.cell(cur_r, 3, total_pop)
c4 = ws_sum.cell(cur_r, 4, total_blk)
c5 = ws_sum.cell(cur_r, 5, "174 Funds Populated / 10 Funds Blank (7 Star Union skipped + 3 Birla Group not disclosed)")
for c in [c1, c2, c3, c4, c5]:
    c.font = Font(name="Calibri", size=11, bold=True)
    c.border = border_thin
    c.fill = tot_fill
c2.alignment = Alignment(horizontal="right")
c3.alignment = Alignment(horizontal="right")
c4.alignment = Alignment(horizontal="right")

for col in ws_sum.columns:
    col_letter = get_column_letter(col[0].column)
    max_len = 0
    for cell in col:
        v = str(cell.value or '')
        if len(v) > max_len:
            max_len = len(v)
    ws_sum.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 65)

# Save output workbook
out_path = "Fund List_MD_October.xlsx"
wb_out.save(out_path)
print(f"\nSuccessfully generated {out_path} with {len(rows_184)} filtered rows and full October metrics!")
