import fitz
import re
import openpyxl
import os
import glob

def clean_ligatures(s: str) -> str:
    if not s:
        return ""
    s = s.replace('\u019e', 'tf').replace('\u01a0', 'ti')
    s = s.replace('\ufb01', 'fi').replace('\ufb02', 'fl')
    s = s.replace('ƞ', 'tf').replace('Ɵ', 'ti')
    return s

def extract_pdf_funds(pdf_path, amc_name):
    doc = fitz.open(pdf_path)
    funds = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        if not re.search(r'(?i)\b(?:modified\s+duration|mod\.?\s*duration|m\.?\s*duration|portfolio\s+duration)\b', text):
            continue
        
        # Extract title and SFIN
        d = page.get_text('dict')
        candidates = []
        sfin_code = ""
        for b in d.get('blocks', []):
            if b.get('type') == 0:
                for l in b.get('lines', []):
                    for s in l.get('spans', []):
                        txt = clean_ligatures(s.get('text', '')).strip()
                        if not txt:
                            continue
                        y = s.get('bbox', [0, 0, 0, 0])[1]
                        sz = s.get('size', 0)
                        m_sfin = re.search(r'\b(UL[I|G]F[A-Z0-9\/\-]+)', txt)
                        if m_sfin and not sfin_code:
                            sfin_code = m_sfin.group(1)
                        if y < 150 and len(txt) > 3:
                            if not any(ign in txt.lower() for ign in ['page', 'factsheet', 'report', 'objective', 'benchmark', 'nav as on', 'monthly', 'date', 'fund manager', 'aum']):
                                candidates.append((sz, -y, txt))
        
        fund_name = "Unknown Fund"
        if candidates:
            candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
            raw_title = candidates[0][2]
            raw_title = re.sub(r'\(?UL[I|G]F.*', '', raw_title)
            raw_title = re.sub(r'\s+as\s+on\s+.*', '', raw_title, flags=re.IGNORECASE)
            fund_name = raw_title.strip(' -:\t\n')

        # Extract modified duration
        blocks = page.get_text('blocks')
        mod_dur = None
        for b in blocks:
            b_txt = clean_ligatures(b[4]).strip()
            m_same = re.search(r'(?i)(?:modified\s+duration|mod\.?\s*duration|m\.?\s*duration|portfolio\s+duration)[^\d\n]*[:\-\s\t]+([0-9]+\.[0-9]+)', b_txt)
            if m_same and mod_dur is None:
                mod_dur = float(m_same.group(1))

        if mod_dur is not None:
            funds.append({
                'amc': amc_name,
                'fund_name': fund_name,
                'sfin': sfin_code,
                'mod_dur': mod_dur,
                'page': page_num + 1,
                'source': os.path.basename(pdf_path)
            })
    return funds

if __name__ == '__main__':
    print('Testing Bajaj extraction...')
    for p in ['Factsheet/Bajaj/Bajaj.pdf', 'Factsheet/Bajaj/Bajaj-Group.pdf']:
        if os.path.exists(p):
            f_list = extract_pdf_funds(p, 'BAJAJ')
            print(f'{p}: found {len(f_list)} funds')
            for f in f_list[:5]:
                print(f"  {f['fund_name']} | SFIN: {f['sfin']} | MD: {f['mod_dur']}")
