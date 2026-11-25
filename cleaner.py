import re
from datetime import datetime

MONTHS = {
    'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
    'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
}

# Common PDF page footer/header bleed patterns (e.g. URLs and insurer names)
PDF_ARTIFACT_PATTERNS = [
    r'(?i)\s+(?:Aditya Birla Sun|HDFC|Star Union(?: Dai-ichi)?|Tata AIA|Bandhan|SBI|ICICI Prudential|Bajaj(?: Allianz)?|Kotak(?: Mahindra)?|Max|Aviva|PNB MetLife|Generali|Canara HSBC|Bharti AXA|Pramerica|Ageas Federal|Edelweiss(?: Tokio)?|Shriram)\s+Life\s+Insurance.*',
    r'(?i)\b(?:https?://|www\.)?\S+\.(?:com|in|org|net|co\.in)\b',
    r'(?i)\s*/\s*india\b'
]

def clean_pdf_company_name(txt: str) -> str:
    """
    Cleans raw PDF_COMPANY_NAME by removing:
    1. PDF footer/header bleed (URLs like 'lifeinsurance.adityabirlacapital.com' and merged insurer names)
    2. DTB Treasury bill patterns
    3. Coupon rates (e.g. '7.80%')
    4. Maturity dates and maturity years
    5. Formatting variations as defined in Insurance_Data_Functions_QC.txt.
    """
    if not txt or not isinstance(txt, str):
        return ""
    txt = txt.strip()
    
    # 0. Strip PDF Header/Footer bleed artifacts (e.g. URLs, insurer names)
    for pat in PDF_ARTIFACT_PATTERNS:
        txt = re.sub(pat, '', txt).strip()

    # Check for DTB pattern: e.g. '0364 DTB...' or '364 DTB 14102027' -> '364 DAYS TREASURY BILL'
    dtb_match = re.sub(r'^\s*0*(91|182|364)\s+DTB\b.*', r'\1 DAYS TREASURY BILL', txt, flags=re.IGNORECASE)
    if dtb_match != txt:
        return dtb_match.strip()
    
    # 1. Strip coupon rate at the beginning (e.g. '7.5%', '7.80% ', '7.50 ')
    x = re.sub(r'^\s*(?:\d+(?:\.\d+)?\s*%|(?:\d+\.\d+)\s*)\s*', '', txt).strip()
    
    # 2. Strip leading date/year patterns:
    pat_lead_date = r'^\s*(?:\d{4}\s*[-/.]\s*[A-Z]{3,9}\s*[-/.]\s*\d{1,2}|\d{1,2}\s*[-/.]\s*[A-Z]{3,9}\s*[-/.]\s*\d{2,4}|\d{4}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2}|\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]?\s*\d{2,4}|\d{8}|\d{4})\s*[-–]?\s*'
    y = re.sub(pat_lead_date, '', x, flags=re.IGNORECASE).strip()
    
    # 3. Strip trailing maturity and date patterns:
    pat_trail_date = r'\s*\(?\s*(?:\b(?:MD|MAT)\b.*|\d{4}\s*[-/.]\s*[A-Z]{3,9}\s*[-/.]\s*\d{1,2}.*|\d{1,2}\s*[-/.]\s*[A-Z]{3,9}\s*[-/.]\s*\d{2,4}.*|\d{1,2}\s+[A-Z]{3,9}\s+\d{2,4}.*|\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]?\s*\d{2,4}.*|\d{8}.*|\d{4}(?:\b|[-/.]).*)$'
    z = re.sub(pat_trail_date, '', y, flags=re.IGNORECASE).strip()
    
    # 4. Strip trailing punctuation, hyphens, en-dashes, and brackets
    r = re.sub(r'[\s\(\-–]+$', '', z).strip()
    
    # 5. T-Bill pattern: ^\d+\s+Days?\s+(T-Bill)$ -> \1
    base = re.sub(r'^\d+\s+Days?\s+(T-Bill)$', r'\1', r, flags=re.IGNORECASE).strip()
    return base


def normalize_for_matching(s: str) -> str:
    """
    Normalizes company names for fuzzy/approximate matching:
    - Replaces OCR pipe characters '|' with 'l' or 'I'
    - Lowercases text
    - Standardizes entity suffixes (Ltd, Limited, Pvt Ltd, etc.)
    - Removes punctuation, dots, hyphens, and compresses whitespace
    """
    if not s or not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r'(?<=[a-zA-Z])\|(?=[a-zA-Z\s])', 'l', s)
    s = re.sub(r'\|', 'l', s)
    s = s.lower()
    
    s = re.sub(r'\b(ltd|l\.t\.d|l\.t\.d\.|limited)\b', 'ltd', s)
    s = re.sub(r'\b(pvt|p\.v\.t|p\.v\.t\.|private)\b', 'pvt', s)
    s = re.sub(r'\b(corp|corporation)\b', 'corp', s)
    s = re.sub(r'\b(co|company)\b', 'co', s)
    s = re.sub(r'\b(and|&)\b', '&', s)
    s = re.sub(r'[\.,\(\)\-\_\[\]]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_coupon(txt: str):
    """
    Extracts coupon percentage rate from raw company string (e.g. '8.00% HDFC...' -> 8.0).
    Returns float or None.
    """
    if not txt or not isinstance(txt, str):
        return None
    txt = txt.strip()
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*%', txt)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return None


def extract_maturity(txt: str):
    """
    Extracts maturity date and maturity year from raw company string.
    Ensures all years are in the 2000s (e.g. 1932/32 -> 2032, 1963/63 -> 2063).
    Returns: (maturity_date_str, maturity_year_int)
    """
    if not txt or not isinstance(txt, str):
        return None, None
    txt = txt.strip()
    
    # 1. Textual Month: '16 Feb 2031', '11- MAY -2036', '27-Jul-32', '27JAN27', '07 SEP 2031'
    m_alpha = re.search(r'(\d{1,2})[-\s]*([A-Za-z]{3})[-\s]*(\d{2,4})', txt)
    if m_alpha:
        day = int(m_alpha.group(1))
        mon_str = m_alpha.group(2).upper()
        yr = int(m_alpha.group(3))
        if yr < 100:
            yr += 2000
        elif yr < 2000:
            yr = 2000 + (yr % 100)
        if mon_str in MONTHS and 1 <= day <= 31:
            month = MONTHS[mon_str]
            try:
                dt = datetime(yr, month, day)
                return dt.strftime('%Y-%m-%d'), yr
            except ValueError:
                pass

    # 2. Numeric date: '23-10-2028', '28/12/2026', '04-08-2036', '28-Jun-27'
    m_num = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})', txt)
    if m_num:
        day = int(m_num.group(1))
        month = int(m_num.group(2))
        yr = int(m_num.group(3))
        if yr < 100:
            yr += 2000
        elif yr < 2000:
            yr = 2000 + (yr % 100)
        if 1 <= month <= 12 and 1 <= day <= 31:
            try:
                dt = datetime(yr, month, day)
                return dt.strftime('%Y-%m-%d'), yr
            except ValueError:
                pass

    # 3. 8-digit date in DTBs: DDMMYYYY e.g. '14102027' -> 2027-10-14
    m_8 = re.search(r'\b(\d{2})(\d{2})(\d{4})\b', txt)
    if m_8:
        day = int(m_8.group(1))
        month = int(m_8.group(2))
        yr = int(m_8.group(3))
        if yr < 2000:
            yr = 2000 + (yr % 100)
        if 1 <= month <= 12 and 1 <= day <= 31 and yr >= 2020:
            try:
                dt = datetime(yr, month, day)
                return dt.strftime('%Y-%m-%d'), yr
            except ValueError:
                pass

    # 4. Standalone maturity year at the end, e.g. '7.80% HDFC Bank Ltd 2033'
    m_yr = re.search(r'\b(20[2-9]\d)\b', txt)
    if m_yr:
        yr = int(m_yr.group(1))
        return None, yr

    return None, None
