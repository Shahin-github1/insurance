import urllib.request
import ssl
import gzip
import json
import re
import openpyxl
import time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Origin': 'https://www.iciciprulife.com',
    'Referer': 'https://www.iciciprulife.com/',
    'Sec-Ch-Ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
    'Sec-Ch-Ua-Mobile': '?0',
    'Sec-Ch-Ua-Platform': '"Windows"',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'cross-site'
}

def fetch_json(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
        content = resp.read()
        if resp.info().get('Content-Encoding') == 'gzip':
            content = gzip.decompress(content)
        return json.loads(content.decode('utf-8'))

# 1. Load the 25 ICICI funds from Sheet2
wb = openpyxl.load_workbook('Fund List_MD_sep (1).xlsx', data_only=True)
ws2 = wb['Sheet2']
icici_funds = []
for r in range(2, ws2.max_row + 1):
    ins = ws2.cell(r, 1).value
    sep = ws2.cell(r, 14).value
    if ins and 'ICICI' in str(ins) and sep is not None and str(sep).strip() != '':
        icici_funds.append({
            'row': r,
            'fund_name': str(ws2.cell(r, 2).value).strip(),
            'sfin': str(ws2.cell(r, 5).value).strip() if ws2.cell(r, 5).value else '',
            'code': str(ws2.cell(r, 4).value).strip() if ws2.cell(r, 4).value else '',
            'sep_md': ws2.cell(r, 14).value
        })

print(f'Total ICICI funds with Sep MD in Sheet2: {len(icici_funds)}')

# 2. Get all product codes from product-wise-fund-performance-details.html
with open('product-wise-fund-performance-details.html', 'r', encoding='utf-8') as f:
    html = f.read()

m = re.search(r'<select[^>]*id=[\'"]products[\'"][^>]*>(.*?)</select>', html, re.DOTALL)
product_codes = []
if m:
    opts = re.findall(r'<option[^>]*value=[\'"]([^\'"]+)[\'"][^>]*>(.*?)</option>', m.group(1))
    product_codes = [v.strip() for v, t in opts if v.strip() and v.strip() != '-']

print(f'Found {len(product_codes)} product codes.')

# 3. Query all product codes to map SFIN -> fundCode
sfin_to_fund = {}
name_to_fund = {}
all_funds_discovered = {}

for pcode in product_codes:
    try:
        url = f'https://buy.icicilife.com/portal/login/funds-by-product.htm?productCode={pcode}'
        data = fetch_json(url)
        for item in data:
            fcode = item.get('LAfundCode')
            fname = item.get('Fund', '').strip()
            sfin = item.get('SFIN', '').strip()
            if sfin:
                # normalize sfin (remove multiple spaces)
                norm_sfin = ' '.join(sfin.split()).upper()
                sfin_to_fund[norm_sfin] = item
            if fname:
                norm_name = ' '.join(fname.split()).upper()
                name_to_fund[norm_name] = item
            if fcode:
                all_funds_discovered[fcode] = item
    except Exception as e:
        print(f'Error fetching product {pcode}: {e}')

print(f'Total unique SFINs discovered: {len(sfin_to_fund)}')
print(f'Total unique fund names discovered: {len(name_to_fund)}')
print(f'Total unique fund codes discovered: {len(all_funds_discovered)}')

# Save the discovered map to json for caching
with open('icici_discovered_funds.json', 'w', encoding='utf-8') as f:
    json.dump({
        'sfin_to_fund': sfin_to_fund,
        'name_to_fund': name_to_fund,
        'all_funds': all_funds_discovered
    }, f, indent=2)

# 4. Match the 25 ICICI funds
matched = []
unmatched = []

for f in icici_funds:
    norm_sfin = ' '.join(f['sfin'].split()).upper()
    norm_name = ' '.join(f['fund_name'].split()).upper()
    
    found_item = None
    if norm_sfin in sfin_to_fund:
        found_item = sfin_to_fund[norm_sfin]
        match_type = 'SFIN'
    elif norm_name in name_to_fund:
        found_item = name_to_fund[norm_name]
        match_type = 'Name'
    else:
        # try partial or fuzzy
        for s, it in sfin_to_fund.items():
            if norm_sfin and (norm_sfin in s or s in norm_sfin):
                found_item = it
                match_type = 'Partial SFIN'
                break
        if not found_item:
            for n, it in name_to_fund.items():
                if norm_name and (norm_name in n or n in norm_name):
                    found_item = it
                    match_type = 'Partial Name'
                    break

    if found_item:
        f['la_fund_code'] = found_item.get('LAfundCode')
        f['match_type'] = match_type
        f['api_fund_name'] = found_item.get('Fund')
        f['api_sfin'] = found_item.get('SFIN')
        matched.append(f)
    else:
        unmatched.append(f)

print(f'\nMatched {len(matched)} / {len(icici_funds)} ICICI funds.')
for m in matched:
    print(f"Row {m['row']}: {m['fund_name']} -> LAfundCode={m['la_fund_code']} (via {m['match_type']})")

if unmatched:
    print(f'\nUnmatched {len(unmatched)} funds:')
    for u in unmatched:
        print(f"Row {u['row']}: {u['fund_name']} | SFIN={u['sfin']} | Sep={u['sep_md']}")
