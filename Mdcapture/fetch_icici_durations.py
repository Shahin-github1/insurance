import urllib.request
import ssl
import gzip
import json
import openpyxl

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

# Load cached discovered funds
with open('icici_discovered_funds.json', 'r', encoding='utf-8') as f:
    cache = json.load(f)
sfin_to_fund = cache['sfin_to_fund']

wb = openpyxl.load_workbook('Fund List_MD_sep (1).xlsx', data_only=True)
ws2 = wb['Sheet2']
icici_funds = []
for r in range(2, ws2.max_row + 1):
    ins = ws2.cell(r, 1).value
    sep = ws2.cell(r, 14).value
    if ins and 'ICICI' in str(ins) and sep is not None and str(sep).strip() != '':
        sfin = str(ws2.cell(r, 5).value).strip() if ws2.cell(r, 5).value else ''
        norm_sfin = ' '.join(sfin.split()).upper()
        fcode = sfin_to_fund.get(norm_sfin, {}).get('LAfundCode')
        icici_funds.append({
            'row': r,
            'fund_name': str(ws2.cell(r, 2).value).strip(),
            'sfin': sfin,
            'code': str(ws2.cell(r, 4).value).strip() if ws2.cell(r, 4).value else '',
            'sep_md': sep,
            'la_fund_code': fcode
        })

print(f'Fetching October Modified Duration for {len(icici_funds)} ICICI funds...')
results = []
for f in icici_funds:
    fcode = f['la_fund_code']
    if not fcode:
        print(f"Row {f['row']}: No fund code found!")
        continue
    url = f'https://buy.icicilife.com/portal/login/funds-portfolio-details.htm?fundCode={fcode}'
    try:
        data = fetch_json(url)
        ytm_mod = data.get('YTM_And_Modified_Duration', [])
        mod_val = None
        ytm_val = None
        if ytm_mod and isinstance(ytm_mod, list) and len(ytm_mod) > 0:
            item = ytm_mod[0]
            mod_val = item.get('modified_value')
            ytm_val = item.get('ytm_value')
        
        # also check average maturity
        avg_mat = None
        if 'Average_maturity_of_debt_portfolio' in data:
            amat = data['Average_maturity_of_debt_portfolio']
            if amat and isinstance(amat, list) and len(amat) > 0:
                avg_mat = amat[0].get('value')

        f['oct_md'] = mod_val
        f['ytm'] = ytm_val
        f['avg_maturity'] = avg_mat
        print(f"Row {f['row']}: {f['fund_name']} (fcode={fcode}) | Sep MD: {f['sep_md']} -> Oct MD: {mod_val} | YTM: {ytm_val}")
        results.append(f)
    except Exception as e:
        print(f"Error for {f['fund_name']} ({fcode}): {e}")
        f['oct_md'] = None
        results.append(f)

with open('icici_october_md_results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2)

print(f'\nFinished fetching {len(results)} ICICI funds. Saved to icici_october_md_results.json')
