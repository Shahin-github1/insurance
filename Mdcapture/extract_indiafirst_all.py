import openpyxl

wb_ind = openpyxl.load_workbook('Factsheet/Indiafirst/India_First_Report_Individual_30-Oct-2026.xlsx', data_only=True)
wb_grp = openpyxl.load_workbook('Factsheet/Indiafirst/India_First_Report_Group_30-Oct-2026.xlsx', data_only=True)

if_funds = [
    ('Bond Fund', wb_grp, 'Bond Fund'),
    ('Debt 1 Fund', wb_ind, 'Debt 1 Fund'),
    ('Debt Fund', wb_ind, 'Debt Fund'),
    ('Debt Fund - Pension', wb_ind, 'Debt Fund - Pension')
]

for name, wb, sname in if_funds:
    ws = wb[sname]
    mod_dur = None
    avg_mat = None
    for r in range(1, min(ws.max_row+1, 120)):
        for c in range(1, min(ws.max_column+1, 15)):
            v = str(ws.cell(r, c).value or '').strip()
            if 'modified duration' in v.lower():
                # check cell below or right
                val_below = ws.cell(r+1, c).value
                val_right = ws.cell(r, c+1).value
                mod_dur = val_below if val_below is not None else val_right
            if 'average maturity' in v.lower():
                val_below = ws.cell(r+1, c).value
                val_right = ws.cell(r, c+1).value
                avg_mat = val_below if val_below is not None else val_right
    print(f'IndiaFirst [{name}]: MD = {mod_dur} | Avg Maturity = {avg_mat}')
