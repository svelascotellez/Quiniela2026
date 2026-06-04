import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\salva\Downloads\Quiniela_Mundial_2026_Final_vacia (1).xlsx', data_only=False)
ws = wb['Eliminatorias']
print('--- Eliminatorias (cols 1-10) ---')
for r in range(1, 120):
    row_vals = []
    has_data = False
    for c in range(1, 12):
        v = ws.cell(row=r, column=c).value
        sv = str(v)[:22] if v is not None else '_'
        row_vals.append(sv)
        if v is not None:
            has_data = True
    if has_data:
        print(f'  R{r:03d}: {" | ".join(row_vals)}')
