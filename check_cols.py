import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\alexc\Downloads\Alejandra\Alejandra\Base_Actualizada.xlsx', data_only=True)
ws = wb['SGA AJ']
rows = list(ws.iter_rows(values_only=True))
header = rows[0]
data = rows[1:]

print(f"Total filas de datos: {len(data)}")
for col_idx, col_name in enumerate(header):
    vals = [r[col_idx] for r in data if r[col_idx] is not None and str(r[col_idx]).strip() != '']
    unique_vals = set(str(v).strip() for v in vals)
    print(f"Col {col_idx} ('{col_name}'): {len(vals)} no vacíos | {len(unique_vals)} únicos")
