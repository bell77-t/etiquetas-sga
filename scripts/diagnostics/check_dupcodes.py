import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\alexc\Downloads\Alejandra\Alejandra\Base_Actualizada.xlsx', data_only=True)
ws = wb['SGA AJ']
rows = list(ws.iter_rows(values_only=True))

dup_codes = ['2FFBAS01', '2FFCOD01', '1DEOXV01']

print("Filas con codigos repetidos:")
for idx, r in enumerate(rows[1:], start=2):
    code = str(r[0]).strip().upper() if r[0] else ""
    if code in dup_codes:
        print(f"  Fila {idx}: Codigo={r[0]} | Nombre={r[1]} | Unidad={r[2]} | Advertencia={r[5]}")
