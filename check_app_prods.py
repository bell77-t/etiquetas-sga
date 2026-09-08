import openpyxl

wb_app = openpyxl.load_workbook(r'C:\Users\alexc\Downloads\Alejandra\Alejandra\aplicacion.xlsm', data_only=True)
ws_data = wb_app['Data']
rows = list(ws_data.iter_rows(values_only=True))
print("aplicacion.xlsm total filas en Data:", len(rows))
# Cabecera
header = rows[0]
print("Data header:", header[:10])

# Contar productos en aplicacion.xlsm
prod_col_idx = None
for idx, c in enumerate(header):
    if c and "PRODUCTO" in str(c).upper():
        prod_col_idx = idx
        break

print(f"Columna de Producto: {prod_col_idx} ('{header[prod_col_idx]}')")

prods = [r[prod_col_idx] for r in rows[1:] if r[prod_col_idx] is not None and str(r[prod_col_idx]).strip() != '']
from collections import Counter
c = Counter([str(p).strip().upper() for p in prods])
print(f"Total registros con producto en Data: {len(prods)}")
print(f"Total productos UNICOS en aplicacion.xlsm: {len(c)}")
