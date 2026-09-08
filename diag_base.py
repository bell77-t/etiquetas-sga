import openpyxl

wb = openpyxl.load_workbook(r'C:\Users\alexc\Downloads\Alejandra\Alejandra\Base_Actualizada.xlsx', data_only=True)
ws = wb['SGA AJ']

rows = list(ws.iter_rows(values_only=True))
print("Total rows:", len(rows))
# Row 0 is header:
print("Header (row 1):", rows[0])

# Rows 1 to end:
data_rows = rows[1:]
print("Total data rows:", len(data_rows))

# Let's check every single row:
non_empty_rows = []
for idx, r in enumerate(data_rows, start=2):
    has_any = any(c is not None and str(c).strip() != '' for c in r)
    code = r[0]
    name = r[1]
    um = r[2]
    frase_h = r[3]
    frase_p = r[4]
    adv = r[5]
    has_sga = r[15] if len(r) > 15 else None
    
    if has_any:
        non_empty_rows.append((idx, code, name, um, adv, has_sga))

print(f"Total non-empty data rows: {len(non_empty_rows)}")

# Check duplicates by code:
from collections import Counter
codes = [str(r[1]).strip().upper() for r in non_empty_rows if r[1] is not None and str(r[1]).strip() != '']
code_counts = Counter(codes)
dup_codes = {k: v for k, v in code_counts.items() if v > 1}
print("\nCodigos repetidos:", dup_codes)

# Check duplicates by normalized product name:
names = [str(r[2]).strip() for r in non_empty_rows if r[2] is not None and str(r[2]).strip() != '']
norm_names = [" ".join(n.upper().split()) for n in names]
name_counts = Counter(norm_names)
dup_names = {k: v for k, v in name_counts.items() if v > 1}
print("\nNombres repetidos:", dup_names)

# Check unique codes:
unique_codes = set(codes)
print(f"Codigos unicos (no nulos): {len(unique_codes)}")

# Check unique names:
unique_names = set(norm_names)
print(f"Nombres unicos (no nulos): {len(unique_names)}")

# Look for blank names or blank codes:
blank_names = [r for r in non_empty_rows if r[2] is None or str(r[2]).strip() == '']
blank_codes = [r for r in non_empty_rows if r[1] is None or str(r[1]).strip() == '']
print(f"\nFilas con nombre vacio: {len(blank_names)}")
print(f"Filas con codigo vacio: {len(blank_codes)}")
for r in blank_codes:
    print(f"  Fila {r[0]}: Nombre='{r[2]}'")

# Check rows where Tiene SGA == 'SI' or 'NO'
sga_si = [r for r in non_empty_rows if str(r[5]).strip().upper() == 'SI']
sga_no = [r for r in non_empty_rows if str(r[5]).strip().upper() == 'NO']
sga_none = [r for r in non_empty_rows if r[5] is None or str(r[5]).strip() not in ('SI', 'NO')]
print(f"\nTiene SGA == SI: {len(sga_si)}")
print(f"Tiene SGA == NO: {len(sga_no)}")
print(f"Tiene SGA == vacio/otro: {len(sga_none)}")

# Look at the very end rows (around row 220-235)
print("\nUltimas 12 filas del archivo Excel:")
for r in non_empty_rows[-12:]:
    print(f"  Fila Excel {r[0]}: Codigo={r[1]} | Nombre={r[2]} | Advertencia={r[4]} | SGA={r[5]}")
