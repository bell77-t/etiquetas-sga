import openpyxl
import re
from collections import Counter

wb = openpyxl.load_workbook(r'C:\Users\alexc\Downloads\Alejandra\Alejandra\Base_Actualizada.xlsx', data_only=True)
ws = wb['SGA AJ']
rows = list(ws.iter_rows(values_only=True))

print("Total filas leidas de la hoja:", len(rows))
header = rows[0]
print("Cabecera:", header[:5])

data_rows = rows[1:]
print("Total filas de datos (sin fila 1 cabecera):", len(data_rows))

nombres_validos = []
filas_con_problema = []

for idx, r in enumerate(data_rows, start=2):
    codigo = r[0]
    nombre = r[1]
    if not any(r):
        filas_con_problema.append((idx, "FILA TOTALMENTE VACIA", r[:3]))
    elif nombre is None or str(nombre).strip() == "":
        filas_con_problema.append((idx, "NOMBRE DE PRODUCTO VACIO/NONE", r[:3]))
    else:
        nombres_validos.append((idx, str(nombre).strip(), str(codigo).strip() if codigo else ""))

print(f"\nFilas con producto valido: {len(nombres_validos)}")
print(f"Filas con problema ({len(filas_con_problema)}):")
for fp in filas_con_problema:
    print(f"  Fila Excel {fp[0]}: {fp[1]} -> {fp[2]}")

norm_nombres = [re.sub(r'\s+', ' ', n[1].upper()) for n in nombres_validos]
counts = Counter(norm_nombres)
dups = {k: v for k, v in counts.items() if v > 1}

print(f"\nProductos con nombres duplicados ({len(dups)}):")
total_repetidos = 0
for k, v in dups.items():
    filas = [n[0] for n in nombres_validos if re.sub(r'\s+', ' ', n[1].upper()) == k]
    codigos = [n[2] for n in nombres_validos if re.sub(r'\s+', ' ', n[1].upper()) == k]
    print(f"  - '{k}': aparece {v} veces en filas Excel {filas} (Codigos: {codigos})")
    total_repetidos += (v - 1)

print(f"\nTotal repeticiones restadas: {total_repetidos}")
print(f"229 filas - {total_repetidos} repeticiones = {len(nombres_validos) - total_repetidos}")
