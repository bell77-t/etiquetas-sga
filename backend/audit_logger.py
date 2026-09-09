"""
audit_logger.py - Registro de Auditoría y Trazabilidad Fitosanitaria (ICA / GlobalGAP)
Almacena en SQLite local cada impresión de etiqueta, despacho de tanques y edición de productos.
"""

import sqlite3
import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = PROJECT_ROOT / "var" / "audit.db"

def init_db():
    """Inicializa la base de datos de auditoría si no existe."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA busy_timeout=5000")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            operario TEXT NOT NULL,
            accion TEXT NOT NULL,
            programa TEXT,
            cultivo TEXT,
            sector_bloque TEXT,
            producto TEXT,
            codigo TEXT,
            dosis REAL,
            volumen_tanque REAL,
            numero_tanque INTEGER,
            total_tanques INTEGER,
            copias INTEGER DEFAULT 1,
            detalles TEXT,
            usuario_id INTEGER,
            rol TEXT,
            ip_origen TEXT,
            sesion_id TEXT,
            resultado TEXT DEFAULT 'EXITO',
            valores_anteriores TEXT,
            valores_nuevos TEXT
        )
    """)
    # Migración compatible con bases SQLite que ya existen.
    existing = {r[1] for r in cursor.execute("PRAGMA table_info(audit_logs)")}
    for name, definition in {
        "usuario_id": "INTEGER", "rol": "TEXT", "ip_origen": "TEXT", "sesion_id": "TEXT",
        "resultado": "TEXT DEFAULT 'EXITO'", "valores_anteriores": "TEXT", "valores_nuevos": "TEXT"
    }.items():
        if name not in existing: cursor.execute(f"ALTER TABLE audit_logs ADD COLUMN {name} {definition}")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_timestamp ON audit_logs(timestamp DESC)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_operario ON audit_logs(operario)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_logs_producto ON audit_logs(producto)")
    conn.commit()
    conn.close()

def log_event(
    operario: str,
    accion: str,
    programa: Optional[str] = None,
    cultivo: Optional[str] = None,
    sector_bloque: Optional[str] = None,
    producto: Optional[str] = None,
    codigo: Optional[str] = None,
    dosis: Optional[float] = None,
    volumen_tanque: Optional[float] = None,
    numero_tanque: Optional[int] = None,
    total_tanques: Optional[int] = None,
    copias: int = 1,
    detalles: Optional[str] = None,
    timestamp: Optional[str] = None,
    usuario_id: Optional[int] = None, rol: Optional[str] = None, ip_origen: Optional[str] = None,
    sesion_id: Optional[str] = None, resultado: str = "EXITO", valores_anteriores: Optional[str] = None,
    valores_nuevos: Optional[str] = None
) -> int:
    """Registra un evento de trazabilidad."""
    init_db()
    if not timestamp:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO audit_logs (
            timestamp, operario, accion, programa, cultivo,
            sector_bloque, producto, codigo, dosis,
            volumen_tanque, numero_tanque, total_tanques, copias, detalles, usuario_id, rol,
            ip_origen, sesion_id, resultado, valores_anteriores, valores_nuevos
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        timestamp, operario or "Operario de Mezclas", accion, programa, cultivo,
        sector_bloque, producto, codigo, dosis,
        volumen_tanque, numero_tanque, total_tanques, copias, detalles, usuario_id, rol,
        ip_origen, sesion_id, resultado, valores_anteriores, valores_nuevos
    ))
    log_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return log_id

def get_logs(limit: int = 100, offset: int = 0, search: Optional[str] = None) -> Dict[str, Any]:
    """Obtiene el historial paginado de auditoría."""
    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT * FROM audit_logs"
    params = []
    if search:
        s = f"%{search}%"
        query += " WHERE operario LIKE ? OR producto LIKE ? OR sector_bloque LIKE ? OR accion LIKE ?"
        params.extend([s, s, s, s])

    # Conteo total
    count_query = query.replace("SELECT *", "SELECT COUNT(*)")
    cursor.execute(count_query, params)
    total = cursor.fetchone()[0]

    # Resultados paginados
    query += " ORDER BY id DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return {
        "total": total,
        "offset": offset,
        "limit": limit,
        "items": rows
    }

def export_csv_data() -> str:
    """Genera CSV con el formato requerido para auditorías ICA y GlobalGAP."""
    import csv
    import io

    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM audit_logs ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    writer.writerow([
        "ID", "Fecha_Hora", "Operario_Responsable", "Accion",
        "Programa_Hoja", "Cultivo", "Sector_Bloque", "Producto_Comercial",
        "Codigo_Interno", "Dosis", "Volumen_Tanque_Litros", "Tanque_Nro",
        "Total_Tanques", "Copias_Impresas", "Detalles_Observaciones"
        , "Usuario_ID", "Rol", "IP_Origen", "Resultado", "Valores_Anteriores", "Valores_Nuevos"
    ])
    for r in rows:
        writer.writerow([
            r["id"], r["timestamp"], r["operario"], r["accion"],
            r["programa"], r["cultivo"], r["sector_bloque"], r["producto"],
            r["codigo"], r["dosis"], r["volumen_tanque"], r["numero_tanque"],
            r["total_tanques"], r["copias"], r["detalles"], r["usuario_id"], r["rol"],
            r["ip_origen"], r["resultado"], r["valores_anteriores"], r["valores_nuevos"]
        ])
    return output.getvalue()
