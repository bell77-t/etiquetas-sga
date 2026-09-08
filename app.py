import os
import io
import re
import socket
import urllib.parse
from html import escape
from pathlib import Path
from typing import Optional, List, Dict, Any, Literal
from fastapi import FastAPI, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, Response, HTMLResponse
from pydantic import BaseModel, Field
import uvicorn
import qrcode

from backend.data_loader import DataManager
from backend import audit_logger, excel_writer, pdf_generator

app = FastAPI(title="Visor y Generador de Etiquetas Fitosanitarias SGA", version="2.0.0")

data_manager = DataManager()
STATIC_DIR = Path(__file__).resolve().parent / "static"


def get_local_ip() -> str:
    """Detecta la dirección IP de red local (LAN) para acceso desde smartphones/tablets."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        s.connect(('10.254.254.254', 1))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


class SetDirectoryRequest(BaseModel):
    directory: str = Field(min_length=1, max_length=1024)


class SetProgramRequest(BaseModel):
    program: str = Field(min_length=1, max_length=32)


class ProductUpdateRequest(BaseModel):
    nombre: str = Field(min_length=1, max_length=200)
    codigo: Optional[str] = Field(default=None, max_length=100)
    palabra_advertencia: Optional[str] = Field(default=None, max_length=50)
    frase_h: Optional[str] = Field(default=None, max_length=5000)
    frase_p: Optional[str] = Field(default=None, max_length=5000)
    um: Optional[str] = Field(default=None, max_length=30)
    pictogramas: List[str] = Field(default_factory=list, max_length=4)
    operario: Optional[str] = Field(default="Usuario Web", max_length=100)


class ExportPdfRequest(BaseModel):
    application_ids: List[str] = Field(default_factory=list, max_length=500)
    program: Optional[str] = Field(default=None, max_length=32)
    tanks_mode: Optional[str] = Field(default="all", max_length=10)
    copies: int = Field(default=1, ge=1, le=500)
    layout: Literal["letter", "thermal"] = "letter"
    operario: Optional[str] = Field(default="Operario de Mezclas", max_length=100)


class AuditLogRequest(BaseModel):
    operario: str = Field(min_length=1, max_length=100)
    accion: str = Field(min_length=1, max_length=100)
    programa: Optional[str] = None
    cultivo: Optional[str] = None
    sector_bloque: Optional[str] = None
    producto: Optional[str] = None
    codigo: Optional[str] = None
    dosis: Optional[float] = None
    volumen_tanque: Optional[float] = None
    numero_tanque: Optional[int] = None
    total_tanques: Optional[int] = None
    copias: int = Field(default=1, ge=1, le=500)
    detalles: Optional[str] = Field(default=None, max_length=5000)


@app.get("/api/summary")
def get_summary(program: Optional[str] = Query(None, max_length=32)):
    """Obtiene resumen de datos cargados, fechas disponibles y productos."""
    try:
        return data_manager.get_summary(program)
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")


@app.post("/api/set-program")
def set_program(req: SetProgramRequest):
    """Cambia el programa de cultivo activo (Data, ALZ, R-S-L)."""
    if req.program not in data_manager.programs:
        raise HTTPException(
            status_code=400, 
            detail=f"Programa '{req.program}' no válido. Disponibles: {list(data_manager.programs.keys())}"
        )
    program_data = data_manager.get_program_data(req.program)
    return {
        "success": True,
        "current_program": req.program,
        "summary": data_manager.get_summary(req.program),
        "applications": program_data["applications"]
    }


@app.post("/api/reload")
def reload_data():
    """Relee en caliente los archivos Excel y pictogramas."""
    try:
        data_manager.load_all()
        return {
            "success": True, 
            "message": "Datos actualizados exitosamente desde Excel",
            "summary": data_manager.get_summary()
        }
    except Exception:
        raise HTTPException(status_code=500, detail="No fue posible recargar los datos. Revise los archivos de entrada.")


@app.post("/api/set-directory")
def set_directory(req: SetDirectoryRequest):
    """Permite cambiar la carpeta de origen de los archivos Excel."""
    p = Path(req.directory).resolve()
    if not p.exists() or not p.is_dir():
        raise HTTPException(status_code=400, detail="El directorio especificado no existe.")
    if not (p / "aplicacion.xlsm").is_file():
        raise HTTPException(status_code=400, detail="El directorio debe contener aplicacion.xlsm.")
    
    previous_directory = data_manager.data_dir
    try:
        data_manager.data_dir = p
        data_manager.load_all()
        return {
            "success": True, 
            "message": f"Directorio actualizado a: {p}",
            "summary": data_manager.get_summary()
        }
    except Exception:
        data_manager.data_dir = previous_directory
        data_manager.load_all()
        raise HTTPException(status_code=500, detail="No fue posible leer el directorio indicado.")


@app.get("/api/applications")
def get_applications(
    program: Optional[str] = Query(None, description="Filtrar o activar programa (Data, ALZ, R-S-L)"),
    fecha: Optional[str] = Query(None, description="Filtrar por fecha ISO YYYY-MM-DD"),
    producto: Optional[str] = Query(None, description="Filtrar por nombre de producto"),
    search: Optional[str] = Query(None, description="Búsqueda de texto libre"),
):
    """Devuelve las aplicaciones fitosanitarias filtradas con sus etiquetas generadas."""
    try:
        results = data_manager.get_program_data(program)["applications"]
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")

    if fecha:
        results = [a for a in results if a["fecha"]["iso"] == fecha]

    if producto:
        prod_upper = producto.strip().upper()
        results = [a for a in results if prod_upper in a["producto"].upper()]

    if search:
        s_upper = search.strip().upper()
        results = [
            a for a in results
            if s_upper in a["producto"].upper() 
            or s_upper in a["sector_bloque"].upper()
            or s_upper in str(a.get("observaciones", "")).upper()
            or s_upper in str(a.get("reentrada", "")).upper()
        ]

    return {
        "count": len(results),
        "total_labels": sum(len(a["etiquetas"]) for a in results),
        "items": results
    }


@app.get("/api/labels")
def get_labels(
    program: Optional[str] = Query(None, description="Filtrar o activar programa (Data, ALZ, R-S-L)"),
    fecha: Optional[str] = Query(None, description="Filtrar por fecha ISO YYYY-MM-DD"),
    producto: Optional[str] = Query(None, description="Filtrar por nombre de producto"),
    app_id: Optional[str] = Query(None, description="Filtrar por ID de aplicación específica"),
    search: Optional[str] = Query(None, description="Búsqueda de texto"),
    limit: Optional[int] = Query(None, ge=1, le=1000, description="Límite de etiquetas para paginación"),
    offset: int = Query(0, description="Offset de inicio"),
):
    """Devuelve la lista plana de etiquetas desglosadas por tanque para vista previa o impresión."""
    try:
        apps = data_manager.get_program_data(program)["applications"]
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")

    if app_id:
        apps = [a for a in apps if a["id"] == app_id]
    if fecha:
        apps = [a for a in apps if a["fecha"]["iso"] == fecha]
    if producto:
        prod_upper = producto.strip().upper()
        apps = [a for a in apps if prod_upper in a["producto"].upper()]
    if search:
        s_upper = search.strip().upper()
        apps = [
            a for a in apps
            if s_upper in a["producto"].upper() 
            or s_upper in a["sector_bloque"].upper()
        ]

    all_labels = []
    for a in apps:
        all_labels.extend(a["etiquetas"])

    total_count = len(all_labels)
    if limit is not None:
        paged_labels = all_labels[offset : offset + limit]
    else:
        paged_labels = all_labels[offset:]

    return {
        "total": total_count,
        "offset": offset,
        "limit": limit,
        "returned": len(paged_labels),
        "labels": paged_labels
    }


def find_pictogram_file(filename: str) -> Optional[Path]:
    """Busca el archivo de pictograma en las rutas de Picto."""
    unquoted = urllib.parse.unquote(filename).strip()
    
    candidate_dirs = [data_manager.picto_dir, STATIC_DIR / "picto"]

    for pdir in candidate_dirs:
        if not pdir.exists():
            continue

        resolved_dir = pdir.resolve()

        def is_file_inside_pictogram_dir(candidate: Path) -> bool:
            """Evita que un parámetro de ruta acceda a archivos fuera de ``pdir``."""
            try:
                candidate.resolve().relative_to(resolved_dir)
                return candidate.is_file()
            except (ValueError, OSError):
                return False
        
        # 1. Coincidencia directa
        exact_path = pdir / unquoted
        if is_file_inside_pictogram_dir(exact_path):
            return exact_path
        
        if not unquoted.lower().endswith(".jpg"):
            with_ext = pdir / f"{unquoted}.jpg"
            if is_file_inside_pictogram_dir(with_ext):
                return with_ext

        # 2. Código GHS (ej: GHS01 a GHS09)
        match = re.search(r"(GHS\d{2})", unquoted.upper())
        if match:
            ghs_code = match.group(1)
            ghs_file = pdir / f"{ghs_code}.jpg"
            if ghs_file.exists() and ghs_file.is_file():
                return ghs_file
            
            for f in pdir.glob("*.jpg"):
                if ghs_code in f.name.upper():
                    return f

        # 3. Búsqueda por subcadena
        norm_req = unquoted.upper().replace(".JPG", "").strip()
        for f in pdir.glob("*.jpg"):
            if norm_req in f.name.upper():
                return f

    return None


@app.get("/picto/{filename:path}")
@app.get("/static/picto/{filename:path}")
@app.get("/api/picto/{filename:path}")
def get_pictogram_image(filename: str):
    """Sirve las imágenes de pictogramas desde la carpeta Picto/."""
    img_path = find_pictogram_file(filename)
    if img_path and img_path.exists() and img_path.is_file():
        return FileResponse(
            img_path, 
            media_type="image/jpeg", 
            headers={"Cache-Control": "public, max-age=86400"}
        )
    raise HTTPException(status_code=404, detail="Imagen de pictograma no encontrada")


@app.post("/api/products/update")
def update_product(req: ProductUpdateRequest):
    """Permite editar y asignar pictogramas, frases H/P o advertencia guardando en Base_Actualizada.xlsx."""
    base_file = data_manager.base_path
    if not base_file.exists():
        # Fallback a Base.xlsx
        base_file = data_manager.data_dir / "Base.xlsx"
        if not base_file.exists():
            raise HTTPException(status_code=404, detail="No se encontró el archivo de base de datos Excel.")

    try:
        update_data = req.dict()
        res = excel_writer.update_product_in_excel(base_file, update_data)
        
        # Registrar evento en auditoría
        audit_logger.log_event(
            operario=req.operario or "Usuario Web",
            accion="EDICION_PRODUCTO",
            producto=req.nombre,
            codigo=req.codigo,
            detalles=f"Actualización de ficha SGA en Excel ({base_file.name})"
        )

        # Recargar en caliente
        data_manager.load_all()

        return {
            "success": True,
            "message": res["message"],
            "summary": data_manager.get_summary()
        }
    except (ValueError, FileNotFoundError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception:
        raise HTTPException(status_code=500, detail="No fue posible actualizar el catálogo.")


@app.post("/api/export-pdf")
def export_pdf(req: ExportPdfRequest):
    """Genera archivo PDF vectorial multietiqueta (Carta o Rollo Térmico 100x150 mm) con ReportLab."""
    try:
        selected_program = req.program or data_manager.current_program
        apps = data_manager.get_program_data(selected_program)["applications"]
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")
    if req.application_ids:
        id_set = set(req.application_ids)
        apps = [a for a in apps if a["id"] in id_set]

    if not apps:
        raise HTTPException(status_code=400, detail="No hay aplicaciones seleccionadas para generar PDF.")

    limit = None
    if req.tanks_mode and req.tanks_mode != "all":
        try:
            limit = int(req.tanks_mode)
        except ValueError:
            limit = None

    copies = max(1, req.copies)
    all_labels = []
    
    local_ip = get_local_ip()
    qr_base_url = f"http://{local_ip}:8000"

    for a in apps:
        tanks = a["etiquetas"][:limit] if limit else a["etiquetas"]
        for t in tanks:
            for _ in range(copies):
                all_labels.append(t)
                
        # Registrar en auditoría
        audit_logger.log_event(
            operario=req.operario or "Operario de Mezclas",
            accion="DESCARGA_PDF",
            programa=selected_program,
            cultivo=a.get("cultivo"),
            sector_bloque=a.get("sector_bloque"),
            producto=a.get("producto"),
            dosis=a.get("dosis"),
            volumen_tanque=a.get("litros_total"),
            total_tanques=len(a["etiquetas"]),
            copias=copies,
            detalles=f"Formato: {req.layout}"
        )

    try:
        pdf_buf = pdf_generator.generate_pdf(
            all_labels,
            data_manager.picto_dir,
            qr_base_url=qr_base_url,
            layout=req.layout
        )
        filename = f"etiquetas_{selected_program}_{req.layout}.pdf"
        return Response(
            content=pdf_buf.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except Exception:
        raise HTTPException(status_code=500, detail="No fue posible generar el PDF.")


@app.get("/api/audit/logs")
def get_audit_logs(
    limit: int = Query(50, ge=1, le=500),
    offset: int = Query(0, ge=0),
    search: Optional[str] = Query(None)
):
    """Devuelve los registros históricos de impresiones y trazabilidad."""
    return audit_logger.get_logs(limit=limit, offset=offset, search=search)


@app.post("/api/audit/log")
def create_audit_log(req: AuditLogRequest):
    """Registra una acción de impresión realizada desde el navegador."""
    log_id = audit_logger.log_event(
        operario=req.operario,
        accion=req.accion,
        programa=req.programa,
        cultivo=req.cultivo,
        sector_bloque=req.sector_bloque,
        producto=req.producto,
        codigo=req.codigo,
        dosis=req.dosis,
        volumen_tanque=req.volumen_tanque,
        numero_tanque=req.numero_tanque,
        total_tanques=req.total_tanques,
        copias=req.copias,
        detalles=req.detalles
    )
    return {"success": True, "log_id": log_id}


@app.get("/api/audit/export")
def export_audit_csv():
    """Descarga el registro de auditoría en formato CSV para certificaciones ICA / GlobalGAP."""
    csv_data = audit_logger.export_csv_data()
    return Response(
        content=csv_data.encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=auditoria_trazabilidad_ica.csv"}
    )


@app.get("/api/network-info")
def get_network_info():
    """Obtiene la dirección IP de red local para conectar tablets o celulares en caseta."""
    ip = get_local_ip()
    return {
        "local_ip": ip,
        "port": 8000,
        "network_url": f"http://{ip}:8000",
        "qr_url": f"/api/qr_connect"
    }


@app.get("/api/qr_connect")
def get_qr_connect():
    """Genera código QR para conectar smartphones o tablets en la misma red WiFi."""
    ip = get_local_ip()
    url = f"http://{ip}:8000"
    qr = qrcode.QRCode(box_size=5, border=1)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return Response(content=buf.getvalue(), media_type="image/png")


@app.get("/api/qr/{code_or_name}")
def get_product_qr(code_or_name: str):
    """Genera código QR que apunta a la Ficha de Seguridad Digital de ese producto."""
    ip = get_local_ip()
    unquoted = urllib.parse.unquote(code_or_name).strip()
    url = f"http://{ip}:8000/ficha/{urllib.parse.quote(unquoted)}"
    qr = qrcode.QRCode(box_size=3, border=1)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return Response(content=buf.getvalue(), media_type="image/png")


@app.get("/ficha/{code_or_name}", response_class=HTMLResponse)
def get_ficha_seguridad(code_or_name: str):
    """Página web móvil para consulta rápida de seguridad y primeros auxilios escaneando el código QR."""
    unquoted = urllib.parse.unquote(code_or_name).strip().upper()
    
    # Buscar producto en catálogo base
    prod = None
    if f"COD:{unquoted}" in data_manager.base_catalog:
        prod = data_manager.base_catalog[f"COD:{unquoted}"]
    elif unquoted in data_manager.base_catalog:
        prod = data_manager.base_catalog[unquoted]
    else:
        for k, p in data_manager.base_catalog.items():
            if not k.startswith("COD:"):
                if unquoted in p["nombre"].upper() or (p.get("codigo") and unquoted in p["codigo"].upper()):
                    prod = p
                    break
                    
    if not prod:
        prod = {
            "nombre": unquoted,
            "codigo": "N/A",
            "palabra_advertencia": "ATENCIÓN",
            "frase_h": "No se encontró registro de frases H en la base de datos.",
            "frase_p": "P102 Manténgase fuera del alcance de los niños.\nP270 No comer, beber ni fumar durante su utilización.",
            "pictogramas": []
        }

    nombre = escape(str(prod["nombre"]))
    codigo = escape(str(prod.get("codigo", "N/A")))
    adv = escape(str(prod.get("palabra_advertencia", "PELIGRO")).upper())
    is_danger = adv == "PELIGRO"
    pictos = [p for p in prod.get("pictogramas", []) if p.get("has_image")]
    frase_h = escape(str(prod.get("frase_h", "Sin frases H.")))
    frase_p = escape(str(prod.get("frase_p", "Sin frases P.")))

    pictos_html = ""
    for p in pictos:
        url = escape(str(p.get("url") or f"/picto/{urllib.parse.quote(str(p.get('filename', '')))}"), quote=True)
        code = escape(str(p.get("code", "")))
        label = escape(str(p.get("label") or p.get("name") or ""))
        pictos_html += f"""
        <div class="flex flex-col items-center bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
            <img src="{url}" alt="{code}" class="w-16 h-16 object-contain mb-1">
            <span class="text-xs font-bold text-slate-800">{code}</span>
            <span class="text-[10px] text-slate-500 text-center">{label}</span>
        </div>
        """

    if not pictos_html:
        pictos_html = "<div class='col-span-2 text-center text-slate-400 py-3 text-sm'>Sin pictogramas de peligro reportados</div>"

    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ficha FDS &bull; {nombre}</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800 font-sans pb-10">
    <div class="max-w-md mx-auto bg-white shadow-md border-b border-slate-200 p-4 sticky top-0 z-10 flex items-center justify-between">
        <div>
            <span class="text-[10px] font-bold tracking-widest uppercase text-emerald-700 bg-emerald-100 px-2 py-0.5 rounded-full">SGA / FDS Rápida</span>
            <h1 class="text-base font-bold text-slate-900 leading-tight mt-1">{nombre}</h1>
        </div>
        <span class="text-xs font-mono font-bold bg-slate-100 text-slate-700 px-2.5 py-1 rounded border border-slate-200">{codigo}</span>
    </div>

    <main class="max-w-md mx-auto p-4 space-y-4">
        <div class="flex items-center justify-between p-3.5 rounded-xl border {'bg-red-50 border-red-200 text-red-900' if is_danger else 'bg-amber-50 border-amber-200 text-amber-900'}">
            <div class="flex items-center gap-2">
                <span class="text-2xl">{'⚠️' if is_danger else '⚡'}</span>
                <div>
                    <span class="text-[11px] uppercase tracking-wide block font-semibold">Palabra de Advertencia</span>
                    <strong class="text-base tracking-wide font-black">{adv}</strong>
                </div>
            </div>
            <span class="text-xs font-semibold px-2 py-1 bg-white/70 rounded-lg">Norma SGA</span>
        </div>

        <section class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <h2 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Pictogramas de Peligro GHS</h2>
            <div class="grid grid-cols-2 gap-2.5">
                {pictos_html}
            </div>
        </section>

        <section class="bg-rose-50 border border-rose-200 p-4 rounded-xl shadow-sm text-rose-950">
            <h2 class="text-xs font-bold uppercase tracking-wider text-rose-800 flex items-center gap-1.5 mb-2">
                <span>🚨</span> Teléfonos de Emergencia Toxicológica 24 Horas
            </h2>
            <div class="space-y-2 text-xs">
                <div class="bg-white/80 p-2.5 rounded-lg border border-rose-200">
                    <span class="text-[11px] text-rose-700 font-semibold block">CISPROQUIM (Colombia - Línea Nacional Gratuita 24/7):</span>
                    <a href="tel:018000916012" class="text-sm font-bold text-rose-900 underline">📞 01 8000 916012</a> / <a href="tel:6012886012" class="text-sm font-bold text-rose-900 underline">(601) 288 6012</a>
                </div>
                <div class="grid grid-cols-2 gap-2 text-center pt-1">
                    <a href="tel:119" class="bg-white/80 p-2 rounded-lg border border-rose-200 font-bold text-rose-900">🚒 Bomberos: 119</a>
                    <a href="tel:132" class="bg-white/80 p-2 rounded-lg border border-rose-200 font-bold text-rose-900">🏥 Cruz Roja: 132</a>
                </div>
            </div>
        </section>

        <section class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
            <h2 class="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                <span>🩺</span> Guía Rápida de Primeros Auxilios
            </h2>
            <div class="text-xs space-y-2 text-slate-700">
                <p><strong>Inhalación:</strong> Trasladar inmediatamente a la víctima al aire libre. Mantenerla abrigada y en reposo.</p>
                <p><strong>Contacto Piel:</strong> Quitar ropa contaminada. Lavar la piel afectada con abundante agua y jabón mínimo 15 minutos.</p>
                <p><strong>Contacto Ojos:</strong> Enjuagar con abundante agua corriente durante 15 minutos manteniendo párpados abiertos.</p>
                <p><strong>Ingestión:</strong> <u>NO provocar el vómito</u> a menos que lo indique personal médico especializado.</p>
            </div>
        </section>

        <section class="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
            <div>
                <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">Frases de Peligro (Frase H)</h3>
                <p class="text-xs font-medium text-slate-800 bg-slate-50 p-2.5 rounded-lg border border-slate-200 whitespace-pre-line">{frase_h}</p>
            </div>
            <div>
                <h3 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">Consejos de Prudencia (Frase P)</h3>
                <p class="text-xs font-medium text-slate-800 bg-slate-50 p-2.5 rounded-lg border border-slate-200 whitespace-pre-line">{frase_p}</p>
            </div>
        </section>

        <div class="text-center pt-2">
            <a href="/" class="inline-flex items-center gap-2 bg-emerald-700 text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-md hover:bg-emerald-800 transition">
                <span>&larr; Volver a SGA Label Studio</span>
            </a>
        </div>
    </main>
</body>
</html>
"""
    return HTMLResponse(content=html_content)


@app.get("/manifest.json")
def get_manifest():
    """Manifest para instalación PWA en teléfonos y tablets."""
    return {
        "name": "SGA Label Studio",
        "short_name": "SGA Studio",
        "description": "Visor y Generador de Etiquetas Fitosanitarias SGA",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#0f172a",
        "theme_color": "#059669",
        "icons": [
            {
                "src": "/static/picto/GHS07.jpg",
                "sizes": "192x192",
                "type": "image/jpeg"
            }
        ]
    }


# Montar archivos estáticos
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

if data_manager.picto_dir.exists():
    app.mount("/picto_static", StaticFiles(directory=str(data_manager.picto_dir)), name="picto_static")


@app.get("/")
def serve_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return {"message": "Visor de Etiquetas SGA - Frontend no cargado"}


if __name__ == "__main__":
    ip = get_local_ip()
    print(f"Iniciando servidor en red local:")
    print(f"  - Localhost:    http://127.0.0.1:8000")
    print(f"  - Red Local IP: http://{ip}:8000")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
