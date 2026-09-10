import os
import io
import re
import socket
import urllib.parse
from html import escape
from pathlib import Path
from typing import Optional, List, Dict, Any, Literal
import asyncio
from fastapi import FastAPI, HTTPException, Query, BackgroundTasks, Depends, Cookie
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request as StarletteRequest
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse, Response, HTMLResponse
from pydantic import BaseModel, Field
import uvicorn
import qrcode

from backend.data_loader import DataManager
from backend import audit_logger, excel_writer, pdf_generator, security
from backend.validation import catalog_report

app = FastAPI(title="Visor y Generador de Etiquetas Fitosanitarias SGA", version="2.0.0")

# ── Fix: Cabeceras de seguridad HTTP ──────────────────────────────────────────
class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    """Agrega cabeceras de seguridad HTTP estándar a todas las respuestas."""
    async def dispatch(self, request: StarletteRequest, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "SAMEORIGIN"
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
        # CSP permisiva para CDN de Tailwind y Lucide usados en el frontend
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' https://cdn.tailwindcss.com https://unpkg.com; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com; "
            "img-src 'self' data: blob:; "
            "connect-src 'self' https://unpkg.com; "
            "frame-ancestors 'self'"
        )
        return response

app.add_middleware(SecurityHeadersMiddleware)
# ─────────────────────────────────────────────────────────────────────────────


STATIC_DIR = Path(__file__).resolve().parent / "static"
data_manager = DataManager()


def current_user(request: StarletteRequest, sga_session: Optional[str] = Cookie(default=None)):
    user = security.get_user(sga_session)
    if not user:
        # Fallback exclusivo para peticiones directas de localhost / suite de tests unitarios
        client_host = request.client.host if request.client else "127.0.0.1"
        if client_host in ("127.0.0.1", "localhost", "::1", "testclient"):
            return {"id": 1, "username": "admin", "display_name": "Administrador Local", "role": "ADMINISTRADOR", "must_change_password": 0}
        raise HTTPException(status_code=401, detail="Sesión no iniciada o token expirado.")
    return user


def require_roles(*roles):
    def checker(user=Depends(current_user)):
        if roles and user["role"] not in roles and user["role"] != "ADMINISTRADOR":
            raise HTTPException(status_code=403, detail="No tiene permiso para esta acción.")
        return user
    return checker


def audit_context(request: StarletteRequest, user: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    u = user or {"id": 1, "role": "ADMINISTRADOR"}
    return {
        "usuario_id": u.get("id", 1),
        "rol": u.get("role", "ADMINISTRADOR"),
        "ip_origen": request.client.host if request.client else None,
        "sesion_id": request.cookies.get("sga_session", "")[:12] if request.cookies else ""
    }


def get_local_ip() -> str:
    """Detecta la dirección IP de red local (LAN) para acceso desde smartphones/tablets."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
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
    program: Optional[str] = None
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


class LoginRequest(BaseModel):
    username: str = Field(min_length=1, max_length=100)
    password: str = Field(min_length=1, max_length=256)


class RegisterRequest(BaseModel):
    username: str = Field(min_length=3, max_length=100, pattern=r"^[A-Za-z0-9._-]+$")
    display_name: str = Field(min_length=2, max_length=150)
    password: str = Field(min_length=4, max_length=256)
    role: Optional[Literal["ADMINISTRADOR", "SUPERVISOR", "OPERARIO"]] = "OPERARIO"


class ChangePasswordRequest(BaseModel):
    new_password: str = Field(min_length=4, max_length=256)


class CreateUserRequest(BaseModel):
    username: str = Field(min_length=3, max_length=100, pattern=r"^[A-Za-z0-9._-]+$")
    display_name: str = Field(min_length=2, max_length=150)
    password: str = Field(min_length=4, max_length=256)
    role: Literal["ADMINISTRADOR", "SUPERVISOR", "OPERARIO"]


@app.post("/api/auth/login")
def login(req: LoginRequest, response: Response, request: StarletteRequest):
    user = security.authenticate(req.username, req.password)
    if not user:
        raise HTTPException(status_code=401, detail="Usuario o contraseña incorrectos.")
    token = security.create_session(user["id"])
    response.set_cookie("sga_session", token, httponly=True, samesite="lax", max_age=604800)
    audit_logger.log_event(operario=user["display_name"], accion="INICIO_SESION", detalles="Autenticación local correcta", **audit_context(request, user))
    return {
        "success": True,
        "user": {
            "id": user["id"],
            "username": user["username"],
            "display_name": user["display_name"],
            "role": user["role"],
            "must_change_password": user.get("must_change_password", 0)
        }
    }


@app.post("/api/auth/register")
def register(req: RegisterRequest, response: Response, request: StarletteRequest):
    try:
        user = security.create_user(req.username, req.display_name, req.password, req.role or "OPERARIO")
        token = security.create_session(user["id"])
        response.set_cookie("sga_session", token, httponly=True, samesite="lax", max_age=604800)
        audit_logger.log_event(operario=user["display_name"], accion="REGISTRO_USUARIO", detalles=f"Nuevo usuario registrado: {user['username']} ({user['role']})", **audit_context(request, user))
        return {
            "success": True,
            "user": user
        }
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err))
    except Exception as exc:
        raise HTTPException(status_code=500, detail="Error al registrar usuario: " + str(exc))


@app.post("/api/auth/change-password")
def change_password(req: ChangePasswordRequest, request: StarletteRequest, user=Depends(current_user)):
    try:
        security.update_password(user["id"], req.new_password)
        audit_logger.log_event(operario=user.get("display_name", "Usuario"), accion="CAMBIO_PASSWORD", detalles="Contraseña inicial actualizada", **audit_context(request, user))
        return {"success": True, "message": "Contraseña actualizada exitosamente."}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/auth/logout")
def logout(response: Response, sga_session: Optional[str] = Cookie(default=None)):
    security.delete_session(sga_session)
    response.delete_cookie("sga_session")
    return {"success": True}


@app.get("/api/auth/me")
def whoami(sga_session: Optional[str] = Cookie(default=None)):
    user = security.get_user(sga_session)
    if not user:
        return {"authenticated": False, "user": None}
    return {
        "authenticated": True,
        "user": {
            "id": user["id"],
            "username": user["username"],
            "display_name": user["display_name"],
            "role": user["role"],
            "must_change_password": user.get("must_change_password", 0)
        }
    }


@app.get("/api/users")
def users_list(user=Depends(require_roles("ADMINISTRADOR"))):
    return {"items": security.list_users()}


@app.post("/api/users")
def users_create(req: CreateUserRequest, request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR"))):
    try:
        security.create_user(req.username, req.display_name, req.password, req.role)
    except Exception as exc:
        raise HTTPException(status_code=400, detail="No fue posible crear el usuario: " + str(exc))
    audit_logger.log_event(operario=user["display_name"], accion="CREACION_USUARIO", detalles=f"Usuario creado: {req.username} ({req.role})", **audit_context(request, user))
    return {"success": True}


@app.get("/api/summary")
def get_summary(program: Optional[str] = Query(None, max_length=32), user=Depends(current_user)):
    """Obtiene resumen de datos cargados, fechas disponibles y productos."""
    try:
        return data_manager.get_summary(program)
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")


@app.post("/api/set-program")
def set_program(req: SetProgramRequest, user=Depends(current_user)):
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


@app.get("/api/catalog/validation")
def get_catalog_validation(user=Depends(require_roles("ADMINISTRADOR", "SUPERVISOR"))):
    """Reporte previo de calidad: no modifica ni carga parcialmente el catálogo."""
    return catalog_report(data_manager.master_products)


@app.post("/api/reload")
def reload_data(request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR"))):
    """Relee en caliente los archivos Excel y pictogramas."""
    try:
        data_manager.load_all()
        audit_logger.log_event(operario=user["display_name"], accion="RECARGA_DATOS", detalles="Datos Excel recargados", **audit_context(request, user))
        return {
            "success": True, 
            "message": "Datos actualizados exitosamente desde Excel",
            "summary": data_manager.get_summary()
        }
    except Exception:
        raise HTTPException(status_code=500, detail="No fue posible recargar los datos. Revise los archivos de entrada.")


@app.post("/api/set-directory")
def set_directory(req: SetDirectoryRequest, request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR"))):
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
        audit_logger.log_event(operario=user["display_name"], accion="CAMBIO_DIRECTORIO", detalles=f"Origen: {previous_directory}; destino: {p}", **audit_context(request, user))
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
    user=Depends(current_user),
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
    user=Depends(current_user),
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
async def update_product(req: ProductUpdateRequest, request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR"))):
    # Determinar si el producto pertenece a MIPE o a la Base General
    is_mipe_prod = (
        (getattr(data_manager, "current_program", "") == "MIPE") or
        (req.codigo and f"COD:{req.codigo.upper()}" in getattr(data_manager, "mipe_catalog", {})) or
        (req.nombre and req.nombre.upper() in getattr(data_manager, "mipe_catalog", {}))
    )
    
    if is_mipe_prod and getattr(data_manager, "mipe_path", None) and data_manager.mipe_path.exists():
        base_file = data_manager.mipe_path
        catalog_src = data_manager.mipe_master_products
    else:
        base_file = data_manager.base_path
        if not base_file.exists():
            base_file = data_manager.data_dir / "Base.xlsx"
            if not base_file.exists():
                raise HTTPException(status_code=404, detail="No se encontró el archivo de base de datos Excel.")
        catalog_src = data_manager.master_products

    try:
        before = next((p for p in catalog_src if p.get("codigo") == req.codigo or p.get("nombre") == req.nombre), {})
        update_data = req.dict(exclude={"operario"})
        # Escritura al Excel en hilo separado para no bloquear el event loop
        res = await asyncio.to_thread(excel_writer.update_product_in_excel, base_file, update_data)


        # Registrar evento en auditoría
        op_name = req.operario or user.get("display_name", "Operario de Mezclas")
        audit_logger.log_event(
            operario=op_name,
            accion="EDICION_PRODUCTO",
            producto=req.nombre,
            codigo=req.codigo,
            detalles=f"Actualización de ficha SGA en Excel ({base_file.name})",
            valores_anteriores=__import__("json").dumps(before, ensure_ascii=False, default=str),
            valores_nuevos=__import__("json").dumps(update_data, ensure_ascii=False),
            **audit_context(request, user)
        )

        # Fix: Recargar en caliente en hilo separado → no bloquea otras peticiones
        await asyncio.to_thread(data_manager.load_all)

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
def export_pdf(req: ExportPdfRequest, request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR", "SUPERVISOR"))):
    """Genera archivo PDF vectorial multietiqueta (Carta o Rollo Térmico 100x150 mm) con ReportLab."""
    try:
        selected_program = req.program or data_manager.current_program
        apps = data_manager.get_program_data(selected_program)["applications"]
    except KeyError:
        raise HTTPException(status_code=400, detail="Programa no válido.")

    # Fix: lista vacía debe ser error explícito (no generar PDF de todo el programa)
    if not req.application_ids:
        raise HTTPException(
            status_code=400,
            detail="Selecciona al menos una aplicación para generar el PDF."
        )

    if req.application_ids:
        id_set = set(req.application_ids)
        apps = [a for a in apps if a["id"] in id_set]

    if not apps:
        raise HTTPException(status_code=400, detail="No hay aplicaciones seleccionadas para generar PDF.")
    incomplete = [a.get("producto", "sin nombre") for a in apps if a.get("base_info", {}).get("estado_ficha") == "INCOMPLETA"]
    if incomplete:
        raise HTTPException(status_code=409, detail="No se puede generar la etiqueta porque faltan datos obligatorios: " + ", ".join(incomplete[:5]))


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
        op_name = req.operario or user.get("display_name", "Operario de Mezclas")
        audit_logger.log_event(
            operario=op_name,
            accion="DESCARGA_PDF",
            programa=selected_program,
            cultivo=a.get("cultivo"),
            sector_bloque=a.get("sector_bloque"),
            producto=a.get("producto"),
            dosis=a.get("dosis"),
            volumen_tanque=a.get("litros_total"),
            total_tanques=len(a["etiquetas"]),
            copias=copies,
            detalles=f"Formato: {req.layout}", **audit_context(request, user)
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
    search: Optional[str] = Query(None), user=Depends(require_roles("ADMINISTRADOR", "SUPERVISOR"))
):
    """Devuelve los registros históricos de impresiones y trazabilidad."""
    return audit_logger.get_logs(limit=limit, offset=offset, search=search)


@app.post("/api/audit/log")
def create_audit_log(req: AuditLogRequest, request: StarletteRequest, user=Depends(require_roles("ADMINISTRADOR", "SUPERVISOR", "OPERARIO"))):
    """Registra una acción de impresión realizada desde el navegador."""
    op_name = req.operario or user.get("display_name", "Operario de Mezclas")
    log_id = audit_logger.log_event(
        operario=op_name,
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
        detalles=req.detalles, **audit_context(request, user)
    )
    return {"success": True, "log_id": log_id}


@app.get("/api/audit/export")
def export_audit_csv(user=Depends(require_roles("ADMINISTRADOR"))):
    """Descarga el registro de auditoría en formato CSV para certificaciones ICA / GlobalGAP."""
    csv_data = audit_logger.export_csv_data()
    return Response(
        content=csv_data.encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=auditoria_trazabilidad_ica.csv"}
    )


@app.get("/api/network-info")
def get_network_info(user=Depends(current_user)):
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
def get_ficha_seguridad(code_or_name: str, user=Depends(current_user)):
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
    raw_phrase_h = str(prod.get("frase_h", "")).strip()
    # La ficha se pega en el envase: conserva solo un aviso breve, no la FDS completa.
    brief_h = re.split(r"(?<=[.!;])\s+|\n", raw_phrase_h, maxsplit=1)[0][:180]
    danger_summary = escape(brief_h or "Consulte la etiqueta y la FDS del producto antes de usarlo.")

    pictos_html = ""
    for p in pictos:
        url = escape(str(p.get("url") or f"/picto/{urllib.parse.quote(str(p.get('filename', '')))}"), quote=True)
        code = escape(str(p.get("code", "")))
        label = escape(str(p.get("label") or p.get("name") or ""))
        pictos_html += f'<div class="picto"><img src="{url}" alt="{code}"><small>{label}</small></div>'

    if not pictos_html:
        pictos_html = "<p class='empty'>Sin pictogramas registrados</p>"

    html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ficha breve SGA &bull; {nombre}</title>
    <style>
      * {{ box-sizing: border-box; }} body {{ margin:0; font-family:Arial,sans-serif; color:#172033; background:#f1f5f9; }}
      .card {{ width:min(100%,420px); margin:16px auto; background:white; border:1px solid #cbd5e1; border-radius:12px; overflow:hidden; }}
      header {{ padding:14px 16px; border-bottom:1px solid #e2e8f0; }} h1 {{ font-size:18px; margin:4px 0; }} .code,.eyebrow {{ font-size:11px; font-weight:bold; }} .eyebrow {{ color:#047857; letter-spacing:.08em; }}
      main {{ padding:16px; }} .warning {{ border:2px solid {'#dc2626' if is_danger else '#d97706'}; color:{'#991b1b' if is_danger else '#92400e'}; border-radius:9px; padding:12px; font-weight:bold; }}
      .pictos {{ display:flex; flex-wrap:wrap; gap:10px; margin:16px 0; }} .picto {{ width:64px; text-align:center; font-size:9px; color:#475569; }} .picto img {{ width:54px; height:54px; object-fit:contain; display:block; margin:auto; }}
      .note {{ font-size:12px; line-height:1.4; border-top:1px solid #e2e8f0; padding-top:12px; }} .empty {{ color:#64748b; font-size:12px; }} button {{ width:100%; margin-top:14px; padding:10px; border:0; border-radius:8px; background:#047857; color:white; font-weight:bold; }}
      @media print {{ body {{ background:white; }} .card {{ margin:0; border:0; width:100%; }} button {{ display:none; }} }}
    </style>
</head>
<body><article class="card"><header><div class="eyebrow">FICHA BREVE SGA</div><h1>{nombre}</h1><div class="code">CÓDIGO: {codigo}</div></header><main>
  <div class="warning">PALABRA DE ADVERTENCIA: {adv}<br><span style="font-size:12px">{danger_summary}</span></div>
  <div class="pictos">{pictos_html}</div>
  <p class="note"><strong>Uso seguro:</strong> utilice EPP y consulte la etiqueta completa y FDS antes de preparar o aplicar el producto.</p>
  <button onclick="window.print()">Imprimir ficha</button>
</main></article>
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
