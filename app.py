import os
import re
import urllib.parse
from pathlib import Path
from typing import Optional, List
from fastapi import FastAPI, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

from data_loader import DataManager

app = FastAPI(title="Visor y Generador de Etiquetas Fitosanitarias SGA", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

data_manager = DataManager()
STATIC_DIR = Path(__file__).resolve().parent / "static"

# Localizar directorio Picto
PICTO_DIR = data_manager.data_dir / "Picto"
if not PICTO_DIR.exists():
    PICTO_DIR = Path(__file__).resolve().parent / "Picto"
if not PICTO_DIR.exists():
    PICTO_DIR = Path(r"C:\Users\alexc\Downloads\Alejandra\Alejandra\Picto")


class SetDirectoryRequest(BaseModel):
    directory: str


@app.get("/api/summary")
def get_summary():
    """Obtiene resumen de datos cargados, fechas disponibles y productos."""
    return data_manager.get_summary()


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
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al recargar datos: {str(e)}")


@app.post("/api/set-directory")
def set_directory(req: SetDirectoryRequest):
    """Permite cambiar la carpeta de origen de los archivos Excel."""
    p = Path(req.directory)
    if not p.exists() or not p.is_dir():
        raise HTTPException(status_code=400, detail="El directorio especificado no existe.")
    
    try:
        data_manager.data_dir = p
        data_manager.load_all()
        return {
            "success": True, 
            "message": f"Directorio actualizado a: {p}",
            "summary": data_manager.get_summary()
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al leer desde el nuevo directorio: {str(e)}")


@app.get("/api/applications")
def get_applications(
    fecha: Optional[str] = Query(None, description="Filtrar por fecha ISO YYYY-MM-DD"),
    producto: Optional[str] = Query(None, description="Filtrar por nombre de producto"),
    search: Optional[str] = Query(None, description="Búsqueda de texto libre"),
):
    """Devuelve las aplicaciones fitosanitarias filtradas con sus etiquetas generadas."""
    results = data_manager.applications

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
    fecha: Optional[str] = Query(None, description="Filtrar por fecha ISO YYYY-MM-DD"),
    producto: Optional[str] = Query(None, description="Filtrar por nombre de producto"),
    app_id: Optional[str] = Query(None, description="Filtrar por ID de aplicación específica"),
    search: Optional[str] = Query(None, description="Búsqueda de texto"),
    limit: Optional[int] = Query(None, description="Límite de etiquetas para paginación"),
    offset: int = Query(0, description="Offset de inicio"),
):
    """Devuelve la lista plana de etiquetas desglosadas por tanque para vista previa o impresión."""
    apps = data_manager.applications

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
    
    candidate_dirs = [
        data_manager.data_dir / "Picto",
        Path(__file__).resolve().parent / "static" / "picto",
        Path(__file__).resolve().parent / "Picto",
        Path(r"C:\Users\alexc\Downloads\Alejandra\Alejandra\Picto"),
    ]

    for pdir in candidate_dirs:
        if not pdir.exists():
            continue
        
        # 1. Coincidencia directa
        exact_path = pdir / unquoted
        if exact_path.exists() and exact_path.is_file():
            return exact_path
        
        if not unquoted.lower().endswith(".jpg"):
            with_ext = pdir / f"{unquoted}.jpg"
            if with_ext.exists() and with_ext.is_file():
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


# Montar archivos estáticos
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

if PICTO_DIR.exists():
    app.mount("/picto_static", StaticFiles(directory=str(PICTO_DIR)), name="picto_static")


@app.get("/")
def serve_index():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return FileResponse(index_file)
    return {"message": "Visor de Etiquetas SGA - Frontend no cargado"}


if __name__ == "__main__":
    print("Iniciando servidor en http://localhost:8000 ...")
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=False)
