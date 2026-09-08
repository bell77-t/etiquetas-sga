"""
ghs_engine.py - Motor de Inferencia Inteligente de Pictogramas SGA / GHS
Analiza frases H (Indicaciones de Peligro) y asigna automáticamente los pictogramas correspondientes.
"""

import re
from typing import List, Dict, Any, Optional

# Mapeo oficial SGA / GHS de códigos H a pictogramas
GHS_H_MAPPING = {
    # GHS01: Explosivo
    "GHS01": {
        "name": "Explosivo",
        "filename": "GHS01.jpg",
        "codes": ["H200", "H201", "H202", "H203", "H204", "H205", "H240", "H241"],
        "keywords": ["explosivo", "inestable", "detonacion"]
    },
    # GHS02: Inflamable
    "GHS02": {
        "name": "Inflamable",
        "filename": "GHS02.jpg",
        "codes": ["H220", "H221", "H222", "H223", "H224", "H225", "H226", "H228", "H242", "H250", "H251", "H252", "H260", "H261"],
        "keywords": ["inflamable", "combustion", "pirofórico", "fuego"]
    },
    # GHS03: Comburente
    "GHS03": {
        "name": "Comburente",
        "filename": "GHS03.jpg",
        "codes": ["H270", "H271", "H272"],
        "keywords": ["comburente", "oxigeno", "agravar un incendio"]
    },
    # GHS04: Gas a Presión
    "GHS04": {
        "name": "Gas a Presión",
        "filename": "GHS04.jpg",
        "codes": ["H280", "H281"],
        "keywords": ["gas a presion", "gas comprimido", "gas licuado"]
    },
    # GHS05: Corrosión (Ojos, Piel, Metales)
    "GHS05": {
        "name": "Corrosión",
        "filename": "GHS05.jpg",
        "codes": ["H290", "H314", "H318"],
        "keywords": ["corrosivo", "lesiones oculares graves", "quemaduras graves", "daño ocular", "corroe"]
    },
    # GHS06: Toxicidad Aguda Severa (Calavera)
    "GHS06": {
        "name": "Toxicidad Aguda Severa",
        "filename": "GHS06.jpg",
        "codes": ["H300", "H301", "H310", "H311", "H330", "H331"],
        "keywords": ["mortal en caso de ingestion", "mortal por inhalacion", "mortal en contacto", "toxico en caso de ingestion", "toxico por inhalacion"]
    },
    # GHS07: Toxicidad Leve / Irritación / Sensibilización
    "GHS07": {
        "name": "Atención / Irritante",
        "filename": "GHS07.jpg",
        "codes": ["H302", "H312", "H332", "H315", "H317", "H319", "H335", "H336"],
        "keywords": ["nocivo en caso de ingestion", "nocivo en contacto", "nocivo por inhalacion", "irritacion cutanea", "irritacion ocular", "alergica en la piel", "somnolencia", "vertigo"]
    },
    # GHS08: Peligro para la Salud Crónico (Carcinógeno, Mutágeno, Toxicidad Órganos)
    "GHS08": {
        "name": "Peligro Crónico a la Salud",
        "filename": "GHS08.jpg",
        "codes": ["H304", "H334", "H340", "H341", "H350", "H351", "H360", "H361", "H370", "H371", "H372", "H373"],
        "keywords": ["cancér", "cancer", "mutacion", "fertilidad", "feto", "dificultades respiratorias", "asma", "órganos", "organos"]
    },
    # GHS09: Medio Ambiente Acuático
    "GHS09": {
        "name": "Medio Ambiente",
        "filename": "GHS09.jpg",
        "codes": ["H400", "H410", "H411", "H412", "H413"],
        "keywords": ["acuatico", "acuático", "vida acuatica", "medio ambiente", "organismos acuaticos"]
    }
}

def normalize_text_simple(text: str) -> str:
    """Normaliza texto para comparación eliminando tildes y mayúsculas."""
    if not text:
        return ""
    import unicodedata
    nfkd = unicodedata.normalize('NFKD', str(text))
    cleaned = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return cleaned.lower()

def infer_ghs_pictograms(frase_h: str, palabra_advertencia: str = "", max_pictos: int = 4) -> List[Dict[str, Any]]:
    """
    Analiza Frase H y opcionalmente Palabra de Advertencia para deducir pictogramas GHS.
    Retorna una lista de diccionarios compatibles con la estructura de producto en data_loader.
    """
    if not frase_h or not isinstance(frase_h, str):
        return []

    norm_h = normalize_text_simple(frase_h)
    
    # Extraer todos los códigos tipo Hxxx (ej: H318, H410, H302)
    found_codes = set(re.findall(r"\b(h\d{3}[a-z]?)\b", norm_h, re.IGNORECASE))
    found_codes_upper = {c.upper() for c in found_codes}

    detected_pictos = []
    
    # 1. Reglas de precedencia y detección
    for ghs_id, info in GHS_H_MAPPING.items():
        matched = False
        reason = ""
        
        # Coincidencia por código directo
        common_codes = found_codes_upper.intersection(set(info["codes"]))
        if common_codes:
            matched = True
            reason = f"Códigos: {', '.join(sorted(common_codes))}"
        else:
            # Coincidencia por palabra clave relevante
            for kw in info["keywords"]:
                if kw in norm_h:
                    matched = True
                    reason = f"Frase: '{kw}'"
                    break
                    
        if matched:
            detected_pictos.append({
                "code": ghs_id,
                "name": info["name"],
                "filename": info["filename"],
                "raw_text": f"{ghs_id} ({info['name']})",
                "has_image": True,
                "inferred": True,
                "reason": reason
            })

    # 2. Regla de compatibilidad SGA oficial:
    # Si GHS06 (Calavera) está presente, no debe aparecer GHS07 (Exclamación) para toxicidad aguda.
    has_ghs06 = any(p["code"] == "GHS06" for p in detected_pictos)
    if has_ghs06:
        detected_pictos = [p for p in detected_pictos if p["code"] != "GHS07" or "H315" in norm_h or "H319" in norm_h]

    # Limitar al máximo estándar de pictogramas en etiqueta (usualmente máx 4)
    return detected_pictos[:max_pictos]
