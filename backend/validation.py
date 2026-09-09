"""Reglas explícitas de calidad para fichas SGA y archivos cargados."""
from typing import Any, Dict, List

REQUIRED_PRODUCT_FIELDS = ("codigo", "nombre", "palabra_advertencia", "frase_h", "frase_p")

def validate_product(product: Dict[str, Any]) -> List[str]:
    missing = [field for field in REQUIRED_PRODUCT_FIELDS if not str(product.get(field) or "").strip()]
    pictos = product.get("pictogramas") or []
    if not any(p.get("has_image") and not p.get("inferred", False) for p in pictos if isinstance(p, dict)):
        missing.append("pictogramas_validados")
    return missing

def catalog_report(products: List[Dict[str, Any]]) -> Dict[str, Any]:
    invalid = [{"codigo": p.get("codigo"), "nombre": p.get("nombre"), "faltantes": validate_product(p)} for p in products if validate_product(p)]
    return {"validos": len(products)-len(invalid), "requieren_atencion": len(invalid), "registros": invalid}
