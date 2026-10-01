"""Parámetros predeterminados del motor de vouching (modificables por proyecto)."""
from __future__ import annotations

import copy

DEFAULT_PARAMETERS: dict = {
    # Regla de valor: DIFERENCIA_PORCENTAJE <= TOLERANCIA_VALOR
    "TOLERANCIA_VALOR": 0.10,
    # Política cuando el valor esperado es cero: REVISION_MANUAL | EXACTA
    "POLITICA_VALOR_CERO": "REVISION_MANUAL",
    # Similitud mínima de nombres (0-100) tras normalizar
    "UMBRAL_NOMBRE": 85,
    # Días de diferencia aceptados entre fecha esperada y extraída
    "TOLERANCIA_DIAS": 5,
    # Puntaje mínimo (0-1) para relacionar un documento con una partida
    "UMBRAL_RELACION": 0.60,
    # Confianza OCR mínima (0-100) para aceptar automáticamente un campo
    "UMBRAL_CONFIANZA_OCR": 60,
    # Máximo de soportes a combinar para alcanzar el valor contabilizado
    "MAX_SOPORTES_SUMA": 4,
    "EXIGIR_FECHA_EN_TOLERANCIA": True,
    "MONEDA_BASE": "COP",
    # Pesos por criterio (prioridad). Solo se promedian los criterios evaluables.
    "PESOS": {
        "nit": 30,
        "numero": 30,
        "valor": 20,
        "contrato": 15,
        "orden_compra": 15,
        "fecha": 12,
        "nombre": 10,
        "archivo": 25,
        "tipo": 5,
        "concepto": 4,
        "moneda": 3,
        "subtotal": 3,
        "iva": 3,
        "retenciones": 2,
    },
}

PRIORITY_LABELS = {
    "nit": "Muy alta",
    "numero": "Muy alta",
    "valor": "Alta",
    "contrato": "Alta",
    "orden_compra": "Alta",
    "fecha": "Media-alta",
    "nombre": "Media",
    "archivo": "Alta (referencia explícita)",
    "tipo": "Complementaria",
    "concepto": "Complementaria",
    "moneda": "Complementaria",
    "subtotal": "Complementaria",
    "iva": "Complementaria",
    "retenciones": "Complementaria",
}

CRITERION_LABELS = {
    "nit": "NIT / identificación",
    "numero": "Número documental",
    "valor": "Valor",
    "contrato": "Contrato",
    "orden_compra": "Orden de compra",
    "fecha": "Fecha",
    "nombre": "Nombre del tercero",
    "archivo": "Archivo referenciado",
    "tipo": "Tipo documental",
    "concepto": "Concepto / descripción",
    "moneda": "Moneda",
    "subtotal": "Subtotal",
    "iva": "IVA",
    "retenciones": "Retenciones",
}


def merged_parameters(overrides: dict | None) -> dict:
    """Combina los parámetros del proyecto con los predeterminados."""
    params = copy.deepcopy(DEFAULT_PARAMETERS)
    if not overrides:
        return params
    for k, v in overrides.items():
        if k == "PESOS" and isinstance(v, dict):
            params["PESOS"].update({kk: float(vv) for kk, vv in v.items() if kk in params["PESOS"]})
        elif k in params:
            params[k] = v
    return params


def validate_parameters(params: dict) -> list[str]:
    errors = []
    tol = params.get("TOLERANCIA_VALOR")
    if not isinstance(tol, (int, float)) or not 0 <= float(tol) <= 1:
        errors.append("TOLERANCIA_VALOR debe estar entre 0 y 1 (ej. 0.10 = 10 %)")
    un = params.get("UMBRAL_NOMBRE")
    if not isinstance(un, (int, float)) or not 0 <= float(un) <= 100:
        errors.append("UMBRAL_NOMBRE debe estar entre 0 y 100")
    ur = params.get("UMBRAL_RELACION")
    if not isinstance(ur, (int, float)) or not 0 <= float(ur) <= 1:
        errors.append("UMBRAL_RELACION debe estar entre 0 y 1")
    if params.get("POLITICA_VALOR_CERO") not in ("REVISION_MANUAL", "EXACTA"):
        errors.append("POLITICA_VALOR_CERO debe ser REVISION_MANUAL o EXACTA")
    for k, v in (params.get("PESOS") or {}).items():
        if not isinstance(v, (int, float)) or v < 0:
            errors.append(f"Peso inválido para {k}")
    return errors
