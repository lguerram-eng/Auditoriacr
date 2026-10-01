"""Clasificación documental explicable por palabras clave ponderadas.

Las palabras clave predeterminadas pueden ampliarse por proyecto con la hoja
TIPOS_DOCUMENTO del Excel. Si la IA está habilitada para el proyecto, su
clasificación se usa solo cuando las reglas no son concluyentes.
"""
from __future__ import annotations

import re

from ..normalization import parse_nit
from .fields import fold

DOC_TYPES: dict[str, str] = {
    "FACTURA_VENTA": "Factura de venta",
    "FACTURA_COMPRA": "Factura de compra",
    "FACTURA_ELECTRONICA": "Factura electrónica",
    "RECIBO_CAJA": "Recibo de caja (RC)",
    "COMPROBANTE_EGRESO": "Comprobante de egreso (CE)",
    "CONTRATO": "Contrato",
    "OTROSI": "Otrosí",
    "ORDEN_COMPRA": "Orden de compra (OC)",
    "CERTIFICACION": "Certificación",
    "CUENTA_COBRO": "Cuenta de cobro",
    "EXTRACTO_BANCARIO": "Extracto bancario",
    "COMPROBANTE_CONTABLE": "Comprobante contable",
    "NOTA_DEBITO": "Nota débito",
    "NOTA_CREDITO": "Nota crédito",
    "DOCUMENTO_EQUIVALENTE": "Documento equivalente",
    "SOPORTE_PAGO": "Soporte de pago",
    "OTRO": "Otro documento relevante",
}

# Alias aceptados en la columna TIPO_DOCUMENTO del Excel
TYPE_ALIASES = {
    "FACTURA": "FACTURA",
    "FV": "FACTURA_VENTA",
    "FC": "FACTURA_COMPRA",
    "FE": "FACTURA_ELECTRONICA",
    "RC": "RECIBO_CAJA",
    "RECIBO DE CAJA": "RECIBO_CAJA",
    "CE": "COMPROBANTE_EGRESO",
    "COMPROBANTE DE EGRESO": "COMPROBANTE_EGRESO",
    "OC": "ORDEN_COMPRA",
    "ORDEN DE COMPRA": "ORDEN_COMPRA",
    "NC": "NOTA_CREDITO",
    "ND": "NOTA_DEBITO",
    "DS": "DOCUMENTO_EQUIVALENTE",
    "DOCUMENTO SOPORTE": "DOCUMENTO_EQUIVALENTE",
    "CC": "CUENTA_COBRO",
    "CUENTA DE COBRO": "CUENTA_COBRO",
    "OTRO SI": "OTROSI",
}

FAMILY = {
    "FACTURA_VENTA": "FACTURA",
    "FACTURA_COMPRA": "FACTURA",
    "FACTURA_ELECTRONICA": "FACTURA",
    "DOCUMENTO_EQUIVALENTE": "FACTURA",
    "CUENTA_COBRO": "FACTURA",
    "CONTRATO": "CONTRATO",
    "OTROSI": "CONTRATO",
    "COMPROBANTE_EGRESO": "PAGO",
    "SOPORTE_PAGO": "PAGO",
    "EXTRACTO_BANCARIO": "PAGO",
}

# (expresión, peso). Los títulos (primeras líneas) puntúan el triple.
KEYWORDS: dict[str, list[tuple[str, float]]] = {
    "FACTURA": [(r"FACTURA(?:\s+DE\s+VENTA)?", 3), (r"FACTURA\s+ELECTRONICA", 2), (r"CUFE", 2), (r"RESOLUCION\s+(?:DIAN|DE\s+FACTURACION)", 1), (r"SUBTOTAL", 0.5), (r"(?<![A-Z])IVA(?![A-Z])", 0.5)],
    "RECIBO_CAJA": [(r"RECIBO\s+DE\s+CAJA", 4), (r"RECIBIMOS\s+DE", 2), (r"RECIBI\s+DE", 2)],
    "COMPROBANTE_EGRESO": [(r"COMPROBANTE\s+DE\s+EGRESO", 4), (r"PAGADO\s+A", 2), (r"(?<![A-Z])EGRESO(?![A-Z])", 1)],
    "CONTRATO": [(r"CONTRATO\s+DE\s+[A-Z]+", 3), (r"CLAUSULA", 2), (r"CONTRATANTE", 1.5), (r"CONTRATISTA", 1.5), (r"OBJETO", 0.5)],
    "OTROSI": [(r"OTRO\s?SI", 5), (r"MODIFICATORIO", 2)],
    "ORDEN_COMPRA": [(r"ORDEN\s+DE\s+COMPRA", 3), (r"PURCHASE\s+ORDER", 3)],
    "CERTIFICACION": [(r"(?<![A-Z])CERTIFICA(?:CION|DO)?(?![A-Z])", 3), (r"HACE\s+CONSTAR", 2), (r"SE\s+EXPIDE", 1)],
    "CUENTA_COBRO": [(r"CUENTA\s+DE\s+COBRO", 5), (r"DEBE\s+A", 2)],
    "EXTRACTO_BANCARIO": [(r"EXTRACTO", 3), (r"SALDO\s+ANTERIOR", 2), (r"SALDO\s+(?:FINAL|ACTUAL)", 2), (r"MOVIMIENTOS", 1)],
    "COMPROBANTE_CONTABLE": [(r"COMPROBANTE\s+(?:CONTABLE|DE\s+DIARIO)", 4), (r"NOTA\s+DE\s+CONTABILIDAD", 4), (r"DEBITOS?\s+CREDITOS?", 1)],
    "NOTA_DEBITO": [(r"NOTA\s+DEBITO", 5)],
    "NOTA_CREDITO": [(r"NOTA\s+CREDITO", 5)],
    "DOCUMENTO_EQUIVALENTE": [(r"DOCUMENTO\s+EQUIVALENTE", 5), (r"DOCUMENTO\s+SOPORTE", 4)],
    "SOPORTE_PAGO": [(r"COMPROBANTE\s+DE\s+(?:PAGO|TRANSFERENCIA)", 4), (r"SOPORTE\s+DE\s+PAGO", 4), (r"CONSTANCIA\s+DE\s+PAGO", 4), (r"TRANSFERENCIA\s+(?:EXITOSA|ELECTRONICA)", 2), (r"(?<![A-Z])PSE(?![A-Z])", 1), (r"PAGO\s+EXITOSO", 2)],
}


def normalize_type(raw: str | None) -> str | None:
    """Convierte el texto de la columna TIPO_DOCUMENTO a un código interno."""
    if not raw:
        return None
    t = fold(str(raw)).strip().replace("_", " ")
    t = re.sub(r"\s+", " ", t)
    if t in TYPE_ALIASES:
        return TYPE_ALIASES[t]
    code = t.replace(" ", "_")
    if code in DOC_TYPES:
        return code
    for c, label in DOC_TYPES.items():
        if fold(label).startswith(t) or t.startswith(fold(label).split(" (")[0]):
            return c
    if "FACTURA" in t:
        return "FACTURA"
    return "OTRO"


def type_family(code: str | None) -> str | None:
    if not code:
        return None
    return FAMILY.get(code, code)


def classify(text: str, emitter_nit: str | None, client_nit: str | None, extra_keywords: dict[str, list[str]] | None = None) -> tuple[str, float, str]:
    """Devuelve (código, confianza 0-1, explicación)."""
    folded = fold(text or "")
    head = "\n".join(folded.splitlines()[:12])
    scores: dict[str, float] = {}
    reasons: dict[str, list[str]] = {}
    kw = {k: list(v) for k, v in KEYWORDS.items()}
    for code, words in (extra_keywords or {}).items():
        kw.setdefault(code, []).extend((re.escape(fold(w)), 3) for w in words if w)
    for code, patterns in kw.items():
        for pat, weight in patterns:
            n_head = len(re.findall(pat, head))
            n_body = len(re.findall(pat, folded))
            if n_body:
                s = weight * min(n_body, 3) + 2 * weight * min(n_head, 1)
                scores[code] = scores.get(code, 0) + s
                reasons.setdefault(code, []).append(pat.replace("\\s+", " ").replace("(?<![A-Z])", "").replace("(?![A-Z])", ""))
    if not scores:
        return "OTRO", 0.2, "Sin palabras clave reconocidas"
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best, top = ranked[0]
    # Un "otrosí" o una nota contienen la palabra contrato/factura: priorizarlos
    second = ranked[1][1] if len(ranked) > 1 else 0
    confidence = round(min(0.99, 0.5 + 0.5 * (top - second) / max(top, 1)), 2)
    if best == "FACTURA":
        if "CUFE" in folded or "FACTURA ELECTRONICA" in folded:
            best = "FACTURA_ELECTRONICA"
        else:
            e, c = parse_nit(emitter_nit), parse_nit(client_nit)
            best = "FACTURA_VENTA" if e and c and e.base == c.base else "FACTURA_COMPRA"
    explanation = f"Palabras clave: {', '.join(reasons.get(ranked[0][0], [])[:4])} (puntaje {top:.1f}"
    explanation += f"; siguiente {ranked[1][0]} {second:.1f})" if len(ranked) > 1 else ")"
    return best, confidence, explanation
