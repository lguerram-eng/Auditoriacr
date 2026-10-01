"""Estructuración asistida por IA (opcional) con la API de Claude.

Se activa solo si:
- MUENRA_LLM_ENABLED=true y existe ANTHROPIC_API_KEY (o MUENRA_ANTHROPIC_API_KEY), y
- el proyecto autoriza explícitamente el procesamiento con IA (allow_ai_processing).

Principios:
- La IA nunca sobrescribe valores obtenidos de forma determinística (XML / reglas);
  solo completa campos faltantes o sugiere el tipo documental si las reglas no son
  concluyentes.
- Cada campo devuelto debe citar el texto literal de evidencia; si la cita no se
  encuentra en el documento el campo se descarta (control anti-alucinación).
- Los documentos nunca se usan para entrenamiento: la API comercial de Anthropic
  no entrena con datos de clientes y el proyecto registra allow_training=False.
"""
from __future__ import annotations

import json
import logging
import os

from ...config import get_settings
from .base import ExtractionResult, FieldEvidence, Method
from .classifier import DOC_TYPES
from .fields import fold

log = logging.getLogger(__name__)

AI_FIELDS = [
    "numero_documento", "fecha_emision", "fecha_vencimiento", "emisor_nombre", "receptor_nombre", "emisor_nit",
    "receptor_nit", "subtotal", "base_gravable", "iva", "retenciones", "descuentos", "valor_total", "moneda",
    "numero_contrato", "orden_compra", "centro_costo", "concepto", "forma_pago", "cuenta_bancaria", "firmantes",
    "vigencia_contrato", "objeto_contractual",
]

MAX_CHARS = 150_000

SCHEMA = {
    "type": "object",
    "properties": {
        "tipo_documento": {"type": "string", "enum": list(DOC_TYPES)},
        "campos": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "campo": {"type": "string", "enum": AI_FIELDS},
                    "valor": {"type": "string"},
                    "evidencia": {"type": "string"},
                },
                "required": ["campo", "valor", "evidencia"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["tipo_documento", "campos"],
    "additionalProperties": False,
}

SYSTEM = (
    "Eres un asistente de auditoría financiera colombiana. Extraes datos de documentos soporte "
    "(facturas, recibos de caja, comprobantes de egreso, contratos, órdenes de compra, etc.). "
    "Reglas: usa solo información presente en el texto; en 'evidencia' copia literalmente el fragmento "
    "del documento (máximo 120 caracteres) que contiene el valor; fechas en formato AAAA-MM-DD; valores "
    "numéricos sin símbolo de moneda con punto decimal; si un campo no aparece, no lo incluyas. "
    "No emitas juicios sobre fraude."
)


def ai_enabled(project_allows: bool) -> bool:
    s = get_settings()
    key = s.anthropic_api_key or os.environ.get("ANTHROPIC_API_KEY")
    return bool(s.llm_enabled and project_allows and key)


def enrich_with_ai(result: ExtractionResult, project_allows: bool) -> list[str]:
    """Completa campos faltantes. Devuelve advertencias (nunca lanza excepción)."""
    if not ai_enabled(project_allows):
        return []
    text = result.full_text
    if not text.strip():
        return []
    warnings = []
    if len(text) > MAX_CHARS:
        warnings.append(f"IA: el texto supera {MAX_CHARS} caracteres; se enviaron solo los primeros {MAX_CHARS}")
        text = text[:MAX_CHARS]
    try:
        import anthropic
    except ImportError:
        return warnings + ["IA: el paquete 'anthropic' no está instalado"]
    try:
        s = get_settings()
        client = anthropic.Anthropic(api_key=s.anthropic_api_key or None)
        response = client.beta.messages.create(
            model=s.llm_model,
            max_tokens=8000,
            system=SYSTEM,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            output_config={"effort": "low", "format": {"type": "json_schema", "schema": SCHEMA}},
            messages=[{"role": "user", "content": f"Documento:\n<documento>\n{text}\n</documento>"}],
        )
        if response.stop_reason == "refusal":
            return warnings + ["IA: la solicitud fue rechazada por el modelo; se conservan solo las reglas"]
        payload = next((b.text for b in response.content if b.type == "text"), "")
        data = json.loads(payload)
    except anthropic.RateLimitError:
        return warnings + ["IA: límite de tasa alcanzado; se conservan solo las reglas"]
    except anthropic.APIStatusError as exc:
        return warnings + [f"IA: error del servicio ({exc.status_code})"]
    except anthropic.APIConnectionError:
        return warnings + ["IA: sin conexión con el servicio"]
    except (json.JSONDecodeError, StopIteration) as exc:
        return warnings + [f"IA: respuesta no válida ({exc.__class__.__name__})"]

    if (result.doc_type_confidence or 0) < 0.6 and data.get("tipo_documento"):
        result.doc_type = data["tipo_documento"]
        result.doc_type_confidence = 0.75
        result.classification_reason = (result.classification_reason or "") + " | Clasificación sugerida por IA"

    for item in data.get("campos", []):
        name, value, evidence = item.get("campo"), item.get("valor"), item.get("evidencia", "")
        if not name or not value or result.get(name):
            continue
        located = _locate(result, evidence)
        if not located:
            continue  # la evidencia no existe en el documento: se descarta
        page, line = located
        result.fields.append(
            FieldEvidence(name, value, value, page.number, line.bbox, line.text, Method.AI, min(70.0, line.conf or 70.0))
        )
    return warnings


def _locate(result: ExtractionResult, evidence: str):
    needle = fold(evidence or "").strip()
    if len(needle) < 3:
        return None
    for page in result.pages:
        for line in page.lines:
            hay = fold(line.text)
            if needle in hay or (len(hay) >= 6 and hay in needle):
                return page, line
    return None
