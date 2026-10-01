"""Indicadores del tablero principal."""
from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from sqlalchemy.orm import Session

from .. import models
from ..models import Status

SUPPORTED = (Status.COINCIDE, Status.COINCIDE_TOLERANCIA)
WITH_DIFFERENCES = (Status.EXCEPCION,)
NEEDS_REVIEW = (Status.EXCEPCION, Status.REVISION_MANUAL, Status.POSIBLE_DUPLICADO, Status.ILEGIBLE, Status.ERROR_LECTURA)


def kpis(db: Session, project_id: int) -> dict:
    refs = {r.id: r for r in db.query(models.ReferenceItem).filter_by(project_id=project_id)}
    results = db.query(models.VouchingResult).filter_by(project_id=project_id).all()
    docs = db.query(models.Document).filter_by(project_id=project_id).all()

    by_status: dict[str, dict] = defaultdict(lambda: {"partidas": 0, "valor": Decimal(0)})
    for st in Status.ALL:
        if st != Status.NO_REFERENCIADO:
            by_status[st]
    total_value = sum((r.expected_value or Decimal(0)) for r in refs.values())
    supported_value = Decimal(0)
    diff_value = Decimal(0)
    abs_diffs = Decimal(0)
    with_support = 0
    reviewed = 0
    pending_review = 0
    for res in results:
        ref = refs.get(res.reference_item_id)
        val = (ref.expected_value if ref else None) or Decimal(0)
        st = res.status
        by_status[st]["partidas"] += 1
        by_status[st]["valor"] += val
        if st in SUPPORTED:
            supported_value += val
        if st in WITH_DIFFERENCES:
            diff_value += val
            abs_diffs += res.abs_difference or Decimal(0)
        if st not in (Status.SIN_SOPORTE, Status.PENDIENTE):
            with_support += 1
        if res.review_decision:
            reviewed += 1
        elif st in NEEDS_REVIEW:
            pending_review += 1
    n = len(results)
    processed_docs = sum(1 for d in docs if d.processing_status not in ("PENDIENTE", "PROCESANDO"))
    done_rows = sum(1 for r in results if r.status != Status.PENDIENTE and (r.status not in NEEDS_REVIEW or r.review_decision))
    progress_parts = []
    if docs:
        progress_parts.append(processed_docs / len(docs))
    if n:
        progress_parts.append(done_rows / n)
    doc_status = defaultdict(int)
    for d in docs:
        doc_status[d.vouching_status] += 1
    return {
        "aplicativo": "Muenra Vouching",
        "total_partidas": n,
        "total_documentos": len(docs),
        "documentos_procesados": processed_docs,
        "partidas_con_soporte": with_support,
        "partidas_sin_soporte": by_status[Status.SIN_SOPORTE]["partidas"],
        "coincidencias": by_status[Status.COINCIDE]["partidas"],
        "coincidencias_tolerancia": by_status[Status.COINCIDE_TOLERANCIA]["partidas"],
        "excepciones": by_status[Status.EXCEPCION]["partidas"],
        "revision_manual": by_status[Status.REVISION_MANUAL]["partidas"],
        "posibles_duplicados": by_status[Status.POSIBLE_DUPLICADO]["partidas"],
        "pendientes_revision": pending_review,
        "revisadas": reviewed,
        "documentos_ilegibles": sum(1 for d in docs if d.processing_status == "ILEGIBLE"),
        "documentos_error": sum(1 for d in docs if d.processing_status == "ERROR"),
        "soportes_no_referenciados": doc_status.get(Status.NO_REFERENCIADO, 0),
        "valor_poblacion": total_value,
        "valor_soportado": supported_value,
        "valor_con_diferencias": diff_value,
        "suma_diferencias_absolutas": abs_diffs,
        "porcentaje_cobertura": round(float(supported_value / total_value * 100), 2) if total_value else 0.0,
        "porcentaje_avance": round(sum(progress_parts) / len(progress_parts) * 100, 2) if progress_parts else 0.0,
        "por_estado": {k: v for k, v in by_status.items()},
        "documentos_por_estado": dict(doc_status),
    }
