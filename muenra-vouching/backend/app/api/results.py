"""Resultados de vouching, excepciones, revisión humana y exportación."""
from __future__ import annotations

from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, Request
from fastapi.encoders import jsonable_encoder
from fastapi.responses import Response
from sqlalchemy.orm import Session

from .. import models, schemas
from ..audit import log_event
from ..database import get_db
from ..models import Status
from ..security import client_ip, project_for_user, require
from ..services import exporter
from ..services.extraction.classifier import DOC_TYPES
from ..services.normalization import parse_amount, parse_date
from ..worker import KIND_MATCH, enqueue

router = APIRouter(tags=["Resultados y revisión"])

EXCEPTION_STATES = [Status.EXCEPCION, Status.REVISION_MANUAL, Status.POSIBLE_DUPLICADO, Status.ILEGIBLE, Status.ERROR_LECTURA, Status.SIN_SOPORTE]


def _row(res: models.VouchingResult, links: list[models.MatchLink], docs: dict, users: dict) -> dict:
    r = res.reference_item
    prim = [ln for ln in links if ln.role in ("PRINCIPAL", "REPRESENTACION_GRAFICA")]
    return {
        "id": res.id, "partida_id": r.id, "fila": r.row_number, "id_muestra": r.sample_id, "tipo_esperado": r.doc_type,
        "numero_esperado": r.doc_number, "fecha_esperada": r.doc_date, "tercero_esperado": r.third_party, "nit_esperado": r.nit,
        "valor_esperado": r.expected_value, "moneda": r.currency, "contrato": r.contract, "orden_compra": r.purchase_order,
        "concepto": r.concept, "estado": res.status, "estado_automatico": res.auto_status, "estado_manual": res.manual_status,
        "motivos": res.reasons, "puntaje": res.score, "numero_extraido": res.extracted_number, "fecha_extraida": res.extracted_date,
        "tercero_extraido": res.extracted_party, "nit_extraido": res.extracted_nit, "valor_extraido": res.extracted_value,
        "diferencia_absoluta": res.abs_difference, "diferencia_porcentual": res.pct_difference, "tolerancia": res.tolerance_applied,
        "similitud_nombre": res.name_similarity, "coincidencia_nit": res.nit_match, "coincidencia_dv": res.dv_match,
        "coincidencia_numero": res.number_match, "diferencia_dias": res.days_difference, "confianza_ocr": res.ocr_confidence,
        "pagina_evidencia": res.evidence_page, "texto_evidencia": res.evidence_text,
        "documentos": [{"id": ln.document_id, "archivo": docs[ln.document_id].filename if ln.document_id in docs else None,
                        "hash": docs[ln.document_id].sha256 if ln.document_id in docs else None, "rol": ln.role,
                        "tipo": DOC_TYPES.get(docs[ln.document_id].doc_type or "", None) if ln.document_id in docs else None,
                        "puntaje": ln.score, "manual": ln.is_manual, "link_id": ln.id} for ln in links],
        "archivos": "; ".join(docs[ln.document_id].filename for ln in prim if ln.document_id in docs),
        "decision_revisor": res.review_decision, "revisor": users.get(res.reviewer_id), "fecha_revision": res.reviewed_at,
        "comentario_revisor": res.review_comment,
    }


def _load(db: Session, project_id: int, results: list[models.VouchingResult]):
    links = db.query(models.MatchLink).filter_by(project_id=project_id).all()
    by_ref: dict[int, list] = {}
    for ln in links:
        by_ref.setdefault(ln.reference_item_id, []).append(ln)
    docs = {d.id: d for d in db.query(models.Document).filter_by(project_id=project_id)}
    users = {u.id: u.email for u in db.query(models.User)}
    return by_ref, docs, users


@router.get("/projects/{project_id}/results")
def list_results(
    project_id: int,
    status: list[str] | None = Query(None),
    q: str | None = None,
    exceptions_only: bool = False,
    user: models.User = Depends(require("ver")),
    db: Session = Depends(get_db),
):
    p = project_for_user(db, project_id, user)
    results = (
        db.query(models.VouchingResult).join(models.ReferenceItem)
        .filter(models.VouchingResult.project_id == p.id).order_by(models.ReferenceItem.row_number).all()
    )
    if exceptions_only:
        results = [r for r in results if r.status in EXCEPTION_STATES]
    if status:
        results = [r for r in results if r.status in status]
    by_ref, docs, users = _load(db, p.id, results)
    rows = [_row(r, by_ref.get(r.reference_item_id, []), docs, users) for r in results]
    if q:
        ql = q.lower()
        rows = [r for r in rows if any(ql in str(v).lower() for k, v in r.items() if k in (
            "id_muestra", "numero_esperado", "tercero_esperado", "nit_esperado", "archivos", "numero_extraido", "tercero_extraido", "concepto"))]
    return jsonable_encoder(rows)


def _result_for_user(db: Session, result_id: int, user: models.User) -> models.VouchingResult:
    res = db.get(models.VouchingResult, result_id)
    if not res:
        raise HTTPException(404, "Resultado no encontrado")
    project_for_user(db, res.project_id, user)
    return res


@router.get("/results/{result_id}")
def result_detail(result_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    res = _result_for_user(db, result_id, user)
    by_ref, docs, users = _load(db, res.project_id, [res])
    links = by_ref.get(res.reference_item_id, [])
    out = _row(res, links, docs, users)
    out["explicacion"] = res.explanation
    out["criterios_por_documento"] = {ln.document_id: ln.criteria for ln in links}
    evs = db.query(models.ReviewEvent).filter_by(result_id=res.id).order_by(models.ReviewEvent.id.desc()).all()
    out["historial"] = [{"fecha": e.created_at, "usuario": e.user_email, "accion": e.action, "campo": e.field_name,
                         "anterior": e.old_value, "nuevo": e.new_value, "comentario": e.comment} for e in evs]
    # Candidatos para relación manual
    out["candidatos"] = [{"id": d.id, "archivo": d.filename, "tipo": DOC_TYPES.get(d.doc_type or "", d.doc_type), "estado": d.vouching_status}
                         for d in docs.values() if d.id not in {ln.document_id for ln in links}]
    return jsonable_encoder(out)


def _event(db, res, user, action, old=None, new=None, comment=None, field=None, document_id=None):
    db.add(models.ReviewEvent(
        project_id=res.project_id if res else field.document.project_id, result_id=res.id if res else None,
        field_id=field.id if field else None, document_id=document_id, sample_id=res.reference_item.sample_id if res else None,
        user_id=user.id, user_email=user.email, action=action, field_name=field.field_name if field else None,
        old_value=None if old is None else str(old), new_value=None if new is None else str(new), comment=comment,
    ))


@router.post("/results/{result_id}/review")
def review_result(result_id: int, body: schemas.ResultReviewIn, request: Request, user: models.User = Depends(require("revisar")), db: Session = Depends(get_db)):
    res = _result_for_user(db, result_id, user)
    if body.status and body.status not in Status.ALL:
        raise HTTPException(422, "Estado inválido")
    if (body.decision == "RECHAZADO" or (body.status and body.status != res.auto_status)) and not (body.comment and body.comment.strip()):
        raise HTTPException(422, "Se requiere un comentario para rechazar o cambiar el estado")
    old = f"{res.status} / {res.review_decision or '-'}"
    res.review_decision = body.decision
    res.manual_status = body.status if body.status and body.status != res.auto_status else None
    res.reviewer_id = user.id
    res.reviewed_at = datetime.now(timezone.utc)
    if body.comment:
        res.review_comment = body.comment
    _event(db, res, user, f"RESULTADO_{body.decision}", old, f"{res.status} / {body.decision}", body.comment)
    log_event(db, "RESULTADO_REVISADO", user, res.project_id, "result", res.id,
              {"id_muestra": res.reference_item.sample_id, "decision": body.decision, "estado": res.status}, client_ip(request))
    return {"ok": True, "estado": res.status}


@router.post("/results/{result_id}/comment")
def comment_result(result_id: int, body: schemas.CommentIn, request: Request, user: models.User = Depends(require("revisar")), db: Session = Depends(get_db)):
    res = _result_for_user(db, result_id, user)
    _event(db, res, user, "COMENTARIO", comment=body.comment)
    log_event(db, "COMENTARIO", user, res.project_id, "result", res.id, {"id_muestra": res.reference_item.sample_id}, client_ip(request))
    return {"ok": True}


def _enqueue_match(db: Session, project_id: int, user: models.User) -> None:
    if not db.query(models.Job).filter_by(project_id=project_id, kind=KIND_MATCH, status="EN_COLA").first():
        enqueue(db, project_id, KIND_MATCH, {"user_id": user.id}, commit=False)


@router.post("/results/{result_id}/links")
def add_manual_link(result_id: int, body: schemas.ManualLinkIn, request: Request, user: models.User = Depends(require("revisar")), db: Session = Depends(get_db)):
    res = _result_for_user(db, result_id, user)
    doc = db.get(models.Document, body.document_id)
    if not doc or doc.project_id != res.project_id:
        raise HTTPException(404, "Documento no encontrado en el proyecto")
    db.add(models.MatchLink(project_id=res.project_id, run_id=res.run_id or _ensure_run(db, res.project_id), reference_item_id=res.reference_item_id,
                            document_id=doc.id, role=body.role, score=1.0, criteria={"manual": True, "comentario": body.comment}, is_manual=True))
    _event(db, res, user, "RELACION_MANUAL", None, f"{doc.filename} ({body.role})", body.comment, document_id=doc.id)
    _enqueue_match(db, res.project_id, user)
    log_event(db, "RELACION_MANUAL", user, res.project_id, "result", res.id, {"documento": doc.filename, "rol": body.role}, client_ip(request))
    return {"ok": True}


def _ensure_run(db: Session, project_id: int) -> int:
    run = models.MatchRun(project_id=project_id, parameters={}, engine_version="manual")
    db.add(run)
    db.flush()
    return run.id


@router.delete("/links/{link_id}")
def delete_link(link_id: int, request: Request, user: models.User = Depends(require("revisar")), db: Session = Depends(get_db)):
    ln = db.get(models.MatchLink, link_id)
    if not ln:
        raise HTTPException(404, "Relación no encontrada")
    project_for_user(db, ln.project_id, user)
    res = db.query(models.VouchingResult).filter_by(reference_item_id=ln.reference_item_id).first()
    doc = db.get(models.Document, ln.document_id)
    if not ln.is_manual:
        raise HTTPException(400, "Solo se pueden retirar relaciones manuales; para relaciones automáticas use la decisión del revisor")
    db.delete(ln)
    if res:
        _event(db, res, user, "RELACION_RETIRADA", doc.filename if doc else ln.document_id, None, document_id=ln.document_id)
    _enqueue_match(db, ln.project_id, user)
    log_event(db, "RELACION_RETIRADA", user, ln.project_id, "link", link_id, ip=client_ip(request))
    return {"ok": True}


@router.post("/fields/{field_id}/review")
def review_field(field_id: int, body: schemas.FieldReviewIn, request: Request, user: models.User = Depends(require("revisar")), db: Session = Depends(get_db)):
    """Aceptar, rechazar o corregir un valor extraído. El valor original nunca se modifica."""
    f = db.get(models.ExtractedField, field_id)
    if not f:
        raise HTTPException(404, "Campo no encontrado")
    doc = f.document
    project_for_user(db, doc.project_id, user)
    if body.action == "CORREGIR":
        val = (body.corrected_value or "").strip()
        if not val:
            raise HTTPException(422, "Debe indicar el valor corregido")
        if f.field_name in ("valor_total", "subtotal", "iva", "retenciones", "descuentos", "base_gravable"):
            amt = parse_amount(val)
            if amt is None:
                raise HTTPException(422, "El valor corregido no es numérico")
            val = str(amt)
        if f.field_name in ("fecha_emision", "fecha_vencimiento"):
            d = parse_date(val)
            if d is None:
                raise HTTPException(422, "La fecha corregida no es válida")
            val = d.isoformat()
        if not (body.comment and body.comment.strip()):
            raise HTTPException(422, "Se requiere un comentario para corregir un valor")
    if body.action == "RECHAZAR" and not (body.comment and body.comment.strip()):
        raise HTTPException(422, "Se requiere un comentario para rechazar un valor")
    old = f.effective_value
    f.review_status = {"ACEPTAR": "ACEPTADO", "RECHAZAR": "RECHAZADO", "CORREGIR": "CORREGIDO"}[body.action]
    f.corrected_value = val if body.action == "CORREGIR" else None
    f.reviewed_by = user.id
    f.reviewed_at = datetime.now(timezone.utc)
    f.review_comment = body.comment
    links = db.query(models.MatchLink).filter_by(document_id=doc.id).all()
    res = db.query(models.VouchingResult).filter_by(reference_item_id=links[0].reference_item_id).first() if links else None
    db.add(models.ReviewEvent(
        project_id=doc.project_id, result_id=res.id if res else None, field_id=f.id, document_id=doc.id,
        sample_id=res.reference_item.sample_id if res else None, user_id=user.id, user_email=user.email,
        action=f"CAMPO_{f.review_status}", field_name=f.field_name, old_value=old, new_value=f.effective_value, comment=body.comment,
    ))
    if body.action != "ACEPTAR":
        _enqueue_match(db, doc.project_id, user)
    log_event(db, "CAMPO_REVISADO", user, doc.project_id, "field", f.id,
              {"documento": doc.filename, "campo": f.field_name, "accion": f.review_status}, client_ip(request))
    return {"ok": True, "valor_original": f.value, "valor_efectivo": f.effective_value, "estado_revision": f.review_status}


@router.get("/projects/{project_id}/export")
def export_excel(project_id: int, request: Request, sheets: str | None = Query(None, description="Lista separada por comas: " + ",".join(exporter.SHEETS)),
                 user: models.User = Depends(require("exportar")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    only = [s.strip() for s in sheets.split(",") if s.strip() in exporter.SHEETS] if sheets else None
    data = exporter.build_workbook(db, p, only)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M")
    name = f"Muenra_Vouching_{p.code}_{stamp}.xlsx"
    log_event(db, "EXPORTACION_EXCEL", user, p.id, detail={"hojas": only or list(exporter.SHEETS), "archivo": name}, ip=client_ip(request))
    return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@router.get("/export/sheets")
def export_sheets(_: models.User = Depends(require("ver"))):
    return exporter.SHEETS
