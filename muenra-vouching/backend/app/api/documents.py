"""Importación del Excel de referencia, carga masiva de documentos y procesamiento."""
from __future__ import annotations

import io
import json
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, HTTPException, Query, Request, UploadFile
from fastapi.encoders import jsonable_encoder
from fastapi.responses import Response, StreamingResponse
from sqlalchemy.orm import Session

from .. import models, schemas
from ..audit import log_event
from ..config import get_settings
from ..database import get_db
from ..security import client_ip, project_for_user, require
from ..services import storage
from ..services.excel_import import ImportRejected, import_workbook
from ..services.extraction.classifier import DOC_TYPES
from ..services.file_security import DOCUMENT_TYPES, REFERENCE_TYPES, FileRejected, sanitize_filename, validate_file
from ..worker import KIND_DOCUMENT, KIND_MATCH, enqueue

router = APIRouter(tags=["Documentos e importación"])

CHUNK = 1024 * 1024


async def _read_limited(upload: UploadFile) -> bytes:
    limit = get_settings().max_file_mb * 1024 * 1024
    buf = io.BytesIO()
    while True:
        chunk = await upload.read(CHUNK)
        if not chunk:
            break
        buf.write(chunk)
        if buf.tell() > limit:
            raise FileRejected(f"El archivo supera el límite de {get_settings().max_file_mb} MB")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Importación del Excel
# ---------------------------------------------------------------------------


@router.post("/projects/{project_id}/imports")
async def import_reference(
    project_id: int,
    request: Request,
    file: UploadFile = File(...),
    dry_run: bool = Query(False, description="Solo validar y devolver el informe de integridad"),
    user: models.User = Depends(require("importar")),
    db: Session = Depends(get_db),
):
    p = project_for_user(db, project_id, user)
    name = sanitize_filename(file.filename or "referencia.xlsx")
    try:
        data = await _read_limited(file)
        validate_file(name, data, REFERENCE_TYPES)
    except FileRejected as exc:
        log_event(db, "IMPORTACION_RECHAZADA", user, p.id, detail={"archivo": name, "motivo": str(exc)}, ip=client_ip(request), success=False)
        raise HTTPException(422, str(exc))
    sha = storage.sha256(data)
    try:
        batch, report = import_workbook(db, p, name, data, sha, user, dry_run=dry_run)
    except ImportRejected as exc:
        db.rollback()
        log_event(db, "IMPORTACION_RECHAZADA", user, p.id, detail={"archivo": name, "errores": exc.report["errores"][:10]}, ip=client_ip(request), success=False)
        return Response(content=json.dumps(jsonable_encoder({"importado": False, "informe": exc.report})), status_code=422, media_type="application/json")
    except Exception as exc:  # archivo ilegible
        db.rollback()
        raise HTTPException(422, f"No fue posible leer el archivo: {exc.__class__.__name__}")
    if batch:
        log_event(db, "IMPORTACION_EXCEL", user, p.id, "import", batch.id,
                  {"archivo": name, "sha256": sha, "partidas": batch.row_count, "valor_total": str(batch.total_value), "estado": batch.status}, client_ip(request))
    return jsonable_encoder({"importado": batch is not None, "importacion_id": batch.id if batch else None, "informe": report})


@router.get("/projects/{project_id}/imports")
def list_imports(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    rows = db.query(models.ImportBatch).filter_by(project_id=p.id).order_by(models.ImportBatch.id.desc()).all()
    return jsonable_encoder([{
        "id": b.id, "archivo": b.filename, "sha256": b.sha256, "estado": b.status, "partidas": b.row_count, "valor_total": b.total_value,
        "activa": b.is_active, "fecha": b.created_at, "informe": b.integrity_report,
    } for b in rows])


@router.get("/projects/{project_id}/references")
def list_references(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    rows = db.query(models.ReferenceItem).filter_by(project_id=p.id).order_by(models.ReferenceItem.row_number).all()
    return [schemas.ReferenceOut.model_validate(r).model_dump(mode="json") for r in rows]


# ---------------------------------------------------------------------------
# Documentos
# ---------------------------------------------------------------------------


@router.post("/projects/{project_id}/documents")
async def upload_documents(
    project_id: int,
    request: Request,
    files: list[UploadFile] = File(...),
    process: bool = Query(True, description="Encolar el procesamiento inmediatamente"),
    user: models.User = Depends(require("cargar")),
    db: Session = Depends(get_db),
):
    p = project_for_user(db, project_id, user)
    s = get_settings()
    if len(files) > s.max_files_per_upload:
        raise HTTPException(422, f"Máximo {s.max_files_per_upload} archivos por carga")
    accepted, rejected = [], []
    for up in files:
        name = sanitize_filename(up.filename or "documento")
        try:
            data = await _read_limited(up)
            check = validate_file(name, data, DOCUMENT_TYPES)
        except FileRejected as exc:
            rejected.append({"archivo": name, "motivo": str(exc)})
            log_event(db, "ARCHIVO_RECHAZADO", user, p.id, "document", None, {"archivo": name, "motivo": str(exc)}, client_ip(request), success=False, commit=False)
            continue
        sha = storage.sha256(data)
        existing = db.query(models.Document).filter_by(project_id=p.id, sha256=sha).order_by(models.Document.id).first()
        doc = models.Document(
            project_id=p.id, filename=name, stored_path=storage.save(p.id, data, check.file_type), sha256=sha, size_bytes=len(data),
            file_type=check.file_type, mime=check.mime, av_status=check.av_status, uploaded_by=user.id,
            duplicate_of_id=existing.id if existing else None,
        )
        db.add(doc)
        db.flush()
        if process:
            enqueue(db, p.id, KIND_DOCUMENT, {"document_id": doc.id}, commit=False)
        accepted.append({"id": doc.id, "archivo": name, "sha256": sha, "tamano": len(data), "tipo": check.file_type,
                         "antivirus": check.av_status, "duplicado_de": existing.id if existing else None})
    if process and accepted:
        _enqueue_match(db, p.id, user)
    log_event(db, "DOCUMENTOS_CARGADOS", user, p.id, detail={"aceptados": len(accepted), "rechazados": len(rejected),
              "archivos": [a["archivo"] for a in accepted][:200]}, ip=client_ip(request), commit=False)
    db.commit()
    return {"aceptados": accepted, "rechazados": rejected, "total_recibidos": len(files)}


def _enqueue_match(db: Session, project_id: int, user: models.User | None) -> None:
    pending = db.query(models.Job).filter_by(project_id=project_id, kind=KIND_MATCH, status="EN_COLA").first()
    if not pending:
        enqueue(db, project_id, KIND_MATCH, {"user_id": user.id if user else None}, commit=False)


@router.post("/projects/{project_id}/process")
def process_pending(project_id: int, request: Request, reprocess_errors: bool = False, user: models.User = Depends(require("procesar")), db: Session = Depends(get_db)):
    """Encola documentos pendientes (y opcionalmente con error) y una ejecución del motor de vouching."""
    p = project_for_user(db, project_id, user)
    states = ["PENDIENTE"] + (["ERROR", "ILEGIBLE"] if reprocess_errors else [])
    queued_ids = {j.payload.get("document_id") for j in db.query(models.Job).filter(models.Job.project_id == p.id, models.Job.kind == KIND_DOCUMENT, models.Job.status.in_(["EN_COLA", "EN_PROCESO"]))}
    n = 0
    for d in db.query(models.Document).filter(models.Document.project_id == p.id, models.Document.processing_status.in_(states)):
        if d.id not in queued_ids:
            enqueue(db, p.id, KIND_DOCUMENT, {"document_id": d.id}, commit=False)
            n += 1
    _enqueue_match(db, p.id, user)
    log_event(db, "PROCESAMIENTO_SOLICITADO", user, p.id, detail={"documentos": n}, ip=client_ip(request), commit=False)
    db.commit()
    return {"documentos_encolados": n, "vouching_encolado": True}


@router.post("/projects/{project_id}/match")
def run_match(project_id: int, request: Request, user: models.User = Depends(require("procesar")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    _enqueue_match(db, p.id, user)
    log_event(db, "VOUCHING_SOLICITADO", user, p.id, ip=client_ip(request), commit=False)
    db.commit()
    return {"vouching_encolado": True}


@router.get("/projects/{project_id}/jobs")
def jobs_status(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    jobs = db.query(models.Job).filter_by(project_id=p.id).order_by(models.Job.id.desc()).limit(1000).all()
    counts: dict[str, int] = {}
    for j in jobs:
        counts[f"{j.kind}:{j.status}"] = counts.get(f"{j.kind}:{j.status}", 0) + 1
    docs = db.query(models.Document).filter_by(project_id=p.id).all()
    by_state: dict[str, int] = {}
    for d in docs:
        by_state[d.processing_status] = by_state.get(d.processing_status, 0) + 1
    return jsonable_encoder({
        "trabajos": counts,
        "documentos_por_estado": by_state,
        "total_documentos": len(docs),
        "en_curso": any(j.status in ("EN_COLA", "EN_PROCESO") for j in jobs),
        "ultimos": [{"id": j.id, "tipo": j.kind, "estado": j.status, "intentos": j.attempts, "error": j.error, "creado": j.created_at,
                     "inicio": j.started_at, "fin": j.finished_at, "documento_id": j.payload.get("document_id")} for j in jobs[:50]],
    })


def _doc_summary(d: models.Document) -> dict:
    f = {x.field_name: x.effective_value for x in d.fields}
    return {
        "id": d.id, "archivo": d.filename, "sha256": d.sha256, "tamano": d.size_bytes, "formato": d.file_type,
        "estado_tecnico": d.processing_status, "estado": d.vouching_status, "tipo": d.doc_type,
        "tipo_nombre": DOC_TYPES.get(d.doc_type or "", d.doc_type), "confianza_tipo": d.doc_type_confidence, "paginas": d.page_count,
        "metodo": d.extraction_method, "confianza_ocr": d.ocr_confidence, "antivirus": d.av_status, "duplicado_de": d.duplicate_of_id,
        "grupo": d.group_key, "cargado": d.uploaded_at, "procesado": d.processed_at, "observacion": d.error,
        "numero": f.get("numero_documento"), "fecha": f.get("fecha_emision"), "emisor": f.get("emisor_nombre"),
        "nit_emisor": f.get("emisor_nit"), "valor_total": f.get("valor_total"),
    }


@router.get("/projects/{project_id}/documents")
def list_documents(project_id: int, status: str | None = None, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    q = db.query(models.Document).filter_by(project_id=p.id)
    if status:
        q = q.filter(models.Document.vouching_status == status)
    return jsonable_encoder([_doc_summary(d) for d in q.order_by(models.Document.id).all()])


def _doc_for_user(db: Session, document_id: int, user: models.User) -> models.Document:
    d = db.get(models.Document, document_id)
    if not d:
        raise HTTPException(404, "Documento no encontrado")
    project_for_user(db, d.project_id, user)
    return d


@router.get("/documents/{document_id}")
def document_detail(document_id: int, request: Request, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    d = _doc_for_user(db, document_id, user)
    links = db.query(models.MatchLink).filter_by(document_id=d.id).all()
    refs = {r.id: r for r in db.query(models.ReferenceItem).filter(models.ReferenceItem.id.in_([ln.reference_item_id for ln in links] or [0]))}
    users = {u.id: u.email for u in db.query(models.User)}
    log_event(db, "DOCUMENTO_CONSULTADO", user, d.project_id, "document", d.id, ip=client_ip(request))
    return jsonable_encoder({
        **_doc_summary(d),
        "motivo_clasificacion": d.classification_reason,
        "tablas": d.tables,
        "paginas_detalle": [{"pagina": pg.page_number, "metodo": pg.method, "confianza_ocr": pg.ocr_confidence, "ancho": pg.width,
                             "alto": pg.height, "lineas": pg.lines} for pg in d.pages],
        "campos": [{
            "id": f.id, "campo": f.field_name, "valor": f.value, "normalizado": f.normalized_value, "valor_efectivo": f.effective_value,
            "pagina": f.page, "region": f.bbox, "evidencia": f.evidence_text, "metodo": f.method, "confianza": f.confidence,
            "extraido": f.extracted_at, "estado_revision": f.review_status, "valor_corregido": f.corrected_value,
            "revisor": users.get(f.reviewed_by), "fecha_revision": f.reviewed_at, "comentario": f.review_comment,
        } for f in sorted(d.fields, key=lambda x: x.field_name)],
        "relaciones": [{"id": ln.id, "partida_id": ln.reference_item_id, "id_muestra": refs[ln.reference_item_id].sample_id if ln.reference_item_id in refs else None,
                        "rol": ln.role, "puntaje": ln.score, "manual": ln.is_manual} for ln in links],
    })


@router.get("/documents/{document_id}/pages/{page}/image")
def page_image(document_id: int, page: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    from PIL import Image

    from ..services.extraction.readers import image_frames, render_pdf_page

    d = _doc_for_user(db, document_id, user)
    if d.file_type not in ("pdf", "png", "jpg", "tiff"):
        raise HTTPException(404, "Este formato se visualiza como texto")
    data = storage.load(d.stored_path)
    try:
        if d.file_type == "pdf":
            img = render_pdf_page(data, page - 1, dpi=110)
        else:
            frames = image_frames(data)
            img = frames[page - 1].convert("RGB")
            if max(img.size) > 1800:
                img.thumbnail((1800, 1800), Image.LANCZOS)
    except (IndexError, ValueError):
        raise HTTPException(404, "Página no encontrada")
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return Response(buf.getvalue(), media_type="image/png", headers={"Cache-Control": "private, max-age=600"})


@router.get("/documents/{document_id}/file")
def download_file(document_id: int, request: Request, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    d = _doc_for_user(db, document_id, user)
    data = storage.load(d.stored_path)
    log_event(db, "DOCUMENTO_DESCARGADO", user, d.project_id, "document", d.id, ip=client_ip(request))
    return StreamingResponse(io.BytesIO(data), media_type=d.mime, headers={
        "Content-Disposition": f"attachment; filename*=UTF-8''{quote(d.filename)}",
        "X-Content-Type-Options": "nosniff",
    })


@router.post("/documents/{document_id}/reprocess")
def reprocess(document_id: int, request: Request, user: models.User = Depends(require("procesar")), db: Session = Depends(get_db)):
    d = _doc_for_user(db, document_id, user)
    d.processing_status = "PENDIENTE"
    enqueue(db, d.project_id, KIND_DOCUMENT, {"document_id": d.id}, commit=False)
    _enqueue_match(db, d.project_id, user)
    log_event(db, "DOCUMENTO_REPROCESADO", user, d.project_id, "document", d.id, ip=client_ip(request), commit=False)
    db.commit()
    return {"ok": True}


@router.delete("/documents/{document_id}")
def delete_document(document_id: int, request: Request, user: models.User = Depends(require("cargar")), db: Session = Depends(get_db)):
    d = _doc_for_user(db, document_id, user)
    pid, name, sha, path = d.project_id, d.filename, d.sha256, d.stored_path
    db.query(models.Document).filter_by(duplicate_of_id=d.id).update({"duplicate_of_id": None})
    db.delete(d)
    storage.delete(path)
    _enqueue_match(db, pid, user)
    log_event(db, "DOCUMENTO_ELIMINADO", user, pid, "document", document_id, {"archivo": name, "sha256": sha}, client_ip(request), commit=False)
    db.commit()
    return {"ok": True}
