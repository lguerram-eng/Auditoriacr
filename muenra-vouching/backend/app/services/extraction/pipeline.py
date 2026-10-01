"""Orquestación del procesamiento de un documento soporte."""
from __future__ import annotations

import logging
import time
from datetime import datetime, timezone

from sqlalchemy.orm import Session

from ... import models
from ...audit import mask
from .. import storage
from .base import ExtractionResult, FieldEvidence, Method, UnreadableDocument
from .classifier import classify
from .fields import extract_fields
from .llm import enrich_with_ai
from .readers import read_docx, read_image, read_pdf
from .xml_extractor import InvalidXML, extract_xml

log = logging.getLogger(__name__)

MIN_TEXT_CHARS = 25
MIN_OCR_CONFIDENCE_LEGIBLE = 35.0
ENGINE_VERSION = "extraccion-1.0"


def extract(data: bytes, file_type: str, client_nit: str | None = None, extra_keywords: dict | None = None, allow_ai: bool = False) -> ExtractionResult:
    """Extrae texto, campos y clasificación de un archivo en memoria (sin base de datos)."""
    if file_type == "xml":
        result = extract_xml(data)
    else:
        if file_type == "pdf":
            pages, tables, warnings = read_pdf(data)
            ocr_pages = sum(1 for p in pages if p.method == Method.OCR)
            method = Method.OCR if pages and ocr_pages == len(pages) else ("PDF_MIXTO" if ocr_pages else Method.PDF_TEXT)
        elif file_type in ("png", "jpg", "tiff"):
            pages, tables, warnings = read_image(data), [], []
            method = Method.OCR
        elif file_type == "docx":
            (pages, tables), warnings = read_docx(data), []
            method = Method.DOCX
        else:
            raise ValueError(f"Tipo no soportado: {file_type}")
        result = ExtractionResult(file_type=file_type, method=method, pages=pages, tables=tables, warnings=warnings)
        text = result.full_text
        if len(text.strip()) < MIN_TEXT_CHARS:
            raise UnreadableDocument("El documento no contiene texto legible")
        if result.method in (Method.OCR,) and (result.ocr_confidence or 0) < MIN_OCR_CONFIDENCE_LEGIBLE:
            raise UnreadableDocument(f"Confianza OCR muy baja ({result.ocr_confidence})")
        # 1) extracción preliminar para obtener el NIT del emisor; 2) clasificación; 3) extracción definitiva
        prelim = extract_fields(pages)
        emitter = next((f.normalized for f in prelim if f.name == "emisor_nit"), None)
        code, conf, why = classify(text, emitter, client_nit, extra_keywords)
        result.doc_type, result.doc_type_confidence, result.classification_reason = code, conf, why
        result.fields = extract_fields(pages, code)
        first = pages[0] if pages else None
        result.fields.insert(
            0,
            FieldEvidence(
                "tipo_documento", code, code, 1, first.lines[0].bbox if first and first.lines else None,
                why, Method.RULES, round(conf * 100, 1),
            ),
        )
    result.warnings.extend(enrich_with_ai(result, allow_ai))
    return result


def process_document(db: Session, document_id: int) -> models.Document:
    doc = db.get(models.Document, document_id)
    if doc is None:
        raise ValueError(f"Documento {document_id} no existe")
    project = db.get(models.Project, doc.project_id)
    doc.processing_status = "PROCESANDO"
    doc.error = None
    db.commit()
    started = time.monotonic()
    try:
        data = storage.load(doc.stored_path)
        if storage.sha256(data) != doc.sha256:
            raise RuntimeError("El hash SHA-256 del archivo almacenado no coincide con el original (integridad comprometida)")
        extra = {
            t.code: t.keywords for t in db.query(models.DocumentTypeDef).filter_by(project_id=doc.project_id).all() if t.keywords
        }
        result = extract(data, doc.file_type, project.client_nit if project else None, extra, bool(project and project.allow_ai_processing))
        _persist(db, doc, result)
        doc.processing_status = "PROCESADO"
        doc.vouching_status = models.Status.PENDIENTE
    except UnreadableDocument as exc:
        _clear(db, doc)
        doc.processing_status = "ILEGIBLE"
        doc.vouching_status = models.Status.ILEGIBLE
        doc.error = str(exc)
    except (InvalidXML, Exception) as exc:  # noqa: BLE001 - cualquier fallo queda registrado en el documento
        log.warning("Error procesando documento %s: %s", document_id, mask(str(exc)))
        db.rollback()
        doc = db.get(models.Document, document_id)
        _clear(db, doc)
        doc.processing_status = "ERROR"
        doc.vouching_status = models.Status.ERROR_LECTURA
        doc.error = f"{exc.__class__.__name__}: {exc}"[:2000]
    doc.processed_at = datetime.now(timezone.utc)
    doc.processing_ms = int((time.monotonic() - started) * 1000)
    db.commit()
    return doc


def _clear(db: Session, doc: models.Document) -> None:
    doc.pages.clear()
    doc.fields.clear()
    doc.tables = []
    doc.full_text = None
    db.flush()


def _persist(db: Session, doc: models.Document, result: ExtractionResult) -> None:
    _clear(db, doc)
    now = datetime.now(timezone.utc)
    for p in result.pages:
        doc.pages.append(
            models.DocumentPage(
                page_number=p.number, width=p.width, height=p.height, method=p.method,
                ocr_confidence=p.ocr_confidence, text=p.text, lines=[ln.to_dict() for ln in p.lines],
            )
        )
    for f in result.fields:
        doc.fields.append(
            models.ExtractedField(
                field_name=f.name, value=f.value, normalized_value=f.normalized, page=f.page, bbox=f.bbox,
                evidence_text=(f.evidence_text or "")[:2000], method=f.method, confidence=f.confidence, extracted_at=now,
            )
        )
    doc.page_count = len(result.pages)
    doc.extraction_method = result.method
    doc.ocr_confidence = result.ocr_confidence
    doc.full_text = result.full_text
    doc.tables = result.tables
    doc.doc_type = result.doc_type
    doc.doc_type_confidence = result.doc_type_confidence
    doc.classification_reason = result.classification_reason
    if result.warnings:
        doc.error = " | ".join(result.warnings)[:2000]
