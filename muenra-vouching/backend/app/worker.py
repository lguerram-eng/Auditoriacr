"""Cola de procesamiento asíncrono respaldada en la base de datos.

- En desarrollo (MUENRA_INLINE_WORKER=true) corre como hilo dentro de la API.
- En Docker corre como servicio independiente: ``python -m app.worker``.
- En PostgreSQL usa ``SELECT ... FOR UPDATE SKIP LOCKED`` para permitir varios
  trabajadores en paralelo sin tomar el mismo trabajo dos veces.
"""
from __future__ import annotations

import logging
import threading
import time
from datetime import datetime, timedelta, timezone

from sqlalchemy import select

from . import database, models
from .audit import configure_logging, log_event, mask
from .config import get_settings

log = logging.getLogger("muenra.worker")

KIND_DOCUMENT = "PROCESAR_DOCUMENTO"
KIND_MATCH = "EJECUTAR_VOUCHING"
MAX_ATTEMPTS = 3


def enqueue(db, project_id: int, kind: str, payload: dict | None = None, commit: bool = True) -> models.Job:
    job = models.Job(project_id=project_id, kind=kind, payload=payload or {})
    db.add(job)
    if commit:
        db.commit()
    return job


def _claim(db) -> models.Job | None:
    """Toma el siguiente trabajo: primero documentos; el vouching solo cuando el proyecto no tiene documentos pendientes."""
    for kind in (KIND_DOCUMENT, KIND_MATCH):
        stmt = select(models.Job).where(models.Job.status == "EN_COLA", models.Job.kind == kind).order_by(models.Job.id)
        if database.engine.dialect.name == "postgresql":
            stmt = stmt.with_for_update(skip_locked=True)
        for job in db.execute(stmt.limit(20)).scalars():
            if kind == KIND_MATCH:
                pending_docs = (
                    db.query(models.Job)
                    .filter(models.Job.project_id == job.project_id, models.Job.kind == KIND_DOCUMENT, models.Job.status.in_(["EN_COLA", "EN_PROCESO"]))
                    .count()
                )
                if pending_docs:
                    continue
            job.status = "EN_PROCESO"
            job.started_at = datetime.now(timezone.utc)
            job.attempts += 1
            db.commit()
            return job
    db.rollback()
    return None


def run_job(db, job: models.Job) -> None:
    from .services.extraction.pipeline import process_document
    from .services.matching import run_matching

    try:
        if job.kind == KIND_DOCUMENT:
            process_document(db, int(job.payload["document_id"]))
        elif job.kind == KIND_MATCH:
            user = db.get(models.User, job.payload.get("user_id")) if job.payload.get("user_id") else None
            run = run_matching(db, job.project_id, user)
            log_event(db, "VOUCHING_EJECUTADO", user, job.project_id, "match_run", run.id, {"conciliacion": run.reconciliation})
        job = db.merge(job)
        job.status = "COMPLETADO"
        job.error = None
    except Exception as exc:  # noqa: BLE001
        db.rollback()
        job = db.merge(job)
        log.exception("Trabajo %s falló: %s", job.id, mask(str(exc)))
        job.error = f"{exc.__class__.__name__}: {exc}"[:2000]
        job.status = "EN_COLA" if job.attempts < MAX_ATTEMPTS else "FALLIDO"
    job.finished_at = datetime.now(timezone.utc)
    db.commit()


def work_once() -> bool:
    """Procesa un trabajo. Devuelve True si había trabajo."""
    db = database.SessionLocal()
    try:
        job = _claim(db)
        if job is None:
            return False
        run_job(db, job)
        return True
    finally:
        db.close()


def drain(max_jobs: int = 10_000) -> int:
    """Procesa todos los trabajos pendientes (útil para pruebas y CLI)."""
    n = 0
    while n < max_jobs and work_once():
        n += 1
    return n


def recover_stale(minutes: int = 30) -> None:
    db = database.SessionLocal()
    try:
        limit = datetime.now(timezone.utc) - timedelta(minutes=minutes)
        for job in db.query(models.Job).filter(models.Job.status == "EN_PROCESO"):
            started = job.started_at.replace(tzinfo=timezone.utc) if job.started_at and job.started_at.tzinfo is None else job.started_at
            if started is None or started < limit:
                job.status = "EN_COLA"
        db.commit()
    finally:
        db.close()


_stop = threading.Event()


def loop(stop: threading.Event | None = None) -> None:
    stop = stop or _stop
    poll = get_settings().worker_poll_seconds
    recover_stale()
    while not stop.is_set():
        try:
            if not work_once():
                stop.wait(poll)
        except Exception:  # pragma: no cover - el bucle nunca debe morir
            log.exception("Error en el bucle del trabajador")
            stop.wait(poll)


def start_inline() -> threading.Thread:
    t = threading.Thread(target=loop, name="muenra-worker", daemon=True)
    t.start()
    return t


if __name__ == "__main__":  # pragma: no cover
    configure_logging()
    log.info("Trabajador de Muenra Vouching iniciado")
    loop()
