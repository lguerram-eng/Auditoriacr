"""Proyectos de vouching, tablero, configuración, retención e historial."""
from __future__ import annotations

from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.encoders import jsonable_encoder
from sqlalchemy.orm import Session

from .. import models, schemas
from ..audit import log_event
from ..database import get_db
from ..models import Role
from ..security import client_ip, get_current_user, has_permission, project_for_user, require
from ..services import dashboard, storage
from ..services.matching import reconcile
from ..services.settings_defaults import CRITERION_LABELS, DEFAULT_PARAMETERS, PRIORITY_LABELS, merged_parameters, validate_parameters

router = APIRouter(tags=["Proyectos"])


def _out(p: models.Project, db: Session) -> dict:
    d = schemas.ProjectOut.model_validate(p).model_dump(mode="json")
    d["member_ids"] = [m.user_id for m in p.members]
    d["partidas"] = db.query(models.ReferenceItem).filter_by(project_id=p.id).count()
    d["documentos"] = db.query(models.Document).filter_by(project_id=p.id).count()
    return d


@router.get("/projects")
def list_projects(user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    q = db.query(models.Project)
    if user.role != Role.ADMIN:
        q = q.join(models.ProjectMember).filter(models.ProjectMember.user_id == user.id)
    return [_out(p, db) for p in q.order_by(models.Project.created_at.desc()).all()]


@router.post("/projects", status_code=201)
def create_project(body: schemas.ProjectIn, request: Request, user: models.User = Depends(require("proyecto.crear")), db: Session = Depends(get_db)):
    if db.query(models.Project).filter_by(code=body.code).first():
        raise HTTPException(409, "Ya existe un proyecto con ese código")
    p = models.Project(
        code=body.code, name=body.name, client_name=body.client_name, client_nit=body.client_nit, period=body.period,
        description=body.description, allow_ai_processing=body.allow_ai_processing, allow_training=False,
        retention_days=body.retention_days, settings={}, created_by=user.id,
    )
    db.add(p)
    db.flush()
    members = set(body.member_ids) | {user.id}
    for uid in members:
        if db.get(models.User, uid):
            db.add(models.ProjectMember(project_id=p.id, user_id=uid))
    log_event(db, "PROYECTO_CREADO", user, p.id, "project", p.id, {"codigo": p.code, "cliente": p.client_name}, client_ip(request))
    db.refresh(p)
    return _out(p, db)


@router.get("/projects/{project_id}")
def get_project(project_id: int, request: Request, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    log_event(db, "PROYECTO_CONSULTADO", user, p.id, "project", p.id, ip=client_ip(request))
    return _out(p, db)


@router.patch("/projects/{project_id}")
def update_project(project_id: int, body: schemas.ProjectUpdate, request: Request, user: models.User = Depends(require("proyecto.editar")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    changes = body.model_dump(exclude_unset=True)
    members = changes.pop("member_ids", None)
    for k, v in changes.items():
        setattr(p, k, v)
    if changes.get("status") == "CERRADO":
        p.closed_at = datetime.now(timezone.utc)
    if members is not None:
        p.members.clear()
        db.flush()
        for uid in set(members) | {user.id}:
            if db.get(models.User, uid):
                p.members.append(models.ProjectMember(user_id=uid))
        changes["miembros"] = sorted(set(members) | {user.id})
    p.allow_training = False
    log_event(db, "PROYECTO_ACTUALIZADO", user, p.id, "project", p.id, changes, client_ip(request))
    return _out(p, db)


@router.delete("/projects/{project_id}")
def delete_project(project_id: int, request: Request, user: models.User = Depends(require("proyecto.eliminar")), db: Session = Depends(get_db)):
    """Eliminación definitiva del proyecto y de sus archivos cifrados (política de retención)."""
    p = project_for_user(db, project_id, user)
    code = p.code
    storage.delete_project_files(p.id)
    db.delete(p)
    log_event(db, "PROYECTO_ELIMINADO", user, None, "project", project_id, {"codigo": code}, client_ip(request))
    return {"ok": True}


@router.post("/projects/retention/apply")
def apply_retention(request: Request, user: models.User = Depends(require("proyecto.eliminar")), db: Session = Depends(get_db)):
    """Elimina proyectos CERRADOS cuyo plazo de retención venció."""
    now = datetime.now(timezone.utc)
    removed = []
    for p in db.query(models.Project).filter(models.Project.status == "CERRADO", models.Project.closed_at.isnot(None)).all():
        closed = p.closed_at if p.closed_at.tzinfo else p.closed_at.replace(tzinfo=timezone.utc)
        if closed + timedelta(days=p.retention_days) < now:
            removed.append(p.code)
            storage.delete_project_files(p.id)
            db.delete(p)
    log_event(db, "RETENCION_APLICADA", user, detail={"proyectos_eliminados": removed}, ip=client_ip(request))
    return {"eliminados": removed}


@router.get("/projects/{project_id}/dashboard")
def project_dashboard(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    data = dashboard.kpis(db, p.id)
    run = db.query(models.MatchRun).filter_by(project_id=p.id).order_by(models.MatchRun.id.desc()).first()
    jobs = db.query(models.Job).filter(models.Job.project_id == p.id, models.Job.status.in_(["EN_COLA", "EN_PROCESO"])).count()
    data["ultima_ejecucion"] = {"id": run.id, "fecha": run.finished_at or run.started_at} if run else None
    data["trabajos_pendientes"] = jobs
    return jsonable_encoder(data)


@router.get("/dashboard")
def global_dashboard(user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    q = db.query(models.Project)
    if user.role != Role.ADMIN:
        q = q.join(models.ProjectMember).filter(models.ProjectMember.user_id == user.id)
    out = []
    for p in q.all():
        k = dashboard.kpis(db, p.id)
        out.append({"proyecto": _out(p, db), **{key: k[key] for key in (
            "total_partidas", "total_documentos", "coincidencias", "coincidencias_tolerancia", "excepciones", "pendientes_revision",
            "partidas_sin_soporte", "valor_poblacion", "valor_soportado", "porcentaje_cobertura", "porcentaje_avance")}})
    return jsonable_encoder(out)


@router.get("/projects/{project_id}/reconciliation")
def reconciliation(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    return jsonable_encoder(reconcile(db, p.id))


@router.get("/projects/{project_id}/settings")
def get_settings_(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    return {
        "parametros": merged_parameters(p.settings),
        "predeterminados": DEFAULT_PARAMETERS,
        "criterios": CRITERION_LABELS,
        "prioridades": PRIORITY_LABELS,
        "permite_ia": p.allow_ai_processing,
        "permite_entrenamiento": False,
        "retencion_dias": p.retention_days,
    }


@router.put("/projects/{project_id}/settings")
def put_settings(project_id: int, body: schemas.SettingsIn, request: Request, user: models.User = Depends(require("configurar")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    merged = merged_parameters(body.parameters)
    errors = validate_parameters(merged)
    if errors:
        raise HTTPException(422, "; ".join(errors))
    before = merged_parameters(p.settings)
    p.settings = {k: v for k, v in merged.items()}
    log_event(db, "CONFIGURACION_ACTUALIZADA", user, p.id, "project", p.id, {"antes": before, "despues": merged}, client_ip(request))
    return {"parametros": merged}


@router.get("/projects/{project_id}/audit")
def project_audit(project_id: int, limit: int = 500, user: models.User = Depends(get_current_user), db: Session = Depends(get_db)):
    if not has_permission(user, "auditoria.proyecto"):
        raise HTTPException(403, "Permiso requerido: auditoria.proyecto")
    p = project_for_user(db, project_id, user)
    rows = db.query(models.AuditLog).filter_by(project_id=p.id).order_by(models.AuditLog.id.desc()).limit(min(limit, 5000)).all()
    return jsonable_encoder([_audit_row(a) for a in rows])


@router.get("/audit")
def global_audit(limit: int = 500, action: str | None = None, user: models.User = Depends(require("auditoria.global")), db: Session = Depends(get_db)):
    q = db.query(models.AuditLog)
    if action:
        q = q.filter(models.AuditLog.action == action)
    return jsonable_encoder([_audit_row(a) for a in q.order_by(models.AuditLog.id.desc()).limit(min(limit, 5000)).all()])


def _audit_row(a: models.AuditLog) -> dict:
    return {"id": a.id, "fecha": a.created_at, "usuario": a.user_email, "accion": a.action, "entidad": a.entity,
            "entidad_id": a.entity_id, "detalle": a.detail, "ip": a.ip, "exito": a.success, "proyecto_id": a.project_id}


@router.get("/projects/{project_id}/review-history")
def review_history(project_id: int, user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    p = project_for_user(db, project_id, user)
    evs = db.query(models.ReviewEvent).filter_by(project_id=p.id).order_by(models.ReviewEvent.id.desc()).limit(5000).all()
    return jsonable_encoder([{
        "id": e.id, "fecha": e.created_at, "usuario": e.user_email, "accion": e.action, "id_muestra": e.sample_id,
        "documento_id": e.document_id, "campo": e.field_name, "valor_anterior": e.old_value, "valor_nuevo": e.new_value, "comentario": e.comment,
    } for e in evs])
