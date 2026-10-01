"""Usuarios y permisos (solo administrador)."""
from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session

from .. import models, schemas
from ..audit import log_event
from ..database import get_db
from ..models import Role
from ..security import PERMISSIONS, client_ip, hash_password, password_problems, require

router = APIRouter(prefix="/users", tags=["Usuarios y permisos"])


@router.get("")
def list_users(user: models.User = Depends(require("ver")), db: Session = Depends(get_db)):
    users = db.query(models.User).order_by(models.User.email).all()
    if user.role != Role.ADMIN:
        # Los demás roles solo ven nombre y correo (para asignar miembros)
        return [{"id": u.id, "email": u.email, "full_name": u.full_name, "role": u.role} for u in users if u.is_active]
    return [schemas.UserOut.model_validate(u).model_dump(mode="json") for u in users]


@router.get("/roles")
def roles(_: models.User = Depends(require("ver"))):
    return {"roles": list(Role.ALL), "permisos": {r: sorted(p) for r, p in PERMISSIONS.items()}}


@router.post("", status_code=201)
def create_user(body: schemas.UserCreate, request: Request, admin: models.User = Depends(require("usuarios")), db: Session = Depends(get_db)):
    if body.role not in Role.ALL:
        raise HTTPException(422, "Rol inválido")
    problems = password_problems(body.password)
    if problems:
        raise HTTPException(422, "Contraseña débil: " + ", ".join(problems))
    email = body.email.strip().lower()
    if db.query(models.User).filter_by(email=email).first():
        raise HTTPException(409, "Ya existe un usuario con ese correo")
    u = models.User(email=email, full_name=body.full_name, role=body.role, password_hash=hash_password(body.password), must_change_password=True)
    db.add(u)
    db.flush()
    log_event(db, "USUARIO_CREADO", admin, entity="user", entity_id=u.id, detail={"email": email, "rol": body.role}, ip=client_ip(request))
    return schemas.UserOut.model_validate(u).model_dump(mode="json")


@router.patch("/{user_id}")
def update_user(user_id: int, body: schemas.UserUpdate, request: Request, admin: models.User = Depends(require("usuarios")), db: Session = Depends(get_db)):
    u = db.get(models.User, user_id)
    if not u:
        raise HTTPException(404, "Usuario no encontrado")
    changes = {}
    if body.full_name is not None:
        u.full_name = body.full_name
        changes["nombre"] = body.full_name
    if body.role is not None:
        if body.role not in Role.ALL:
            raise HTTPException(422, "Rol inválido")
        if u.id == admin.id and body.role != Role.ADMIN:
            raise HTTPException(400, "No puede quitarse a sí mismo el rol de administrador")
        changes["rol"] = f"{u.role} -> {body.role}"
        u.role = body.role
    if body.is_active is not None:
        if u.id == admin.id and not body.is_active:
            raise HTTPException(400, "No puede desactivarse a sí mismo")
        u.is_active = body.is_active
        changes["activo"] = body.is_active
    if body.password:
        problems = password_problems(body.password)
        if problems:
            raise HTTPException(422, "Contraseña débil: " + ", ".join(problems))
        u.password_hash = hash_password(body.password)
        u.must_change_password = True
        changes["password"] = "restablecida"
    log_event(db, "USUARIO_ACTUALIZADO", admin, entity="user", entity_id=u.id, detail=changes, ip=client_ip(request))
    return schemas.UserOut.model_validate(u).model_dump(mode="json")
