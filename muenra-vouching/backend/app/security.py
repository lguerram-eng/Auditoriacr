"""Autenticación (JWT + bcrypt), roles, permisos y aislamiento por proyecto."""
from __future__ import annotations

import hashlib
import re
import secrets
import uuid
from datetime import datetime, timedelta, timezone

import bcrypt
import jwt
from fastapi import Depends, HTTPException, Request, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy.orm import Session

from . import models
from .config import get_settings
from .database import get_db
from .models import Role

ALGORITHM = "HS256"

PERMISSIONS: dict[str, set[str]] = {
    Role.ADMIN: {
        "ver", "exportar", "revisar", "proyecto.crear", "proyecto.editar", "proyecto.eliminar", "importar", "cargar",
        "procesar", "configurar", "usuarios", "auditoria.global", "auditoria.proyecto",
    },
    Role.AUDITOR: {"ver", "exportar", "revisar", "proyecto.crear", "proyecto.editar", "importar", "cargar", "procesar", "configurar", "auditoria.proyecto"},
    Role.REVISOR: {"ver", "exportar", "revisar"},
    Role.CONSULTA: {"ver", "exportar"},
}

bearer = HTTPBearer(auto_error=False)


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8")[:72], bcrypt.gensalt(rounds=12)).decode()


def verify_password(password: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8")[:72], hashed.encode())
    except ValueError:
        return False


def password_problems(password: str) -> list[str]:
    problems = []
    if len(password) < 10:
        problems.append("mínimo 10 caracteres")
    if not re.search(r"[A-ZÁÉÍÓÚÑ]", password):
        problems.append("al menos una mayúscula")
    if not re.search(r"[a-záéíóúñ]", password):
        problems.append("al menos una minúscula")
    if not re.search(r"\d", password):
        problems.append("al menos un número")
    if not re.search(r"[^\w\s]", password):
        problems.append("al menos un símbolo")
    return problems


def create_access_token(user: models.User) -> str:
    s = get_settings()
    now = datetime.now(timezone.utc)
    payload = {
        "sub": str(user.id),
        "role": user.role,
        "iat": now,
        "exp": now + timedelta(minutes=s.access_token_minutes),
        "jti": uuid.uuid4().hex,
        "iss": "muenra-vouching",
    }
    return jwt.encode(payload, s.secret_key, algorithm=ALGORITHM)


def new_reset_token() -> tuple[str, str]:
    token = secrets.token_urlsafe(32)
    return token, hashlib.sha256(token.encode()).hexdigest()


def hash_token(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


def client_ip(request: Request | None) -> str | None:
    if request is None:
        return None
    fwd = request.headers.get("x-forwarded-for")
    return fwd.split(",")[0].strip() if fwd else (request.client.host if request.client else None)


def get_current_user(
    creds: HTTPAuthorizationCredentials | None = Depends(bearer), db: Session = Depends(get_db)
) -> models.User:
    if creds is None:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "No autenticado", headers={"WWW-Authenticate": "Bearer"})
    try:
        payload = jwt.decode(creds.credentials, get_settings().secret_key, algorithms=[ALGORITHM], issuer="muenra-vouching")
    except jwt.PyJWTError:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Sesión inválida o expirada", headers={"WWW-Authenticate": "Bearer"})
    user = db.get(models.User, int(payload["sub"]))
    if not user or not user.is_active:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Usuario inactivo")
    return user


def has_permission(user: models.User, perm: str) -> bool:
    return perm in PERMISSIONS.get(user.role, set())


def require(perm: str):
    def dep(user: models.User = Depends(get_current_user)) -> models.User:
        if not has_permission(user, perm):
            raise HTTPException(status.HTTP_403_FORBIDDEN, f"Permiso requerido: {perm}")
        return user

    return dep


def project_for_user(db: Session, project_id: int, user: models.User) -> models.Project:
    """Devuelve el proyecto si el usuario tiene acceso (separación por cliente/encargo)."""
    project = db.get(models.Project, project_id)
    if project is None:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Proyecto no encontrado")
    if user.role == Role.ADMIN:
        return project
    member = db.query(models.ProjectMember).filter_by(project_id=project_id, user_id=user.id).first()
    if not member:
        # 404 en lugar de 403 para no revelar la existencia de proyectos ajenos
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Proyecto no encontrado")
    return project
