"""Autenticación, recuperación y cambio de contraseña."""
from __future__ import annotations

import logging
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage

from fastapi import APIRouter, Depends, HTTPException, Request
from sqlalchemy.orm import Session

from .. import models, schemas
from ..audit import log_event
from ..config import APP_NAME, get_settings
from ..database import get_db
from ..security import (
    PERMISSIONS,
    client_ip,
    create_access_token,
    get_current_user,
    hash_password,
    hash_token,
    new_reset_token,
    password_problems,
    verify_password,
)

router = APIRouter(prefix="/auth", tags=["Autenticación"])
log = logging.getLogger(__name__)


def _aware(dt):
    return dt.replace(tzinfo=timezone.utc) if dt is not None and dt.tzinfo is None else dt


@router.post("/login")
def login(body: schemas.LoginIn, request: Request, db: Session = Depends(get_db)):
    s = get_settings()
    email = body.email.strip().lower()
    user = db.query(models.User).filter(models.User.email == email).first()
    now = datetime.now(timezone.utc)
    ip = client_ip(request)
    if user and user.locked_until and _aware(user.locked_until) > now:
        log_event(db, "LOGIN_BLOQUEADO", user, ip=ip, success=False)
        raise HTTPException(423, "Cuenta bloqueada temporalmente por intentos fallidos. Intente más tarde.")
    if not user or not user.is_active or not verify_password(body.password, user.password_hash):
        if user:
            user.failed_attempts += 1
            if user.failed_attempts >= s.max_login_attempts:
                user.locked_until = now + timedelta(minutes=s.lockout_minutes)
                user.failed_attempts = 0
        log_event(db, "LOGIN_FALLIDO", user, detail={"email": email}, ip=ip, success=False)
        raise HTTPException(401, "Credenciales inválidas")
    user.failed_attempts = 0
    user.locked_until = None
    user.last_login_at = now
    log_event(db, "LOGIN", user, ip=ip)
    return {
        "access_token": create_access_token(user),
        "token_type": "bearer",
        "expires_in": s.access_token_minutes * 60,
        "user": schemas.UserOut.model_validate(user).model_dump(mode="json"),
        "permisos": sorted(PERMISSIONS.get(user.role, set())),
    }


@router.get("/me")
def me(user: models.User = Depends(get_current_user)):
    return {**schemas.UserOut.model_validate(user).model_dump(mode="json"), "permisos": sorted(PERMISSIONS.get(user.role, set()))}


@router.post("/logout")
def logout(request: Request, user: models.User = Depends(get_current_user), db: Session = Depends(get_db)):
    log_event(db, "LOGOUT", user, ip=client_ip(request))
    return {"ok": True}


def _send_reset_email(to: str, token: str) -> bool:
    s = get_settings()
    if not s.smtp_host:
        return False
    msg = EmailMessage()
    msg["Subject"] = f"{APP_NAME} — Recuperación de contraseña"
    msg["From"] = s.smtp_from
    msg["To"] = to
    msg.set_content(
        f"Se solicitó restablecer su contraseña en {APP_NAME}.\n\n"
        f"Enlace (válido {s.reset_token_minutes} minutos): {s.public_url}/#/restablecer?token={token}\n\n"
        "Si usted no lo solicitó, ignore este mensaje."
    )
    try:
        with smtplib.SMTP(s.smtp_host, s.smtp_port, timeout=15) as smtp:
            smtp.starttls()
            if s.smtp_user:
                smtp.login(s.smtp_user, s.smtp_password)
            smtp.send_message(msg)
        return True
    except (OSError, smtplib.SMTPException):
        log.warning("No fue posible enviar el correo de recuperación")
        return False


@router.post("/password/forgot")
def forgot(body: schemas.PasswordForgotIn, request: Request, db: Session = Depends(get_db)):
    s = get_settings()
    user = db.query(models.User).filter(models.User.email == body.email.strip().lower(), models.User.is_active.is_(True)).first()
    out = {"mensaje": "Si el correo existe, recibirá instrucciones para restablecer la contraseña."}
    if user:
        token, token_hash = new_reset_token()
        db.add(models.PasswordResetToken(user_id=user.id, token_hash=token_hash, expires_at=datetime.now(timezone.utc) + timedelta(minutes=s.reset_token_minutes)))
        log_event(db, "RECUPERACION_SOLICITADA", user, ip=client_ip(request))
        sent = _send_reset_email(user.email, token)
        if not sent and not s.is_production:
            out["token_desarrollo"] = token  # solo en desarrollo, sin SMTP configurado
    return out


@router.post("/password/reset")
def reset(body: schemas.PasswordResetIn, request: Request, db: Session = Depends(get_db)):
    rec = db.query(models.PasswordResetToken).filter_by(token_hash=hash_token(body.token)).first()
    if not rec or rec.used_at or _aware(rec.expires_at) < datetime.now(timezone.utc):
        raise HTTPException(400, "Token inválido o expirado")
    problems = password_problems(body.new_password)
    if problems:
        raise HTTPException(422, "Contraseña débil: " + ", ".join(problems))
    user = db.get(models.User, rec.user_id)
    user.password_hash = hash_password(body.new_password)
    user.must_change_password = False
    user.failed_attempts = 0
    user.locked_until = None
    rec.used_at = datetime.now(timezone.utc)
    log_event(db, "CONTRASENA_RESTABLECIDA", user, ip=client_ip(request))
    return {"mensaje": "Contraseña actualizada"}


@router.post("/password/change")
def change(body: schemas.PasswordChangeIn, request: Request, user: models.User = Depends(get_current_user), db: Session = Depends(get_db)):
    if not verify_password(body.current_password, user.password_hash):
        raise HTTPException(400, "La contraseña actual no es correcta")
    problems = password_problems(body.new_password)
    if problems:
        raise HTTPException(422, "Contraseña débil: " + ", ".join(problems))
    user.password_hash = hash_password(body.new_password)
    user.must_change_password = False
    log_event(db, "CONTRASENA_CAMBIADA", user, ip=client_ip(request))
    return {"mensaje": "Contraseña actualizada"}
