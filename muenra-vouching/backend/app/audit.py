"""Registro de auditoría (accesos y cambios) y enmascaramiento de datos sensibles."""
from __future__ import annotations

import json
import logging
import re

from sqlalchemy.orm import Session

from . import models

_PATTERNS = [
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), lambda m: m.group(0)[:2] + "***@***"),
    (re.compile(r"\b\d{1,3}(?:[.,]?\d{3}){2,3}(?:-\d)?\b"), lambda m: "***" + m.group(0)[-3:]),  # NIT / cédulas / cuentas
    (re.compile(r"\b\d{6,}\b"), lambda m: "***" + m.group(0)[-3:]),
    (re.compile(r"(?i)(password|contrase[nñ]a|token|secret|api[_-]?key)\s*[=:]\s*\S+"), lambda m: m.group(1) + "=***"),
]


def mask(text: str | None) -> str:
    """Enmascara correos, NIT, números de cuenta y secretos en registros técnicos."""
    if not text:
        return ""
    out = str(text)
    for pat, repl in _PATTERNS:
        out = pat.sub(repl, out)
    return out


class MaskingFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        try:
            record.msg = mask(record.getMessage())
            record.args = ()
        except Exception:  # pragma: no cover
            pass
        return True


def configure_logging() -> None:
    handler = logging.StreamHandler()
    handler.addFilter(MaskingFilter())
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.handlers = [handler]
    root.setLevel(logging.INFO)


SENSITIVE_KEYS = {"password", "new_password", "token", "secret", "password_hash"}


def _clean(detail: dict | None) -> dict:
    if not detail:
        return {}
    out = {}
    for k, v in detail.items():
        if k in SENSITIVE_KEYS:
            out[k] = "***"
        elif isinstance(v, (str, int, float, bool)) or v is None:
            out[k] = v
        else:
            out[k] = json.loads(json.dumps(v, default=str))
    return out


def log_event(
    db: Session,
    action: str,
    user: models.User | None = None,
    project_id: int | None = None,
    entity: str | None = None,
    entity_id=None,
    detail: dict | None = None,
    ip: str | None = None,
    success: bool = True,
    commit: bool = True,
) -> None:
    db.add(
        models.AuditLog(
            user_id=user.id if user else None,
            user_email=user.email if user else (detail or {}).get("email"),
            project_id=project_id,
            action=action,
            entity=entity,
            entity_id=str(entity_id) if entity_id is not None else None,
            detail=_clean(detail),
            ip=ip,
            success=success,
        )
    )
    if commit:
        db.commit()
