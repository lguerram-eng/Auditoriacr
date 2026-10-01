"""Almacenamiento seguro de archivos con cifrado en reposo (Fernet / AES-128-CBC + HMAC).

Los archivos se guardan separados por proyecto. El hash SHA-256 se calcula sobre
el contenido original (antes de cifrar) y se conserva como evidencia.
"""
from __future__ import annotations

import base64
import hashlib
import shutil
import uuid
from functools import lru_cache
from pathlib import Path

from cryptography.fernet import Fernet

from ..config import get_settings


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@lru_cache
def _fernet() -> Fernet:
    s = get_settings()
    key = s.storage_encryption_key
    if not key:
        key = base64.urlsafe_b64encode(hashlib.sha256(("muenra-storage:" + s.secret_key).encode()).digest()).decode()
    return Fernet(key.encode() if isinstance(key, str) else key)


def _root() -> Path:
    root = Path(get_settings().storage_dir)
    root.mkdir(parents=True, exist_ok=True)
    return root


def save(project_id: int, data: bytes, suffix: str = "bin") -> str:
    rel = Path(f"proyecto_{project_id}") / uuid.uuid4().hex[:2] / f"{uuid.uuid4().hex}.{suffix}.enc"
    path = _root() / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_fernet().encrypt(data))
    return rel.as_posix()


def load(rel_path: str) -> bytes:
    path = (_root() / rel_path).resolve()
    if _root().resolve() not in path.parents:
        raise PermissionError("Ruta fuera del almacenamiento")
    return _fernet().decrypt(path.read_bytes())


def delete(rel_path: str) -> None:
    path = (_root() / rel_path).resolve()
    if _root().resolve() in path.parents and path.exists():
        path.unlink()


def delete_project_files(project_id: int) -> None:
    shutil.rmtree(_root() / f"proyecto_{project_id}", ignore_errors=True)
