"""Configuración central de Muenra Vouching.

Todos los valores sensibles se leen de variables de entorno (ver .env.example).
Nunca se deben escribir secretos en el código fuente.
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict

APP_NAME = "Muenra Vouching"
APP_VERSION = "1.0.0"

BASE_DIR = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_prefix="MUENRA_", extra="ignore")

    env: str = Field("development", description="development | production | test")
    database_url: str = f"sqlite:///{BASE_DIR / 'data' / 'muenra.db'}"
    storage_dir: Path = BASE_DIR / "data" / "storage"
    frontend_dir: Path = BASE_DIR.parent / "frontend"

    # Seguridad
    secret_key: str = "cambie-esta-clave-en-produccion-por-una-aleatoria-de-64-caracteres"
    # Clave Fernet (base64 urlsafe, 32 bytes) para cifrado en reposo de los archivos.
    # Si está vacía en desarrollo se deriva de secret_key; en producción es obligatoria.
    storage_encryption_key: str = ""
    access_token_minutes: int = 60
    reset_token_minutes: int = 30
    max_login_attempts: int = 5
    lockout_minutes: int = 15
    cors_origins: str = "http://localhost:8080,http://localhost:8000"

    # Usuario administrador inicial (solo se crea si no existe ningún usuario)
    bootstrap_admin_email: str = "admin@muenra.local"
    bootstrap_admin_password: str = ""

    # Archivos
    max_file_mb: int = 50
    max_files_per_upload: int = 500
    clamav_host: str = ""
    clamav_port: int = 3310
    antivirus_required: bool = False

    # OCR
    ocr_languages: str = "spa+eng"
    ocr_dpi: int = 300
    ocr_min_text_chars: int = 40  # por página: por debajo se considera escaneado

    # Cola de procesamiento
    inline_worker: bool = True  # hilo en el mismo proceso (desarrollo); en Docker se usa el servicio worker
    worker_poll_seconds: float = 2.0

    # Inteligencia artificial (opcional). Se usa solo si el proyecto lo autoriza.
    llm_enabled: bool = False
    llm_model: str = "claude-opus-5-5"
    anthropic_api_key: str = ""

    # Retención
    default_retention_days: int = 3650

    # Correo para recuperación de contraseña
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    smtp_from: str = "no-reply@muenra.local"
    public_url: str = "http://localhost:8000"

    @property
    def is_production(self) -> bool:
        return self.env == "production"


@lru_cache
def get_settings() -> Settings:
    s = Settings()
    if s.is_production:
        if "cambie-esta-clave" in s.secret_key or len(s.secret_key) < 32:
            raise RuntimeError("MUENRA_SECRET_KEY debe configurarse con un valor aleatorio fuerte en producción")
        if not s.storage_encryption_key:
            raise RuntimeError("MUENRA_STORAGE_ENCRYPTION_KEY es obligatoria en producción")
    return s
