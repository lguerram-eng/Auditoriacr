"""Punto de entrada de la API de Muenra Vouching (FastAPI)."""
from __future__ import annotations

import logging
import secrets
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles

from . import database, models
from .api import auth, documents, projects, results, users
from .audit import configure_logging
from .config import APP_NAME, APP_VERSION, get_settings
from .security import hash_password

log = logging.getLogger("muenra")

API_DESCRIPTION = """
**Muenra Vouching** — plataforma de vouching para auditoría financiera.

Flujo: crear proyecto → importar Excel de referencia → cargar documentos →
procesar (XML / PDF / OCR / IA opcional) → motor de coincidencias → revisión humana → exportación.

Autenticación: `POST /api/auth/login` devuelve un token Bearer (JWT).
"""


def bootstrap(db) -> None:
    """Crea el administrador inicial si la base de datos no tiene usuarios."""
    if db.query(models.User).count():
        return
    s = get_settings()
    password = s.bootstrap_admin_password or secrets.token_urlsafe(12) + "aA1!"
    db.add(models.User(email=s.bootstrap_admin_email.lower(), full_name="Administrador", role=models.Role.ADMIN,
                       password_hash=hash_password(password), must_change_password=not bool(s.bootstrap_admin_password)))
    db.commit()
    if not s.bootstrap_admin_password:
        # Se imprime directamente (no por el logger enmascarado) una única vez.
        print(f"\n[{APP_NAME}] Administrador inicial: {s.bootstrap_admin_email} / contraseña temporal: {password}\n", flush=True)


@asynccontextmanager
async def lifespan(app: FastAPI):
    configure_logging()
    s = get_settings()
    if not s.is_production:
        models.Base.metadata.create_all(database.engine)
    db = database.SessionLocal()
    try:
        bootstrap(db)
    finally:
        db.close()
    if s.inline_worker and s.env != "test":
        from .worker import start_inline

        start_inline()
        log.info("Trabajador en línea iniciado")
    yield


def create_app() -> FastAPI:
    s = get_settings()
    app = FastAPI(
        title=f"{APP_NAME} API",
        version=APP_VERSION,
        description=API_DESCRIPTION,
        lifespan=lifespan,
        docs_url="/api/docs",
        redoc_url="/api/redoc",
        openapi_url="/api/openapi.json",
    )
    app.add_middleware(
        CORSMiddleware,
        allow_origins=[o.strip() for o in s.cors_origins.split(",") if o.strip()],
        allow_credentials=False,
        allow_methods=["GET", "POST", "PUT", "PATCH", "DELETE"],
        allow_headers=["Authorization", "Content-Type"],
    )

    @app.middleware("http")
    async def security_headers(request: Request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
        if not request.url.path.startswith("/api/docs") and not request.url.path.startswith("/api/redoc"):
            response.headers["Content-Security-Policy"] = (
                "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
                "font-src 'self' https://fonts.gstatic.com; script-src 'self'; connect-src 'self'; frame-ancestors 'none'"
            )
        if s.is_production:
            response.headers["Strict-Transport-Security"] = "max-age=63072000; includeSubDomains"
        if request.url.path.startswith("/api/"):
            response.headers.setdefault("Cache-Control", "no-store")
        return response

    @app.exception_handler(Exception)
    async def unhandled(request: Request, exc: Exception):  # pragma: no cover
        log.exception("Error no controlado en %s", request.url.path)
        return JSONResponse(status_code=500, content={"detail": "Error interno. El incidente quedó registrado."})

    for r in (auth.router, users.router, projects.router, documents.router, results.router):
        app.include_router(r, prefix="/api")

    @app.get("/api/health", tags=["Sistema"])
    def health():
        from .services.extraction.readers import ocr_available
        from .services.extraction.llm import ai_enabled

        return {"aplicativo": APP_NAME, "version": APP_VERSION, "estado": "ok", "ocr": ocr_available(),
                "ia_configurada": ai_enabled(True), "base_de_datos": database.engine.dialect.name}

    frontend = Path(s.frontend_dir)
    if frontend.exists():
        app.mount("/", StaticFiles(directory=frontend, html=True), name="frontend")
    return app


app = create_app()
