"""Carga el escenario de demostración FICTICIO (usuarios, proyecto, Excel y documentos).

Uso:  python -m app.seed  [--sin-procesar] [--forzar]
Credenciales de demostración (solo para entornos locales):
  admin@muenra.local / Demo#Muenra2026   (administrador)
  auditor@muenra.local / Demo#Muenra2026 (auditor)
  revisor@muenra.local / Demo#Muenra2026 (revisor)
  consulta@muenra.local / Demo#Muenra2026 (consulta)
"""
from __future__ import annotations

import sys

from . import database, models
from .audit import configure_logging
from .config import get_settings
from .security import hash_password
from .services import storage
from .services.excel_import import import_workbook
from .services.file_security import DOCUMENT_TYPES, validate_file
from .worker import KIND_DOCUMENT, KIND_MATCH, drain, enqueue

DEMO_PASSWORD = "Demo#Muenra2026"


def seed(process: bool = True, force: bool = False) -> int:
    from demo.generate_demo import build_reference_workbook, scenario

    if get_settings().is_production and not force:
        raise SystemExit("Entorno de producción: use 'python -m app.seed --forzar' solo en instalaciones de demostración")
    models.Base.metadata.create_all(database.engine)
    db = database.SessionLocal()
    try:
        users = {}
        for email, name, role in (
            ("admin@muenra.local", "Administración Demo", models.Role.ADMIN),
            ("auditor@muenra.local", "Auditora Demo", models.Role.AUDITOR),
            ("revisor@muenra.local", "Revisor Demo", models.Role.REVISOR),
            ("consulta@muenra.local", "Consulta Demo", models.Role.CONSULTA),
        ):
            u = db.query(models.User).filter_by(email=email).first()
            if not u:
                u = models.User(email=email, full_name=name, role=role, password_hash=hash_password(DEMO_PASSWORD))
                db.add(u)
            else:
                u.password_hash, u.is_active, u.must_change_password = hash_password(DEMO_PASSWORD), True, False
            users[role] = u
        db.flush()
        project = db.query(models.Project).filter_by(code="DEMO-2026").first()
        if project:
            print("El proyecto DEMO-2026 ya existe; no se vuelve a cargar.")
            db.commit()
            return project.id
        project = models.Project(code="DEMO-2026", name="Vouching de compras marzo 2026 (ficticio)", client_name="Cliente Demostración S.A.S.",
                                 period="2026-03", description="Escenario ficticio con todos los casos de prueba", settings={},
                                 created_by=users[models.Role.AUDITOR].id)
        db.add(project)
        db.flush()
        for u in users.values():
            db.add(models.ProjectMember(project_id=project.id, user_id=u.id))
        rows, docs, expected = scenario()
        data = build_reference_workbook(rows, expected)
        import_workbook(db, project, "referencia_vouching_demo.xlsx", data, storage.sha256(data), users[models.Role.AUDITOR])
        for name, content in docs.items():
            chk = validate_file(name, content, DOCUMENT_TYPES)
            sha = storage.sha256(content)
            existing = db.query(models.Document).filter_by(project_id=project.id, sha256=sha).first()
            d = models.Document(project_id=project.id, filename=name, stored_path=storage.save(project.id, content, chk.file_type), sha256=sha,
                                size_bytes=len(content), file_type=chk.file_type, mime=chk.mime, av_status=chk.av_status,
                                uploaded_by=users[models.Role.AUDITOR].id, duplicate_of_id=existing.id if existing else None)
            db.add(d)
            db.flush()
            enqueue(db, project.id, KIND_DOCUMENT, {"document_id": d.id}, commit=False)
        enqueue(db, project.id, KIND_MATCH, {"user_id": users[models.Role.AUDITOR].id}, commit=False)
        db.commit()
        pid = project.id
    finally:
        db.close()
    if process:
        n = drain()
        print(f"Procesados {n} trabajos")
    print(f"Proyecto de demostración DEMO-2026 cargado (id {pid}). Usuarios *@muenra.local / {DEMO_PASSWORD}")
    return pid


if __name__ == "__main__":
    configure_logging()
    seed(process="--sin-procesar" not in sys.argv, force="--forzar" in sys.argv)
