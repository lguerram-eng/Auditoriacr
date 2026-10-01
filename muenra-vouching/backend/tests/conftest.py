"""Configuración de pruebas: base SQLite temporal, almacenamiento temporal y sin trabajador en hilo."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

_TMP = Path(tempfile.mkdtemp(prefix="muenra_test_"))
os.environ.update({
    "MUENRA_ENV": "test",
    "MUENRA_DATABASE_URL": f"sqlite:///{_TMP / 'test.db'}",
    "MUENRA_STORAGE_DIR": str(_TMP / "storage"),
    "MUENRA_INLINE_WORKER": "false",
    "MUENRA_BOOTSTRAP_ADMIN_EMAIL": "admin@muenra.test",
    "MUENRA_BOOTSTRAP_ADMIN_PASSWORD": "Admin#Prueba2026",
    "MUENRA_SECRET_KEY": "clave-de-pruebas-unicamente-no-usar-en-produccion-0123456789",
    "MUENRA_LLM_ENABLED": "false",
    "MUENRA_CLAMAV_HOST": "",
})

import pytest  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from app.services.extraction.readers import ocr_available  # noqa: E402

ADMIN = ("admin@muenra.test", "Admin#Prueba2026")


@pytest.fixture(scope="session")
def client():
    from app.main import app

    with TestClient(app) as c:
        yield c


def login(client, email, password) -> dict:
    r = client.post("/api/auth/login", json={"email": email, "password": password})
    assert r.status_code == 200, r.text
    return {"Authorization": f"Bearer {r.json()['access_token']}"}


@pytest.fixture(scope="session")
def admin_headers(client):
    return login(client, *ADMIN)


@pytest.fixture(scope="session")
def demo_data():
    from demo.generate_demo import build_reference_workbook, scenario

    rows, docs, expected = scenario()
    return {"rows": rows, "docs": docs, "expected": expected, "xlsx": build_reference_workbook(rows, expected)}


requires_ocr = pytest.mark.skipif(not ocr_available(), reason="Tesseract no está instalado")
