"""Pruebas de autenticación, permisos por rol, aislamiento por proyecto y seguridad de archivos."""
import pytest

from app.audit import mask
from app.services.file_security import DOCUMENT_TYPES, EICAR, FileRejected, sanitize_filename, validate_file

from .conftest import login

PWD = "Clave#Segura2026"


@pytest.fixture(scope="module")
def users(client, admin_headers):
    out = {}
    for role in ("auditor", "revisor", "consulta"):
        email = f"{role}@muenra.test"
        r = client.post("/api/users", headers=admin_headers, json={"email": email, "full_name": role.title(), "role": role, "password": PWD})
        assert r.status_code == 201, r.text
        out[role] = {"id": r.json()["id"], "headers": login(client, email, PWD)}
    return out


@pytest.fixture(scope="module")
def project(client, users):
    r = client.post("/api/projects", headers=users["auditor"]["headers"], json={
        "code": "SEC-1", "name": "Permisos", "client_name": "Cliente", "member_ids": [users["revisor"]["id"], users["consulta"]["id"]]})
    assert r.status_code == 201, r.text
    return r.json()["id"]


def test_requires_authentication(client):
    assert client.get("/api/projects").status_code == 401
    assert client.get("/api/projects", headers={"Authorization": "Bearer invalido"}).status_code == 401


def test_role_permissions(client, users, project):
    pdf = {"files": ("a.pdf", b"%PDF-1.4\n%%EOF", "application/pdf")}
    # consulta: solo lectura y exportación
    h = users["consulta"]["headers"]
    assert client.get(f"/api/projects/{project}", headers=h).status_code == 200
    assert client.post(f"/api/projects/{project}/documents", headers=h, files=pdf).status_code == 403
    assert client.post("/api/projects", headers=h, json={"code": "X1", "name": "x", "client_name": "y"}).status_code == 403
    assert client.get(f"/api/projects/{project}/export", headers=h).status_code == 200
    # revisor: revisa pero no importa ni configura
    h = users["revisor"]["headers"]
    assert client.post(f"/api/projects/{project}/imports", headers=h, files={"file": ("r.xlsx", b"x", "a/b")}).status_code == 403
    assert client.put(f"/api/projects/{project}/settings", headers=h, json={"parameters": {}}).status_code == 403
    # auditor: no administra usuarios ni ve la auditoría global
    h = users["auditor"]["headers"]
    assert client.post("/api/users", headers=h, json={"email": "z@z.co", "full_name": "Zz", "role": "consulta", "password": PWD}).status_code == 403
    assert client.get("/api/audit", headers=h).status_code == 403
    assert client.get(f"/api/projects/{project}/audit", headers=h).status_code == 200


def test_project_isolation(client, admin_headers, users, project):
    other = client.post("/api/projects", headers=admin_headers, json={"code": "SEC-2", "name": "Ajeno", "client_name": "Otro"}).json()["id"]
    h = users["auditor"]["headers"]
    assert client.get(f"/api/projects/{other}", headers=h).status_code == 404
    assert client.get(f"/api/projects/{other}/results", headers=h).status_code == 404
    assert other not in [p["id"] for p in client.get("/api/projects", headers=h).json()]
    assert project in [p["id"] for p in client.get("/api/projects", headers=users["consulta"]["headers"]).json()]


def test_settings_update_and_validation(client, users, project):
    h = users["auditor"]["headers"]
    r = client.put(f"/api/projects/{project}/settings", headers=h, json={"parameters": {"TOLERANCIA_VALOR": 0.05, "PESOS": {"nit": 40}}})
    assert r.status_code == 200 and r.json()["parametros"]["TOLERANCIA_VALOR"] == 0.05
    assert client.put(f"/api/projects/{project}/settings", headers=h, json={"parameters": {"TOLERANCIA_VALOR": 3}}).status_code == 422


def test_lockout_after_failed_attempts(client, admin_headers):
    client.post("/api/users", headers=admin_headers, json={"email": "bloqueo@muenra.test", "full_name": "Bloqueo", "role": "consulta", "password": PWD})
    for _ in range(5):
        assert client.post("/api/auth/login", json={"email": "bloqueo@muenra.test", "password": "malaClave1!"}).status_code == 401
    assert client.post("/api/auth/login", json={"email": "bloqueo@muenra.test", "password": PWD}).status_code == 423


def test_password_recovery_flow(client, admin_headers):
    client.post("/api/users", headers=admin_headers, json={"email": "olvido@muenra.test", "full_name": "Olvido", "role": "consulta", "password": PWD})
    r = client.post("/api/auth/password/forgot", json={"email": "olvido@muenra.test"})
    token = r.json()["token_desarrollo"]
    assert client.post("/api/auth/password/forgot", json={"email": "noexiste@muenra.test"}).json().get("token_desarrollo") is None
    assert client.post("/api/auth/password/reset", json={"token": token, "new_password": "debil"}).status_code == 422
    assert client.post("/api/auth/password/reset", json={"token": token, "new_password": "Nueva#Clave2026"}).status_code == 200
    assert client.post("/api/auth/password/reset", json={"token": token, "new_password": "Otra#Clave2026"}).status_code == 400  # un solo uso
    login(client, "olvido@muenra.test", "Nueva#Clave2026")


def test_weak_password_rejected(client, admin_headers):
    r = client.post("/api/users", headers=admin_headers, json={"email": "d@muenra.test", "full_name": "Debil", "role": "consulta", "password": "12345678"})
    assert r.status_code == 422


def test_audit_log_records_access(client, admin_headers):
    actions = {a["accion"] for a in client.get("/api/audit", headers=admin_headers).json()}
    assert {"LOGIN", "LOGIN_FALLIDO", "USUARIO_CREADO", "PROYECTO_CREADO"} <= actions


@pytest.mark.parametrize("name,data,msg", [
    ("virus.pdf", b"%PDF-1.4 " + EICAR, "EICAR"),
    ("falso.pdf", b"MZ\x90\x00 ejecutable", "no coincide"),
    ("script.exe", b"MZ", "no permitido"),
    ("activo.pdf", b"%PDF-1.4 /JavaScript (app.alert(1))", "contenido activo"),
    ("vacio.png", b"", "vacío"),
    ("xxe.xml", b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x "y">]><a/>', "entidades"),
])
def test_malicious_files_rejected(name, data, msg):
    with pytest.raises(FileRejected, match=msg):
        validate_file(name, data, DOCUMENT_TYPES)


def test_valid_file_and_filename_sanitization():
    assert validate_file("a.PDF", b"%PDF-1.4\n%%EOF", DOCUMENT_TYPES).file_type == "pdf"
    assert sanitize_filename("../../etc/passwd") == "passwd"
    assert sanitize_filename('a<b>:"c".pdf') == "a_b___c_.pdf"


def test_upload_rejects_and_reports(client, admin_headers):
    pid = client.post("/api/projects", headers=admin_headers, json={"code": "SEC-3", "name": "Carga", "client_name": "Cliente"}).json()["id"]
    files = [("files", ("ok.pdf", b"%PDF-1.4\n%%EOF", "application/pdf")), ("files", ("malo.exe", b"MZ", "application/octet-stream"))]
    r = client.post(f"/api/projects/{pid}/documents?process=false", headers=admin_headers, files=files).json()
    assert len(r["aceptados"]) == 1 and r["rechazados"][0]["archivo"] == "malo.exe"


def test_masking_of_sensitive_data():
    m = mask("Error con NIT 900.123.456-8 cuenta 1234567890 correo persona@empresa.com password=abc123")
    assert "900.123.456" not in m and "1234567890" not in m and "persona@" not in m and "abc123" not in m


def test_storage_is_encrypted_at_rest(client, admin_headers):
    from pathlib import Path

    from app.config import get_settings
    from app.services import storage

    rel = storage.save(999, b"%PDF-1.4 contenido confidencial")
    raw = (Path(get_settings().storage_dir) / rel).read_bytes()
    assert b"confidencial" not in raw
    assert storage.load(rel) == b"%PDF-1.4 contenido confidencial"
    with pytest.raises(PermissionError):
        storage.load("../../etc/passwd")


def test_training_flag_cannot_be_enabled(client, admin_headers):
    r = client.post("/api/projects", headers=admin_headers, json={"code": "SEC-4", "name": "IA", "client_name": "Cliente", "allow_ai_processing": True})
    assert r.json()["allow_training"] is False


def test_security_headers(client):
    r = client.get("/api/health")
    assert r.headers["X-Frame-Options"] == "DENY" and r.headers["X-Content-Type-Options"] == "nosniff"
    assert r.json()["aplicativo"] == "Muenra Vouching"
