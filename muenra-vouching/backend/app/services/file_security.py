"""Validación de archivos: tipo real (magic bytes), tamaño, contenido activo y antivirus.

Defensa en profundidad:
1. Lista blanca de extensiones.
2. Verificación de la firma binaria (no se confía en la extensión ni en el MIME del navegador).
3. Límite de tamaño y protección contra "zip bombs" en DOCX/XLSX.
4. Rechazo de PDF con JavaScript/acciones de lanzamiento y de Office con macros.
5. Escaneo con ClamAV (clamd INSTREAM) cuando está configurado; firma EICAR siempre.
"""
from __future__ import annotations

import io
import re
import socket
import struct
import zipfile
from dataclasses import dataclass

from ..config import get_settings

DOCUMENT_TYPES = {"pdf", "xml", "png", "jpg", "jpeg", "tif", "tiff", "docx"}
REFERENCE_TYPES = {"xlsx", "xls", "csv"}

MIME = {
    "pdf": "application/pdf",
    "xml": "application/xml",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
    "csv": "text/csv",
}

EICAR = b"X5O!P%@AP[4\\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!"
MAX_UNZIPPED = 300 * 1024 * 1024


class FileRejected(Exception):
    pass


@dataclass
class FileCheck:
    file_type: str
    mime: str
    av_status: str


def sanitize_filename(name: str) -> str:
    name = (name or "archivo").replace("\\", "/").split("/")[-1]
    name = re.sub(r"[\x00-\x1f<>:\"|?*]", "_", name).strip(" .")
    return name[:200] or "archivo"


def extension(name: str) -> str:
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def sniff(data: bytes) -> str | None:
    head = data[:16]
    if head.startswith(b"%PDF-"):
        return "pdf"
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if head.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if head.startswith(b"II*\x00") or head.startswith(b"MM\x00*"):
        return "tiff"
    if head.startswith(b"PK\x03\x04"):
        return "zip"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "ole"
    stripped = data[:200].lstrip(b"\xef\xbb\xbf \t\r\n")
    if stripped.startswith(b"<?xml") or stripped.startswith(b"<"):
        return "xml"
    return None


def _check_zip(data: bytes, expected_member: str) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            total = sum(i.file_size for i in zf.infolist())
    except zipfile.BadZipFile as exc:
        raise FileRejected("Archivo Office dañado o no válido") from exc
    if total > MAX_UNZIPPED:
        raise FileRejected("El archivo comprimido excede el tamaño descomprimido permitido")
    if expected_member not in names:
        raise FileRejected("El contenido no corresponde a la extensión del archivo")
    if any(n.lower().endswith("vbaproject.bin") for n in names):
        raise FileRejected("Se rechazan documentos Office con macros")


def validate_file(filename: str, data: bytes, allowed: set[str]) -> FileCheck:
    settings = get_settings()
    ext = extension(filename)
    if ext not in allowed:
        raise FileRejected(f"Tipo de archivo no permitido: .{ext or '(sin extensión)'}")
    if not data:
        raise FileRejected("Archivo vacío")
    if len(data) > settings.max_file_mb * 1024 * 1024:
        raise FileRejected(f"El archivo supera el límite de {settings.max_file_mb} MB")
    kind = sniff(data)
    ok = {
        "pdf": kind == "pdf",
        "png": kind == "png",
        "jpg": kind == "jpg",
        "jpeg": kind == "jpg",
        "tif": kind == "tiff",
        "tiff": kind == "tiff",
        "xml": kind == "xml",
        "docx": kind == "zip",
        "xlsx": kind == "zip",
        "xls": kind == "ole",
        "csv": kind in (None, "xml") and b"\x00" not in data[:4096],
    }[ext]
    if not ok:
        raise FileRejected("El contenido del archivo no coincide con su extensión")
    if ext == "docx":
        _check_zip(data, "word/document.xml")
    if ext == "xlsx":
        _check_zip(data, "xl/workbook.xml")
    if ext == "pdf" and re.search(rb"/(JavaScript|JS|Launch|EmbeddedFile)\b", data):
        raise FileRejected("PDF con contenido activo (JavaScript, acciones o archivos embebidos) rechazado por seguridad")
    if ext == "xml" and re.search(rb"<!ENTITY", data[:20000]):
        raise FileRejected("XML con declaraciones de entidades rechazado por seguridad")
    av = scan(data)
    return FileCheck(file_type="jpg" if ext == "jpeg" else "tiff" if ext == "tif" else ext, mime=MIME[ext], av_status=av)


def scan(data: bytes) -> str:
    """Devuelve LIMPIO | NO_ESCANEADO. Lanza FileRejected si se detecta malware."""
    if EICAR in data:
        raise FileRejected("Archivo infectado (firma de prueba EICAR detectada)")
    settings = get_settings()
    if not settings.clamav_host:
        if settings.antivirus_required:
            raise FileRejected("Antivirus requerido pero no configurado")
        return "NO_ESCANEADO"
    try:
        with socket.create_connection((settings.clamav_host, settings.clamav_port), timeout=30) as s:
            s.sendall(b"zINSTREAM\0")
            for i in range(0, len(data), 1 << 16):
                chunk = data[i : i + (1 << 16)]
                s.sendall(struct.pack("!L", len(chunk)) + chunk)
            s.sendall(struct.pack("!L", 0))
            reply = s.recv(4096).decode(errors="replace")
    except OSError as exc:
        if settings.antivirus_required:
            raise FileRejected("Servicio antivirus no disponible") from exc
        return "NO_ESCANEADO"
    if "FOUND" in reply:
        raise FileRejected(f"Archivo infectado: {reply.split(':', 1)[-1].replace('FOUND', '').strip(chr(0) + ' ')}")
    return "LIMPIO"
