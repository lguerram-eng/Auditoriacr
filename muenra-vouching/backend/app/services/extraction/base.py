"""Estructuras comunes del servicio de extracción documental."""
from __future__ import annotations

from dataclasses import dataclass, field


class Method:
    XML = "XML"
    PDF_TEXT = "PDF_TEXTO"
    OCR = "OCR"
    DOCX = "DOCX"
    AI = "IA"
    RULES = "REGLAS"


@dataclass
class Line:
    """Línea de texto con su región normalizada en la página (0-1)."""

    text: str
    bbox: list[float] | None = None  # [x0, y0, x1, y1] relativos al ancho/alto de la página
    conf: float | None = None  # 0-100

    def to_dict(self) -> dict:
        return {"text": self.text, "bbox": self.bbox, "conf": self.conf}


@dataclass
class Page:
    number: int
    method: str
    lines: list[Line] = field(default_factory=list)
    width: float | None = None
    height: float | None = None
    ocr_confidence: float | None = None

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)


@dataclass
class FieldEvidence:
    """Un campo extraído con su evidencia (página, región, texto, método, confianza)."""

    name: str
    value: str | None
    normalized: str | None = None
    page: int | None = None
    bbox: list[float] | None = None
    evidence_text: str | None = None
    method: str = Method.RULES
    confidence: float | None = None


@dataclass
class ExtractionResult:
    file_type: str
    method: str
    pages: list[Page] = field(default_factory=list)
    fields: list[FieldEvidence] = field(default_factory=list)
    tables: list[dict] = field(default_factory=list)
    doc_type: str | None = None
    doc_type_confidence: float | None = None
    classification_reason: str | None = None
    xml_root: str | None = None
    warnings: list[str] = field(default_factory=list)

    @property
    def full_text(self) -> str:
        return "\n\n".join(p.text for p in self.pages)

    @property
    def ocr_confidence(self) -> float | None:
        vals = [p.ocr_confidence for p in self.pages if p.ocr_confidence is not None]
        return round(sum(vals) / len(vals), 2) if vals else None

    def get(self, name: str) -> FieldEvidence | None:
        for f in self.fields:
            if f.name == name:
                return f
        return None

    def set_field(self, fe: FieldEvidence, overwrite: bool = False) -> None:
        existing = self.get(fe.name)
        if existing is None:
            self.fields.append(fe)
        elif overwrite:
            self.fields.remove(existing)
            self.fields.append(fe)


class UnreadableDocument(Exception):
    """El documento no contiene texto legible (incluso tras OCR)."""
