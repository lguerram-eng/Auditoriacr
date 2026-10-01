"""Modelo de datos relacional de Muenra Vouching.

Convenciones:
- Toda la información de negocio pertenece a un proyecto (``project_id``) para
  garantizar la separación por cliente/encargo.
- Los valores monetarios usan ``Numeric(20, 2)``.
- Los valores extraídos originalmente nunca se sobrescriben: las correcciones
  humanas se guardan en columnas separadas (``corrected_value``) y en el
  historial de revisiones.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal

from sqlalchemy import (
    JSON,
    Boolean,
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    Numeric,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


Money = Numeric(20, 2)


class Role:
    ADMIN = "administrador"
    AUDITOR = "auditor"
    REVISOR = "revisor"
    CONSULTA = "consulta"
    ALL = (ADMIN, AUDITOR, REVISOR, CONSULTA)


class Status:
    """Estados de vouching."""

    COINCIDE = "COINCIDE"
    COINCIDE_TOLERANCIA = "COINCIDE CON TOLERANCIA"
    EXCEPCION = "EXCEPCIÓN"
    REVISION_MANUAL = "REVISIÓN MANUAL"
    SIN_SOPORTE = "SIN SOPORTE"
    NO_REFERENCIADO = "SOPORTE NO REFERENCIADO"
    POSIBLE_DUPLICADO = "POSIBLE DUPLICADO"
    ILEGIBLE = "DOCUMENTO ILEGIBLE"
    ERROR_LECTURA = "ERROR DE LECTURA"
    PENDIENTE = "PENDIENTE"
    ALL = (
        COINCIDE,
        COINCIDE_TOLERANCIA,
        EXCEPCION,
        REVISION_MANUAL,
        SIN_SOPORTE,
        NO_REFERENCIADO,
        POSIBLE_DUPLICADO,
        ILEGIBLE,
        ERROR_LECTURA,
        PENDIENTE,
    )


class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(String(255), unique=True, index=True)
    full_name: Mapped[str] = mapped_column(String(255))
    password_hash: Mapped[str] = mapped_column(String(255))
    role: Mapped[str] = mapped_column(String(30), default=Role.CONSULTA)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0)
    locked_until: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    must_change_password: Mapped[bool] = mapped_column(Boolean, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    last_login_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))


class PasswordResetToken(Base):
    __tablename__ = "password_reset_tokens"
    id: Mapped[int] = mapped_column(primary_key=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    token_hash: Mapped[str] = mapped_column(String(128), unique=True)
    expires_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    used_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))


class Project(Base):
    __tablename__ = "projects"
    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(50), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(255))
    client_name: Mapped[str] = mapped_column(String(255))
    client_nit: Mapped[str | None] = mapped_column(String(30))
    period: Mapped[str | None] = mapped_column(String(50))
    description: Mapped[str | None] = mapped_column(Text)
    status: Mapped[str] = mapped_column(String(30), default="ACTIVO")
    # Parámetros del motor (tolerancias, pesos, umbrales). Ver services/settings_defaults.py
    settings: Mapped[dict] = mapped_column(JSON, default=dict)
    allow_ai_processing: Mapped[bool] = mapped_column(Boolean, default=False)
    # Siempre False: los documentos nunca se usan para entrenamiento. Se guarda
    # explícitamente para que quede como evidencia en las exportaciones.
    allow_training: Mapped[bool] = mapped_column(Boolean, default=False)
    retention_days: Mapped[int] = mapped_column(Integer, default=3650)
    created_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    closed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))

    members: Mapped[list[ProjectMember]] = relationship(cascade="all, delete-orphan", back_populates="project")


class ProjectMember(Base):
    __tablename__ = "project_members"
    __table_args__ = (UniqueConstraint("project_id", "user_id"),)
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("users.id", ondelete="CASCADE"), index=True)
    project: Mapped[Project] = relationship(back_populates="members")
    user: Mapped[User] = relationship()


class ImportBatch(Base):
    __tablename__ = "import_batches"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    filename: Mapped[str] = mapped_column(String(255))
    sha256: Mapped[str] = mapped_column(String(64))
    status: Mapped[str] = mapped_column(String(30))  # VALIDA | CON_ADVERTENCIAS | RECHAZADA
    row_count: Mapped[int] = mapped_column(Integer, default=0)
    total_value: Mapped[Decimal] = mapped_column(Money, default=0)
    integrity_report: Mapped[dict] = mapped_column(JSON, default=dict)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class ReferenceItem(Base):
    """Partida de la población/muestra (hoja REFERENCIA_VOUCHING)."""

    __tablename__ = "reference_items"
    __table_args__ = (UniqueConstraint("project_id", "sample_id"),)
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    import_id: Mapped[int] = mapped_column(ForeignKey("import_batches.id", ondelete="CASCADE"), index=True)
    row_number: Mapped[int] = mapped_column(Integer)
    sample_id: Mapped[str] = mapped_column(String(100))
    doc_type: Mapped[str | None] = mapped_column(String(100))
    doc_number: Mapped[str | None] = mapped_column(String(100))
    doc_date: Mapped[date | None] = mapped_column(Date)
    third_party: Mapped[str | None] = mapped_column(String(255))
    nit: Mapped[str | None] = mapped_column(String(30))
    expected_value: Mapped[Decimal | None] = mapped_column(Money)
    currency: Mapped[str] = mapped_column(String(10), default="COP")
    contract: Mapped[str | None] = mapped_column(String(100))
    purchase_order: Mapped[str | None] = mapped_column(String(100))
    concept: Mapped[str | None] = mapped_column(Text)
    cost_center: Mapped[str | None] = mapped_column(String(100))
    account: Mapped[str | None] = mapped_column(String(50))
    expected_file: Mapped[str | None] = mapped_column(String(255))
    extra: Mapped[dict] = mapped_column(JSON, default=dict)


class EntityAlias(Base):
    __tablename__ = "entity_aliases"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    nit: Mapped[str | None] = mapped_column(String(30))
    canonical_name: Mapped[str] = mapped_column(String(255))
    alias: Mapped[str] = mapped_column(String(255))


class DocumentTypeDef(Base):
    __tablename__ = "document_type_defs"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    code: Mapped[str] = mapped_column(String(50))
    name: Mapped[str] = mapped_column(String(100))
    keywords: Mapped[list] = mapped_column(JSON, default=list)
    is_primary: Mapped[bool] = mapped_column(Boolean, default=True)


class ExtractionFieldDef(Base):
    __tablename__ = "extraction_field_defs"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    field: Mapped[str] = mapped_column(String(60))
    description: Mapped[str | None] = mapped_column(Text)
    required: Mapped[bool] = mapped_column(Boolean, default=False)
    data_type: Mapped[str] = mapped_column(String(20), default="TEXTO")
    doc_types: Mapped[list] = mapped_column(JSON, default=list)


class ExpectedResult(Base):
    """Hoja RESULTADO_ESPERADO: resultado esperado de control para validar el motor."""

    __tablename__ = "expected_results"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    sample_id: Mapped[str] = mapped_column(String(100))
    expected_status: Mapped[str] = mapped_column(String(40))
    expected_file: Mapped[str | None] = mapped_column(String(255))
    note: Mapped[str | None] = mapped_column(Text)


class Document(Base):
    __tablename__ = "documents"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    filename: Mapped[str] = mapped_column(String(255))
    stored_path: Mapped[str] = mapped_column(String(500))
    sha256: Mapped[str] = mapped_column(String(64), index=True)
    size_bytes: Mapped[int] = mapped_column(Integer)
    file_type: Mapped[str] = mapped_column(String(10))  # pdf, xml, png, jpg, tiff, docx
    mime: Mapped[str] = mapped_column(String(100))
    av_status: Mapped[str] = mapped_column(String(30), default="NO_ESCANEADO")
    # Estado técnico: PENDIENTE | PROCESANDO | PROCESADO | ILEGIBLE | ERROR
    processing_status: Mapped[str] = mapped_column(String(20), default="PENDIENTE", index=True)
    error: Mapped[str | None] = mapped_column(Text)
    doc_type: Mapped[str | None] = mapped_column(String(60))
    doc_type_confidence: Mapped[float | None] = mapped_column(Float)
    classification_reason: Mapped[str | None] = mapped_column(Text)
    page_count: Mapped[int] = mapped_column(Integer, default=0)
    extraction_method: Mapped[str | None] = mapped_column(String(30))
    ocr_confidence: Mapped[float | None] = mapped_column(Float)
    full_text: Mapped[str | None] = mapped_column(Text)
    tables: Mapped[list] = mapped_column(JSON, default=list)
    duplicate_of_id: Mapped[int | None] = mapped_column(ForeignKey("documents.id", ondelete="SET NULL"))
    group_key: Mapped[str | None] = mapped_column(String(200), index=True)  # XML + PDF de la misma factura
    vouching_status: Mapped[str] = mapped_column(String(40), default=Status.PENDIENTE)
    uploaded_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    uploaded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    processed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    processing_ms: Mapped[int | None] = mapped_column(Integer)

    pages: Mapped[list[DocumentPage]] = relationship(
        cascade="all, delete-orphan", order_by="DocumentPage.page_number", back_populates="document"
    )
    fields: Mapped[list[ExtractedField]] = relationship(cascade="all, delete-orphan", back_populates="document")


class DocumentPage(Base):
    __tablename__ = "document_pages"
    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    page_number: Mapped[int] = mapped_column(Integer)
    width: Mapped[float | None] = mapped_column(Float)
    height: Mapped[float | None] = mapped_column(Float)
    method: Mapped[str] = mapped_column(String(30))
    ocr_confidence: Mapped[float | None] = mapped_column(Float)
    text: Mapped[str | None] = mapped_column(Text)
    # Líneas con coordenadas normalizadas (0-1): [{"text", "bbox": [x0,y0,x1,y1], "conf"}]
    lines: Mapped[list] = mapped_column(JSON, default=list)
    document: Mapped[Document] = relationship(back_populates="pages")


class ExtractedField(Base):
    __tablename__ = "extracted_fields"
    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    field_name: Mapped[str] = mapped_column(String(60), index=True)
    value: Mapped[str | None] = mapped_column(Text)  # valor original extraído (inmutable)
    normalized_value: Mapped[str | None] = mapped_column(Text)
    page: Mapped[int | None] = mapped_column(Integer)
    bbox: Mapped[list | None] = mapped_column(JSON)
    evidence_text: Mapped[str | None] = mapped_column(Text)
    method: Mapped[str] = mapped_column(String(30))  # XML | PDF_TEXTO | OCR | DOCX | IA
    confidence: Mapped[float | None] = mapped_column(Float)
    extracted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    # Revisión humana
    review_status: Mapped[str] = mapped_column(String(20), default="PENDIENTE")  # ACEPTADO | RECHAZADO | CORREGIDO
    corrected_value: Mapped[str | None] = mapped_column(Text)
    reviewed_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    reviewed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    review_comment: Mapped[str | None] = mapped_column(Text)
    document: Mapped[Document] = relationship(back_populates="fields")

    @property
    def effective_value(self) -> str | None:
        if self.review_status == "RECHAZADO":
            return None
        if self.review_status == "CORREGIDO":
            return self.corrected_value
        return self.normalized_value if self.normalized_value is not None else self.value


class MatchRun(Base):
    __tablename__ = "match_runs"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    started_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    finished_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    parameters: Mapped[dict] = mapped_column(JSON, default=dict)
    reconciliation: Mapped[dict] = mapped_column(JSON, default=dict)
    engine_version: Mapped[str] = mapped_column(String(20))
    triggered_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"))


class MatchLink(Base):
    """Relación documento ↔ partida con la explicación por criterio."""

    __tablename__ = "match_links"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    run_id: Mapped[int] = mapped_column(ForeignKey("match_runs.id", ondelete="CASCADE"), index=True)
    reference_item_id: Mapped[int] = mapped_column(ForeignKey("reference_items.id", ondelete="CASCADE"), index=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    role: Mapped[str] = mapped_column(String(30))  # PRINCIPAL | COMPLEMENTARIO | REPRESENTACION_GRAFICA
    score: Mapped[float] = mapped_column(Float)
    criteria: Mapped[dict] = mapped_column(JSON, default=dict)
    allocated_value: Mapped[Decimal | None] = mapped_column(Money)
    is_manual: Mapped[bool] = mapped_column(Boolean, default=False)


class VouchingResult(Base):
    __tablename__ = "vouching_results"
    __table_args__ = (UniqueConstraint("reference_item_id"),)
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    run_id: Mapped[int | None] = mapped_column(ForeignKey("match_runs.id", ondelete="SET NULL"))
    reference_item_id: Mapped[int] = mapped_column(ForeignKey("reference_items.id", ondelete="CASCADE"))
    auto_status: Mapped[str] = mapped_column(String(40), default=Status.PENDIENTE)
    reasons: Mapped[list] = mapped_column(JSON, default=list)
    score: Mapped[float | None] = mapped_column(Float)
    # Resumen de la comparación (para tablas y exportación)
    extracted_number: Mapped[str | None] = mapped_column(String(100))
    extracted_date: Mapped[date | None] = mapped_column(Date)
    extracted_party: Mapped[str | None] = mapped_column(String(255))
    extracted_nit: Mapped[str | None] = mapped_column(String(30))
    extracted_value: Mapped[Decimal | None] = mapped_column(Money)
    abs_difference: Mapped[Decimal | None] = mapped_column(Money)
    pct_difference: Mapped[float | None] = mapped_column(Float)
    tolerance_applied: Mapped[float | None] = mapped_column(Float)
    name_similarity: Mapped[float | None] = mapped_column(Float)
    nit_match: Mapped[str | None] = mapped_column(String(30))
    dv_match: Mapped[str | None] = mapped_column(String(30))
    number_match: Mapped[str | None] = mapped_column(String(30))
    days_difference: Mapped[int | None] = mapped_column(Integer)
    ocr_confidence: Mapped[float | None] = mapped_column(Float)
    evidence_page: Mapped[int | None] = mapped_column(Integer)
    evidence_text: Mapped[str | None] = mapped_column(Text)
    explanation: Mapped[dict] = mapped_column(JSON, default=dict)
    # Revisión humana (se conserva entre re-ejecuciones del motor)
    manual_status: Mapped[str | None] = mapped_column(String(40))
    review_decision: Mapped[str | None] = mapped_column(String(20))  # APROBADO | RECHAZADO
    reviewer_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    reviewed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    review_comment: Mapped[str | None] = mapped_column(Text)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, onupdate=utcnow)

    reference_item: Mapped[ReferenceItem] = relationship()
    reviewer: Mapped[User | None] = relationship()

    @property
    def status(self) -> str:
        return self.manual_status or self.auto_status


class ReviewEvent(Base):
    """Historial de revisiones (inmutable, solo inserción)."""

    __tablename__ = "review_events"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    result_id: Mapped[int | None] = mapped_column(ForeignKey("vouching_results.id", ondelete="SET NULL"))
    field_id: Mapped[int | None] = mapped_column(ForeignKey("extracted_fields.id", ondelete="SET NULL"))
    document_id: Mapped[int | None] = mapped_column(ForeignKey("documents.id", ondelete="SET NULL"))
    sample_id: Mapped[str | None] = mapped_column(String(100))
    user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id"))
    user_email: Mapped[str | None] = mapped_column(String(255))
    action: Mapped[str] = mapped_column(String(40))
    field_name: Mapped[str | None] = mapped_column(String(60))
    old_value: Mapped[str | None] = mapped_column(Text)
    new_value: Mapped[str | None] = mapped_column(Text)
    comment: Mapped[str | None] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class AuditLog(Base):
    """Registro de accesos y cambios (solo inserción)."""

    __tablename__ = "audit_logs"
    id: Mapped[int] = mapped_column(primary_key=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, index=True)
    user_id: Mapped[int | None] = mapped_column(ForeignKey("users.id", ondelete="SET NULL"))
    user_email: Mapped[str | None] = mapped_column(String(255))
    project_id: Mapped[int | None] = mapped_column(ForeignKey("projects.id", ondelete="SET NULL"), index=True)
    action: Mapped[str] = mapped_column(String(60), index=True)
    entity: Mapped[str | None] = mapped_column(String(60))
    entity_id: Mapped[str | None] = mapped_column(String(60))
    detail: Mapped[dict] = mapped_column(JSON, default=dict)
    ip: Mapped[str | None] = mapped_column(String(64))
    success: Mapped[bool] = mapped_column(Boolean, default=True)


class Job(Base):
    """Cola de procesamiento persistente (respaldada en la base de datos)."""

    __tablename__ = "jobs"
    id: Mapped[int] = mapped_column(primary_key=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id", ondelete="CASCADE"), index=True)
    kind: Mapped[str] = mapped_column(String(30))  # PROCESAR_DOCUMENTO | EJECUTAR_VOUCHING
    payload: Mapped[dict] = mapped_column(JSON, default=dict)
    status: Mapped[str] = mapped_column(String(20), default="EN_COLA", index=True)
    attempts: Mapped[int] = mapped_column(Integer, default=0)
    error: Mapped[str | None] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    started_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    finished_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
