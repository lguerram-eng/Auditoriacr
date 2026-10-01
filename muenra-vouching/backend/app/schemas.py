"""Esquemas Pydantic de entrada/salida de la API."""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field, field_serializer


class ORM(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    @field_serializer("*", when_used="json", check_fields=False)
    def _dec(self, v):
        return float(v) if isinstance(v, Decimal) else v


class LoginIn(BaseModel):
    email: str
    password: str


class TokenOut(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int
    user: "UserOut"


class PasswordForgotIn(BaseModel):
    email: str


class PasswordResetIn(BaseModel):
    token: str
    new_password: str


class PasswordChangeIn(BaseModel):
    current_password: str
    new_password: str


class UserOut(ORM):
    id: int
    email: str
    full_name: str
    role: str
    is_active: bool
    must_change_password: bool = False
    created_at: datetime | None = None
    last_login_at: datetime | None = None


class UserCreate(BaseModel):
    email: str = Field(pattern=r"^[^@\s]+@[^@\s]+\.[^@\s]+$", max_length=255)
    full_name: str = Field(min_length=2, max_length=255)
    role: str
    password: str


class UserUpdate(BaseModel):
    full_name: str | None = None
    role: str | None = None
    is_active: bool | None = None
    password: str | None = None


class ProjectIn(BaseModel):
    code: str = Field(min_length=2, max_length=50, pattern=r"^[A-Za-z0-9_.-]+$")
    name: str = Field(min_length=2, max_length=255)
    client_name: str = Field(min_length=2, max_length=255)
    client_nit: str | None = None
    period: str | None = None
    description: str | None = None
    allow_ai_processing: bool = False
    retention_days: int = Field(3650, ge=30, le=36500)
    member_ids: list[int] = []


class ProjectUpdate(BaseModel):
    name: str | None = None
    client_name: str | None = None
    client_nit: str | None = None
    period: str | None = None
    description: str | None = None
    status: str | None = None
    allow_ai_processing: bool | None = None
    retention_days: int | None = Field(None, ge=30, le=36500)
    member_ids: list[int] | None = None


class ProjectOut(ORM):
    id: int
    code: str
    name: str
    client_name: str
    client_nit: str | None
    period: str | None
    description: str | None
    status: str
    allow_ai_processing: bool
    allow_training: bool
    retention_days: int
    created_at: datetime
    settings: dict


class SettingsIn(BaseModel):
    parameters: dict


class FieldReviewIn(BaseModel):
    action: str = Field(pattern="^(ACEPTAR|RECHAZAR|CORREGIR)$")
    corrected_value: str | None = None
    comment: str | None = None


class ResultReviewIn(BaseModel):
    decision: str = Field(pattern="^(APROBADO|RECHAZADO)$")
    status: str | None = None  # estado manual opcional
    comment: str | None = None


class CommentIn(BaseModel):
    comment: str = Field(min_length=1, max_length=4000)


class ManualLinkIn(BaseModel):
    document_id: int
    role: str = Field("PRINCIPAL", pattern="^(PRINCIPAL|COMPLEMENTARIO)$")
    comment: str | None = None


class ReferenceOut(ORM):
    id: int
    row_number: int
    sample_id: str
    doc_type: str | None
    doc_number: str | None
    doc_date: date | None
    third_party: str | None
    nit: str | None
    expected_value: Decimal | None
    currency: str
    contract: str | None
    purchase_order: str | None
    concept: str | None
    cost_center: str | None
    account: str | None
    expected_file: str | None


TokenOut.model_rebuild()
