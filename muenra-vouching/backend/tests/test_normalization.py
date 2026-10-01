"""Pruebas de normalización de nombres, NIT, valores, fechas y números documentales."""
from datetime import date
from decimal import Decimal

import pytest

from app.services.normalization import (
    build_alias_index,
    compare_nit,
    compute_dv,
    doc_number_score,
    name_similarity,
    normalize_name,
    parse_amount,
    parse_date,
    parse_nit,
)

VARIANTS = ["PROVEEDOR ANDINO SAS", "Proveedor Andino S.A.S.", "PROVEEDOR ANDINO S A S", "Proveedor Andino", "Proveedor Andino S.A.S"]


@pytest.mark.parametrize("name", VARIANTS)
def test_legal_suffix_variants_are_equivalent(name):
    assert normalize_name(name) == "PROVEEDOR ANDINO"
    assert name_similarity(name, "PROVEEDOR ANDINO SAS") == 100.0


@pytest.mark.parametrize("raw,expected", [
    ("Compañía Ñandú Ltda.", "NANDU"),
    ("Tecnología Cumbre S.A.", "TECNOLOGIA CUMBRE"),
    ("Papelería La Estrella E.U.", "PAPELERIA LA ESTRELLA"),
    ("Servicios   Logísticos, del-Caribe LTDA", "SERVICIOS LOGISTICOS DEL CARIBE"),
    ("Acme Inc.", "ACME"),
    ("Global Trade LLC", "GLOBAL TRADE"),
    ("Pérez & Cía", "PEREZ"),
])
def test_normalize_name(raw, expected):
    assert normalize_name(raw) == expected


def test_fuzzy_threshold_and_aliases():
    assert name_similarity("Consultores Boreal SAS", "Consultores Borreal S.A.S.") >= 85
    assert name_similarity("Consultores Boreal SAS", "Distribuidora Nevado SAS") < 60
    idx = build_alias_index([("NEVADO DISTRIBUCIONES", "Distribuidora Nevado SAS")])
    assert name_similarity("NEVADO DISTRIBUCIONES", "Distribuidora Nevado S.A.S.", idx) == 100.0
    assert name_similarity(None, "X") is None


def test_dv_algorithm():
    # Valores calculados con el algoritmo módulo 11 de la DIAN
    assert compute_dv("900123456") == 8
    assert compute_dv("800234567") == compute_dv("800.234.567")
    assert compute_dv("") is None


@pytest.mark.parametrize("raw,base,dv", [
    ("900.123.456-8", "900123456", "8"),
    ("900123456-8", "900123456", "8"),
    ("900 123 456 - 8", "900123456", "8"),
    ("9001234568", "900123456", "8"),  # 10 dígitos con DV válido sin guion
    ("NIT: 900.123.456-8", "900123456", "8"),
    ("1.020.304.050", "1020304050", None),  # cédula sin DV
    (900123456.0, "900123456", None),  # celda numérica de Excel
])
def test_parse_nit(raw, base, dv):
    p = parse_nit(raw)
    assert p.base == base and p.dv == dv
    assert p.original  # conserva el texto original


def test_compare_nit_reports_dv_separately():
    c = compare_nit("900.123.456-8", "900123456-3")
    assert c.base_match == "COINCIDE" and c.dv_match == "NO COINCIDE"
    c = compare_nit("900.123.456-8", "900123456")
    assert c.base_match == "COINCIDE" and c.dv_match == "NO DISPONIBLE"
    c = compare_nit("900.123.456-8", "800.234.567-1")
    assert c.base_match == "NO COINCIDE"
    assert c.expected.original == "900.123.456-8"


@pytest.mark.parametrize("raw,expected", [
    ("$ 1.234.567,89", Decimal("1234567.89")),
    ("1,234,567.89", Decimal("1234567.89")),
    ("1.090.000", Decimal("1090000")),
    ("1.090", Decimal("1090")),
    ("12.50", Decimal("12.50")),
    ("1,5", Decimal("1.5")),
    ("(1.000)", Decimal("-1000")),
    ("COP 750000", Decimal("750000")),
    (1000000, Decimal("1000000")),
    ("abc", None),
    ("", None),
])
def test_parse_amount(raw, expected):
    assert parse_amount(raw) == expected


@pytest.mark.parametrize("raw,expected", [
    ("02/03/2026", date(2026, 3, 2)),
    ("2026-03-02", date(2026, 3, 2)),
    ("15 de marzo de 2026", date(2026, 3, 15)),
    ("Bogotá, 5 de Septiembre del 2026", date(2026, 9, 5)),
    ("marzo 15 de 2026", date(2026, 3, 15)),
    ("15-mar-2026", date(2026, 3, 15)),
    ("31/02/2026", None),
    ("sin fecha", None),
])
def test_parse_date(raw, expected):
    assert parse_date(raw) == expected


def test_doc_number_score():
    assert doc_number_score("FV-1001", "FV1001")[0] == 1.0
    assert doc_number_score("FE-0001234", "FE1234")[0] == 0.95
    assert doc_number_score("1234", "FV-1234")[0] == 0.95
    assert doc_number_score("FV-1001", "FV-1002")[0] == 0.0
    assert doc_number_score(None, "X")[0] is None
