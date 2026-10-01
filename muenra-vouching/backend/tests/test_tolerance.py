"""Pruebas de la regla de coincidencia de valor (tolerancia 10 % y valor esperado cero)."""
from decimal import Decimal

import pytest

from app.services.normalization import compare_values
from app.services.settings_defaults import DEFAULT_PARAMETERS, merged_parameters, validate_parameters


def test_default_tolerance_is_ten_percent():
    assert DEFAULT_PARAMETERS["TOLERANCIA_VALOR"] == 0.10


def test_example_nine_percent_matches():
    c = compare_values(1_000_000, 1_090_000)
    assert c.result == "DENTRO_TOLERANCIA"
    assert c.pct_difference == pytest.approx(0.09)
    assert c.abs_difference == Decimal("90000.00")


def test_example_twelve_percent_is_exception():
    c = compare_values(1_000_000, 1_120_000)
    assert c.result == "FUERA_TOLERANCIA"
    assert c.pct_difference == pytest.approx(0.12)


def test_exact_boundary_and_below():
    assert compare_values(1_000_000, 1_100_000).result == "DENTRO_TOLERANCIA"  # exactamente 10 %
    assert compare_values(1_000_000, 900_000).result == "DENTRO_TOLERANCIA"
    assert compare_values(1_000_000, 1_100_001).result == "FUERA_TOLERANCIA"
    assert compare_values(1_000_000, 1_000_000).result == "EXACTO"


def test_configurable_tolerance():
    assert compare_values(1_000_000, 1_090_000, tolerance=0.05).result == "FUERA_TOLERANCIA"
    assert compare_values(1_000_000, 1_120_000, tolerance=0.15).result == "DENTRO_TOLERANCIA"


def test_negative_expected_uses_absolute_value():
    c = compare_values(-1_000_000, -1_050_000)
    assert c.result == "DENTRO_TOLERANCIA" and c.pct_difference == pytest.approx(0.05)


def test_zero_expected_never_divides():
    c = compare_values(0, 0)
    assert c.result == "EXACTO"
    c = compare_values(0, 50_000)
    assert c.result == "REVISION_MANUAL" and c.pct_difference is None
    c = compare_values(0, 50_000, zero_policy="EXACTA")
    assert c.result == "FUERA_TOLERANCIA"


def test_missing_value():
    assert compare_values(None, 10).result == "NO_DISPONIBLE"
    assert compare_values(10, "ilegible").result == "NO_DISPONIBLE"


def test_parameter_validation():
    assert validate_parameters(merged_parameters({"TOLERANCIA_VALOR": 0.2})) == []
    assert validate_parameters(merged_parameters({"TOLERANCIA_VALOR": 5})) != []
    p = merged_parameters({"PESOS": {"nit": 50, "desconocido": 3}})
    assert p["PESOS"]["nit"] == 50 and "desconocido" not in p["PESOS"]
