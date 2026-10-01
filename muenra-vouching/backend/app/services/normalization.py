"""Normalización de nombres, NIT, números documentales, valores y fechas.

Todas las funciones son puras (sin acceso a base de datos) para facilitar las
pruebas unitarias. El texto original nunca se modifica: estas funciones solo
producen una forma canónica *para comparar*.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from rapidfuzz import fuzz

# ---------------------------------------------------------------------------
# Nombres / razones sociales
# ---------------------------------------------------------------------------

# Sufijos jurídicos que se retiran SOLO para comparar (ya normalizados: sin tildes ni puntos).
LEGAL_SUFFIXES = {
    "SAS",
    "SA",
    "LTDA",
    "LIMITADA",
    "EU",
    "INC",
    "LLC",
    "COMPANY",
    "COMPANIA",
    "CIA",
    "BIC",
    "SCA",
    "SCS",
    "CORP",
    "CORPORATION",
    "LTD",
}
_CONNECTORS_AT_END = {"Y", "DE", "EN", "C"}


def strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))


def basic_normalize(text: str | None) -> str:
    """Mayúsculas, sin tildes, sin puntuación, espacios simples."""
    if not text:
        return ""
    t = strip_accents(str(text)).upper()
    t = t.replace("&", " Y ")
    t = t.replace(".", "")  # S.A.S. -> SAS
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)  # comas, guiones, comillas, paréntesis...
    t = re.sub(r"\s+", " ", t).strip()
    # Une letras sueltas consecutivas: "S A S" -> "SAS", "E U" -> "EU"
    tokens = t.split(" ") if t else []
    merged: list[str] = []
    buf = ""
    for tok in tokens:
        if len(tok) == 1 and tok.isalpha():
            buf += tok
            continue
        if buf:
            merged.append(buf)
            buf = ""
        merged.append(tok)
    if buf:
        merged.append(buf)
    return " ".join(merged)


def normalize_name(text: str | None) -> str:
    """Forma de comparación: normalización básica + retiro de sufijos jurídicos."""
    base = basic_normalize(text)
    if not base:
        return ""
    tokens = [t for t in base.split(" ") if t not in LEGAL_SUFFIXES]
    while tokens and tokens[-1] in _CONNECTORS_AT_END:
        tokens.pop()
    return " ".join(tokens) if tokens else base


def build_alias_index(aliases: list[tuple[str, str]]) -> dict[str, str]:
    """Recibe pares (alias, nombre_canonico) y devuelve índice normalizado alias -> canónico."""
    idx: dict[str, str] = {}
    for alias, canonical in aliases:
        canon = normalize_name(canonical)
        if not canon:
            continue
        idx[canon] = canon
        a = normalize_name(alias)
        if a:
            idx[a] = canon
    return idx


def canonical_name(text: str | None, alias_index: dict[str, str] | None = None) -> str:
    n = normalize_name(text)
    if alias_index and n in alias_index:
        return alias_index[n]
    return n


def name_similarity(a: str | None, b: str | None, alias_index: dict[str, str] | None = None) -> float | None:
    """Similitud 0-100 entre dos razones sociales después de normalizar y aplicar alias."""
    na, nb = canonical_name(a, alias_index), canonical_name(b, alias_index)
    if not na or not nb:
        return None
    if na == nb:
        return 100.0
    score = max(fuzz.ratio(na, nb), fuzz.token_sort_ratio(na, nb))
    # Un nombre contenido en el otro ("PROVEEDOR ANDINO" vs "PROVEEDOR ANDINO COLOMBIA")
    shorter = min(na, nb, key=len)
    if len(shorter.split()) >= 2:
        score = max(score, fuzz.token_set_ratio(na, nb) * 0.92)
    return round(float(score), 2)


# ---------------------------------------------------------------------------
# NIT
# ---------------------------------------------------------------------------

_DV_WEIGHTS = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71]


def compute_dv(base: str) -> int | None:
    """Dígito de verificación según el algoritmo de la DIAN (módulo 11)."""
    digits = re.sub(r"\D", "", base or "")
    if not digits or len(digits) > len(_DV_WEIGHTS):
        return None
    total = sum(int(d) * w for d, w in zip(reversed(digits), _DV_WEIGHTS))
    r = total % 11
    return r if r in (0, 1) else 11 - r


@dataclass(frozen=True)
class NitParts:
    original: str
    base: str
    dv: str | None

    @property
    def dv_valid(self) -> bool | None:
        if self.dv is None:
            return None
        return compute_dv(self.base) == int(self.dv)


def parse_nit(raw: str | int | None) -> NitParts | None:
    """Separa número base y dígito de verificación. Conserva el texto original."""
    if raw is None:
        return None
    original = str(raw).strip()
    if original.endswith(".0") and original[:-2].isdigit():  # celdas numéricas de Excel
        original = original[:-2]
    text = re.sub(r"(?i)\b(NIT|C\.?C\.?|NIT\.)\b[:.\s]*", "", original)
    text = re.sub(r"[.,\s]", "", text)
    if not text:
        return None
    m = re.fullmatch(r"(\d+)[-‐–](\d)", text)
    if m:
        return NitParts(original, m.group(1).lstrip("0") or "0", m.group(2))
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    # Sin guion: 10 dígitos de persona jurídica (8xx/9xx) cuyo último dígito valida como DV.
    if len(digits) == 10 and digits[0] in "89" and compute_dv(digits[:-1]) == int(digits[-1]):
        return NitParts(original, digits[:-1], digits[-1])
    return NitParts(original, digits.lstrip("0") or "0", None)


@dataclass(frozen=True)
class NitComparison:
    base_match: str  # COINCIDE | NO COINCIDE | NO DISPONIBLE
    dv_match: str  # COINCIDE | NO COINCIDE | NO DISPONIBLE
    expected: NitParts | None
    found: NitParts | None


def compare_nit(expected: str | None, found: str | None) -> NitComparison:
    e, f = parse_nit(expected), parse_nit(found)
    if not e or not f:
        return NitComparison("NO DISPONIBLE", "NO DISPONIBLE", e, f)
    base = "COINCIDE" if e.base == f.base else "NO COINCIDE"
    if e.dv is None or f.dv is None:
        dv = "NO DISPONIBLE"
    else:
        dv = "COINCIDE" if e.dv == f.dv else "NO COINCIDE"
    return NitComparison(base, dv, e, f)


# ---------------------------------------------------------------------------
# Números documentales
# ---------------------------------------------------------------------------


def normalize_doc_number(raw: str | None) -> str:
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return re.sub(r"[^A-Z0-9]", "", strip_accents(s).upper())


def _split_prefix(n: str) -> tuple[str, str]:
    m = re.fullmatch(r"([A-Z]*)(\d+)([A-Z0-9]*)", n)
    if not m:
        return n, ""
    return m.group(1), (m.group(2).lstrip("0") or "0") + m.group(3)


def doc_number_score(expected: str | None, found: str | None) -> tuple[float | None, str]:
    """Devuelve (puntaje 0-1, descripción)."""
    e, f = normalize_doc_number(expected), normalize_doc_number(found)
    if not e or not f:
        return None, "NO DISPONIBLE"
    if e == f:
        return 1.0, "COINCIDE"
    pe, ne = _split_prefix(e)
    pf, nf = _split_prefix(f)
    if ne and ne == nf and len(ne) >= 2 and (not pe or not pf or pe == pf):
        return 0.95, "COINCIDE (sin prefijo/ceros)"
    if len(min(e, f, key=len)) >= 4 and (e in f or f in e):
        return 0.7, "PARCIAL"
    return 0.0, "NO COINCIDE"


# ---------------------------------------------------------------------------
# Valores
# ---------------------------------------------------------------------------

_AMOUNT_RE = re.compile(r"\(?-?\s*(?:COP|USD|EUR|US\$|\$)?\s*\d[\d.,\s]*\d|\d")


def parse_amount(raw) -> Decimal | None:
    """Convierte textos como "$ 1.234.567,89", "1,234,567.89" o "(1.000)" a Decimal."""
    if raw is None or raw == "":
        return None
    if isinstance(raw, (int, float, Decimal)) and not isinstance(raw, bool):
        try:
            return Decimal(str(raw)).quantize(Decimal("0.01"))
        except InvalidOperation:
            return None
    s = str(raw).strip()
    negative = s.startswith("(") and s.endswith(")") or s.startswith("-") or s.endswith("-")
    s = re.sub(r"[^\d.,]", "", s)
    if not s or not re.search(r"\d", s):
        return None
    if "." in s and "," in s:
        dec = "." if s.rfind(".") > s.rfind(",") else ","
        thou = "," if dec == "." else "."
        s = s.replace(thou, "").replace(dec, ".")
    elif "," in s or "." in s:
        sep = "," if "," in s else "."
        parts = s.split(sep)
        if len(parts) > 2:
            s = s.replace(sep, "")
        elif len(parts[1]) == 3 and parts[0] != "0":
            s = s.replace(sep, "")  # separador de miles (1.090 -> 1090)
        else:
            s = s.replace(sep, ".")
    try:
        val = Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None
    return -val if negative else val


@dataclass(frozen=True)
class ValueComparison:
    expected: Decimal | None
    extracted: Decimal | None
    abs_difference: Decimal | None
    pct_difference: float | None  # fracción: 0.09 = 9 %
    tolerance: float
    result: str  # EXACTO | DENTRO_TOLERANCIA | FUERA_TOLERANCIA | REVISION_MANUAL | NO_DISPONIBLE
    detail: str


def compare_values(expected, extracted, tolerance: float = 0.10, zero_policy: str = "REVISION_MANUAL") -> ValueComparison:
    """Regla de coincidencia de valor.

    DIFERENCIA_PORCENTAJE = ABS(VALOR_EXTRAIDO - VALOR_ESPERADO) / ABS(VALOR_ESPERADO)
    Coincide si DIFERENCIA_PORCENTAJE <= tolerancia.

    Si el valor esperado es cero no se divide: se exige coincidencia exacta y,
    según ``zero_policy``, un valor distinto de cero se marca como
    REVISION_MANUAL (predeterminado) o FUERA_TOLERANCIA (política EXACTA).
    """
    e = parse_amount(expected)
    x = parse_amount(extracted)
    if e is None or x is None:
        return ValueComparison(e, x, None, None, tolerance, "NO_DISPONIBLE", "Valor no disponible para comparar")
    diff = abs(x - e)
    if e == 0:
        if x == 0:
            return ValueComparison(e, x, diff, 0.0, tolerance, "EXACTO", "Valor esperado cero y extraído cero")
        res = "REVISION_MANUAL" if zero_policy == "REVISION_MANUAL" else "FUERA_TOLERANCIA"
        return ValueComparison(
            e, x, diff, None, tolerance, res, "Valor esperado cero: no se calcula porcentaje; requiere coincidencia exacta"
        )
    pct = float(diff / abs(e))
    if diff == 0:
        res, detail = "EXACTO", "Valores idénticos"
    elif pct <= tolerance + 1e-12:
        res, detail = "DENTRO_TOLERANCIA", f"Diferencia {pct:.2%} dentro de la tolerancia {tolerance:.2%}"
    else:
        res, detail = "FUERA_TOLERANCIA", f"Diferencia {pct:.2%} supera la tolerancia {tolerance:.2%}"
    return ValueComparison(e, x, diff, round(pct, 6), tolerance, res, detail)


# ---------------------------------------------------------------------------
# Fechas
# ---------------------------------------------------------------------------

MONTHS = {
    "ENERO": 1, "ENE": 1, "JANUARY": 1, "JAN": 1,
    "FEBRERO": 2, "FEB": 2, "FEBRUARY": 2,
    "MARZO": 3, "MAR": 3, "MARCH": 3,
    "ABRIL": 4, "ABR": 4, "APRIL": 4, "APR": 4,
    "MAYO": 5, "MAY": 5,
    "JUNIO": 6, "JUN": 6, "JUNE": 6,
    "JULIO": 7, "JUL": 7, "JULY": 7,
    "AGOSTO": 8, "AGO": 8, "AUGUST": 8, "AUG": 8,
    "SEPTIEMBRE": 9, "SETIEMBRE": 9, "SEP": 9, "SEPT": 9, "SET": 9, "SEPTEMBER": 9,
    "OCTUBRE": 10, "OCT": 10, "OCTOBER": 10,
    "NOVIEMBRE": 11, "NOV": 11, "NOVEMBER": 11,
    "DICIEMBRE": 12, "DIC": 12, "DECEMBER": 12, "DEC": 12,
}

DATE_PATTERNS = [
    re.compile(r"\b(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\b"),  # ISO
    re.compile(r"\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b"),  # dd/mm/aaaa
    re.compile(r"\b(\d{1,2})\s*(?:DE\s+)?([A-Z]{3,10})\.?\s*(?:DE(?:L)?\s+|[-/ ])?\s*(\d{4})\b"),  # 15 de marzo de 2026
    re.compile(r"\b([A-Z]{3,10})\.?\s+(\d{1,2})\s*(?:DE|,)?\s*(\d{4})\b"),  # marzo 15 de 2026
]


def parse_date(raw) -> date | None:
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = strip_accents(str(raw)).upper().strip()
    for i, pat in enumerate(DATE_PATTERNS):
        m = pat.search(s)
        if not m:
            continue
        try:
            if i == 0:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            elif i == 1:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if mo > 12 and d <= 12:  # formato mm/dd/aaaa
                    d, mo = mo, d
            elif i == 2:
                mo = MONTHS.get(m.group(2))
                if not mo:
                    continue
                d, y = int(m.group(1)), int(m.group(3))
            else:
                mo = MONTHS.get(m.group(1))
                if not mo:
                    continue
                d, y = int(m.group(2)), int(m.group(3))
            if 1900 <= y <= 2100:
                return date(y, mo, d)
        except ValueError:
            continue
    return None


def normalize_text(text: str | None) -> str:
    return basic_normalize(text)


def normalize_ref(raw: str | None) -> str:
    """Contratos / órdenes de compra: solo alfanuméricos en mayúscula."""
    return normalize_doc_number(raw)
