"""Motor de vouching: relaciona documentos soporte con partidas y determina estados.

El motor es explicable: cada relación conserva el puntaje de cada criterio
(NIT, número, valor, fecha, nombre, contrato, OC, concepto, moneda, tipo y archivo),
el peso aplicado, los valores comparados y el motivo. No existe una única
puntuación de "caja negra".

Fases:
1. Construcción de vistas de documentos (con correcciones humanas aplicadas) y
   agrupación XML + representación gráfica PDF / duplicados.
2. Puntuación por pares partida × unidad documental (con bloqueo para eficiencia).
3. Asignación de soportes principales (1:1 voraz por puntaje), soportes compartidos
   (varias partidas con un documento, sin duplicar valor) y combinación de varios
   soportes cuya suma alcanza el valor contabilizado.
4. Soportes complementarios (contrato, OC, egreso, certificación...).
5. Determinación del estado y motivos por partida; estado de cada documento.
6. Conciliación de conteos y valores antes/después.
"""
from __future__ import annotations

import itertools
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from decimal import Decimal

from rapidfuzz import fuzz
from sqlalchemy.orm import Session

from .. import models
from ..models import Status
from .extraction.classifier import normalize_type, type_family
from .normalization import (
    basic_normalize,
    build_alias_index,
    compare_nit,
    compare_values,
    doc_number_score,
    name_similarity,
    normalize_doc_number,
    normalize_ref,
    parse_amount,
    parse_date,
    parse_nit,
)
from .normalization import format_cop as cop
from .settings_defaults import CRITERION_LABELS, PRIORITY_LABELS, merged_parameters

ENGINE_VERSION = "vouching-1.0"
IDENTITY = ("nit", "numero", "contrato", "orden_compra", "archivo")


# ---------------------------------------------------------------------------
# Vistas
# ---------------------------------------------------------------------------


@dataclass
class FieldRef:
    value: str | None
    field_id: int | None = None
    page: int | None = None
    evidence: str | None = None
    confidence: float | None = None
    method: str | None = None


@dataclass
class DocView:
    id: int
    filename: str
    sha256: str
    file_type: str
    status: str
    doc_type: str | None
    ocr_confidence: float | None
    text_norm: str
    fields: dict[str, FieldRef]
    nits: list[tuple[str, str, str | None, FieldRef]]  # (rol, base, dv, campo)
    names: list[tuple[str, FieldRef]]
    duplicate_of: int | None = None

    def f(self, name: str) -> str | None:
        fr = self.fields.get(name)
        return fr.value if fr else None

    @property
    def value(self) -> Decimal | None:
        return parse_amount(self.f("valor_total"))

    @property
    def number(self) -> str | None:
        return self.f("numero_documento")

    @property
    def date(self) -> date | None:
        return parse_date(self.f("fecha_emision"))

    @property
    def family(self) -> str | None:
        return type_family(self.doc_type)


@dataclass
class Unit:
    """Unidad documental: un documento o un grupo XML + PDF de la misma factura."""

    key: str
    docs: list[DocView]
    duplicate_flag: bool = False
    duplicate_reason: str | None = None

    @property
    def rep(self) -> DocView:
        return sorted(self.docs, key=lambda d: (d.file_type != "xml", d.id))[0]


def build_doc_view(doc: models.Document, client_nit: str | None) -> DocView:
    fields: dict[str, FieldRef] = {}
    for f in doc.fields:
        val = f.effective_value
        if val in (None, ""):
            continue
        if f.field_name in fields and f.method == "IA":
            continue
        fields[f.field_name] = FieldRef(val, f.id, f.page, f.evidence_text, f.confidence, f.method)
    nits = []
    client = parse_nit(client_nit) if client_nit else None
    roles = ["emisor", "receptor"] + sorted(k[:-4] for k in fields if k.endswith("_nit") and k[:-4] not in ("emisor", "receptor"))
    for role in roles:
        fr = fields.get(f"{role}_nit")
        if fr:
            parts = parse_nit(fr.value)
            dv = fields.get(f"{role}_dv")
            if parts and all(parts.base != n[1] for n in nits):
                nits.append((role, parts.base, (dv.value if dv else parts.dv), fr))
    names = [(role, fields[f"{role}_nombre"]) for role in ("emisor", "receptor") if f"{role}_nombre" in fields]
    return DocView(
        id=doc.id,
        filename=doc.filename,
        sha256=doc.sha256,
        file_type=doc.file_type,
        status=doc.processing_status,
        doc_type=doc.doc_type,
        ocr_confidence=doc.ocr_confidence,
        text_norm=normalize_doc_number(doc.full_text or "")[:200000],
        fields=fields,
        nits=[n for n in nits if not (client and n[1] == client.base)] or nits,
        names=names,
        duplicate_of=doc.duplicate_of_id,
    )


def group_units(docs: list[DocView]) -> list[Unit]:
    """Agrupa XML + PDF de la misma factura y marca posibles duplicados."""
    by_key: dict[str, list[DocView]] = defaultdict(list)
    singles: list[Unit] = []
    for d in docs:
        if d.status != "PROCESADO":
            singles.append(Unit(f"doc:{d.id}", [d]))
            continue
        cufe = d.f("cufe")
        num = normalize_doc_number(d.number)
        nit = d.nits[0][1] if d.nits else ""
        if cufe:
            key = f"cufe:{cufe.lower()}"
        elif num and nit and d.family in ("FACTURA", None):
            key = f"fac:{nit}:{num}"
        else:
            singles.append(Unit(f"doc:{d.id}", [d]))
            continue
        by_key[key].append(d)
    # Unifica grupos CUFE con grupos NIT+número de la misma factura (PDF sin CUFE legible)
    for key in [k for k in by_key if k.startswith("cufe:")]:
        rep = by_key[key][0]
        alt = f"fac:{rep.nits[0][1]}:{normalize_doc_number(rep.number)}" if rep.nits and rep.number else None
        if alt and alt in by_key:
            by_key[key].extend(by_key.pop(alt))
    units = singles
    for key, members in by_key.items():
        unit = Unit(key, members)
        types = [m.file_type for m in members]
        hashes = [m.sha256 for m in members]
        if len(set(hashes)) < len(hashes):
            unit.duplicate_flag = True
            unit.duplicate_reason = "Archivo idéntico cargado más de una vez (mismo SHA-256)"
        non_xml = [t for t in types if t != "xml"]
        if types.count("xml") > 1 or len(non_xml) > 1:
            unit.duplicate_flag = True
            unit.duplicate_reason = unit.duplicate_reason or "Más de un documento con el mismo número, emisor y/o CUFE"
        units.append(unit)
    for u in units:
        if any(d.duplicate_of for d in u.docs):
            u.duplicate_flag = True
            u.duplicate_reason = u.duplicate_reason or "Archivo idéntico cargado más de una vez (mismo SHA-256)"
    return units


# ---------------------------------------------------------------------------
# Puntuación por criterio
# ---------------------------------------------------------------------------


@dataclass
class Criterion:
    key: str
    score: float | None
    expected: str | None
    found: str | None
    detail: str
    weight: float = 0.0
    evidence: FieldRef | None = None

    def as_dict(self) -> dict:
        return {
            "criterio": CRITERION_LABELS.get(self.key, self.key),
            "clave": self.key,
            "prioridad": PRIORITY_LABELS.get(self.key),
            "puntaje": None if self.score is None else round(self.score, 4),
            "peso": self.weight,
            "evaluado": self.score is not None,
            "esperado": self.expected,
            "extraido": self.found,
            "detalle": self.detail,
            "pagina": self.evidence.page if self.evidence else None,
            "evidencia": self.evidence.evidence if self.evidence else None,
            "campo_id": self.evidence.field_id if self.evidence else None,
        }


@dataclass
class PairScore:
    ref_id: int
    unit: Unit
    score: float
    criteria: dict[str, Criterion]
    nit_contradiction: bool = False
    strong: bool = False

    def explanation(self) -> dict:
        return {
            "puntaje_total": round(self.score, 4),
            "contradiccion_nit": self.nit_contradiction,
            "identificador_fuerte": self.strong,
            "criterios": [c.as_dict() for c in self.criteria.values()],
        }


def _fmt(v) -> str | None:
    if v is None:
        return None
    if isinstance(v, Decimal):
        return cop(v)
    return str(v)


def value_score(cmp, tol: float) -> float | None:
    if cmp.result == "NO_DISPONIBLE":
        return None
    if cmp.result == "EXACTO":
        return 1.0
    if cmp.result == "DENTRO_TOLERANCIA":
        return 1.0 - 0.3 * ((cmp.pct_difference or 0) / tol if tol else 0)
    if cmp.result == "REVISION_MANUAL":
        return 0.3
    return max(0.0, 0.4 * (1 - min(cmp.pct_difference or 1.0, 1.0)))


class Scorer:
    def __init__(self, params: dict, alias_index: dict[str, str], client_nit: str | None, client_name: str | None):
        self.p = params
        self.w = params["PESOS"]
        self.alias = alias_index
        self.client = parse_nit(client_nit) if client_nit else None
        self.client_name = client_name

    def _ref_mentioned(self, needle: str | None, d: DocView) -> bool:
        n = normalize_ref(needle)
        return len(n) >= 4 and n in d.text_norm

    def score(self, ref: models.ReferenceItem, unit: Unit) -> PairScore:
        d = unit.rep
        p, tol = self.p, float(self.p["TOLERANCIA_VALOR"])
        crit: dict[str, Criterion] = {}
        contradiction = False

        # NIT (prioridad muy alta)
        exp = parse_nit(ref.nit)
        if exp and d.nits:
            match = next((n for n in d.nits if n[1] == exp.base), None)
            if match:
                c = compare_nit(ref.nit, f"{match[1]}-{match[2]}" if match[2] else match[1])
                det = f"NIT base coincide ({match[0]}); DV: {c.dv_match}"
                crit["nit"] = Criterion("nit", 1.0, ref.nit, match[3].value, det, evidence=match[3])
            else:
                contradiction = True
                found = ", ".join(n[3].value for n in d.nits)
                crit["nit"] = Criterion("nit", 0.0, ref.nit, found, "NIT del documento contradice el NIT esperado", evidence=d.nits[0][3])
        else:
            crit["nit"] = Criterion("nit", None, ref.nit, None, "NIT no disponible en partida o documento")

        # Número documental
        sc, det = doc_number_score(ref.doc_number, d.number)
        if (sc is None or sc < 0.7) and self._ref_mentioned(ref.doc_number, d):
            sc, det = 0.6, "Número esperado mencionado en el texto del documento"
        crit["numero"] = Criterion("numero", sc, ref.doc_number, d.number, det, evidence=d.fields.get("numero_documento"))

        # Valor
        cmpv = compare_values(ref.expected_value, d.value, tol, p["POLITICA_VALOR_CERO"])
        crit["valor"] = Criterion("valor", value_score(cmpv, tol), _fmt(cmpv.expected), _fmt(cmpv.extracted), cmpv.detail, evidence=d.fields.get("valor_total"))

        # Fecha
        dd = d.date
        if ref.doc_date and dd:
            days = abs((dd - ref.doc_date).days)
            tdays = int(p["TOLERANCIA_DIAS"])
            s = 1.0 - (0.3 * days / tdays if tdays else 0) if days <= tdays else max(0.0, 0.6 * (1 - (days - tdays) / 90))
            crit["fecha"] = Criterion("fecha", s, ref.doc_date.isoformat(), dd.isoformat(), f"Diferencia de {days} días (tolerancia {tdays})", evidence=d.fields.get("fecha_emision"))
        else:
            crit["fecha"] = Criterion("fecha", None, ref.doc_date.isoformat() if ref.doc_date else None, dd.isoformat() if dd else None, "Fecha no disponible")

        # Nombre normalizado
        best, best_fr = None, None
        for _role, fr in d.names:
            if self.client_name and (name_similarity(fr.value, self.client_name, self.alias) or 0) >= 95 and len(d.names) > 1:
                continue
            s = name_similarity(ref.third_party, fr.value, self.alias)
            if s is not None and (best is None or s > best):
                best, best_fr = s, fr
        if best is not None:
            thr = float(p["UMBRAL_NOMBRE"])
            crit["nombre"] = Criterion("nombre", best / 100, ref.third_party, best_fr.value, f"Similitud {best:.1f} % (umbral {thr:.0f} %)", evidence=best_fr)
        else:
            crit["nombre"] = Criterion("nombre", None, ref.third_party, None, "Nombre no disponible")

        # Contrato y orden de compra
        for key, ref_val, fname in (("contrato", ref.contract, "numero_contrato"), ("orden_compra", ref.purchase_order, "orden_compra")):
            dv_ = d.f(fname)
            if ref_val and dv_:
                ok = normalize_ref(ref_val) == normalize_ref(dv_)
                crit[key] = Criterion(key, 1.0 if ok else 0.0, ref_val, dv_, "Coincide" if ok else "No coincide", evidence=d.fields.get(fname))
            elif ref_val and self._ref_mentioned(ref_val, d):
                crit[key] = Criterion(key, 0.9, ref_val, ref_val, "Referencia encontrada en el texto del documento")
            else:
                crit[key] = Criterion(key, None, ref_val, dv_, "No disponible")

        # Archivo referenciado explícitamente en la partida
        if ref.expected_file:
            wanted = [w.strip().lower() for w in ref.expected_file.replace(",", ";").split(";") if w.strip()]
            names = {x.filename.lower() for x in unit.docs} | {x.filename.lower().rsplit(".", 1)[0] for x in unit.docs}
            ok = any(w in names or w.rsplit(".", 1)[0] in names for w in wanted)
            crit["archivo"] = Criterion("archivo", 1.0 if ok else 0.0, ref.expected_file, d.filename, "Archivo indicado en la partida" if ok else "Archivo distinto al indicado")
        else:
            crit["archivo"] = Criterion("archivo", None, None, d.filename, "La partida no indica archivo")

        # Tipo documental
        rt = normalize_type(ref.doc_type)
        if rt and d.doc_type:
            if rt == d.doc_type:
                s = 1.0
            elif type_family(rt) == d.family or (rt == "FACTURA" and d.family == "FACTURA"):
                s = 0.8
            else:
                s = 0.2
            crit["tipo"] = Criterion("tipo", s, ref.doc_type, d.doc_type, "Tipo documental " + ("compatible" if s >= 0.8 else "diferente"))
        else:
            crit["tipo"] = Criterion("tipo", None, ref.doc_type, d.doc_type, "No disponible")

        # Concepto (complementario)
        dc = d.f("concepto") or d.f("objeto_contractual")
        if ref.concept and dc:
            s = fuzz.token_set_ratio(basic_normalize(ref.concept), basic_normalize(dc)) / 100
            crit["concepto"] = Criterion("concepto", s, ref.concept, dc, f"Similitud {s:.0%}", evidence=d.fields.get("concepto"))
        else:
            crit["concepto"] = Criterion("concepto", None, ref.concept, dc, "No disponible")

        # Moneda
        dm = d.f("moneda")
        if ref.currency and dm:
            ok = ref.currency.upper() == dm.upper()
            crit["moneda"] = Criterion("moneda", 1.0 if ok else 0.0, ref.currency, dm, "Coincide" if ok else "Moneda diferente", evidence=d.fields.get("moneda"))
        else:
            crit["moneda"] = Criterion("moneda", None, ref.currency, dm, "No disponible")

        # Ponderación
        num = den = 0.0
        for k, c in crit.items():
            c.weight = float(self.w.get(k, 0))
            if c.score is not None and c.weight > 0:
                num += c.weight * c.score
                den += c.weight
        total = num / den if den else 0.0
        strong = any((crit[k].score or 0) >= 0.9 for k in IDENTITY) or (
            (crit["valor"].score or 0) >= 0.7 and (crit["nombre"].score or 0) * 100 >= float(p["UMBRAL_NOMBRE"])
        )
        if not strong:
            total *= 0.5
        if den < 40:
            total *= den / 40
        if contradiction and (crit["numero"].score or 0) < 0.95 and (crit["archivo"].score or 0) < 1:
            total = min(total, float(p["UMBRAL_RELACION"]) - 0.01)
        return PairScore(ref.id, unit, round(total, 4), crit, contradiction, strong)


def _blocked(ref: models.ReferenceItem, unit: Unit, ref_nit: str | None) -> bool:
    """Descarta pares sin ninguna señal en común (eficiencia)."""
    d = unit.rep
    if ref.expected_file:
        return False
    if ref_nit and any(n[1] == ref_nit for n in d.nits):
        return False
    rn = normalize_doc_number(ref.doc_number)
    if rn and (rn == normalize_doc_number(d.number) or (len(rn) >= 4 and rn in d.text_norm)):
        return False
    for val in (ref.contract, ref.purchase_order):
        if val and len(normalize_ref(val)) >= 4 and normalize_ref(val) in d.text_norm:
            return False
    e, x = parse_amount(ref.expected_value), d.value
    if e is not None and x is not None and (e == x or (e != 0 and abs(x - e) / abs(e) <= Decimal("0.5"))):
        return False
    if ref.third_party and d.names:
        return False
    return True


# ---------------------------------------------------------------------------
# Ejecución del motor
# ---------------------------------------------------------------------------


@dataclass
class RowOutcome:
    ref: models.ReferenceItem
    primaries: list[tuple[PairScore, Decimal | None]] = field(default_factory=list)  # (par, valor asignado)
    complements: list[PairScore] = field(default_factory=list)
    shared_with: list[int] = field(default_factory=list)
    status: str = Status.PENDIENTE
    reasons: list[str] = field(default_factory=list)
    value_cmp: object | None = None
    best: PairScore | None = None
    special_doc: DocView | None = None


def run_matching(db: Session, project_id: int, user: models.User | None = None) -> models.MatchRun:
    project = db.get(models.Project, project_id)
    params = merged_parameters(project.settings)
    tol = float(params["TOLERANCIA_VALOR"])
    thr = float(params["UMBRAL_RELACION"])

    refs = db.query(models.ReferenceItem).filter_by(project_id=project_id).order_by(models.ReferenceItem.row_number).all()
    docs = db.query(models.Document).filter_by(project_id=project_id).order_by(models.Document.id).all()
    aliases = db.query(models.EntityAlias).filter_by(project_id=project_id).all()
    alias_index = build_alias_index([(a.alias, a.canonical_name) for a in aliases])

    run = models.MatchRun(project_id=project_id, parameters=params, engine_version=ENGINE_VERSION, triggered_by=user.id if user else None)
    db.add(run)
    db.flush()

    views = [build_doc_view(d, project.client_nit) for d in docs]
    units = group_units(views)
    unit_by_doc = {d.id: u for u in units for d in u.docs}
    scorer = Scorer(params, alias_index, project.client_nit, project.client_name)

    # Partidas duplicadas dentro de la población
    dup_keys: dict[tuple, list[int]] = defaultdict(list)
    for r in refs:
        pn = parse_nit(r.nit)
        k = (pn.base if pn else "", normalize_doc_number(r.doc_number), str(parse_amount(r.expected_value)))
        if k[0] or k[1]:
            dup_keys[k].append(r.id)
    dup_refs = {rid for ids in dup_keys.values() if len(ids) > 1 for rid in ids}

    # Puntuación de pares
    usable = [u for u in units if u.rep.status == "PROCESADO"]
    pairs: dict[int, list[PairScore]] = defaultdict(list)
    for r in refs:
        pn = parse_nit(r.nit)
        for u in usable:
            if _blocked(r, u, pn.base if pn else None):
                continue
            ps = scorer.score(r, u)
            if ps.score > 0:
                pairs[r.id].append(ps)
        pairs[r.id].sort(key=lambda x: x.score, reverse=True)

    outcomes = {r.id: RowOutcome(r) for r in refs}
    ref_by_id = {r.id: r for r in refs}
    unit_alloc: dict[str, Decimal] = defaultdict(Decimal)
    unit_rows: dict[str, list[int]] = defaultdict(list)

    def value_bearing(ps: PairScore) -> bool:
        return ps.criteria["valor"].score is not None or ps.criteria["valor"].expected is None

    # Relaciones manuales (forzadas por el revisor)
    manual = db.query(models.MatchLink).filter_by(project_id=project_id, is_manual=True).all()
    manual_keep = []
    for ml in manual:
        u = unit_by_doc.get(ml.document_id)
        r = ref_by_id.get(ml.reference_item_id)
        if not u or not r:
            continue
        ps = scorer.score(r, u)
        ps.criteria["manual"] = Criterion("manual", 1.0, None, None, "Relación confirmada manualmente por el revisor")
        manual_keep.append((r.id, ml.document_id, ml.role))
        if ml.role == "COMPLEMENTARIO":
            outcomes[r.id].complements.append(ps)
        else:
            outcomes[r.id].primaries.append((ps, None))
            unit_rows[u.key].append(r.id)

    # Fase 1: asignación 1:1 voraz por puntaje
    all_pairs = sorted((ps for lst in pairs.values() for ps in lst), key=lambda x: x.score, reverse=True)
    for ps in all_pairs:
        o = outcomes[ps.ref_id]
        if o.primaries or ps.score < thr or not value_bearing(ps) or unit_rows[ps.unit.key]:
            continue
        o.primaries.append((ps, None))
        unit_rows[ps.unit.key].append(ps.ref_id)

    # Fase 2: documento compartido por varias partidas (sin duplicar valor)
    for r in refs:
        o = outcomes[r.id]
        if o.primaries:
            continue
        for ps in pairs[r.id]:
            if ps.score < thr or not unit_rows[ps.unit.key] or ps.nit_contradiction:
                continue
            ident = any((ps.criteria[k].score or 0) >= 0.9 for k in ("nit", "numero", "contrato", "orden_compra"))
            doc_val = ps.unit.rep.value
            if not ident or doc_val is None:
                continue
            group_expected = sum((parse_amount(ref_by_id[x].expected_value) or Decimal(0)) for x in unit_rows[ps.unit.key])
            mine = parse_amount(r.expected_value) or Decimal(0)
            if group_expected + mine <= doc_val * Decimal(str(1 + tol)):
                o.primaries.append((ps, None))
                unit_rows[ps.unit.key].append(r.id)
                break

    # Fase 3: varios soportes cuya suma alcanza el valor contabilizado
    max_k = int(params["MAX_SOPORTES_SUMA"])
    for r in refs:
        o = outcomes[r.id]
        expected = parse_amount(r.expected_value)
        if not expected or expected <= 0:
            continue
        current = [ps for ps, _ in o.primaries]
        if current:
            cmp_cur = compare_values(expected, current[0].unit.rep.value, tol)
            if cmp_cur.result != "FUERA_TOLERANCIA" or (current[0].unit.rep.value or 0) >= expected or len(unit_rows[current[0].unit.key]) > 1:
                continue
        cands = [
            ps for ps in pairs[r.id]
            if not unit_rows[ps.unit.key] and not ps.nit_contradiction and ps.unit.rep.value and 0 < ps.unit.rep.value < expected
            and ((ps.criteria["nit"].score or 0) >= 1 or (ps.criteria["nombre"].score or 0) * 100 >= float(params["UMBRAL_NOMBRE"])
                 or (ps.criteria["contrato"].score or 0) >= 0.9 or (ps.criteria["orden_compra"].score or 0) >= 0.9)
        ][:12]
        best_combo, best_diff = None, None
        base_val = sum((ps.unit.rep.value or 0) for ps in current)
        for k in range(1 if current else 2, max_k - len(current) + 1):
            for combo in itertools.combinations(cands, k):
                total = base_val + sum(ps.unit.rep.value for ps in combo)
                c = compare_values(expected, total, tol)
                if c.result in ("EXACTO", "DENTRO_TOLERANCIA") and (best_diff is None or c.abs_difference < best_diff):
                    best_combo, best_diff = combo, c.abs_difference
            if best_combo is not None and best_diff == 0:
                break
        if best_combo:
            for ps in best_combo:
                o.primaries.append((ps, None))
                unit_rows[ps.unit.key].append(r.id)

    # Asignación de valores (nunca se asigna más que el valor del documento)
    for r in refs:
        o = outcomes[r.id]
        remaining = parse_amount(r.expected_value)
        for i, (ps, _) in enumerate(o.primaries):
            doc_val = ps.unit.rep.value
            if doc_val is None:
                continue
            available = doc_val - unit_alloc[ps.unit.key]
            if len(o.primaries) == 1 and len(unit_rows[ps.unit.key]) == 1:
                alloc = doc_val
            else:
                alloc = max(Decimal(0), min(available, remaining if remaining is not None else available))
            unit_alloc[ps.unit.key] += alloc
            if remaining is not None:
                remaining -= alloc
            o.primaries[i] = (ps, alloc)

    # Fase 4: soportes complementarios
    primary_units = {r_id: {ps.unit.key for ps, _ in o.primaries} for r_id, o in outcomes.items()}
    for r in refs:
        o = outcomes[r.id]
        fams = {ps.unit.rep.family for ps, _ in o.primaries}
        for ps in pairs[r.id]:
            if ps.unit.key in primary_units[r.id] or ps.nit_contradiction:
                continue
            c = ps.criteria
            references_row = any((c[k].score or 0) >= 0.9 for k in ("contrato", "orden_compra", "archivo")) or (c["numero"].score or 0) >= 0.6
            related_type = (c["nit"].score or 0) >= 1 and ps.unit.rep.family not in fams and ps.unit.rep.doc_type in (
                "CONTRATO", "OTROSI", "ORDEN_COMPRA", "CERTIFICACION", "COMPROBANTE_EGRESO", "SOPORTE_PAGO", "RECIBO_CAJA")
            if not (references_row or related_type) or ps.score < thr * 0.8:
                continue
            if unit_rows[ps.unit.key] and not references_row:
                continue  # es soporte principal de otra partida y no referencia a esta
                continue
            if ps.unit.rep.family in fams and unit_rows[ps.unit.key]:
                continue  # otra factura ya usada por otra partida no es complemento
            if any(c.unit.key == ps.unit.key for c in o.complements):
                continue
            o.complements.append(ps)

    # Documentos ilegibles o con error referenciados por nombre de archivo
    special_by_file = {}
    for v in views:
        if v.status in ("ILEGIBLE", "ERROR"):
            special_by_file[v.filename.lower()] = v
            special_by_file[v.filename.lower().rsplit(".", 1)[0]] = v

    # Fase 5: estados
    for r in refs:
        o = outcomes[r.id]
        o.best = pairs[r.id][0] if pairs[r.id] else None
        if r.expected_file:
            for w in r.expected_file.replace(",", ";").split(";"):
                sv = special_by_file.get(w.strip().lower()) or special_by_file.get(w.strip().lower().rsplit(".", 1)[0])
                if sv:
                    o.special_doc = sv
        _decide(o, params, unit_rows, ref_by_id, r.id in dup_refs)

    # Persistencia
    db.query(models.MatchLink).filter_by(project_id=project_id, is_manual=False).delete()
    keep_manual = {(a, b) for a, b, _ in manual_keep}
    linked_docs: dict[int, str] = {}
    for r in refs:
        o = outcomes[r.id]
        for ps, alloc in o.primaries:
            for d in ps.unit.docs:
                role = "PRINCIPAL" if d.id == ps.unit.rep.id else "REPRESENTACION_GRAFICA"
                if (r.id, d.id) not in keep_manual:
                    db.add(models.MatchLink(
                        project_id=project_id, run_id=run.id, reference_item_id=r.id, document_id=d.id, role=role,
                        score=ps.score, criteria=ps.explanation(), allocated_value=alloc if role == "PRINCIPAL" else None,
                    ))
                linked_docs.setdefault(d.id, o.status)
        for ps in o.complements:
            for d in ps.unit.docs:
                if (r.id, d.id) not in keep_manual:
                    db.add(models.MatchLink(
                        project_id=project_id, run_id=run.id, reference_item_id=r.id, document_id=d.id, role="COMPLEMENTARIO",
                        score=ps.score, criteria=ps.explanation(), allocated_value=None,
                    ))
                linked_docs.setdefault(d.id, o.status)
        if o.special_doc:
            linked_docs.setdefault(o.special_doc.id, o.status)
        _save_result(db, project_id, run.id, o, tol)

    dup_docs = {d.id for u in units if u.duplicate_flag for d in u.docs}
    for doc in docs:
        if doc.processing_status == "ILEGIBLE":
            doc.vouching_status = Status.ILEGIBLE
        elif doc.processing_status == "ERROR":
            doc.vouching_status = Status.ERROR_LECTURA
        elif doc.processing_status != "PROCESADO":
            doc.vouching_status = Status.PENDIENTE
        elif doc.id in dup_docs:
            doc.vouching_status = Status.POSIBLE_DUPLICADO
        elif doc.id in linked_docs:
            doc.vouching_status = linked_docs[doc.id]
        else:
            doc.vouching_status = Status.NO_REFERENCIADO
        u = unit_by_doc.get(doc.id)
        doc.group_key = u.key if u and len(u.docs) > 1 else None

    db.flush()
    run.reconciliation = reconcile(db, project_id, unit_alloc)
    run.finished_at = datetime.now(timezone.utc)
    db.commit()
    return run


def _decide(o: RowOutcome, params: dict, unit_rows, ref_by_id, ref_duplicated: bool) -> None:
    r = o.ref
    tol = float(params["TOLERANCIA_VALOR"])
    reasons: list[str] = []
    if ref_duplicated:
        reasons.append("Partida duplicada en la población (mismo NIT, número y valor)")
    if not o.primaries:
        if o.special_doc:
            o.status = Status.ILEGIBLE if o.special_doc.status == "ILEGIBLE" else Status.ERROR_LECTURA
            reasons.append(f"El soporte indicado ({o.special_doc.filename}) no pudo leerse")
        else:
            o.status = Status.POSIBLE_DUPLICADO if ref_duplicated else Status.SIN_SOPORTE
            if o.best and o.best.nit_contradiction and (o.best.criteria["numero"].score or 0) >= 0.7:
                reasons.append(f"Candidato {o.best.unit.rep.filename} descartado: el NIT contradice el NIT esperado")
            elif o.best:
                reasons.append(f"Mejor candidato {o.best.unit.rep.filename} con puntaje {o.best.score:.2f} inferior al umbral {params['UMBRAL_RELACION']}")
            else:
                reasons.append("No se encontró ningún documento relacionado")
            if o.complements:
                reasons.append("Solo existen soportes complementarios (sin documento que soporte el valor)")
        o.reasons = reasons
        return

    main = o.primaries[0][0]
    shared_rows = unit_rows[main.unit.key]
    expected = parse_amount(r.expected_value)
    if len(o.primaries) > 1:
        total = sum((ps.unit.rep.value or Decimal(0)) for ps, _ in o.primaries)
        cmp = compare_values(expected, total, tol, params["POLITICA_VALOR_CERO"])
        reasons.append(f"Valor soportado con {len(o.primaries)} documentos cuya suma es {cop(total)}")
    elif len(shared_rows) > 1:
        group_expected = sum((parse_amount(ref_by_id[x].expected_value) or Decimal(0)) for x in shared_rows)
        cmp = compare_values(group_expected, main.unit.rep.value, tol, params["POLITICA_VALOR_CERO"])
        others = [ref_by_id[x].sample_id for x in shared_rows if x != r.id]
        o.shared_with = [x for x in shared_rows if x != r.id]
        reasons.append(f"Documento compartido con las partidas {', '.join(others)}; se compara la suma de las partidas ({cop(group_expected)}) con el documento")
    else:
        cmp = compare_values(expected, main.unit.rep.value, tol, params["POLITICA_VALOR_CERO"])
    o.value_cmp = cmp

    flags: set[str] = set()
    if any(ps.unit.duplicate_flag for ps, _ in o.primaries):
        flags.add(Status.POSIBLE_DUPLICADO)
        reasons.append(next(ps.unit.duplicate_reason for ps, _ in o.primaries if ps.unit.duplicate_flag))
    if ref_duplicated:
        flags.add(Status.POSIBLE_DUPLICADO)
    for ps, _ in o.primaries:
        c = ps.criteria
        if ps.nit_contradiction:
            flags.add(Status.EXCEPCION)
            reasons.append(f"NIT no coincide: esperado {c['nit'].expected}, encontrado {c['nit'].found}")
        if c["nit"].score == 1.0 and "DV: NO COINCIDE" in c["nit"].detail:
            flags.add(Status.REVISION_MANUAL)
            reasons.append("El NIT base coincide pero el dígito de verificación es diferente")
        if c["moneda"].score == 0:
            flags.add(Status.EXCEPCION)
            reasons.append(f"Moneda diferente: esperada {c['moneda'].expected}, documento {c['moneda'].found}")
        if c["fecha"].score is not None and params.get("EXIGIR_FECHA_EN_TOLERANCIA") and c["fecha"].score < 0.7:
            flags.add(Status.EXCEPCION)
            reasons.append(f"Fecha fuera de tolerancia: {c['fecha'].detail}")
        if c["numero"].score is not None and c["numero"].score < 0.6 and len(o.primaries) == 1 and c["numero"].expected:
            flags.add(Status.REVISION_MANUAL)
            reasons.append(f"Número documental no coincide (esperado {c['numero'].expected}, extraído {c['numero'].found})")
        if c["nit"].score is None and (c["nombre"].score is None or c["nombre"].score * 100 < float(params["UMBRAL_NOMBRE"])):
            flags.add(Status.REVISION_MANUAL)
            reasons.append("Tercero no confirmado: sin NIT y nombre por debajo del umbral de similitud")
        vf = c["valor"].evidence
        if vf and vf.confidence is not None and vf.confidence < float(params["UMBRAL_CONFIANZA_OCR"]):
            flags.add(Status.REVISION_MANUAL)
            reasons.append(f"Confianza de lectura del valor baja ({vf.confidence:.0f})")
        if vf and vf.method == "IA":
            flags.add(Status.REVISION_MANUAL)
            reasons.append("Valor obtenido por IA: requiere confirmación humana")
        if ps.score < float(params["UMBRAL_RELACION"]) and "manual" not in c and len(o.primaries) == 1:
            flags.add(Status.REVISION_MANUAL)
            reasons.append(f"Relación con puntaje {ps.score:.2f} inferior al umbral (relación manual o por archivo)")

    if cmp.result == "FUERA_TOLERANCIA":
        flags.add(Status.EXCEPCION)
        reasons.append(f"Valor: {cmp.detail}")
    elif cmp.result in ("REVISION_MANUAL", "NO_DISPONIBLE"):
        flags.add(Status.REVISION_MANUAL)
        reasons.append(f"Valor: {cmp.detail}")

    for st in (Status.POSIBLE_DUPLICADO, Status.EXCEPCION, Status.REVISION_MANUAL):
        if st in flags:
            o.status = st
            break
    else:
        o.status = Status.COINCIDE if cmp.result == "EXACTO" else Status.COINCIDE_TOLERANCIA
        if o.status == Status.COINCIDE_TOLERANCIA:
            reasons.append(f"Valor: {cmp.detail}")
    o.reasons = list(dict.fromkeys(reasons))


def _save_result(db: Session, project_id: int, run_id: int, o: RowOutcome, tol: float) -> None:
    res = db.query(models.VouchingResult).filter_by(reference_item_id=o.ref.id).one_or_none()
    if res is None:
        res = models.VouchingResult(project_id=project_id, reference_item_id=o.ref.id)
        db.add(res)
    res.run_id = run_id
    res.auto_status = o.status
    res.reasons = o.reasons
    res.tolerance_applied = tol
    main = o.primaries[0][0] if o.primaries else None
    cmp = o.value_cmp
    for attr in ("extracted_number", "extracted_date", "extracted_party", "extracted_nit", "extracted_value", "abs_difference",
                 "pct_difference", "name_similarity", "nit_match", "dv_match", "number_match", "days_difference",
                 "ocr_confidence", "evidence_page", "evidence_text", "score"):
        setattr(res, attr, None)
    if main:
        d = main.unit.rep
        c = main.criteria
        res.score = main.score
        res.extracted_number = d.number
        res.extracted_date = d.date
        res.extracted_party = c["nombre"].found or (d.names[0][1].value if d.names else None)
        res.extracted_nit = c["nit"].found
        res.extracted_value = cmp.extracted if cmp else d.value
        res.abs_difference = cmp.abs_difference if cmp else None
        res.pct_difference = cmp.pct_difference if cmp else None
        res.name_similarity = round(c["nombre"].score * 100, 2) if c["nombre"].score is not None else None
        nc = compare_nit(o.ref.nit, c["nit"].found if c["nit"].score == 1.0 else (d.nits[0][3].value if d.nits else None))
        res.nit_match = nc.base_match
        res.dv_match = nc.dv_match
        res.number_match = c["numero"].detail
        if c["fecha"].score is not None and o.ref.doc_date and d.date:
            res.days_difference = (d.date - o.ref.doc_date).days
        vf = c["valor"].evidence
        res.ocr_confidence = vf.confidence if vf and vf.confidence is not None else d.ocr_confidence
        res.evidence_page = vf.page if vf else 1
        res.evidence_text = (vf.evidence if vf else None)
    res.explanation = {
        "estado": o.status,
        "motivos": o.reasons,
        "comparacion_valor": None if not cmp else {
            "valor_esperado": _fmt(cmp.expected), "valor_extraido": _fmt(cmp.extracted),
            "diferencia_absoluta": _fmt(cmp.abs_difference), "diferencia_porcentual": cmp.pct_difference,
            "tolerancia": cmp.tolerance, "resultado": cmp.result, "detalle": cmp.detail,
        },
        "soportes_principales": [
            {"documento_id": ps.unit.rep.id, "archivo": ps.unit.rep.filename, "valor_asignado": _fmt(alloc),
             "documentos_grupo": [d.filename for d in ps.unit.docs], **ps.explanation()}
            for ps, alloc in o.primaries
        ],
        "soportes_complementarios": [
            {"documento_id": ps.unit.rep.id, "archivo": ps.unit.rep.filename, "tipo": ps.unit.rep.doc_type, "puntaje_total": ps.score}
            for ps in o.complements
        ],
        "partidas_que_comparten_documento": o.shared_with,
        "mejor_candidato_descartado": (o.best.explanation() | {"archivo": o.best.unit.rep.filename}) if (o.best and not o.primaries) else None,
    }


def reconcile(db: Session, project_id: int, unit_alloc: dict | None = None) -> dict:
    """Conciliación de conteos y valores antes y después del procesamiento."""
    refs = db.query(models.ReferenceItem).filter_by(project_id=project_id).all()
    results = db.query(models.VouchingResult).filter_by(project_id=project_id).all()
    docs = db.query(models.Document).filter_by(project_id=project_id).all()
    imp = db.query(models.ImportBatch).filter_by(project_id=project_id, is_active=True).order_by(models.ImportBatch.id.desc()).first()
    links = db.query(models.MatchLink).filter_by(project_id=project_id).all()

    ref_ids = {r.id for r in refs}
    res_ids = {r.reference_item_id for r in results}
    total_expected = sum((r.expected_value or Decimal(0)) for r in refs)
    by_status: dict[str, dict] = defaultdict(lambda: {"partidas": 0, "valor": Decimal(0)})
    ref_val = {r.id: r.expected_value or Decimal(0) for r in refs}
    for res in results:
        b = by_status[res.status]
        b["partidas"] += 1
        b["valor"] += ref_val.get(res.reference_item_id, Decimal(0))
    sum_status_values = sum((b["valor"] for b in by_status.values()), Decimal(0))

    doc_status: dict[str, int] = defaultdict(int)
    for d in docs:
        doc_status[d.vouching_status] += 1
    hashes = defaultdict(int)
    for d in docs:
        hashes[d.sha256] += 1

    # Valor soportado: suma de asignaciones por documento principal, sin duplicar
    alloc_by_doc: dict[int, Decimal] = defaultdict(Decimal)
    for ln in links:
        if ln.role == "PRINCIPAL" and ln.allocated_value is not None:
            alloc_by_doc[ln.document_id] += ln.allocated_value
    doc_values = {}
    for d in docs:
        v = next((f.effective_value for f in d.fields if f.field_name == "valor_total"), None)
        doc_values[d.id] = parse_amount(v)
    over_allocated = [doc_id for doc_id, a in alloc_by_doc.items() if doc_values.get(doc_id) is not None and a > doc_values[doc_id] + Decimal("0.01")]

    checks = {
        "partidas_importadas": imp.row_count if imp else len(refs),
        "partidas_en_base": len(refs),
        "partidas_con_resultado": len(res_ids & ref_ids),
        "conteo_partidas_cuadra": (imp.row_count if imp else len(refs)) == len(refs) == len(res_ids & ref_ids),
        "valor_importado": str(imp.total_value if imp else total_expected),
        "valor_en_base": str(total_expected),
        "valor_por_estados": str(sum_status_values),
        "valor_cuadra": (Decimal(str(imp.total_value)) if imp else total_expected) == total_expected == sum_status_values,
        "documentos_cargados": len(docs),
        "documentos_por_estado": dict(doc_status),
        "documentos_clasificados": sum(doc_status.values()),
        "conteo_documentos_cuadra": sum(doc_status.values()) == len(docs),
        "archivos_identicos_repetidos": sum(c - 1 for c in hashes.values() if c > 1),
        "valor_asignado_a_soportes": str(sum(alloc_by_doc.values(), Decimal(0))),
        "documentos_sobreasignados": over_allocated,
        "sin_duplicacion_de_valor": not over_allocated,
        "partidas_por_estado": {k: {"partidas": v["partidas"], "valor": str(v["valor"])} for k, v in by_status.items()},
    }
    # Validación contra RESULTADO_ESPERADO (si se cargó)
    expected = db.query(models.ExpectedResult).filter_by(project_id=project_id).all()
    if expected:
        res_by_sample = {r.reference_item.sample_id: r for r in results}
        ok = 0
        diffs = []
        for e in expected:
            got = res_by_sample.get(e.sample_id)
            if got and got.status == e.expected_status:
                ok += 1
            else:
                diffs.append({"id_muestra": e.sample_id, "esperado": e.expected_status, "obtenido": got.status if got else None})
        checks["validacion_resultado_esperado"] = {"total": len(expected), "coinciden": ok, "precision": round(ok / len(expected), 4), "diferencias": diffs}
    return checks
