"""Importación del Excel de referencia con validación e informe de integridad.

Hojas reconocidas por nombre (sin distinguir mayúsculas/tildes):
CONFIGURACION, REFERENCIA_VOUCHING (obligatoria), ENTIDADES_ALIAS,
TIPOS_DOCUMENTO, CAMPOS_EXTRACCION y RESULTADO_ESPERADO.
También se acepta un CSV con las columnas de REFERENCIA_VOUCHING.
"""
from __future__ import annotations

import csv
import io
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal

from sqlalchemy.orm import Session

from .. import models
from .extraction.classifier import normalize_type
from .extraction.fields import fold
from .normalization import format_cop, normalize_doc_number, parse_amount, parse_date, parse_nit
from .settings_defaults import DEFAULT_PARAMETERS, merged_parameters, validate_parameters

SHEETS = ["CONFIGURACION", "REFERENCIA_VOUCHING", "ENTIDADES_ALIAS", "TIPOS_DOCUMENTO", "CAMPOS_EXTRACCION", "RESULTADO_ESPERADO"]

REFERENCE_COLUMNS: dict[str, list[str]] = {
    "ID_MUESTRA": ["ID_MUESTRA", "ID", "MUESTRA", "ID_PARTIDA"],
    "TIPO_DOCUMENTO": ["TIPO_DOCUMENTO", "TIPO_DOC", "TIPO"],
    "NUMERO_DOCUMENTO": ["NUMERO_DOCUMENTO", "NUMERO", "NO_DOCUMENTO", "NUM_DOCUMENTO", "DOCUMENTO", "NRO_DOCUMENTO"],
    "FECHA": ["FECHA", "FECHA_DOCUMENTO", "FECHA_CONTABLE", "FECHA_EMISION"],
    "TERCERO": ["TERCERO", "NOMBRE_TERCERO", "RAZON_SOCIAL", "PROVEEDOR", "CLIENTE", "NOMBRE"],
    "NIT": ["NIT", "NIT_TERCERO", "NIT_IDENTIFICACION", "IDENTIFICACION", "ID_TERCERO", "NIT_CC"],
    "VALOR_ESPERADO": ["VALOR_ESPERADO", "VALOR_TOTAL", "VALOR", "VALOR_CONTABLE", "IMPORTE", "TOTAL"],
    "MONEDA": ["MONEDA", "DIVISA"],
    "CONTRATO": ["CONTRATO", "NUMERO_CONTRATO"],
    "ORDEN_COMPRA": ["ORDEN_COMPRA", "OC", "NUMERO_OC"],
    "CONCEPTO": ["CONCEPTO", "DESCRIPCION", "DETALLE"],
    "CENTRO_COSTO": ["CENTRO_COSTO", "CENTRO_DE_COSTO", "CECO"],
    "CUENTA_CONTABLE": ["CUENTA_CONTABLE", "CUENTA"],
    "ARCHIVO_SOPORTE": ["ARCHIVO_SOPORTE", "ARCHIVO", "ARCHIVO_ESPERADO", "SOPORTE"],
    # Plantilla simple: una sola columna para contrato u orden de compra
    "CONTRATO_OC": ["CONTRATO_OC", "CONTRATO_U_OC", "CONTRATO_ORDEN_COMPRA", "REFERENCIA"],
    # Valores esperados complementarios (se comparan como referencia)
    "SUBTOTAL": ["SUBTOTAL", "SUB_TOTAL", "BASE"],
    "IVA": ["IVA", "VALOR_IVA"],
    "RETENCIONES": ["RETENCIONES", "RETENCION", "RETEFUENTE"],
    "TOLERANCIA_VALOR": ["TOLERANCIA_VALOR", "TOLERANCIA"],
    "OBSERVACIONES": ["OBSERVACIONES", "OBSERVACION", "NOTAS"],
}
# Hojas aceptadas como población cuando no existe REFERENCIA_VOUCHING (plantilla simple)
REFERENCE_SHEET_ALIASES = ["REFERENCIA_VOUCHING", "CARGA", "POBLACION", "MUESTRA", "DOCUMENTOS"]
REQUIRED_COLUMNS = ["ID_MUESTRA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO", "FECHA", "TERCERO", "NIT", "VALOR_ESPERADO"]

CONFIG_KEYS = {
    "TOLERANCIA_VALOR", "POLITICA_VALOR_CERO", "UMBRAL_NOMBRE", "TOLERANCIA_DIAS", "UMBRAL_RELACION",
    "UMBRAL_CONFIANZA_OCR", "MAX_SOPORTES_SUMA", "EXIGIR_FECHA_EN_TOLERANCIA", "MONEDA_BASE",
}


class ImportRejected(Exception):
    def __init__(self, report: dict):
        super().__init__("La carga fue rechazada por errores de integridad")
        self.report = report


@dataclass
class ParsedWorkbook:
    sheets: dict[str, list[list]] = field(default_factory=dict)
    sheet_names: list[str] = field(default_factory=list)


def _key(text) -> str:
    return fold(str(text or "")).strip().replace(" ", "_").replace("-", "_").replace(".", "")


def read_workbook(filename: str, data: bytes) -> ParsedWorkbook:
    ext = filename.rsplit(".", 1)[-1].lower()
    wb = ParsedWorkbook()
    if ext == "csv":
        text = data.decode("utf-8-sig", errors="replace")
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=";,\t|") if text.strip() else csv.excel
        rows = [r for r in csv.reader(io.StringIO(text), dialect)]
        wb.sheets["REFERENCIA_VOUCHING"] = rows
        wb.sheet_names = ["REFERENCIA_VOUCHING (CSV)"]
        return wb
    if ext == "xls":
        import xlrd

        book = xlrd.open_workbook(file_contents=data)
        for sh in book.sheets():
            rows = []
            for r in range(sh.nrows):
                row = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        row.append(xlrd.xldate.xldate_as_datetime(cell.value, book.datemode))
                    else:
                        row.append(cell.value)
                rows.append(row)
            wb.sheet_names.append(sh.name)
            wb.sheets[_key(sh.name)] = rows
        return wb
    from openpyxl import load_workbook

    book = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    for ws in book.worksheets:
        wb.sheet_names.append(ws.title)
        wb.sheets[_key(ws.title)] = [list(r) for r in ws.iter_rows(values_only=True)]
    book.close()
    return wb


def _known_columns(row: list) -> int:
    keys = {_key(c) for c in row if c not in (None, "")}
    return sum(1 for aliases in REFERENCE_COLUMNS.values() if keys & set(aliases))


def _table(rows: list[list], reference: bool = False) -> tuple[list[str], list[tuple[int, dict]]]:
    """Devuelve el encabezado y las filas como dicts.

    En la hoja de población el encabezado es la primera fila (de las 30 primeras) con al
    menos 3 columnas reconocidas, lo que permite títulos e instrucciones encima de la tabla.
    En las demás hojas es la primera fila no vacía.
    """
    header_idx = None
    if reference:
        header_idx = next((i for i, r in enumerate(rows[:30]) if _known_columns(r) >= 3), None)
    if header_idx is None:
        header_idx = next((i for i, r in enumerate(rows) if any(c not in (None, "") for c in r)), None)
    if header_idx is None:
        return [], []
    header = [_key(c) for c in rows[header_idx]]
    out = []
    for i, r in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if not any(c not in (None, "") and str(c).strip() != "" for c in r):
            continue
        out.append((i, {header[j]: r[j] for j in range(min(len(header), len(r))) if header[j]}))
    return header, out


def _resolve_columns(header: list[str]) -> dict[str, str]:
    mapping = {}
    for canonical, aliases in REFERENCE_COLUMNS.items():
        for a in aliases:
            if a in header:
                mapping[canonical] = a
                break
    return mapping


def _s(v) -> str | None:
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def _num(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("%", "")
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def parse_configuration(rows: list[list]) -> tuple[dict, dict, list[str]]:
    """Devuelve (parámetros del motor, datos del proyecto, advertencias)."""
    _, table = _table(rows)
    params: dict = {}
    project: dict = {}
    warnings: list[str] = []
    pesos = {}
    for rownum, r in table:
        name = _key(r.get("PARAMETRO") or r.get("CLAVE") or next(iter(r.values()), ""))
        raw = r.get("VALOR")
        if raw is None and len(r) >= 2:
            raw = list(r.values())[1]
        if not name:
            continue
        if name.startswith("PESO_"):
            v = _num(raw)
            if v is None:
                warnings.append(f"CONFIGURACION fila {rownum}: peso {name} no numérico")
            else:
                pesos[name[5:].lower()] = v
            continue
        if name in ("TOLERANCIA_VALOR", "UMBRAL_RELACION"):
            v = _num(raw)
            if v is not None and v > 1:
                v = v / 100  # 10 -> 0.10
            if isinstance(raw, str) and "%" in raw and v is not None and v > 1:
                v = v / 100
            params[name] = v
        elif name in ("UMBRAL_NOMBRE", "UMBRAL_CONFIANZA_OCR"):
            v = _num(raw)
            params[name] = v * 100 if v is not None and v <= 1 else v
        elif name in ("TOLERANCIA_DIAS", "MAX_SOPORTES_SUMA"):
            v = _num(raw)
            params[name] = int(v) if v is not None else None
        elif name == "EXIGIR_FECHA_EN_TOLERANCIA":
            params[name] = str(raw).strip().upper() in ("SI", "SÍ", "TRUE", "1", "VERDADERO", "S")
        elif name in ("POLITICA_VALOR_CERO", "MONEDA_BASE"):
            params[name] = _key(raw)
        elif name in ("NIT_CLIENTE", "NOMBRE_CLIENTE", "PERIODO", "CONTROL_TOTAL_REGISTROS", "CONTROL_TOTAL_VALOR"):
            project[name] = raw
        else:
            warnings.append(f"CONFIGURACION fila {rownum}: parámetro desconocido '{name}' (ignorado)")
    if pesos:
        params["PESOS"] = pesos
    params = {k: v for k, v in params.items() if v is not None}
    return params, project, warnings


def find_reference_sheet(wb: ParsedWorkbook) -> str | None:
    """Hoja de partidas: por nombre conocido o, si no, la primera con los encabezados obligatorios."""
    for name in REFERENCE_SHEET_ALIASES:
        if name in wb.sheets:
            return name
    for name, rows in wb.sheets.items():
        if name in SHEETS:
            continue
        header, _ = _table(rows, reference=True)
        if all(c in _resolve_columns(header) for c in ("ID_MUESTRA", "VALOR_ESPERADO")):
            return name
    return None


def split_contract_or_order(value: str) -> tuple[str | None, str | None]:
    """Columna CONTRATO_OC: decide por el prefijo si es contrato u orden de compra."""
    k = _key(value)
    if k.startswith(("OC", "ORDEN", "PO", "OS")):
        return None, value
    return value, None


def validate_workbook(wb: ParsedWorkbook) -> tuple[dict, dict]:
    """Valida el libro y devuelve (informe de integridad, datos normalizados)."""
    errors: list[str] = []
    warnings: list[str] = []
    info: list[str] = []
    row_issues: list[dict] = []
    normalized: dict = {"references": [], "aliases": [], "types": [], "fields": [], "expected": [], "params": {}, "project": {}}

    found = {s: s in wb.sheets for s in SHEETS}
    ref_sheet = find_reference_sheet(wb)
    if ref_sheet is None:
        errors.append("No se encontró la hoja de partidas: use REFERENCIA_VOUCHING o CARGA (plantilla simple) con los encabezados de la plantilla")
        return _report(wb, found, errors, warnings, info, row_issues, normalized), normalized
    found["REFERENCIA_VOUCHING"] = True
    if ref_sheet != "REFERENCIA_VOUCHING":
        info.append(f"Partidas leídas de la hoja {ref_sheet} (plantilla simple)")
    simple = ref_sheet != "REFERENCIA_VOUCHING" and not any(found[x] for x in SHEETS if x != "REFERENCIA_VOUCHING")
    for s, ok in found.items():
        if simple:
            break
        if not ok and s != "REFERENCIA_VOUCHING":
            info.append(f"Hoja {s} no encontrada: se usarán valores predeterminados")

    if found["CONFIGURACION"]:
        params, proj, w = parse_configuration(wb.sheets["CONFIGURACION"])
        warnings.extend(w)
        perrs = validate_parameters(merged_parameters(params))
        errors.extend(f"CONFIGURACION: {e}" for e in perrs)
        normalized["params"], normalized["project"] = params, proj

    header, table = _table(wb.sheets[ref_sheet], reference=True)
    mapping = _resolve_columns(header)
    missing = [c for c in REQUIRED_COLUMNS if c not in mapping]
    if missing:
        errors.append(f"Columnas obligatorias faltantes en {ref_sheet}: {', '.join(missing)}")
        return _report(wb, {ref_sheet: True} if simple else found, errors, warnings, info, row_issues, normalized, mapping=mapping), normalized

    def issue(row, col, msg, level="ADVERTENCIA"):
        row_issues.append({"fila": row, "columna": col, "nivel": level, "detalle": msg})

    ids = Counter()
    business = Counter()
    total = Decimal(0)
    for rownum, r in table:
        g = {c: r.get(src) for c, src in mapping.items()}
        sid = _s(g.get("ID_MUESTRA"))
        if not sid:
            issue(rownum, "ID_MUESTRA", "ID_MUESTRA vacío", "ERROR")
            continue
        ids[sid] += 1
        for col in REQUIRED_COLUMNS:
            if col != "ID_MUESTRA" and _s(g.get(col)) is None:
                issue(rownum, col, "Valor vacío en columna obligatoria")
        raw_val = g.get("VALOR_ESPERADO")
        value = parse_amount(raw_val) if raw_val not in (None, "") else None
        if raw_val not in (None, "") and value is None:
            issue(rownum, "VALOR_ESPERADO", f"Valor no numérico: '{raw_val}'", "ERROR")
        if value is not None:
            total += value
            if value < 0:
                issue(rownum, "VALOR_ESPERADO", "Valor negativo")
            if value == 0:
                issue(rownum, "VALOR_ESPERADO", "Valor esperado cero: se exigirá coincidencia exacta / revisión manual")
        raw_date = g.get("FECHA")
        d = parse_date(raw_date) if raw_date not in (None, "") else None
        if raw_date not in (None, "") and d is None:
            issue(rownum, "FECHA", f"Fecha inválida: '{raw_date}'")
        if d and not (date(1990, 1, 1) <= d <= date(2100, 12, 31)):
            issue(rownum, "FECHA", f"Fecha fuera de rango razonable: {d}")
            d = None
        nit_raw = _s(g.get("NIT"))
        nit = parse_nit(nit_raw) if nit_raw else None
        if nit_raw and not nit:
            issue(rownum, "NIT", f"NIT no válido: '{nit_raw}'")
        elif nit and nit.dv is not None and nit.dv_valid is False:
            issue(rownum, "NIT", f"Dígito de verificación no corresponde al NIT {nit_raw}")
        tipo = _s(g.get("TIPO_DOCUMENTO"))
        if tipo and normalize_type(tipo) == "OTRO" and _key(tipo) not in ("OTRO", "OTROS"):
            issue(rownum, "TIPO_DOCUMENTO", f"Tipo documental no reconocido: '{tipo}' (se tratará como OTRO)")
        moneda = (_s(g.get("MONEDA")) or normalized["params"].get("MONEDA_BASE") or DEFAULT_PARAMETERS["MONEDA_BASE"]).upper()
        if len(moneda) != 3:
            issue(rownum, "MONEDA", f"Moneda no estándar ISO 4217: '{moneda}'")
        bk = ((nit.base if nit else ""), normalize_doc_number(_s(g.get("NUMERO_DOCUMENTO"))), str(value))
        business[bk] += 1
        extra = {k: (v.isoformat() if isinstance(v, (date, datetime)) else v) for k, v in r.items() if k not in mapping.values() and v not in (None, "")}
        contract, order = _s(g.get("CONTRATO")), _s(g.get("ORDEN_COMPRA"))
        combined = _s(g.get("CONTRATO_OC"))
        if combined and not (contract or order):
            contract, order = split_contract_or_order(combined)
        for col in ("SUBTOTAL", "IVA", "RETENCIONES"):
            raw = g.get(col)
            if raw in (None, ""):
                continue
            amt = parse_amount(raw)
            if amt is None:
                issue(rownum, col, f"Valor no numérico: '{raw}'")
            else:
                extra[col] = str(amt)
        raw_tol = g.get("TOLERANCIA_VALOR")
        if raw_tol not in (None, ""):
            t = _num(raw_tol)
            if t is not None and t > 1:
                t = t / 100  # 10 -> 10 %
            if t is None or not 0 <= t <= 1:
                issue(rownum, "TOLERANCIA_VALOR", f"Tolerancia inválida: '{raw_tol}' (use 0,10 o 10 %); se aplica la del proyecto")
            else:
                extra["TOLERANCIA_VALOR"] = t
        obs = _s(g.get("OBSERVACIONES"))
        if obs:
            extra["OBSERVACIONES"] = obs
            if _key(obs) in ("EJEMPLO", "FILA_DE_EJEMPLO"):
                issue(rownum, "OBSERVACIONES", "Fila marcada como EJEMPLO: elimínela si no corresponde a un documento real")
        normalized["references"].append({
            "row_number": rownum, "sample_id": sid, "doc_type": tipo, "doc_number": _s(g.get("NUMERO_DOCUMENTO")),
            "doc_date": d, "third_party": _s(g.get("TERCERO")), "nit": nit_raw, "expected_value": value, "currency": moneda,
            "contract": contract, "purchase_order": order, "concept": _s(g.get("CONCEPTO")),
            "cost_center": _s(g.get("CENTRO_COSTO")), "account": _s(g.get("CUENTA_CONTABLE")),
            "expected_file": _s(g.get("ARCHIVO_SOPORTE")), "extra": extra,
        })
    dups = [k for k, c in ids.items() if c > 1]
    if dups:
        errors.append(f"ID_MUESTRA duplicados: {', '.join(dups[:20])}" + (" ..." if len(dups) > 20 else ""))
    dup_business = [k for k, c in business.items() if c > 1 and (k[0] or k[1])]
    if dup_business:
        warnings.append(f"{len(dup_business)} combinaciones NIT + número + valor repetidas (posibles partidas duplicadas)")
    if any(i["nivel"] == "ERROR" for i in row_issues):
        errors.append(f"{sum(1 for i in row_issues if i['nivel'] == 'ERROR')} filas con errores bloqueantes (ver detalle)")

    # Conciliación de cantidad de registros y suma de VALOR_ESPERADO contra los totales de control
    proj = normalized["project"]
    recon = {"registros_leidos": len(table), "registros_validos": len(normalized["references"]), "suma_valor_esperado": str(total)}
    if proj.get("CONTROL_TOTAL_REGISTROS") not in (None, ""):
        ctrl = int(_num(proj["CONTROL_TOTAL_REGISTROS"]) or 0)
        recon["control_total_registros"] = ctrl
        recon["registros_cuadran"] = ctrl == len(normalized["references"])
        if not recon["registros_cuadran"]:
            errors.append(f"La cantidad de registros ({len(normalized['references'])}) no concilia con CONTROL_TOTAL_REGISTROS ({ctrl})")
    if proj.get("CONTROL_TOTAL_VALOR") not in (None, ""):
        ctrl_v = parse_amount(proj["CONTROL_TOTAL_VALOR"])
        recon["control_total_valor"] = str(ctrl_v)
        recon["valor_cuadra"] = ctrl_v is not None and abs(ctrl_v - total) <= Decimal("0.01")
        if not recon["valor_cuadra"]:
            errors.append(f"La suma de VALOR_ESPERADO ({format_cop(total)}) no concilia con CONTROL_TOTAL_VALOR ({format_cop(ctrl_v)})" if ctrl_v is not None else "CONTROL_TOTAL_VALOR no es numérico")

    # Hojas auxiliares
    if found["ENTIDADES_ALIAS"]:
        _, t = _table(wb.sheets["ENTIDADES_ALIAS"])
        for rownum, r in t:
            canon = _s(r.get("NOMBRE_CANONICO") or r.get("NOMBRE") or r.get("RAZON_SOCIAL"))
            alias = _s(r.get("ALIAS") or r.get("VARIANTE"))
            if not canon or not alias:
                warnings.append(f"ENTIDADES_ALIAS fila {rownum}: nombre canónico o alias vacío")
                continue
            normalized["aliases"].append({"nit": _s(r.get("NIT")), "canonical_name": canon, "alias": alias})
    if found["TIPOS_DOCUMENTO"]:
        _, t = _table(wb.sheets["TIPOS_DOCUMENTO"])
        for rownum, r in t:
            code = _s(r.get("CODIGO") or r.get("TIPO"))
            if not code:
                continue
            kws = [k.strip() for k in str(r.get("PALABRAS_CLAVE") or "").replace(",", ";").split(";") if k.strip()]
            normalized["types"].append({
                "code": normalize_type(code) or _key(code), "name": _s(r.get("NOMBRE") or r.get("DESCRIPCION")) or code,
                "keywords": kws, "is_primary": str(r.get("SOPORTA_VALOR") or "SI").strip().upper() in ("SI", "SÍ", "S", "TRUE", "1"),
            })
    if found["CAMPOS_EXTRACCION"]:
        _, t = _table(wb.sheets["CAMPOS_EXTRACCION"])
        for rownum, r in t:
            f = _s(r.get("CAMPO"))
            if not f:
                continue
            normalized["fields"].append({
                "field": f.lower(), "description": _s(r.get("DESCRIPCION")),
                "required": str(r.get("OBLIGATORIO") or "").strip().upper() in ("SI", "SÍ", "S", "TRUE", "1"),
                "data_type": (_s(r.get("TIPO_DATO")) or "TEXTO").upper(),
                "doc_types": [x.strip() for x in str(r.get("TIPOS_DOCUMENTO") or "").replace(",", ";").split(";") if x.strip()],
            })
    if found["RESULTADO_ESPERADO"]:
        _, t = _table(wb.sheets["RESULTADO_ESPERADO"])
        known = {r["sample_id"] for r in normalized["references"]}
        for rownum, r in t:
            sid = _s(r.get("ID_MUESTRA"))
            st = _s(r.get("ESTADO_ESPERADO") or r.get("ESTADO"))
            if not sid or not st:
                continue
            st_norm = next((s for s in models.Status.ALL if fold(s) == fold(st).strip()), None)
            if not st_norm:
                warnings.append(f"RESULTADO_ESPERADO fila {rownum}: estado '{st}' no reconocido")
                continue
            if sid not in known:
                warnings.append(f"RESULTADO_ESPERADO fila {rownum}: ID_MUESTRA {sid} no existe en REFERENCIA_VOUCHING")
            normalized["expected"].append({"sample_id": sid, "expected_status": st_norm, "expected_file": _s(r.get("ARCHIVO")), "note": _s(r.get("OBSERVACION"))})

    sheets_view = {ref_sheet: True} if simple else found
    return _report(wb, sheets_view, errors, warnings, info, row_issues, normalized, mapping=mapping, recon=recon, total=total), normalized


def _report(wb, found, errors, warnings, info, row_issues, normalized, mapping=None, recon=None, total=Decimal(0)) -> dict:
    status = "RECHAZADA" if errors else ("CON_ADVERTENCIAS" if warnings or row_issues else "VALIDA")
    return {
        "aplicativo": "Muenra Vouching",
        "generado": datetime.now().isoformat(timespec="seconds"),
        "estado": status,
        "hojas_encontradas": wb.sheet_names,
        "hojas_reconocidas": found,
        "columnas_mapeadas": mapping or {},
        "errores": errors,
        "advertencias": warnings,
        "informacion": info,
        "incidencias_por_fila": row_issues[:2000],
        "total_incidencias": len(row_issues),
        "conciliacion": recon or {},
        "resumen": {
            "partidas": len(normalized["references"]),
            "valor_total": str(total),
            "alias": len(normalized["aliases"]),
            "tipos_documento": len(normalized["types"]),
            "campos_extraccion": len(normalized["fields"]),
            "resultados_esperados": len(normalized["expected"]),
            "parametros": normalized["params"],
        },
    }


def import_workbook(db: Session, project: models.Project, filename: str, data: bytes, sha: str, user: models.User | None, dry_run: bool = False) -> tuple[models.ImportBatch | None, dict]:
    wb = read_workbook(filename, data)
    report, norm = validate_workbook(wb)
    if dry_run:
        return None, report
    if report["estado"] == "RECHAZADA":
        raise ImportRejected(report)

    # Reemplaza la población anterior del proyecto (queda registrado en la auditoría)
    for old in db.query(models.ImportBatch).filter_by(project_id=project.id, is_active=True):
        old.is_active = False
    db.query(models.MatchLink).filter_by(project_id=project.id).delete()
    db.query(models.VouchingResult).filter_by(project_id=project.id).delete()
    db.query(models.ReferenceItem).filter_by(project_id=project.id).delete()
    for model in (models.EntityAlias, models.DocumentTypeDef, models.ExtractionFieldDef, models.ExpectedResult):
        db.query(model).filter_by(project_id=project.id).delete()

    total = sum((r["expected_value"] or Decimal(0)) for r in norm["references"])
    batch = models.ImportBatch(
        project_id=project.id, filename=filename, sha256=sha, status=report["estado"], row_count=len(norm["references"]),
        total_value=total, integrity_report=report, created_by=user.id if user else None,
    )
    db.add(batch)
    db.flush()
    for r in norm["references"]:
        db.add(models.ReferenceItem(project_id=project.id, import_id=batch.id, **r))
    for a in norm["aliases"]:
        db.add(models.EntityAlias(project_id=project.id, **a))
    for t in norm["types"]:
        db.add(models.DocumentTypeDef(project_id=project.id, **t))
    for f in norm["fields"]:
        db.add(models.ExtractionFieldDef(project_id=project.id, **f))
    for e in norm["expected"]:
        db.add(models.ExpectedResult(project_id=project.id, **e))
    if norm["params"]:
        project.settings = {**(project.settings or {}), **norm["params"]}
    proj = norm["project"]
    if proj.get("NIT_CLIENTE"):
        project.client_nit = _s(proj["NIT_CLIENTE"])
    if proj.get("NOMBRE_CLIENTE"):
        project.client_name = _s(proj["NOMBRE_CLIENTE"])
    if proj.get("PERIODO"):
        project.period = _s(proj["PERIODO"])
    db.flush()
    # Resultados en estado PENDIENTE para todas las partidas
    for ref in db.query(models.ReferenceItem).filter_by(project_id=project.id):
        db.add(models.VouchingResult(project_id=project.id, reference_item_id=ref.id, auto_status=models.Status.PENDIENTE, reasons=[], explanation={}))
    db.commit()
    return batch, report
