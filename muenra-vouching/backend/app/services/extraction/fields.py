"""Extracción de campos por reglas (etiquetas + expresiones regulares) con evidencia.

Trabaja sobre líneas con coordenadas producidas por los lectores (PDF/OCR/DOCX).
Cada campo devuelto conserva la página, la región (bbox), el texto de evidencia,
el método y la confianza de la línea de origen.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from ..normalization import parse_amount, parse_date, parse_nit
from .base import FieldEvidence, Line, Page


def fold(text: str) -> str:
    """Mayúsculas sin tildes conservando la longitud (posiciones 1:1 con el original)."""
    out = []
    for ch in text:
        base = unicodedata.normalize("NFKD", ch)[0]
        up = base.upper()
        out.append(up if len(up) == 1 else base)
    return "".join(out)


AMOUNT_RE = re.compile(
    r"(?<![\w.,])\(?-?\s*(?:\$|COP|USD|US\$|EUR)?\s*(?:\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)\)?(?![\w%])"
)
PERCENT_RE = re.compile(r"\d{1,2}(?:[.,]\d+)?\s*%")
DATE_RE = re.compile(
    r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{4}|\d{1,2}\s*[-/]?\s*(?:DE\s+)?[A-Z]{3,10}\.?\s*(?:DE(?:L)?\s+|[-/ ])\s*\d{4}|[A-Z]{3,10}\s+\d{1,2}\s*(?:DE|,)?\s*\d{4}"
)
NIT_RE = re.compile(r"\d{1,3}(?:[.,]\d{3}){2,3}\s*-\s*\d\b|\d{1,3}(?:[.,]\d{3}){2,3}(?!\d)|\d{6,10}\s*-\s*\d\b|\d{6,10}(?!\d)")
DOCNUM_RE = re.compile(r"[A-Z]{0,6}\s?-?\s?\d[\dA-Z/-]*")
REF_RE = re.compile(r"[A-Z]{0,8}[-/ ]?\d[\dA-Z/-]*")
CUFE_RE = re.compile(r"[0-9A-F]{96}")
ACCOUNT_RE = re.compile(r"\d[\d-]{5,}\d")

NUM_LABEL = r"\s*(?:ELECTRONICA\s*)?(?:DE\s+VENTA\s*)?(?:N[O0]\.?|N[°º]\.?|NRO\.?|NUM\.?|NUMERO|#)\s*[:.]?\s*"

NUMBER_LABELS: dict[str, list[str]] = {
    "FACTURA": [r"FACTURA" + NUM_LABEL, r"FACTURA\s+(?:ELECTRONICA\s+)?(?:DE\s+VENTA\s+)?(?=[A-Z]{1,5}-?\d)"],
    "RECIBO_CAJA": [r"RECIBO\s+DE\s+CAJA" + NUM_LABEL, r"\bRC" + NUM_LABEL],
    "COMPROBANTE_EGRESO": [r"COMPROBANTE\s+DE\s+EGRESO" + NUM_LABEL, r"\bCE" + NUM_LABEL, r"EGRESO" + NUM_LABEL],
    "CUENTA_COBRO": [r"CUENTA\s+DE\s+COBRO" + NUM_LABEL],
    "NOTA_CREDITO": [r"NOTA\s+CREDITO" + NUM_LABEL],
    "NOTA_DEBITO": [r"NOTA\s+DEBITO" + NUM_LABEL],
    "DOCUMENTO_EQUIVALENTE": [r"DOCUMENTO\s+(?:EQUIVALENTE|SOPORTE)" + NUM_LABEL],
    "ORDEN_COMPRA": [r"ORDEN\s+DE\s+COMPRA" + NUM_LABEL],
    "CONTRATO": [r"CONTRATO(?:\s+DE\s+[A-Z ]{3,40}?)?" + NUM_LABEL],
    "OTROSI": [r"OTRO\s?SI" + NUM_LABEL],
    "CERTIFICACION": [r"CERTIFICA(?:DO|CION)" + NUM_LABEL],
    "COMPROBANTE_CONTABLE": [r"COMPROBANTE(?:\s+CONTABLE|\s+DE\s+DIARIO)?" + NUM_LABEL],
    "SOPORTE_PAGO": [r"(?:REFERENCIA|COMPROBANTE|TRANSACCION|OPERACION)" + NUM_LABEL],
    "EXTRACTO_BANCARIO": [r"EXTRACTO" + NUM_LABEL],
}
GENERIC_NUMBER_LABEL = [r"(?<![A-Z])(?:N[O0]\.|N[°º]|NRO\.?|NUMERO)\s*[:.]?\s*"]

TYPE_FAMILY = {
    "FACTURA_VENTA": "FACTURA",
    "FACTURA_COMPRA": "FACTURA",
    "FACTURA_ELECTRONICA": "FACTURA",
}

_STRIP_LABEL_RE = re.compile(
    r"(?i)^.*?(?:(?:raz[oó]n social|se[nñ]or(?:es)?|pagado a|recibimos de|recib[ií] de|a favor de|debe a|beneficiario)\s*[:.]?"
    r"|(?:cliente|adquiri?ente|proveedor|emisor|contratista|contratante|vendedor)\s*:)\s*"
)

RECEIVER_HINTS = re.compile(
    r"CLIENTE\s*:|NIT\s+(?:DEL\s+)?CLIENTE|ADQUIRI?ENTE\s*:|SENOR(?:ES)?\b|RECEPTOR\s*:|COMPRADOR\s*:|PAGADO\s+A\b|BENEFICIARIO|"
    r"RECIBIMOS\s+DE|RECIBI\s+DE|CONTRATISTA\s*:|A\s+FAVOR\s+DE|PAGAR\s+A\b|DEBE\s+A\b|FACTURAR\s+A|ADQUIRENTE|PROVEEDOR\s*:"
)


@dataclass
class _Hit:
    page: Page
    index: int
    line: Line
    value: str
    evidence: str
    bbox: list[float] | None
    conf: float | None


class LineIndex:
    def __init__(self, pages: list[Page]):
        self.pages = pages
        self.items: list[tuple[Page, int, Line, str]] = []
        for p in pages:
            for i, line in enumerate(p.lines):
                self.items.append((p, i, line, fold(line.text)))

    def neighbors(self, page: Page, index: int) -> list[Line]:
        """Líneas a la derecha en la misma fila (True) y luego la siguiente línea (False)."""
        base = page.lines[index]
        out: list[tuple[Line, bool]] = []
        if base.bbox:
            cy = (base.bbox[1] + base.bbox[3]) / 2
            right = [
                ln for ln in page.lines
                if ln is not base and ln.bbox and ln.bbox[0] >= base.bbox[2] - 0.01 and ln.bbox[1] <= cy <= ln.bbox[3]
            ]
            out.extend((ln, True) for ln in sorted(right, key=lambda ln: ln.bbox[0]))
        if index + 1 < len(page.lines):
            nxt = page.lines[index + 1]
            if all(nxt is not ln for ln, _ in out):
                out.append((nxt, False))
        return out

    def find(self, labels: list[str], value_re: re.Pattern | None, pick: str = "first", clean=None) -> _Hit | None:
        """Busca la primera etiqueta (en orden de prioridad) y su valor."""
        for label in labels:
            lab = re.compile(label)
            hits: list[_Hit] = []
            for page, idx, line, folded in self.items:
                for m in lab.finditer(folded):
                    rest_orig = line.text[m.end():]
                    rest_fold = folded[m.end():]
                    hit = self._value(page, idx, line, rest_orig, rest_fold, value_re, clean)
                    if hit:
                        hits.append(hit)
                        break
            if hits:
                return hits[-1] if pick == "last" else hits[0]
        return None

    def _value(self, page, idx, line, rest_orig, rest_fold, value_re, clean):
        def match_in(orig: str, folded: str, whole_line: bool = False):
            if value_re is None:
                v = orig.strip(" :.-\t")
                if clean and v:
                    v = clean(v)
                return v or None
            src = PERCENT_RE.sub(lambda mm: " " * len(mm.group(0)), folded) if value_re is AMOUNT_RE else folded
            for vm in value_re.finditer(src):
                if whole_line and len(src.strip()) - len(vm.group(0).strip()) > 8:
                    return None  # la línea siguiente contiene otra información: no es el valor de esta etiqueta
                val = orig[vm.start():vm.end()].strip()
                if value_re is AMOUNT_RE:
                    val = PERCENT_RE.sub("", val).strip()
                if clean:
                    val = clean(val)
                if val:
                    return val
            return None

        v = match_in(rest_orig, rest_fold)
        if v:
            return _Hit(page, idx, line, v, line.text, line.bbox, line.conf)
        if rest_orig.strip(" :.-\t") and value_re is None:
            return None
        for nb, same_row in self.neighbors(page, idx)[:3]:
            v = match_in(nb.text, fold(nb.text), whole_line=not same_row)
            if v:
                return _Hit(page, idx, line, v, f"{line.text} {nb.text}", _union(line.bbox, nb.bbox), _min_conf(line.conf, nb.conf))
        return None


def _union(a, b):
    if not a:
        return b
    if not b:
        return a
    return [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])]


def _min_conf(a, b):
    vals = [v for v in (a, b) if v is not None]
    return min(vals) if vals else None


def _fe(name: str, hit: _Hit, normalized: str | None = None) -> FieldEvidence:
    return FieldEvidence(
        name=name,
        value=hit.value,
        normalized=normalized if normalized is not None else hit.value,
        page=hit.page.number,
        bbox=hit.bbox,
        evidence_text=hit.evidence,
        method=hit.page.method,
        confidence=hit.conf,
    )


def _clean_party(val: str) -> str:
    val = re.split(r"(?i)\b(?:NIT|N\.I\.T|C\.?C\.?|IDENTIFICACI[OÓ]N|CEDULA|C[EÉ]DULA)\b", val)[0]
    return val.strip(" :.-,\t")


def _clean_docnum(val: str) -> str:
    val = val.replace(" ", "").strip("-/")
    return val if re.search(r"\d", val) else ""


def _amount_fe(idx: LineIndex, name: str, labels: list[str], pick: str = "first") -> FieldEvidence | None:
    hit = idx.find(labels, AMOUNT_RE, pick=pick)
    if not hit:
        return None
    amt = parse_amount(hit.value)
    if amt is None:
        return None
    return _fe(name, hit, str(amt))


def extract_fields(pages: list[Page], doc_type: str | None = None) -> list[FieldEvidence]:
    idx = LineIndex(pages)
    out: list[FieldEvidence] = []

    def add(fe: FieldEvidence | None):
        if fe and fe.value not in (None, "") and not any(f.name == fe.name for f in out):
            out.append(fe)

    # Número documental: primero etiquetas del tipo detectado, luego el resto
    family = TYPE_FAMILY.get(doc_type or "", doc_type or "")
    ordered = []
    if family in NUMBER_LABELS:
        ordered.extend(NUMBER_LABELS[family])
    for k, labels in NUMBER_LABELS.items():
        if k != family and k not in ("CONTRATO", "ORDEN_COMPRA", "OTROSI"):
            ordered.extend(labels)
    ordered.extend(GENERIC_NUMBER_LABEL)
    hit = idx.find(ordered, DOCNUM_RE, clean=_clean_docnum)
    if hit:
        add(_fe("numero_documento", hit))

    # Fechas
    hit = idx.find(
        [
            r"FECHA\s+(?:DE\s+)?(?:EMISION|EXPEDICION|GENERACION|FACTURA(?:CION)?|ELABORACION|SUSCRIPCION|PAGO|TRANSACCION)",
            r"FECHA\s+(?:DEL?\s+)?(?:RECIBO|COMPROBANTE|DOCUMENTO|CERTIFICADO|CONTRATO)",
            r"(?<![A-Z])FECHA(?!\s+(?:DE\s+)?VENC)",
            r"(?:BOGOTA|MEDELLIN|CALI|BARRANQUILLA|CARTAGENA|BUCARAMANGA)(?:\s+D\.?C\.?)?\s*,",
        ],
        DATE_RE,
    )
    if hit:
        d = parse_date(fold(hit.value))
        if d:
            add(_fe("fecha_emision", hit, d.isoformat()))
    hit = idx.find([r"FECHA\s+(?:DE\s+)?VENCIMIENTO", r"(?<![A-Z])VENCE", r"VENCIMIENTO"], DATE_RE)
    if hit:
        d = parse_date(fold(hit.value))
        if d:
            add(_fe("fecha_vencimiento", hit, d.isoformat()))

    # Valores
    add(_amount_fe(idx, "valor_total", [
        r"TOTAL\s+A\s+PAGAR", r"NETO\s+A\s+PAGAR", r"VALOR\s+TOTAL(?:\s+A\s+PAGAR)?", r"TOTAL\s+FACTURA",
        r"VALOR\s+(?:A\s+)?PAGAD[OA]", r"TOTAL\s+PAGADO", r"VALOR\s+A\s+PAGAR", r"VALOR\s+DEL\s+CONTRATO",
        r"LA\s+(?:SUMA|CANTIDAD)\s+DE", r"VALOR\s+(?:DE\s+LA\s+)?TRANSFERENCIA", r"VALOR\s+RECIBIDO",
        r"MONTO", r"(?<![A-Z])VALOR\s*[:$]",
    ]))
    if not any(f.name == "valor_total" for f in out):
        add(_amount_fe(idx, "valor_total", [r"(?<![A-Z])TOTAL(?![A-Z])"], pick="last"))
    add(_amount_fe(idx, "subtotal", [r"SUB\s?-?TOTAL", r"VALOR\s+ANTES\s+DE\s+IVA"]))
    add(_amount_fe(idx, "base_gravable", [r"BASE\s+GRAVABLE", r"BASE\s+IMPONIBLE", r"TOTAL\s+BASE"]))
    add(_amount_fe(idx, "iva", [r"(?<![A-Z])IVA(?!\s*RETENIDO)(?![A-Z])", r"IMPUESTO\s+AL\s+VALOR\s+AGREGADO"]))
    add(_amount_fe(idx, "retenciones", [r"TOTAL\s+RETENCIONES", r"RETENCION(?:ES)?(?:\s+EN\s+LA\s+FUENTE)?", r"RETE\s?FUENTE", r"RETE\s?ICA", r"RETE\s?IVA"]))
    add(_amount_fe(idx, "descuentos", [r"DESCUENTOS?"]))

    # Moneda
    full = fold("\n".join(p.text for p in pages))
    hit = idx.find([r"MONEDA", r"DIVISA"], re.compile(r"COP|USD|EUR|PESOS?|DOLAR(?:ES)?"))
    if hit:
        val = fold(hit.value)
        add(_fe("moneda", hit, "USD" if "USD" in val or "DOLAR" in val else "EUR" if "EUR" in val else "COP"))
    elif re.search(r"\bUSD\b|US\$|DOLARES", full):
        add(FieldEvidence("moneda", "USD", "USD", method="REGLAS", confidence=70.0, evidence_text="Mención de USD en el documento"))
    elif "$" in full or "COP" in full or "PESOS" in full:
        add(FieldEvidence("moneda", "COP", "COP", method="REGLAS", confidence=70.0, evidence_text="Símbolo $ / pesos en el documento"))

    # CUFE / CUDE (puede estar partido en varias líneas por el OCR)
    m = CUFE_RE.search(re.sub(r"[^0-9A-F]", "", full.split("CUFE", 1)[-1])[:200]) if "CUFE" in full or "CUDE" in full else None
    if m:
        for page, i, line, folded in idx.items:
            if "CUFE" in folded or "CUDE" in folded:
                add(FieldEvidence("cufe", m.group(0).lower(), m.group(0).lower(), page.number, line.bbox, line.text, page.method, line.conf))
                break

    # Terceros y NIT
    _extract_parties(idx, add)

    # Referencias
    hit = idx.find([r"CONTRATO(?:\s+DE\s+[A-Z ]{3,40}?)?\s*(?:N[O0]\.?|N[°º]\.?|NRO\.?|NUMERO|#)\s*[:.]?\s*", r"CONTRATO\s*[:]\s*"], REF_RE, clean=_clean_docnum)
    if hit:
        add(_fe("numero_contrato", hit))
    hit = idx.find([r"ORDEN\s+DE\s+COMPRA\s*(?:N[O0]\.?|N[°º]\.?|NRO\.?|NUMERO|#)?\s*[:.]?\s*", r"(?<![A-Z])O\.?\s?C\.?\s*(?:N[O0]\.?|N[°º]|NRO\.?|#)\s*[:.]?\s*"], REF_RE, clean=_clean_docnum)
    if hit:
        add(_fe("orden_compra", hit))
    hit = idx.find([r"CENTRO\s+DE\s+COSTOS?\s*[:.]?", r"(?<![A-Z])C\.?\s?COSTO\s*[:.]?", r"(?<![A-Z])CECO\s*[:.]?"], re.compile(r"[A-Z0-9][A-Z0-9 -]{0,40}"))
    if hit:
        add(_fe("centro_costo", hit, hit.value.strip()))
    hit = idx.find([r"POR\s+CONCEPTO\s+DE\s*[:]?", r"CONCEPTO\s*[:]?", r"DESCRIPCION\s*[:]?", r"DETALLE\s*[:]?"], None)
    if hit:
        add(_fe("concepto", hit))
    hit = idx.find([r"FORMA\s+DE\s+PAGO\s*[:]?", r"MEDIO\s+DE\s+PAGO\s*[:]?", r"CONDICIONES?\s+DE\s+PAGO\s*[:]?"], None)
    if hit:
        add(_fe("forma_pago", hit))
    hit = idx.find([r"CUENTA\s+(?:DE\s+)?(?:AHORROS|CORRIENTE|BANCARIA)\s*(?:N[O0]\.?|N[°º]|NRO\.?|#)?\s*[:.]?\s*", r"(?<![A-Z])CTA\.?\s*(?:N[O0]\.?|#)?\s*"], ACCOUNT_RE)
    if hit:
        add(_fe("cuenta_bancaria", hit))
    hit = idx.find([r"OBJETO(?:\s+DEL\s+CONTRATO)?\s*[:.]", r"PRIMERA\s*[.:-]\s*OBJETO\s*[:.]?"], None)
    if hit:
        add(_fe("objeto_contractual", hit))
    hit = idx.find([r"VIGENCIA\s*[:.]?", r"PLAZO(?:\s+DE\s+EJECUCION)?\s*[:.]?", r"DURACION\s*[:.]?"], None)
    if hit:
        add(_fe("vigencia_contrato", hit))

    signers = []
    sig_re = re.compile(r"FIRMADO\s+POR\s*[:]?|FIRMA\s*[:]|REPRESENTANTE\s+LEGAL\s*[:]|ELABORADO\s+POR\s*[:]?|APROBADO\s+POR\s*[:]?|REVISADO\s+POR\s*[:]?")
    first_sig = None
    for page, i, line, folded in idx.items:
        m = sig_re.search(folded)
        if m:
            name = line.text[m.end():].strip(" :.-")
            if not name:
                continue
            signers.append(name)
            first_sig = first_sig or (page, line)
    if signers and first_sig:
        page, line = first_sig
        add(FieldEvidence("firmantes", "; ".join(dict.fromkeys(signers)), None, page.number, line.bbox, line.text, page.method, line.conf))
    return out


def _extract_parties(idx: LineIndex, add) -> None:
    label = re.compile(r"(?<![A-Z])(?:NIT|N\.\s?I\.\s?T\.?|C\.\s?C\.|CEDULA|IDENTIFICACION|DOCUMENTO\s+DE\s+IDENTIDAD)\s*(?:N[O0]\.?|N[°º]|NRO\.?)?\s*[:.]?\s*")
    occurrences = []
    for page, i, line, folded in idx.items:
        for m in label.finditer(folded):
            rest = folded[m.end():]
            nm = NIT_RE.match(rest.lstrip())
            if not nm:
                continue
            offset = m.end() + (len(rest) - len(rest.lstrip()))
            raw = line.text[offset + nm.start(): offset + nm.end()]
            context = fold(" ".join(ln.text for ln in page.lines[max(0, i - 2): i + 1]))
            occurrences.append((page, i, line, raw, bool(RECEIVER_HINTS.search(context))))
    emitter = next((o for o in occurrences if not o[4]), None)
    receiver = next((o for o in occurrences if o[4] and o is not emitter), None)
    if receiver is None:
        receiver = next((o for o in occurrences if o is not emitter), None)
    others = [o for o in occurrences if o is not emitter and o is not receiver]
    seen = set()
    for n, occ in enumerate(others, start=1):
        parts = parse_nit(occ[3])
        if parts and parts.base not in seen:
            seen.add(parts.base)
            add(FieldEvidence(f"otro{n}_nit", occ[3].strip(), parts.base, occ[0].number, occ[2].bbox, occ[2].text, occ[0].method, occ[2].conf))
    for role, occ in (("emisor", emitter), ("receptor", receiver)):
        if not occ:
            continue
        page, i, line, raw, _ = occ
        parts = parse_nit(raw)
        if not parts:
            continue
        add(FieldEvidence(f"{role}_nit", raw.strip(), parts.base, page.number, line.bbox, line.text, page.method, line.conf))
        if parts.dv is not None:
            add(FieldEvidence(f"{role}_dv", parts.dv, parts.dv, page.number, line.bbox, line.text, page.method, line.conf))
        # Nombre: misma línea antes de "NIT" o etiqueta explícita en las 2 líneas previas
        name = _clean_party(line.text)
        name = _STRIP_LABEL_RE.sub("", name).strip(" :.-")
        if len(re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", name)) < 3 and i > 0:
            prev = page.lines[i - 1].text
            name = _STRIP_LABEL_RE.sub("", _clean_party(prev)).strip(" :.-")
            ev_line = page.lines[i - 1]
        else:
            ev_line = line
        if len(re.sub(r"[^A-Za-zÁÉÍÓÚÑáéíóúñ]", "", name)) >= 3:
            add(FieldEvidence(f"{role}_nombre", name, None, page.number, ev_line.bbox, ev_line.text, page.method, ev_line.conf))
    # Nombres con etiqueta explícita cuando no hubo NIT
    hit = idx.find([r"(?:SENOR(?:ES)?|CLIENTE|ADQUIRI?ENTE|PAGADO\s+A|BENEFICIARIO|RECIBIMOS\s+DE|RECIBI\s+DE|A\s+FAVOR\s+DE|DEBE\s+A)\s*[:.]?\s*"], None, clean=_clean_party)
    if hit:
        add(_fe("receptor_nombre", hit))
    hit = idx.find([r"RAZON\s+SOCIAL\s*[:.]?\s*", r"(?:EMISOR|PROVEEDOR|VENDEDOR|CONTRATISTA)\s*[:.]\s*"], None, clean=_clean_party)
    if hit:
        add(_fe("emisor_nombre", hit))
