"""Extracción determinística de XML de facturación electrónica (UBL 2.1 DIAN).

Soporta Invoice, CreditNote, DebitNote y el contenedor AttachedDocument (que
incluye la factura embebida en un CDATA). El análisis es seguro: sin entidades
externas, sin DTD y sin acceso a red (protección XXE / "billion laughs").
"""
from __future__ import annotations

from decimal import Decimal

from lxml import etree

from ..normalization import parse_amount, parse_nit
from .base import ExtractionResult, FieldEvidence, Line, Method, Page

SAFE_PARSER = etree.XMLParser(
    resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False, remove_comments=True
)

PAYMENT_MEANS = {"1": "Contado", "2": "Crédito"}


class InvalidXML(Exception):
    pass


def _parse(data: bytes):
    try:
        root = etree.fromstring(data, parser=SAFE_PARSER)
    except etree.XMLSyntaxError as exc:
        raise InvalidXML(f"XML mal formado: {exc}") from exc
    if root is None:
        raise InvalidXML("XML vacío")
    return root


def _local(el) -> str:
    return etree.QName(el).localname


def _xp(node, path: str):
    """XPath independiente de namespaces: 'A/B/C' -> *[local-name()='A']/..."""
    parts = []
    for p in path.split("/"):
        if p in ("", "."):
            parts.append(p)
        elif p.startswith("@"):
            parts.append(p)
        else:
            parts.append(f"*[local-name()='{p}']")
    return node.xpath("/".join(parts))


def _first_text(node, *paths: str) -> tuple[str | None, str | None, object | None]:
    for path in paths:
        found = _xp(node, path)
        for el in found:
            text = el.text.strip() if el.text and el.text.strip() else None
            if text:
                return text, path, el
    return None, None, None


def unwrap_attached_document(root):
    """Si es un AttachedDocument, devuelve la factura embebida."""
    if _local(root) != "AttachedDocument":
        return root, None
    desc = _xp(root, "Attachment/ExternalReference/Description")
    for d in desc:
        if d.text and "<" in d.text:
            try:
                inner = _parse(d.text.strip().encode("utf-8"))
                return inner, "AttachedDocument"
            except InvalidXML:
                continue
    return root, "AttachedDocument"


def _party(node, role: str):
    base = f"{role}/Party"
    name, name_path, _ = _first_text(
        node,
        f"{base}/PartyTaxScheme/RegistrationName",
        f"{base}/PartyLegalEntity/RegistrationName",
        f"{base}/PartyName/Name",
    )
    nit_text, nit_path, nit_el = _first_text(
        node, f"{base}/PartyTaxScheme/CompanyID", f"{base}/PartyLegalEntity/CompanyID", f"{base}/PartyIdentification/ID"
    )
    dv = nit_el.get("schemeID") if nit_el is not None else None
    return name, name_path, nit_text, nit_path, dv


def _sum_amounts(node, path: str) -> Decimal | None:
    vals = [parse_amount(el.text) for el in _xp(node, path) if el.text]
    vals = [v for v in vals if v is not None]
    return sum(vals, Decimal("0")) if vals else None


def extract_xml(data: bytes) -> ExtractionResult:
    root = _parse(data)
    doc, wrapper = unwrap_attached_document(root)
    kind = _local(doc)
    result = ExtractionResult(file_type="xml", method=Method.XML, xml_root=kind)
    if wrapper:
        result.warnings.append("Factura obtenida desde AttachedDocument (contenedor DIAN)")

    lines: list[Line] = []

    def add(name: str, value, path: str | None, normalized: str | None = None):
        if value is None or value == "":
            return
        value = str(value)
        evidence = f"{path} = {value}" if path else value
        lines.append(Line(text=evidence, bbox=None, conf=100.0))
        result.fields.append(
            FieldEvidence(
                name=name,
                value=value,
                normalized=normalized if normalized is not None else value,
                page=1,
                bbox=None,
                evidence_text=evidence,
                method=Method.XML,
                confidence=100.0,
            )
        )

    number, p, _ = _first_text(doc, "ID")
    add("numero_documento", number, p)
    v, p, _ = _first_text(doc, "IssueDate")
    add("fecha_emision", v, p)
    v, p, _ = _first_text(doc, "DueDate", "PaymentMeans/PaymentDueDate")
    add("fecha_vencimiento", v, p)
    v, p, el = _first_text(doc, "UUID")
    if v:
        scheme = (el.get("schemeName") or "").upper() if el is not None else ""
        add("cude" if "CUDE" in scheme else "cufe", v, p)
        if "CUDE" in scheme:
            add("cufe", v, p)  # se expone también como CUFE/CUDE para la comparación
    v, p, _ = _first_text(doc, "DocumentCurrencyCode")
    add("moneda", v, p)

    name, np, nit, nitp, dv = _party(doc, "AccountingSupplierParty")
    add("emisor_nombre", name, np)
    if nit:
        parts = parse_nit(f"{nit}-{dv}" if dv else nit)
        add("emisor_nit", nit, nitp, parts.base if parts else None)
        if dv:
            add("emisor_dv", dv, f"{nitp}/@schemeID")
    name, np, nit, nitp, dv = _party(doc, "AccountingCustomerParty")
    add("receptor_nombre", name, np)
    if nit:
        parts = parse_nit(f"{nit}-{dv}" if dv else nit)
        add("receptor_nit", nit, nitp, parts.base if parts else None)
        if dv:
            add("receptor_dv", dv, f"{nitp}/@schemeID")

    for fname, path in (
        ("subtotal", "LegalMonetaryTotal/LineExtensionAmount"),
        ("base_gravable", "LegalMonetaryTotal/TaxExclusiveAmount"),
        ("descuentos", "LegalMonetaryTotal/AllowanceTotalAmount"),
        ("valor_total", "LegalMonetaryTotal/PayableAmount"),
    ):
        v, p, _ = _first_text(doc, path)
        if v is None and fname == "valor_total":
            v, p, _ = _first_text(doc, "RequestedMonetaryTotal/PayableAmount", "LegalMonetaryTotal/TaxInclusiveAmount")
        amt = parse_amount(v) if v else None
        add(fname, v, p, str(amt) if amt is not None else None)

    # IVA: TaxTotal cuyo TaxScheme/ID = 01
    iva = Decimal("0")
    found_iva = False
    for tt in _xp(doc, "TaxTotal"):
        scheme_ids = [e.text for e in _xp(tt, "TaxSubtotal/TaxCategory/TaxScheme/ID")]
        if "01" in scheme_ids or not scheme_ids:
            amt_el = _xp(tt, "TaxAmount")
            if amt_el and amt_el[0].text:
                iva += parse_amount(amt_el[0].text) or 0
                found_iva = True
    if found_iva:
        add("iva", str(iva), "TaxTotal[TaxScheme/ID=01]/TaxAmount", str(iva))
    ret = _sum_amounts(doc, "WithholdingTaxTotal/TaxAmount")
    if ret is not None:
        add("retenciones", str(ret), "WithholdingTaxTotal/TaxAmount", str(ret))

    v, p, _ = _first_text(doc, "OrderReference/ID")
    add("orden_compra", v, p)
    v, p, _ = _first_text(doc, "ContractDocumentReference/ID")
    add("numero_contrato", v, p)
    v, p, _ = _first_text(doc, "PaymentMeans/ID")
    if v:
        add("forma_pago", PAYMENT_MEANS.get(v, v), p)
    v, p, _ = _first_text(doc, "Note", "InvoiceLine/Item/Description", "CreditNoteLine/Item/Description")
    add("concepto", v, p)
    v, p, _ = _first_text(doc, "BillingReference/InvoiceDocumentReference/ID")
    add("documento_referencia", v, p)

    # Ítems como tabla (valores y cantidades contenidos en tablas)
    rows = []
    for line_el in _xp(doc, "InvoiceLine") + _xp(doc, "CreditNoteLine") + _xp(doc, "DebitNoteLine"):
        desc, _, _ = _first_text(line_el, "Item/Description")
        qty, _, _ = _first_text(line_el, "InvoicedQuantity", "CreditedQuantity", "DebitedQuantity")
        price, _, _ = _first_text(line_el, "Price/PriceAmount")
        total, _, _ = _first_text(line_el, "LineExtensionAmount")
        rows.append([desc or "", qty or "", price or "", total or ""])
    if rows:
        result.tables.append(
            {"page": 1, "header": ["Descripción", "Cantidad", "Valor unitario", "Valor total"], "rows": rows, "method": Method.XML}
        )

    type_code, _, _ = _first_text(doc, "InvoiceTypeCode")
    if kind == "CreditNote":
        result.doc_type = "NOTA_CREDITO"
    elif kind == "DebitNote":
        result.doc_type = "NOTA_DEBITO"
    elif kind == "Invoice" and type_code == "05":
        result.doc_type = "DOCUMENTO_EQUIVALENTE"
    elif kind == "Invoice":
        result.doc_type = "FACTURA_ELECTRONICA"
    else:
        result.doc_type = "OTRO"
    result.doc_type_confidence = 1.0
    result.classification_reason = f"Elemento raíz XML <{kind}>" + (f", InvoiceTypeCode={type_code}" if type_code else "")
    add("tipo_documento", result.doc_type, "local-name(/*)")

    result.pages.append(Page(number=1, method=Method.XML, lines=lines, ocr_confidence=None))
    return result
