"""Pruebas de extracción: XML DIAN, PDF con texto, OCR, DOCX y clasificación."""
import pytest

from app.services.extraction.classifier import classify, normalize_type
from app.services.extraction.pipeline import extract
from app.services.extraction.xml_extractor import InvalidXML, extract_xml
from demo.generate_demo import A, B, CLIENT, cufe_for, invoice, render_image, render_pdf, to_png, to_scanned_pdf, ubl_invoice

from datetime import date

from .conftest import requires_ocr


def fields(result):
    return {f.name: f for f in result.fields}


@pytest.mark.parametrize("attached", [False, True])
def test_xml_deterministic_extraction(attached):
    cufe = cufe_for("FE-77")
    r = extract_xml(ubl_invoice("FE-77", B, date(2026, 3, 5), 1_190_000, cufe, attached=attached))
    f = fields(r)
    assert r.doc_type == "FACTURA_ELECTRONICA" and r.method == "XML"
    assert f["numero_documento"].value == "FE-77"
    assert f["fecha_emision"].value == "2026-03-05"
    assert f["valor_total"].normalized == "1190000.00"
    assert f["iva"].normalized == "190000.00"
    assert f["emisor_nit"].normalized == B.base
    assert f["emisor_dv"].value == B.nit[-1]
    assert f["cufe"].value == cufe
    assert f["orden_compra"].value == "OC-5500"
    # Evidencia por campo: método, confianza, página y ruta XML
    assert f["valor_total"].method == "XML" and f["valor_total"].confidence == 100.0
    assert f["valor_total"].page == 1 and "PayableAmount" in f["valor_total"].evidence_text
    assert r.tables and r.tables[0]["rows"]


def test_xml_rejects_entities_and_garbage():
    xxe = b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><Invoice><ID>&e;</ID></Invoice>'
    r = extract_xml(xxe)  # el analizador no resuelve entidades externas
    assert "root" not in (fields(r).get("numero_documento").value if fields(r).get("numero_documento") else "")
    with pytest.raises(InvalidXML):
        extract_xml(b"<Invoice><ID>1</ID>")


def test_pdf_text_extraction_with_coordinates():
    pdf = render_pdf(invoice("FV-9", A, date(2026, 3, 2), 1_000_000))
    r = extract(pdf, "pdf", CLIENT[1])
    f = fields(r)
    assert r.method == "PDF_TEXTO"
    assert f["numero_documento"].value == "FV-9"
    assert f["valor_total"].normalized == "1000000.00"
    assert f["iva"].normalized == "159664.00"
    assert f["emisor_nit"].normalized == A.base
    bbox = f["valor_total"].bbox
    assert bbox and all(0 <= v <= 1 for v in bbox) and bbox[0] < bbox[2] and bbox[1] < bbox[3]
    assert f["valor_total"].page == 1 and "TOTAL A PAGAR" in f["valor_total"].evidence_text
    assert r.doc_type == "FACTURA_COMPRA"


@requires_ocr
def test_ocr_on_image_and_scanned_pdf():
    lines = invoice("OCR-55", A, date(2026, 3, 18), 2_500_000, issuer_label="PROVEEDOR ANDINO S A S")
    for data, kind in ((to_png(render_image(lines, seed=3)), "png"), (to_scanned_pdf(render_image(lines, seed=4)), "pdf")):
        r = extract(data, kind, CLIENT[1])
        f = fields(r)
        assert r.method == "OCR"
        assert r.ocr_confidence and r.ocr_confidence > 70
        assert f["numero_documento"].value == "OCR-55"
        assert f["valor_total"].normalized == "2500000.00"
        assert f["valor_total"].method == "OCR" and f["valor_total"].confidence > 60
        assert f["emisor_nit"].normalized == A.base


@requires_ocr
def test_blank_image_is_unreadable():
    from PIL import Image

    from app.services.extraction.base import UnreadableDocument

    with pytest.raises(UnreadableDocument):
        extract(to_png(Image.new("L", (800, 1000), 255)), "png")


def test_docx_contract():
    from demo.generate_demo import J, contract_docx

    r = extract(contract_docx("CT-1", J, 4_200_000), "docx", CLIENT[1])
    f = fields(r)
    assert r.doc_type == "CONTRATO"
    assert f["numero_contrato"].value == "CT-1"
    assert f["valor_total"].normalized == "4200000.00"
    assert "Montaje" in f["objeto_contractual"].value
    assert "seis" in f["vigencia_contrato"].value
    assert "firmantes" in f
    assert r.tables


@pytest.mark.parametrize("text,code", [
    ("RECIBO DE CAJA No. 12\nRecibimos de: X", "RECIBO_CAJA"),
    ("COMPROBANTE DE EGRESO No. 5\nPagado a: Y", "COMPROBANTE_EGRESO"),
    ("ORDEN DE COMPRA No. OC-1", "ORDEN_COMPRA"),
    ("OTROSÍ No. 2 AL CONTRATO DE SUMINISTRO", "OTROSI"),
    ("CERTIFICACIÓN\nEl suscrito revisor fiscal certifica", "CERTIFICACION"),
    ("CUENTA DE COBRO No. 7\nDebe a", "CUENTA_COBRO"),
    ("EXTRACTO BANCARIO\nSaldo anterior\nSaldo final", "EXTRACTO_BANCARIO"),
    ("NOTA CRÉDITO No. NC-3", "NOTA_CREDITO"),
    ("NOTA DÉBITO No. ND-3", "NOTA_DEBITO"),
    ("DOCUMENTO SOPORTE EN ADQUISICIONES No. DS-1", "DOCUMENTO_EQUIVALENTE"),
    ("COMPROBANTE DE PAGO\nTransferencia exitosa", "SOPORTE_PAGO"),
    ("COMPROBANTE CONTABLE No. 45\nDébitos Créditos", "COMPROBANTE_CONTABLE"),
    ("FACTURA ELECTRÓNICA DE VENTA\nCUFE: abc", "FACTURA_ELECTRONICA"),
    ("Texto sin relación", "OTRO"),
])
def test_classifier(text, code):
    assert classify(text, None, None)[0] == code


def test_invoice_sale_vs_purchase_by_client_nit():
    assert classify("FACTURA DE VENTA No. 1", "900100200", "900.100.200-0")[0] == "FACTURA_VENTA"
    assert classify("FACTURA DE VENTA No. 1", "800234567", "900.100.200-0")[0] == "FACTURA_COMPRA"


def test_normalize_type_aliases():
    assert normalize_type("RC") == "RECIBO_CAJA"
    assert normalize_type("ce") == "COMPROBANTE_EGRESO"
    assert normalize_type("Factura electrónica") == "FACTURA_ELECTRONICA"
    assert normalize_type("Orden de compra") == "ORDEN_COMPRA"
