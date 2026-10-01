# -*- coding: utf-8 -*-
"""Construye el papel de trabajo de auditoría de Obligaciones Financieras (cuenta 21 / detalle 62)
de GUAICARAMO S.A.S. con corte 31-jul-2026.

Fuentes (carpeta UPLOADS):
  - Obl_Financieros.xlsx  -> Sheet1: movimiento auxiliar cuenta 21 (ene-jul 2026); Hoja1: balance de prueba (29-sep-2026)
  - 62._Obligaciones_Financieras.xlsx -> detalle de obligaciones del cliente a jul-26
  - Balance_2026_GUUU.xlsx -> balance de prueba (24-ago-2026)
  - ilovepdf_merged_1.pdf -> extractos / certificados bancarios (datos transcritos en CERTS)
"""
import datetime as dt
import re
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

UP = sys.argv[1] if len(sys.argv) > 1 else "/root/.claude/uploads/9b3ddf83-4eba-59aa-a915-ef35201f01de"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Cruce_Obligaciones_Financieras_62_Jul2026.xlsx"
F_MOV = f"{UP}/fef0ced2-Obl_Financieros.xlsx"
F_DET = f"{UP}/2b0f39e2-62._Obligaciones_Financieras.xlsx"
F_BAL = f"{UP}/a3876b2a-Balance_2026_GUUU.xlsx"

CORTE = dt.date(2026, 7, 31)

# ----------------------------------------------------------------------------- estilos
FN = "Arial Narrow"
f_base = Font(name=FN, size=8)
f_in = Font(name=FN, size=8)
f_link = Font(name=FN, size=8)
f_bold = Font(name=FN, size=8, bold=True)
f_hdr = Font(name=FN, size=8, bold=True, color="FFFFFF")
f_title = Font(name=FN, size=10, bold=True, color="002060")
f_sub = Font(name=FN, size=8, italic=True, color="595959")
f_sec = Font(name=FN, size=8, bold=True, color="0070C0")
fill_hdr = PatternFill("solid", fgColor="002060")
fill_sec = PatternFill(fill_type=None)
fill_tot = PatternFill("solid", fgColor="DDEBF7")
YELLOW = "FFFF99"
fill_key = PatternFill("solid", fgColor=YELLOW)
fill_bad = PatternFill("solid", fgColor="FFC7CE")
thin = Side(style="thin", color="808080")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
NUM = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'
NUM0 = NUM
PCT = '0.00%;-0.00%;"-"'
PB = '#,##0.0" pb";-#,##0.0" pb";"-"'
DATE = "yyyy-mm-dd"

wb = openpyxl.Workbook()
wb.remove(wb.active)


def ws_new(name, title, subtitle, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = True
    ws["A1"] = title
    ws["A1"].font = f_title
    ws["A2"] = subtitle
    ws["A2"].font = f_sub
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def put(ws, r, c, v, font=None, fmt=None, fill=None, bold=False, wrap=False, border=True, align=None):
    cell = ws.cell(row=r, column=c, value=v)
    if font is None:
        font = f_base
        if isinstance(v, str) and v.startswith("="):
            font = f_link if "!" in v else f_base
    if bold:
        font = Font(name=font.name, size=font.size, bold=True, color=font.color)
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if border:
        cell.border = box
    if wrap or align:
        cell.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    return cell


def header(ws, r, labels, c0=1, height=30):
    for i, lab in enumerate(labels):
        cell = put(ws, r, c0 + i, lab, font=f_hdr, fill=fill_hdr, wrap=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = height


def section(ws, r, text, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=r, column=c).fill = fill_sec
    ws.cell(row=r, column=1, value=text).font = f_sec


def note(ws, r, c, text):
    ws.cell(row=r, column=c).comment = Comment(text, "Auditoría")


# ============================================================================ 1. MAYOR (movimiento)
src = openpyxl.load_workbook(F_MOV, data_only=True)
mov_rows = []
gran_total = None
for i, r in enumerate(src["Sheet1"].iter_rows(values_only=True), start=1):
    if i == 1:
        gran_total = (r[5], r[6], r[7])
        continue
    if r[1] is None:
        continue
    mov_rows.append((i,) + tuple(r))

ID_RULES = [  # (regex sobre descripción, ID)
    (r"1159383796|1155297469|115297469|CRED 7469", "BOG-1159383796"),
    (r"830624|CRED 0624", "DAV-830624"),
    (r"607691|CRED 7691", "DAV-607691"),
    (r"569032|CRED 9032", "DAV-569032"),
    (r"748906|CRED 8906|CREDI 8906", "DAV-8906"),
    (r"1019441|3000101441", "DAV-L1019441"),
    (r"1019448", "DAV-L1019448"),
    (r"1021275|1275-3", "DAV-L1021275"),
    (r"453436746", "BOG-453436746"),
    (r"856670629|CRED 0629", "BOG-856670629"),
    (r"858177259|CRED 7259", "BOG-858177259"),
    (r"556449224|LEAS 9224", "BOG-L556449224"),
    (r"556449180|LEAS 9180", "BOG-L556449180"),
    (r"556449117|LEAS 9117", "BOG-L556449117"),
    (r"556449064|LEAS 9064", "BOG-L556449064"),
    (r"556449288|LEAS 9288", "BOG-L556449288"),
    (r"557285005|55728505|LEAS 5005", "BOG-L557285005"),
    (r"631309940|63130940|CRED 9940", "POP-631309940"),
    (r"0141|CRED 4476", "BBVA-0141"),
    (r"1260105867|12660105867|CRED 5867", "BCL-1260105867"),
    (r"362095|LEAS 2095", "BCL-L362095"),
    (r"336817", "BCL-L336817"),
    (r"331330", "BCL-L331330"),
    (r"386735", "BCL-L386735"),
]


def assign_id(acct, terc, desc):
    if acct == "21050301":
        return "TC"
    if "VALLEY" in (terc or ""):
        return "N/A-COMISION"
    if "CRUCE DE DOCUMENTOS" in desc:
        return "N/A-CRUCE"
    if "DAVIVIENDA" in (terc or "") and "CREDI 8906" in desc:
        return "DAV-8906"
    for pat, ident in ID_RULES:
        if re.search(pat, desc):
            return ident
    if acct == "21300102" and "BANCOLOMBIA" in (terc or ""):
        return "BCL-L362095"
    return "SIN-ID"


CLASE = {
    "21050100": "Capital crédito CP",
    "21050201": "Capital crédito LP",
    "21050301": "Tarjeta de crédito",
    "21202040": "Capital leasing CP",
    "21202140": "Capital leasing LP",
    "21300101": "Interés OF por pagar",
    "21300102": "Interés leasing por pagar",
}

ws = ws_new("Mayor_Mov", "Movimiento auxiliar cuenta 21 – Obligaciones financieras (ene–jul 2026)",
            "Fuente: Obl_Financieros.xlsx / Sheet1 (Siesa). Columnas N–P asignadas por auditoría (ID de obligación, clase y tipo de movimiento).",
            [7, 10, 24, 30, 16, 16, 16, 11, 11, 13, 22, 60, 13, 17, 21, 20])
MH = 4
header(ws, MH, ["Fila origen", "Cuenta", "Nombre cuenta", "Tercero", "Débito", "Crédito", "Neto (D-C)", "Fecha",
                "NIT", "Documento", "Tipo documento", "Descripción", "Usuario", "ID obligación (auditoría)",
                "Clase", "Tipo movimiento"])
r = MH + 1
for row in mov_rows:
    (i, _a, acct, nom, _cc, terc, deb, cre, _net, fecha, nit, _t, doc, tdoc, desc, _x, user, *_rest) = row
    acct = str(acct).strip()
    desc = (desc or "").strip()
    ident = assign_id(acct, terc, desc)
    if ident.startswith("N/A") or ident == "TC":
        tipo = "Reclasificación / no obligación" if ident != "TC" else "Tarjeta de crédito"
    elif acct.startswith("2130"):
        tipo = "Causación" if (cre or 0) > 0 else ("Ajuste redondeo" if (deb or 0) < 1000 else "Reversión por pago")
    else:
        if (deb or 0) > 0 and (cre or 0) > 0:
            tipo = "Renovación / prórroga"
        elif (cre or 0) > 0:
            tipo = "Desembolso / nueva obligación"
        else:
            tipo = "Pago de capital"
    vals = [i, acct, nom, terc, deb or 0, cre or 0, f"=E{r}-F{r}", fecha, nit, (doc or "").strip(), tdoc, desc, user,
            ident, CLASE.get(acct, "Otra"), tipo]
    for c, v in enumerate(vals, start=1):
        fnt = f_in if c in (1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13) else (f_bold if c == 14 else f_base)
        fmt = NUM if c in (5, 6, 7) else (DATE if c == 8 else None)
        put(ws, r, c, v, font=fnt, fmt=fmt)
    r += 1
M_FIRST, M_LAST = MH + 1, r - 1
put(ws, r, 4, "TOTAL MOVIMIENTO", bold=True, fill=fill_tot)
for c in (5, 6, 7):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{M_FIRST}:{L}{M_LAST})", fmt=NUM, bold=True, fill=fill_tot)
M_TOT = r
ws.freeze_panes = ws.cell(row=MH + 1, column=5)
ws.auto_filter.ref = f"A{MH}:P{M_LAST}"


def MR(col):
    return f"Mayor_Mov!${col}${M_FIRST}:${col}${M_LAST}"


R_CTA, R_DEB, R_CRE, R_FEC, R_DOC, R_ID = MR("B"), MR("E"), MR("F"), MR("H"), MR("J"), MR("N")

# ============================================================================ 2. BALANCES
def load_bal(f, off):
    wsb = openpyxl.load_workbook(f, data_only=True)["Hoja1"]
    d = {}
    for rr in wsb.iter_rows(values_only=True):
        if off:
            a, n, v = rr[1], rr[2], (rr[3], rr[4], rr[5], rr[6])
        else:
            a, n, v = rr[0], rr[1], (rr[7], rr[9], rr[11], rr[13])
        a = str(a or "").strip()
        if a and a[0].isdigit():
            d[a] = (str(n).strip(), [x if isinstance(x, (int, float)) else 0 for x in v])
    return d


BA = load_bal(F_MOV, 0)  # balance 29-sep-2026 (archivo Obl_Financieros / Hoja1)
BB = load_bal(F_BAL, 1)  # balance 24-ago-2026 (Balance_2026_GUUU)
n_diff_all = sum(1 for k in BA if k in BB and any(abs(x - y) > 0.005 for x, y in zip(BA[k][1], BB[k][1])))
BAL_ACCTS = ["1", "2", "3", "4", "5", "6", "7", "21", "2105", "210501", "21050100", "210502", "21050201", "21050301",
             "2120", "212020", "21202040", "212021", "21202140", "2130", "213001", "21300101", "21300102",
             "5305", "530520", "53052001", "53052002", "53052010", "53052020", "53052030", "53052040"]

ws = ws_new("Balance_21", "Balances de prueba – cuentas 21, 5305 y totales de clase",
            "Fuentes: Obl_Financieros.xlsx/Hoja1 (emitido 29-sep-2026) y Balance_2026_GUUU.xlsx (emitido 24-ago-2026). Libro NIIF, cifras en pesos; saldos crédito con signo negativo.",
            [11, 42, 18, 18, 18, 18, 18, 18, 18, 18, 14, 14, 14, 14])
BH = 5
put(ws, 4, 3, "Balance 29-sep-2026 (Obl_Financieros/Hoja1)", bold=True, border=False)
put(ws, 4, 7, "Balance 24-ago-2026 (Balance_2026_GUUU)", bold=True, border=False)
put(ws, 4, 11, "Diferencias (29-sep menos 24-ago)", bold=True, border=False)
header(ws, BH, ["Cuenta", "Descripción", "Saldo inicial 2026/01", "Débitos", "Créditos", "Saldo final 2026/07",
                "Saldo inicial 2026/01", "Débitos", "Créditos", "Saldo final 2026/07", "Dif. SI", "Dif. Déb.",
                "Dif. Créd.", "Dif. SF"])
BAL_ROW = {}
r = BH + 1
for a in BAL_ACCTS:
    put(ws, r, 1, a, font=f_in)
    put(ws, r, 2, BA[a][0], font=f_in)
    for j in range(4):
        put(ws, r, 3 + j, BA[a][1][j], font=f_in, fmt=NUM)
        put(ws, r, 7 + j, BB[a][1][j], font=f_in, fmt=NUM)
        put(ws, r, 11 + j, f"={get_column_letter(3 + j)}{r}-{get_column_letter(7 + j)}{r}", fmt=NUM)
    BAL_ROW[a] = r
    r += 1
B_FIRST, B_LAST = BH + 1, r - 1
r += 1
put(ws, r, 1, "Comparación integral de las dos versiones del balance (todas las cuentas):", bold=True, border=False)
r += 1
put(ws, r, 1, "Cuentas comparadas", border=True)
put(ws, r, 3, len(BA), font=f_in, fmt=NUM0)
r += 1
put(ws, r, 1, "Cuentas con alguna diferencia (SI, Déb., Créd. o SF)")
put(ws, r, 3, n_diff_all, font=f_in, fmt=NUM0)
note(ws, r, 3, "Comparación cuenta a cuenta de las 1.268 cuentas de ambos archivos realizada por auditoría (script). Resultado: 0 diferencias.")
BAL_NDIFF = f"Balance_21!$C${r}"


def BAL(acct, col):  # col: 'SI','D','C','SF'
    c = {"SI": "C", "D": "D", "C": "E", "SF": "F"}[col]
    return f"Balance_21!${c}${BAL_ROW[acct]}"


# ============================================================================ 3. CRUCE MOVIMIENTO VS BALANCE
ws = ws_new("Cruce_Mov_Balance", "Cruce movimiento auxiliar (cuenta 21) vs balance de prueba – Integridad",
            "SI balance + Débitos movimiento − Créditos movimiento = SF calculado; se compara con SF del balance. Signo: crédito negativo.",
            [11, 34, 18, 18, 18, 18, 18, 16, 18, 14, 18, 14, 12])
header(ws, 4, ["Cuenta", "Descripción", "Saldo inicial balance", "Débitos movimiento", "Créditos movimiento",
               "SF calculado", "SF balance", "Diferencia SF", "Débitos balance", "Dif. débitos", "Créditos balance",
               "Dif. créditos", "Estado"])
LEAF = ["21050100", "21050201", "21050301", "21202040", "21202140", "21300101", "21300102"]
r = 5
for a in LEAF:
    put(ws, r, 1, a, font=f_in)
    put(ws, r, 2, f"=INDEX(Balance_21!$B${B_FIRST}:$B${B_LAST},MATCH(A{r},Balance_21!$A${B_FIRST}:$A${B_LAST},0))")
    put(ws, r, 3, "=" + BAL(a, "SI"), fmt=NUM)
    put(ws, r, 4, f"=SUMIFS({R_DEB},{R_CTA},A{r})", fmt=NUM)
    put(ws, r, 5, f"=SUMIFS({R_CRE},{R_CTA},A{r})", fmt=NUM)
    put(ws, r, 6, f"=C{r}+D{r}-E{r}", fmt=NUM)
    put(ws, r, 7, "=" + BAL(a, "SF"), fmt=NUM)
    put(ws, r, 8, f"=F{r}-G{r}", fmt=NUM)
    put(ws, r, 9, "=" + BAL(a, "D"), fmt=NUM)
    put(ws, r, 10, f"=D{r}-I{r}", fmt=NUM)
    put(ws, r, 11, "=" + BAL(a, "C"), fmt=NUM)
    put(ws, r, 12, f"=E{r}-K{r}", fmt=NUM)
    put(ws, r, 13, f'=IF(MAX(ABS(H{r}),ABS(J{r}),ABS(L{r}))<=Resumen!$C$6,"OK","DIFERENCIA")')
    r += 1
put(ws, r, 2, "TOTAL CUENTA 21", bold=True, fill=fill_tot)
for c in range(3, 13):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}5:{L}{r - 1})", fmt=NUM, bold=True, fill=fill_tot)
put(ws, r, 13, f'=IF(MAX(ABS(H{r}),ABS(J{r}),ABS(L{r}))<=Resumen!$C$6,"OK","DIFERENCIA")', bold=True, fill=fill_tot)
CMB_TOT = r
r += 1
put(ws, r, 2, "Total cuenta 21 según balance (mayor)")
put(ws, r, 3, "=" + BAL("21", "SI"), fmt=NUM)
put(ws, r, 9, "=" + BAL("21", "D"), fmt=NUM)
put(ws, r, 11, "=" + BAL("21", "C"), fmt=NUM)
put(ws, r, 7, "=" + BAL("21", "SF"), fmt=NUM)
r += 1
put(ws, r, 2, "Diferencia suma de auxiliares vs cuenta 21")
for c in (3, 7, 9, 11):
    L = get_column_letter(c)
    put(ws, r, c, f"={L}{CMB_TOT}-{L}{r - 1}", fmt=NUM)
r += 2

section(ws, r, "B. Integridad del archivo de movimiento ('Gran total' del reporte vs suma de líneas)", 13)
r += 1
header(ws, r, ["", "Concepto", "Gran total reporte", "Suma líneas", "Diferencia"])
r += 1
for lab, val, col in (("Débitos", gran_total[0], "E"), ("Créditos", gran_total[1], "F"), ("Neto", gran_total[2], "G")):
    put(ws, r, 2, lab)
    put(ws, r, 3, val, font=f_in, fmt=NUM)
    put(ws, r, 4, f"=Mayor_Mov!{col}{M_TOT}", fmt=NUM)
    put(ws, r, 5, f"=C{r}-D{r}", fmt=NUM)
    r += 1
put(ws, r, 2, "Nº de líneas del movimiento")
put(ws, r, 4, f"=COUNTA({R_CTA})", fmt=NUM0)
r += 2

section(ws, r, "C. Consistencia jerárquica del balance y relación pasivo–gasto", 13)
r += 1
header(ws, r, ["", "Prueba", "Valor A", "Valor B", "Diferencia", "", "Comentario"])
r += 1
tests = [
    ("SF 21 = 2105 + 2120 + 2130", "=" + BAL("21", "SF"), f"={BAL('2105','SF')}+{BAL('2120','SF')}+{BAL('2130','SF')}", ""),
    ("SF 2105 = 210501 + 210502", "=" + BAL("2105", "SF"), f"={BAL('210501','SF')}+{BAL('210502','SF')}", ""),
    ("SF 210502 = 21050201 + 21050301 (tarjeta de crédito agrupada en LP)", "=" + BAL("210502", "SF"),
     f"={BAL('21050201','SF')}+{BAL('21050301','SF')}", "Tarjeta de crédito (pasivo corriente) agrupada dentro de 'Obligaciones financieras a largo plazo' – ver hallazgo de clasificación."),
    ("SF 2120 = 21202040 + 21202140", "=" + BAL("2120", "SF"), f"={BAL('21202040','SF')}+{BAL('21202140','SF')}", ""),
    ("SF 2130 = 21300101 + 21300102", "=" + BAL("2130", "SF"), f"={BAL('21300101','SF')}+{BAL('21300102','SF')}", ""),
    ("Ecuación: Activo + Pasivo + Patrimonio + Resultado (4+5+6+7) = 0", f"={BAL('1','SF')}+{BAL('2','SF')}+{BAL('3','SF')}",
     f"=-({BAL('4','SF')}+{BAL('5','SF')}+{BAL('6','SF')}+{BAL('7','SF')})", "Resultado del periodo ene–jul incluido vía clases 4 a 7."),
    ("Intereses OF por pagar (21300101) = gasto causado neto (53052030)", "=-" + BAL("21300101", "SF"), "=" + BAL("53052030", "SF"),
     "El cliente reversa la causación al pagar y lleva el pago a 53052020; el saldo de 53052030 debe igualar el pasivo."),
    ("Intereses leasing por pagar (21300102) = gasto causado neto (53052040)", "=-" + BAL("21300102", "SF"), "=" + BAL("53052040", "SF"), ""),
    ("SI 21300101 = créditos 53052020 (reclasificación apertura)", "=-" + BAL("21300101", "SI"), "=" + BAL("53052020", "C"),
     "Asiento de apertura: Db 53052030 / Cr 53052020 por el saldo inicial causado."),
    ("SI 21300102 = créditos 53052010 (reclasificación apertura)", "=-" + BAL("21300102", "SI"), "=" + BAL("53052010", "C"), ""),
]
for lab, a_, b_, com in tests:
    put(ws, r, 2, lab, wrap=True)
    put(ws, r, 3, a_, fmt=NUM)
    put(ws, r, 4, b_, fmt=NUM)
    put(ws, r, 5, f"=C{r}-D{r}", fmt=NUM)
    put(ws, r, 7, com, wrap=True)
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=13)
    ws.row_dimensions[r].height = 26
    r += 1
CMB_ECUACION = r - 5  # fila de la ecuación patrimonial
r += 1
section(ws, r, "D. Comparación de versiones del balance", 13)
r += 1
put(ws, r, 2, "Cuentas con diferencias entre balance 24-ago y 29-sep")
put(ws, r, 3, "=" + BAL_NDIFF, fmt=NUM0)
put(ws, r, 4, f'=IF(C{r}=0,"Sin cambios posteriores al cierre de jul-26","Revisar ajustes posteriores")')
ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
ws.freeze_panes = "C5"

# ============================================================================ 4. DETALLE 62
d_ws = openpyxl.load_workbook(F_DET, data_only=True).active
drows = [rr for rr in d_ws.iter_rows(values_only=True)]
rates = {drows[i][1]: drows[i][2] for i in range(1, 7)}
DET_ID = {
    "LEASING 1019441": "DAV-L1019441", "LEASING 1019448": "DAV-L1019448", "LEASING 10212753": "DAV-L1021275",
    "457300569032": "DAV-569032", "457300830624": "DAV-830624", "457300607691": "DAV-607691",
    "453436746": "BOG-453436746", "856670629": "BOG-856670629", "858177259": "BOG-858177259",
    "1159383796": "BOG-1159383796", "LEASING 556449224": "BOG-L556449224", "LEASING 556449180": "BOG-L556449180",
    "LEASING 556449117": "BOG-L556449117", "LEASING 556449064": "BOG-L556449064", "LEASING 556449288": "BOG-L556449288",
    "LEASING 557285005": "BOG-L557285005", "631309940-2": "POP-631309940", "252059": "BBVA-0141",
    "1260105867": "BCL-1260105867", "LEASING 362095": "BCL-L362095", "LEASING 336817": "BCL-L336817",
    "LEASING 331330": "BCL-L331330",
}
det = []
subtot_cli = []
bank = None
for rr in drows[7:]:
    if rr[1] in ("BANCO DAVIVIENDA", "BANCO DE BOGOTA", "BANCO POPULAR", "BBVA", "BANCOLOMBIA S.A."):
        bank = rr[1]
        continue
    if rr[1] == "Total Obligaciones":
        total_cli = rr[3]
        continue
    key = str(rr[1]).strip() if rr[1] is not None else None
    key = key if key is None else key
    if key in DET_ID or (key and key.isdigit() and key.lstrip("0") in [k.lstrip("0") for k in DET_ID]):
        k2 = key if key in DET_ID else [k for k in DET_ID if k.lstrip("0") == key.lstrip("0")][0]
        det.append((bank, rr[1], DET_ID[k2], rr))
    elif rr[3] is not None:
        subtot_cli.append((bank, rr[3], rr[14]))

ws = ws_new("Detalle_62", "Detalle de obligaciones financieras del cliente (papel 62) a 31-jul-2026 – recálculo",
            "Fuente: 62._Obligaciones_Financieras.xlsx. Columnas A–Q: datos del cliente (azul). Columnas R en adelante: recálculo de auditoría.",
            [17, 19, 17, 9, 26, 17, 9, 17, 9, 10, 7, 11, 16, 16, 8, 11, 9, 11, 11, 10, 15, 11, 16, 16, 10])
put(ws, 4, 1, "Tasas de referencia 31-jul-26 (cliente)", bold=True, border=False)
RATE_CELL = {}
rr_ = 5
for k, v in rates.items():
    put(ws, rr_, 1, k, font=f_in)
    put(ws, rr_, 2, v, font=f_in, fmt="0.00000" if k not in ("TRM",) else NUM)
    RATE_CELL[k] = f"Detalle_62!$B${rr_}"
    rr_ += 1
DH = 13
header(ws, DH, ["Banco", "Nº obligación (cliente)", "ID auditoría", "Tipo", "Descripción", "Saldo deuda jul-26",
                "% deuda (cliente)", "Tipo tasa", "Puntos adic.", "Tasa EA cliente", "Veces año", "Tipo pago",
                "Corto plazo (cliente)", "Largo plazo (cliente)", "Cuotas rest.", "Vencimiento", "Tasa pond. cliente",
                "Índice aplicado", "Tasa EA recalculada", "Dif. EA (pb)", "CP + LP − Saldo", "Estado CP/LP",
                "CP corregido", "LP corregido", "% deuda recalc."], height=40)
r = DH + 1
D_FIRST = r
for bank, num, ident, rr in det:
    tipo = "Leasing" if ("-L" in ident or ident.startswith("N/A")) else "Crédito"
    vals = [bank, num, ident, tipo, rr[2], rr[3], rr[4], rr[5], rr[6], rr[7], rr[8], rr[9], rr[10] or 0,
            rr[11] or 0, rr[12], rr[13], rr[14]]
    for c, v in enumerate(vals, start=1):
        fmt = {6: NUM, 7: PCT, 9: PCT, 10: PCT, 13: NUM, 14: NUM, 16: DATE, 17: PCT}.get(c)
        put(ws, r, c, v, font=f_bold if c == 3 else f_in, fmt=fmt)
    put(ws, r, 18, f'=IF(LEFT(H{r},11)="IBR 6 MESES",{RATE_CELL["IBR 6 MESES"]},IF(LEFT(H{r},11)="IBR 3 MESES",'
                   f'{RATE_CELL["IBR 3 MESES"]},IF(H{r}="DTF",{RATE_CELL["DTF"]},{RATE_CELL["IBR"]})))', fmt="0.000%")
    put(ws, r, 19, f"=(1+(R{r}+I{r})/K{r})^K{r}-1", fmt="0.000%")
    put(ws, r, 20, f"=(J{r}-S{r})*10000", fmt=PB)
    put(ws, r, 21, f"=M{r}+N{r}-F{r}", fmt=NUM)
    put(ws, r, 22, f'=IF(ABS(U{r})>Resumen!$C$6,"ERROR",IF(OR(M{r}<0,N{r}<0),"ERROR","OK"))')
    put(ws, r, 23, f"=MIN(MAX(M{r},0),F{r})", fmt=NUM)
    put(ws, r, 24, f"=F{r}-W{r}", fmt=NUM)
    put(ws, r, 25, f"=F{r}/SUMIFS($F${D_FIRST}:$F${D_FIRST + len(det) - 1},$A${D_FIRST}:$A${D_FIRST + len(det) - 1},A{r})", fmt=PCT)
    r += 1
D_LAST = r - 1
put(ws, r, 5, "TOTAL OBLIGACIONES (recalculado)", bold=True, fill=fill_tot)
for c in (6, 13, 14, 21, 23, 24):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{D_FIRST}:{L}{D_LAST})", fmt=NUM, bold=True, fill=fill_tot)
D_TOT = r
r += 1
put(ws, r, 5, "Total obligaciones según cliente")
put(ws, r, 6, total_cli, font=f_in, fmt=NUM)
r += 1
put(ws, r, 5, "Diferencia (recalculado − cliente)")
put(ws, r, 6, f"=F{D_TOT}-F{r - 1}", fmt=NUM)
D_TOTDIF = r
r += 2


def DR(col):
    return f"Detalle_62!${col}${D_FIRST}:${col}${D_LAST}"


section(ws, r, "Subtotales por banco – recálculo vs cliente", 25)
r += 1
header(ws, r, ["Banco", "", "", "", "", "Subtotal recalculado", "", "", "", "", "", "", "Subtotal cliente",
               "Diferencia", "", "", "Tasa pond. recalc.", "Tasa pond. cliente", "Dif. (pb)"])
r += 1
for bank_, sub_cli, tp_cli in subtot_cli:
    put(ws, r, 1, bank_, font=f_in)
    put(ws, r, 6, f"=SUMIFS($F${D_FIRST}:$F${D_LAST},$A${D_FIRST}:$A${D_LAST},A{r})", fmt=NUM)
    put(ws, r, 13, sub_cli, font=f_in, fmt=NUM)
    put(ws, r, 14, f"=F{r}-M{r}", fmt=NUM)
    put(ws, r, 17, f"=SUMPRODUCT(($A${D_FIRST}:$A${D_LAST}=A{r})*$F${D_FIRST}:$F${D_LAST}*$J${D_FIRST}:$J${D_LAST})/F{r}", fmt="0.000%")
    put(ws, r, 18, tp_cli, font=f_in, fmt="0.000%")
    put(ws, r, 19, f"=(R{r}-Q{r})*10000", fmt=PB)
    r += 1
r += 1
put(ws, r, 1, "Tasa EA promedio ponderada total (cliente)", bold=True)
put(ws, r, 6, f"=SUMPRODUCT({DR('F')},{DR('J')})/F{D_TOT}", fmt="0.000%")
D_TPOND = f"Detalle_62!$F${r}"
r += 1
put(ws, r, 1, "Tasa EA promedio ponderada total (recalculada)", bold=True)
put(ws, r, 6, f"=SUMPRODUCT({DR('F')},{DR('S')})/F{D_TOT}", fmt="0.000%")
r += 1
put(ws, r, 1, "Nota: la tasa EA recalculada usa la convención del cliente (1 + (índice + puntos)/n)^n − 1 con n = 'Veces año'. "
              "Diferencias indican que el cliente usó una periodicidad distinta a la declarada.", font=f_sub, border=False)
ws.freeze_panes = ws.cell(row=DH + 1, column=4)

# ============================================================================ 5. CERTIFICADOS
# (pág, banco, obligación, ID, tipo extracto, fecha corte, capital s/cert, abono capital posterior hasta 31-jul,
#  tasa EA cert, vencimiento cert, observación)
CERTS = [
    (1, "Davivienda", "7000457300830624", "DAV-830624", "Extracto crédito corporativo", dt.date(2026, 3, 21),
     "=7499998630.82-127.57", 0, 0.1551, dt.date(2044, 3, 21),
     "Saldo anterior 21-sep-25 $7.499.998.630,82 − abono capital 21-mar-26 $127,57. Próximo pago 21-sep-26 sólo intereses $565.738.188,80 (IBR semestral + 3,00)."),
    (2, "Banco de Bogotá", "00856670629", "BOG-856670629", "Extracto Sustitución Finagro", dt.date(2026, 5, 27),
     1639499992.59, 182166667.00, 0.1494, None,
     "Cuota 12 (21-jun-26): capital $182.166.667 + intereses $59.390.887,23. Valor aprobado $2.186 MM, plazo 60 meses."),
    (4, "Bancolombia", "01260105867", "BCL-1260105867", "Detalle del crédito (Sucursal Virtual)", dt.date(2026, 8, 11),
     3333333334.00, 0, 0.1417, dt.date(2028, 4, 22),
     "Capital vigente al 11-ago-26; sin abonos entre 31-jul y 11-ago. Intereses corrientes causados $24.607.407; próximo pago 22-oct-26. IBR + 1,5."),
    (6, "Banco de Bogotá", "00453436746", "BOG-453436746", "Extracto Ordinaria Comercial", dt.date(2026, 6, 22),
     2999999977.39, 62500000.00, 0.1696, None,
     "Cuota 100 (17-jul-26): capital $62.500.000 + intereses $38.894.999,71. Aprobado $7.500 MM, plazo 146 meses."),
    (8, "Banco de Bogotá", "00858177259", "BOG-858177259", "Extracto Sustitución Finagro", dt.date(2026, 7, 31),
     1124999996.41, 0, 0.1665, None,
     "Corte 31-jul-26 (fecha de auditoría). Próxima cuota 25-ago-26: capital $125.000.000 + intereses $45.166.249,86."),
    (10, "Banco de Bogotá", "01155297469 → prorrogado 1159383796", "BOG-1159383796", "Extracto Sustitución Finagro",
     dt.date(2026, 6, 17), 10000000000.00, 0, 0.1239, None,
     "Cuota 2 (14-jul-26) por $10.299.617.737,55 incluye capital $10.000 MM. Contabilidad registra prórroga (NB-1016, 29-jul) al crédito 1159383796 – sin certificado del nuevo crédito."),
    (12, "Davivienda", "7000457300607691", "DAV-607691", "Extracto crédito corporativo", dt.date(2026, 1, 28),
     3959998435.90, "=272300000-272299382.05", 0.1258, dt.date(2041, 7, 28),
     "Pago 28-jul-26 de $272.300.000 vs intereses facturados $272.299.382,05: excedente $617,95 se aplica a capital."),
    (13, "Davivienda", "7000457300569032", "DAV-569032", "Extracto crédito corporativo", dt.date(2026, 1, 28),
     2999995943.80, "=206288000-206287213.20", 0.1258, dt.date(2041, 7, 28),
     "Pago 28-jul-26 de $206.288.000 vs intereses facturados $206.287.213,20: excedente $786,80 se aplica a capital."),
    ("p.14 + LEASING_1275-3_DAVIVIENDA.pdf", "Davivienda Leasing", "000030001021275-3", "DAV-L1021275", "Factura leasing", dt.date(2026, 7, 6),
     None, 0, 0.1441, None,
     "Canon 006 (23-jul-26) sólo intereses $64.041.851; 16 cuotas pactadas. El extracto NO informa saldo de capital → ver saldo implícito en Recalculo_Intereses."),
    (16, "Banco de Bogotá Leasing", "00557285005", "BOG-L557285005", "Estado de cuenta leasing", dt.date(2026, 7, 12),
     1047113003.63, 62866117.12, 0.1604, dt.date(2029, 7, 22),
     "Saldo capital tras pago 22-abr-26; canon 22-jul-26 capital $62.866.117,12 / costo financiero $39.578.631,88. Opción de compra $16.220.758."),
    (17, "Banco de Bogotá Leasing", "00556449064", "BOG-L556449064", "Estado de cuenta leasing", dt.date(2026, 7, 17),
     280274638.64, 26602533.70, 0.1425, dt.date(2028, 7, 27),
     "Saldo tras pago 27-abr-26; canon 27-jul-26 capital $26.602.533,70 / costo financiero $9.497.274,30. DTF."),
    (18, "Banco de Bogotá Leasing", "00556449117", "BOG-L556449117", "Estado de cuenta leasing", dt.date(2026, 7, 17),
     280274638.73, 26602533.71, 0.1425, dt.date(2028, 7, 27),
     "Saldo tras pago 27-abr-26; canon 27-jul-26 capital $26.602.533,71 / costo financiero $9.497.274,29. DTF."),
    (19, "Banco de Bogotá Leasing", "00556449180", "BOG-L556449180", "Estado de cuenta leasing", dt.date(2026, 7, 13),
     417650348.81, 21538319.01, 0.1422, dt.date(2030, 1, 23),
     "Saldo tras pago 23-abr-26; canon 23-jul-26 capital $21.538.319,01 / costo financiero $14.118.857,99. DTF."),
    (20, "Banco de Bogotá Leasing", "00556449224", "BOG-L556449224", "Estado de cuenta leasing", dt.date(2026, 7, 13),
     417650348.67, 21538319.02, 0.1422, dt.date(2030, 1, 23),
     "Saldo tras pago 23-abr-26; canon 23-jul-26 capital $21.538.319,02 / costo financiero $14.118.857,98. DTF."),
    (21, "Banco de Bogotá Leasing", "00556449288", "BOG-L556449288", "Estado de cuenta leasing", dt.date(2026, 7, 10),
     417736780.24, 21542838.16, 0.1422, dt.date(2030, 1, 20),
     "Saldo tras pago 20-abr-26; canon 20-jul-26 capital $21.542.838,16 / costo financiero $14.121.779,84. DTF."),
    ("p.22 + LEASING_9441-2_DAVIVIENDA.pdf", "Davivienda Leasing", "000030001019441-2", "DAV-L1019441", "Factura leasing", dt.date(2026, 6, 25),
     None, 0, 0.1603, None,
     "Canon 033 (13-jul-26): capital $12.327.298, intereses $5.062.372, seguro $1.263.646. Último pago 11-jun-26 canon 032 $1.029.562 (sólo intereses). Sin saldo de capital."),
    ("p.24 + LEASING_9448-3_DAVIVIENDA.pdf", "Davivienda Leasing", "000030001019448-3", "DAV-L1019448", "Factura leasing", dt.date(2026, 6, 18),
     None, 0, 0.1593, None,
     "Canon 031 (6-jul-26): capital $8.867.342, intereses $4.177.893. Último pago 5-jun-26 canon 030: capital $8.901.235, intereses $4.402.982. Sin saldo de capital."),
]
ws = ws_new("Certificados", "Cruce detalle 62 vs certificados / extractos bancarios (PDF adjunto)",
            "Fuente: ilovepdf_merged_1.pdf (25 págs.). Las págs. 3, 5, 7, 9, 11, 15, 23 y 25 son anexos/instructivos sin cifras. Capital banco 31-jul = capital certificado − abonos a capital entre corte y 31-jul.",
            [5, 20, 20, 24, 16, 22, 11, 17, 16, 17, 17, 16, 10, 10, 10, 11, 11, 11, 70])
CH = 4
header(ws, CH, ["#", "Pág. PDF", "Banco", "Obligación (certificado)", "ID auditoría", "Tipo documento", "Fecha corte",
                "Capital s/certificado", "Abonos capital posteriores al corte (≤31-jul)", "Capital banco 31-jul-26",
                "Saldo detalle 62", "Diferencia detalle − banco", "Tasa EA certificado", "Tasa EA detalle",
                "Dif. (pb)", "Venc. certificado", "Venc. detalle", "¿Venc. coincide?", "Observaciones"], height=48)
r = CH + 1
C_FIRST = r
CERT_ROW = {}
for n, (pg, bk, ob, ident, td, fc, cap, abono, ea, venc, obs) in enumerate(CERTS, start=1):
    put(ws, r, 1, n)
    put(ws, r, 2, pg, font=f_in, wrap=True)
    put(ws, r, 3, bk, font=f_in)
    put(ws, r, 4, ob, font=f_in)
    put(ws, r, 5, ident, font=f_bold)
    put(ws, r, 6, td, font=f_in)
    put(ws, r, 7, fc, font=f_in, fmt=DATE)
    put(ws, r, 8, cap, font=f_in, fmt=NUM)
    put(ws, r, 9, abono, font=f_in, fmt=NUM)
    put(ws, r, 10, f'=IF(H{r}="","N/D",H{r}-I{r})', fmt=NUM)
    put(ws, r, 11, f"=SUMIFS({DR('F')},{DR('C')},E{r})", fmt=NUM)
    put(ws, r, 12, f'=IF(H{r}="","N/D",K{r}-J{r})', fmt=NUM)
    put(ws, r, 13, ea, font=f_in, fmt=PCT)
    put(ws, r, 14, f"=INDEX({DR('J')},MATCH(E{r},{DR('C')},0))", fmt=PCT)
    put(ws, r, 15, f"=(N{r}-M{r})*10000", fmt=PB)
    put(ws, r, 16, venc, font=f_in, fmt=DATE)
    put(ws, r, 17, f"=INDEX({DR('P')},MATCH(E{r},{DR('C')},0))", fmt=DATE)
    put(ws, r, 18, f'=IF(P{r}="","N/D",IF(P{r}=Q{r},"Sí","NO"))')
    put(ws, r, 19, obs, font=f_in, wrap=True)
    ws.row_dimensions[r].height = 36
    CERT_ROW[ident] = r
    r += 1
C_LAST = r - 1
put(ws, r, 6, "TOTAL obligaciones con saldo certificado", bold=True, fill=fill_tot)
for c in (10, 11):
    L = get_column_letter(c)
    put(ws, r, c, f'=SUMIFS({L}{C_FIRST}:{L}{C_LAST},$H${C_FIRST}:$H${C_LAST},"<>")', fmt=NUM, bold=True, fill=fill_tot)
put(ws, r, 12, f"=SUM(L{C_FIRST}:L{C_LAST})", fmt=NUM, bold=True, fill=fill_tot)
C_TOT = r
r += 1
put(ws, r, 6, "Diferencia detalle − banco: leasings Banco de Bogotá")
put(ws, r, 12, f'=SUMIFS(L{C_FIRST}:L{C_LAST},E{C_FIRST}:E{C_LAST},"BOG-L*")', fmt=NUM, fill=fill_key)
C_BOGL = f"Certificados!$L${r}"
r += 1
put(ws, r, 6, "Diferencia detalle − banco: créditos (redondeos)")
put(ws, r, 12, f'=L{C_TOT}-L{r - 1}', fmt=NUM)
C_CRED = f"Certificados!$L${r}"
r += 2

section(ws, r, "B. Cobertura de confirmación – obligaciones del detalle 62 SIN certificado/extracto en el PDF", 19)
r += 1
header(ws, r, ["", "", "Banco", "Obligación", "ID auditoría", "Tipo", "", "", "", "", "Saldo detalle 62"])
r += 1
B_FIRST = r
cert_ids = {c[3] for c in CERTS}
for bank, num, ident, rr in det:
    if ident in cert_ids:
        continue
    put(ws, r, 3, bank, font=f_in)
    put(ws, r, 4, str(num), font=f_in)
    put(ws, r, 5, ident, font=f_bold)
    put(ws, r, 6, "Leasing" if "-L" in ident else "Crédito")
    put(ws, r, 11, f"=SUMIFS({DR('F')},{DR('C')},E{r})", fmt=NUM)
    r += 1
put(ws, r, 3, "Bancolombia", font=f_in)
put(ws, r, 4, "Leasing 386735 (no está en detalle 62)", font=f_in)
put(ws, r, 5, "BCL-L386735", font=f_bold)
put(ws, r, 6, "Leasing")
put(ws, r, 11, f'=SUMIFS({R_CRE},{R_ID},E{r})-SUMIFS({R_DEB},{R_ID},E{r})', fmt=NUM)
note(ws, r, 11, "Valor según mayor (NB-00001014, 28-jul-26), no incluido en el detalle 62.")
r += 1
put(ws, r, 6, "Total sin certificado (detalle 62)", bold=True, fill=fill_tot)
put(ws, r, 11, f"=SUM(K{B_FIRST}:K{r - 2})", fmt=NUM, bold=True, fill=fill_tot)
C_SINCERT = r
r += 1
put(ws, r, 6, "% del saldo del detalle 62 cubierto con certificado/extracto")
put(ws, r, 11, f"=1-K{C_SINCERT}/Detalle_62!$F${D_TOT}", fmt=PCT, fill=fill_key)
C_COBERT = f"Certificados!$K${r}"
r += 2

section(ws, r, "C. Cruce de pagos/intereses según certificado vs registro contable (comprobantes del mayor)", 19)
r += 1
header(ws, r, ["", "", "ID auditoría", "Concepto", "Documento contable", "Fecha", "", "Valor s/certificado",
               "Valor s/contabilidad", "Diferencia", "", "", "", "", "", "", "", "", "Comentario"])
r += 1
P_FIRST = r
PAGOS = [  # (ID, concepto, doc, campo ('D' débito / 'C' crédito / valor fijo), valor certificado, comentario)
    ("DAV-830624", "Pago 21-mar-26 total aplicado (int. 445.207.872,43 + cap. 127,57)", "NI-00004268", "D", 445208000.00,
     "El cliente reversa el total como interés; los $127,57 abonados a capital no se registran."),
    ("DAV-607691", "Pago 28-ene-26 total aplicado (int. 243.932.730,53 + cap. 269,47)", "NI-00004061", "D", 243933000.00,
     "Capital $269,47 no registrado."),
    ("DAV-569032", "Pago 28-ene-26 total aplicado (int. 184.797.372,21 + cap. 627,79)", "NI-00004063", "D", 184798000.00,
     "Capital $627,79 no registrado."),
    ("DAV-607691", "Intereses 28-ene a 28-jul-26 facturados", "NI-00004743", "D", 272299382.05,
     "Contabilidad reversa $272.300.000 (pago redondeado); $617,95 corresponden a capital."),
    ("DAV-569032", "Intereses 28-ene a 28-jul-26 facturados", "NI-00004745", "D", 206287213.20,
     "Contabilidad reversa $206.288.000; $786,80 corresponden a capital."),
    ("BOG-856670629", "Cuota 11 – capital", "CL-00000383", "D", 182166667.53, "Diferencia de redondeo del banco."),
    ("BOG-856670629", "Cuota 11 – intereses", "NI-00004260", "D", 56089116.47, ""),
    ("BOG-856670629", "Cuota 12 – capital", "CL-00000410", "D", 182166667.00, ""),
    ("BOG-856670629", "Cuota 12 – intereses", "NI-00004662", "D", 59390887.23, ""),
    ("BOG-453436746", "Cuota 99 – capital", "CL-00000409", "D", 62500000.09, ""),
    ("BOG-453436746", "Cuota 100 – capital", "CL-00000422", "D", 62500000.00,
     "Intereses cuota 100 $38.894.999,71 coinciden con la descripción del CL-422 (gasto directo 53052020)."),
    ("BOG-858177259", "Cuota mayo – capital", "CL-00000407", "D", 125000000.56, ""),
    ("BOG-858177259", "Cuota mayo – intereses", "NI-00004526", "D", 46619930.44, ""),
    ("BOG-1159383796", "Crédito 1155297469 – intereses cuota 1", "NI-00004402", "D", 254610555.55, ""),
    ("BOG-1159383796", "Crédito 1155297469 – intereses cuota 2 (14-jul)", "NI-00004736", "D", 299617500.00, ""),
    ("BOG-L557285005", "Canon abr – capital", "CL-00000391", "D", 62648860.03, ""),
    ("BOG-L557285005", "Canon abr – costo financiero", "NI-00004396", "D", 36783056.97, ""),
    ("BOG-L557285005", "Canon jul – capital", "CL-00000429", "D", 62866117.12, ""),
    ("BOG-L557285005", "Canon jul – costo financiero", "NI-00004740", "D", 39578631.88, ""),
    ("BOG-L556449064", "Canon abr – capital", "CL-00000394", "D", 26079255.17, ""),
    ("BOG-L556449064", "Canon jul – capital", "CL-00000426", "D", 26602533.70, ""),
    ("BOG-L556449064", "Canon jul – costo financiero", "NI-00004737", "D", 9497274.30, ""),
    ("BOG-L556449117", "Canon abr – capital", "CL-00000389", "D", 26079255.16, ""),
    ("BOG-L556449117", "Canon jul – capital", "CL-00000424", "D", 26602533.71, ""),
    ("BOG-L556449117", "Canon jul – costo financiero", "NI-00004734", "D", 9497274.29, ""),
    ("BOG-L556449180", "Canon abr – capital", "CL-00000392", "D", 21315470.38, ""),
    ("BOG-L556449180", "Canon jul – capital", "CL-00000428", "D", 21538319.01, ""),
    ("BOG-L556449180", "Canon jul – costo financiero", "NI-00004739", "D", 14118857.99, ""),
    ("BOG-L556449224", "Canon abr – capital", "CL-00000393", "D", 21315470.39,
     "La descripción del CL-393 indica 'CAPITAL 34.950.056' (valor total del canon) – error de glosa."),
    ("BOG-L556449224", "Canon jul – capital", "CL-00000427", "D", 21538319.02, ""),
    ("BOG-L556449224", "Canon jul – costo financiero", "NI-00004738", "D", 14118857.98, ""),
    ("BOG-L556449288", "Canon abr – capital", "CL-00000390", "D", 21319941.88, ""),
    ("BOG-L556449288", "Canon jul – capital", "CL-00000430", "D", 21542838.16, ""),
    ("BOG-L556449288", "Canon jul – costo financiero", "NI-00004741", "D", 14121779.84, ""),
    ("DAV-L1021275", "Canon 005 (23-abr) – intereses", "NI-00004400", "D", 55427870.00, ""),
    ("DAV-L1021275", "Canon 006 (23-jul) – intereses", "NI-00004744", "D", 64041851.00, ""),
    ("DAV-L1019448", "Canon 030 (5-jun) – capital", "CL-00000415", "D", 8901235.00, ""),
    ("DAV-L1019448", "Canon 031 (6-jul) – capital", "CL-00000433", "D", 8867342.00, ""),
    ("DAV-L1019441", "Canon 032 (11-jun) – capital", "CL-00000416", "D", 0.00,
     "El banco reporta el canon 032 pagado el 11-jun por $1.029.562 sólo intereses (capital 0); contabilidad registró capital $11.973.677. Verificar aplicación del pago."),
    ("DAV-L1019441", "Canon 033 (13-jul) – capital", "CL-00000434", "D", 12327298.00, ""),
]
for ident, conc, doc, campo, vcert, com in PAGOS:
    put(ws, r, 3, ident, font=f_bold)
    put(ws, r, 4, conc, font=f_in, wrap=True)
    put(ws, r, 5, doc, font=f_in)
    put(ws, r, 6, f"=INDEX({R_FEC},MATCH(E{r},{R_DOC},0))", fmt=DATE)
    put(ws, r, 8, vcert, font=f_in, fmt=NUM)
    rng = R_DEB if campo == "D" else R_CRE
    put(ws, r, 9, f"=SUMIFS({rng},{R_DOC},E{r},{R_ID},C{r})", fmt=NUM)
    put(ws, r, 10, f"=I{r}-H{r}", fmt=NUM)
    put(ws, r, 19, com, font=f_in, wrap=True)
    ws.row_dimensions[r].height = 24
    r += 1
put(ws, r, 4, "TOTAL", bold=True, fill=fill_tot)
for c in (8, 9, 10):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{P_FIRST}:{L}{r - 1})", fmt=NUM, bold=True, fill=fill_tot)
C_PAGOS_TOT = r
r += 2
section(ws, r, "D. Documentos Davivienda Leasing recibidos por separado (LEASING_1275-3 / 9441-2 / 9448-3_DAVIVIENDA.pdf)", 19)
r += 1
header(ws, r, ["", "Archivo", "Contrato", "ID auditoría", "Fecha corte", "Canon / fecha", "Capital canon", "Intereses canon",
               "Seguro", "Tasa EA cobrada", "Último pago (fecha)", "Último pago (valor)", "¿Informa saldo capital?",
               "Coincide intereses con papel", "Coincide capital con papel", "Coincide tasa con papel", "", "", "Conclusión"], height=40)
r += 1
DOCS_ROWS = []
CERT_WS = ws
DOCS_DAV = [
    ("LEASING_1275-3_DAVIVIENDA.pdf", "000030001021275-3", "DAV-L1021275", dt.date(2026, 7, 6), "006 / 2026-07-23", 0, 64041851, 0, 0.1441,
     dt.date(2026, 4, 23), 55427870),
    ("LEASING_9441-2_DAVIVIENDA.pdf", "000030001019441-2", "DAV-L1019441", dt.date(2026, 6, 25), "033 / 2026-07-13", 12327298, 5062372, 1263646, 0.1603,
     dt.date(2026, 6, 11), 1029562),
    ("LEASING_9448-3_DAVIVIENDA.pdf", "000030001019448-3", "DAV-L1019448", dt.date(2026, 6, 18), "031 / 2026-07-06", 8867342, 4177893, 1263646, 0.1593,
     dt.date(2026, 6, 5), 14567863),
]
for arch, con, ident, fc, can, capc, intc, seg, ea, fup, vup in DOCS_DAV:
    DOCS_ROWS.append((r, ident))
    put(ws, r, 2, arch, font=f_in, wrap=True)
    put(ws, r, 3, con, font=f_in)
    put(ws, r, 4, ident, font=f_bold)
    put(ws, r, 5, fc, font=f_in, fmt=DATE)
    put(ws, r, 6, can, font=f_in)
    put(ws, r, 7, capc, font=f_in, fmt=NUM)
    put(ws, r, 8, intc, font=f_in, fmt=NUM)
    put(ws, r, 9, seg, font=f_in, fmt=NUM)
    put(ws, r, 10, ea, font=f_in, fmt=PCT)
    put(ws, r, 11, fup, font=f_in, fmt=DATE)
    put(ws, r, 12, vup, font=f_in, fmt=NUM)
    put(ws, r, 13, "NO", font=f_in, bold=True, fill=fill_bad)
    put(ws, r, 16, f'=IF(J{r}=M{CERT_ROW[ident]},"Sí","NO")')
    put(ws, r, 19, "Mismo documento que el PDF consolidado. Sólo factura del canon: el saldo de capital sigue sin certificar (PEND-1)."
        + (" Reconfirma canon 032 pagado sólo con intereses $1.029.562 (PEND-7)." if ident == "DAV-L1019441" else ""),
        font=f_in, wrap=True)
    ws.row_dimensions[r].height = 36
    r += 1
ws.freeze_panes = "F5"

# -*- coding: utf-8 -*-
# Bloque insertado en build_cruce_62.py antes de "6. CRUCE DETALLE VS MAYOR".
# Construye: Mayor_Dic25 (detalle del mayor al 31-dic-25 tomado del PT 2025) y Saldos_x_Obligacion.

# ============================================================================ 5b. MAYOR AL 31-DIC-2025 (PT 2025)
F_EJ = f"{UP}/e6eb8aed-ejemplo.xlsx"
ej = openpyxl.load_workbook(F_EJ, data_only=True).active

L4 = {"0624": "DAV-830624", "7691": "DAV-607691", "8906": "DAV-8906", "9032": "DAV-569032", "5867": "BCL-1260105867",
      "7259": "BOG-858177259", "1275": "DAV-L1021275", "5005": "BOG-L557285005", "9064": "BOG-L556449064",
      "9117": "BOG-L556449117", "9180": "BOG-L556449180", "9224": "BOG-L556449224", "9288": "BOG-L556449288",
      "0629": "BOG-856670629", "9940": "POP-631309940", "4476": "BBVA-0141", "2095": "BCL-L362095"}


def id_dic25(fila, aux, terc, nota, doc):
    nota = str(nota or "")
    doc = str(doc or "")
    if aux.startswith("2105020103"):
        return "TC"
    if aux.startswith("2130"):
        m = re.search(r"(?:CRED|LEAS)[- ]?(\d{4})", nota)
        return L4.get(m.group(1), "SIN-ID") if m else "SIN-ID"
    rules = [("1260105867", "BCL-1260105867"), ("858906", "DAV-8906"), ("9600277262", "BBVA-0141"),
             ("Obf.158009", "BOG-453436746"), ("858177259", "BOG-858177259"), ("830624", "BOG-X"),
             ("856670629", "BOG-856670629"), ("POPULAR", "POP-631309940"), ("336817", "BCL-L336817"),
             ("TURBOKRAFT", "BCL-L336817"), ("331330", "BCL-L331330"), ("1019441", "DAV-L1019441"),
             ("1019448", "DAV-L1019448"), ("362095", "BCL-L362095"), ("CALDERAS JCT", "DAV-L1021275"),
             ("GECOLSA", "DAV-LGECOLSA"), ("556449064", "BOG-L556449064"), ("556449117", "BOG-L556449117"),
             ("556449180", "BOG-L556449180"), ("556449224", "BOG-L556449224"), ("556449288", "BOG-L556449288"),
             ("557285005", "BOG-L557285005")]
    for pat, ident in rules:
        if pat in nota:
            return "DAV-830624" if ident == "BOG-X" else ident
    if doc.endswith("569032"):
        return "DAV-569032"
    if doc.endswith("607691"):
        return "DAV-607691"
    if doc in ("53436746",):
        return "BOG-453436746"
    if doc in ("85817725",):
        return "BOG-858177259"
    if "RECLASIFICACION SALDO DE LARGO" in nota:  # mapeo del PT 2025 (E289 incluye I68; E290 incluye I69)
        return {68: "BOG-453436746", 69: "BOG-856670629"}.get(fila, "SIN-ID")
    return "SIN-ID"


AUX2026 = {"2105020103": "21050301"}
ws = ws_new("Mayor_Dic25", "Detalle del mayor por obligación al 31-dic-2025 (saldo inicial 2026)",
            "Fuente: ejemplo.xlsx (PT Obligaciones Financieras al 31-dic-2025, sección 2 'Detalle suministrado por Cliente'). ID asignado por auditoría según la nota y el documento cruce.",
            [8, 12, 13, 12, 26, 11, 17, 62, 12, 18, 22])
header(ws, 4, ["Fila PT 2025", "Auxiliar 2025", "Cuenta 2026", "NIT", "Razón social", "Fecha docto.", "Total COP (crédito +)",
               "Notas", "Doc. cruce", "ID obligación", "Clase"])
r = 5
MD_FIRST = r
for fila in range(68, 264):
    aux = ej.cell(fila, 4).value
    if aux is None:
        continue
    aux = str(aux)
    terc, nota, doc = ej.cell(fila, 7).value, ej.cell(fila, 10).value, ej.cell(fila, 11).value
    ident = id_dic25(fila, aux, terc, nota, doc)
    cta = AUX2026.get(aux, aux)
    vals = [fila, aux, cta, ej.cell(fila, 6).value, terc, ej.cell(fila, 8).value, ej.cell(fila, 9).value, nota, str(doc), ident,
            CLASE.get(cta, "Otra")]
    for c, v in enumerate(vals, start=1):
        put(ws, r, c, v, font=f_bold if c == 10 else f_in, fmt=NUM if c == 7 else (DATE if c == 6 else None))
    r += 1
MD_LAST = r - 1
put(ws, r, 5, "TOTAL", bold=True, fill=fill_tot)
put(ws, r, 7, f"=SUM(G{MD_FIRST}:G{MD_LAST})", fmt=NUM, bold=True, fill=fill_tot)
r += 2
section(ws, r, "Cuadre del detalle 31-dic-25 contra el saldo inicial del balance 2026", 11)
r += 1
header(ws, r, ["", "Cuenta", "", "", "Descripción", "", "Detalle 31-dic-25", "Saldo inicial balance (signo +)", "Diferencia"])
r += 1
for a in LEAF:
    put(ws, r, 2, a, font=f_in)
    put(ws, r, 5, BA[a][0])
    put(ws, r, 7, f"=SUMIFS($G${MD_FIRST}:$G${MD_LAST},$C${MD_FIRST}:$C${MD_LAST},B{r})", fmt=NUM)
    put(ws, r, 8, f"=-{BAL(a, 'SI')}", fmt=NUM)
    put(ws, r, 9, f"=G{r}-H{r}", fmt=NUM)
    r += 1
put(ws, r, 5, "Partidas sin ID de obligación", bold=True)
put(ws, r, 7, f'=COUNTIF($J${MD_FIRST}:$J${MD_LAST},"SIN-ID")', fmt=NUM0)
ws.freeze_panes = "B5"


def MD(col):
    return f"Mayor_Dic25!${col}${MD_FIRST}:${col}${MD_LAST}"


# ============================================================================ 5c. SALDOS POR OBLIGACIÓN
EXT_DIC25 = {  # 'Valor Extracto' del PT al 31-dic-2025 (ejemplo.xlsx, sección 3)
    "DAV-8906": 10000000000, "DAV-569032": 2999996571.59, "DAV-830624": 7499998630.82, "DAV-607691": 3959998705.37,
    "DAV-L1019441": 481870906, "DAV-L1019448": 393351698, "DAV-L1021275": 1876684264,
    "BOG-453436746": 3374999979, "BOG-856670629": "=2003833327.94-182166667", "BOG-858177259": 1375000000,
    "BOG-L556449224": 459771024, "BOG-L556449180": 459771024, "BOG-L556449117": 331775059, "BOG-L556449064": 331775059,
    "BOG-L556449288": 459866292, "BOG-L557285005": 1170839109, "POP-631309940": 575000000, "BBVA-0141": 4700000000,
    "BCL-1260105867": 5000000000, "BCL-L362095": 495530373, "BCL-L336817": 849599858, "BCL-L331330": 605798177,
    "BOG-1159383796": 0,
}
SX_IDS = [("BANCO DAVIVIENDA", ["DAV-8906", "DAV-569032", "DAV-830624", "DAV-607691", "DAV-L1019441", "DAV-L1019448",
                                "DAV-L1021275", "DAV-LGECOLSA"]),
          ("BANCO DE BOGOTA", ["BOG-453436746", "BOG-856670629", "BOG-858177259", "BOG-1159383796", "BOG-L556449224",
                               "BOG-L556449180", "BOG-L556449117", "BOG-L556449064", "BOG-L556449288", "BOG-L557285005"]),
          ("BANCO POPULAR", ["POP-631309940"]), ("BBVA", ["BBVA-0141"]),
          ("BANCOLOMBIA S.A.", ["BCL-1260105867", "BCL-L362095", "BCL-L336817", "BCL-L331330", "BCL-L386735"]),
          ("PARTIDAS AJENAS EN 2120 (neto cero)", ["N/A-COMISION", "N/A-CRUCE"])]
SX_DESC = {"DAV-8906": "Crédito tesorería 457300748906 (cancelado 28-ene-26)", "DAV-LGECOLSA": "Leasing 2 motogeneradores GECOLSA (no está en detalle 62)",
           "BCL-L386735": "Leasing 386735 camión VW Constellation (nuevo 28-jul-26)", "BOG-1159383796": "Crédito 1155297469 → prorrogado 1159383796",
           "N/A-COMISION": "Comisiones Valley Trading registradas y reclasificadas en 21202040", "N/A-CRUCE": "NI-4731 cruce anticipo calderas en 21202140"}
ws = ws_new("Saldos_x_Obligacion", "Saldos de capital por obligación al 31-jul-2026 – mayor vs detalle 62 vs tabla de amortización vs banco",
            "Saldo mayor = detalle 31-dic-25 (PT 2025) + créditos 2026 − débitos 2026 (Mayor_Mov). Saldo banco = certificado jul-26; si no hay, extracto 31-dic-25 + movimientos 2026 (estimado, marca [E]).",
            [17, 15, 8, 34, 17, 16, 16, 17, 17, 17, 17, 17, 22, 15, 15, 15, 15, 30])
header(ws, 4, ["ID auditoría", "Banco", "Tipo", "Descripción", "Saldo mayor 31-dic-25", "Créditos 2026 (desembolsos)",
               "Débitos 2026 (pagos)", "Saldo mayor 31-jul-26", "Saldo detalle 62", "Saldo tabla amortización 31-jul",
               "Extracto 31-dic-25 (PT 2025)", "Saldo banco 31-jul-26", "Fuente saldo banco", "Mayor − detalle 62",
               "Mayor − banco", "Detalle 62 − banco", "Tabla − banco", "Conclusión"], height=48)
r = 5
SX_FIRST = r
SX_ROW = {}
CAPC = ("21050100", "21050201", "21202040", "21202140")
for bank_, ids in SX_IDS:
    for ident in ids:
        tipo = "Leasing" if ("-L" in ident or ident.startswith("N/A")) else "Crédito"
        meta = next((d for d in det if d[2] == ident), None)
        put(ws, r, 1, ident, font=f_bold)
        put(ws, r, 2, bank_, font=f_in)
        put(ws, r, 3, tipo)
        put(ws, r, 4, SX_DESC.get(ident, f"{meta[1]} – {meta[3][2]}" if meta else ""), font=f_in)
        put(ws, r, 5, f'=SUMIFS({MD("G")},{MD("J")},A{r},{MD("C")},"<>2130*",{MD("C")},"<>21050301")', fmt=NUM)
        put(ws, r, 6, "=" + "+".join(f'SUMIFS({R_CRE},{R_ID},$A{r},{R_CTA},"{a}")' for a in CAPC), fmt=NUM)
        put(ws, r, 7, "=" + "+".join(f'SUMIFS({R_DEB},{R_ID},$A{r},{R_CTA},"{a}")' for a in CAPC), fmt=NUM)
        put(ws, r, 8, f"=E{r}+F{r}-G{r}", fmt=NUM)
        put(ws, r, 9, f"=SUMIFS({DR('F')},{DR('C')},A{r})", fmt=NUM)
        put(ws, r, 10, f'=IFERROR(INDEX(Rev_Tablas_Amort!$N:$N,MATCH(A{r},Rev_Tablas_Amort!$A:$A,0)),"Sin tabla")', fmt=NUM)
        ext = EXT_DIC25.get(ident)
        put(ws, r, 11, ext if ext is not None else "N/D", font=f_in, fmt=NUM)
        if ident in CERT_ROW and CERTS[[c[3] for c in CERTS].index(ident)][6] is not None:
            put(ws, r, 12, f"=Certificados!J{CERT_ROW[ident]}", fmt=NUM)
            put(ws, r, 13, "[C] Certificado / extracto 2026")
        else:
            put(ws, r, 12, f'=IF(ISNUMBER(K{r}),K{r}+F{r}-G{r},"N/D")', fmt=NUM)
            put(ws, r, 13, f'=IF(ISNUMBER(K{r}),"[E] Extracto dic-25 + mov. 2026","[P] Sin soporte bancario")')
        put(ws, r, 14, f"=H{r}-I{r}", fmt=NUM)
        put(ws, r, 15, f'=IF(ISNUMBER(L{r}),H{r}-L{r},"N/D")', fmt=NUM)
        put(ws, r, 16, f'=IF(ISNUMBER(L{r}),I{r}-L{r},"N/D")', fmt=NUM)
        put(ws, r, 17, f'=IF(AND(ISNUMBER(J{r}),ISNUMBER(L{r})),J{r}-L{r},"N/D")', fmt=NUM)
        put(ws, r, 18, f'=IF(NOT(ISNUMBER(L{r})),"Pendiente soporte bancario",IF(ABS(O{r})<=Resumen!$C$6,'
                       f'IF(ABS(N{r})<=Resumen!$C$6,"OK","Mayor OK – detalle 62 desactualizado"),"Diferencia mayor vs banco"))', wrap=True)
        SX_ROW[ident] = r
        r += 1
SX_LAST = r - 1


def SX(col):
    return f"Saldos_x_Obligacion!${col}${SX_FIRST}:${col}${SX_LAST}"


for tp, accts in (("Crédito", ("21050100", "21050201")), ("Leasing", ("21202040", "21202140"))):
    put(ws, r, 1, f"TOTAL {tp.upper()}", bold=True, fill=fill_tot)
    for c in range(5, 18):
        if c in (11, 13, 18):
            continue
        L = get_column_letter(c)
        put(ws, r, c, f'=SUMIFS({L}{SX_FIRST}:{L}{SX_LAST},$C${SX_FIRST}:$C${SX_LAST},"{tp}")', fmt=NUM, bold=True, fill=fill_tot)
    r += 1
    put(ws, r, 1, f"Balance {tp} (signo +)")
    put(ws, r, 5, f"=-({BAL(accts[0], 'SI')}+{BAL(accts[1], 'SI')})", fmt=NUM)
    put(ws, r, 6, f"={BAL(accts[0], 'C')}+{BAL(accts[1], 'C')}", fmt=NUM)
    put(ws, r, 7, f"={BAL(accts[0], 'D')}+{BAL(accts[1], 'D')}", fmt=NUM)
    put(ws, r, 8, f"=-({BAL(accts[0], 'SF')}+{BAL(accts[1], 'SF')})", fmt=NUM)
    r += 1
    put(ws, r, 1, "Diferencia (por obligación − balance)")
    for c in (5, 6, 7, 8):
        L = get_column_letter(c)
        put(ws, r, c, f"={L}{r - 2}-{L}{r - 1}", fmt=NUM, fill=fill_key)
    if tp == "Crédito":
        SX_CRED = (r - 2, r)
    else:
        SX_LEAS = (r - 2, r)
    r += 2
put(ws, r, 1, "Nota: [C] cotejado con certificado/extracto 2026; [E] estimado con el extracto al 31-dic-25 del PT 2025 más los movimientos de capital 2026 del mayor (supone que el banco aplicó los mismos abonos); [P] sin soporte bancario.",
    font=f_sub, border=False)
ws.freeze_panes = "E5"

# ============================================================================ 6. CRUCE DETALLE VS MAYOR
ws = ws_new("Cruce_Detalle_Mayor", "Cruce detalle 62 vs mayor (cuentas 2105 y 2120) – Integridad y exactitud",
            "Saldos del mayor tomados del balance con signo invertido (pasivo positivo). Movimientos de capital por obligación desde Mayor_Mov.",
            [18, 20, 10, 18, 18, 18, 18, 18, 60])
put(ws, 4, 1, "A. Totales", font=f_sec, border=False)
header(ws, 5, ["Concepto", "Cuentas mayor", "", "Saldo detalle 62", "Saldo mayor 31-jul", "Diferencia (mayor − detalle)", "Estado"])
put(ws, 6, 1, "Créditos (capital)")
put(ws, 6, 2, "21050100 + 21050201")
put(ws, 6, 4, f'=SUMIFS({DR("F")},{DR("D")},"Crédito")', fmt=NUM)
put(ws, 6, 5, f"=-({BAL('21050100','SF')}+{BAL('21050201','SF')})", fmt=NUM)
put(ws, 7, 1, "Leasing financiero (capital)")
put(ws, 7, 2, "21202040 + 21202140")
put(ws, 7, 4, f'=SUMIFS({DR("F")},{DR("D")},"Leasing")', fmt=NUM)
put(ws, 7, 5, f"=-({BAL('21202040','SF')}+{BAL('21202140','SF')})", fmt=NUM)
put(ws, 8, 1, "TOTAL", bold=True, fill=fill_tot)
put(ws, 8, 4, "=D6+D7", fmt=NUM, bold=True, fill=fill_tot)
put(ws, 8, 5, "=E6+E7", fmt=NUM, bold=True, fill=fill_tot)
for rr_ in (6, 7, 8):
    put(ws, rr_, 6, f"=E{rr_}-D{rr_}", fmt=NUM, fill=fill_key if rr_ == 7 else None)
    put(ws, rr_, 7, f'=IF(ABS(F{rr_})<=Resumen!$C$6,"OK","DIFERENCIA")')
put(ws, 9, 1, "Comprobación: total detalle = 'Total Obligaciones' del cliente")
put(ws, 9, 4, f"=Detalle_62!F{D_TOT}-D8", fmt=NUM)
put(ws, 10, 1, "Tarjetas de crédito (21050301) e intereses por pagar (2130) no forman parte del detalle 62.", font=f_sub, border=False)

put(ws, 12, 1, "B. Explicación de la diferencia en leasing (mayor − detalle) por obligación – ver Saldos_x_Obligacion", font=f_sec, border=False)
header(ws, 13, ["Partida", "", "", "", "", "Valor", "", "", "Soporte"])
SXN = lambda pat: f'SUMIFS({SX("N")},{SX("A")},"{pat}")'
EXPL = [
    ("1. Leasing Bancolombia 386735 (camión VW Constellation) registrado el 28-jul-26 y omitido en el detalle 62", "=" + SXN("BCL-L386735"),
     "NB-00001014: Cr 21202040 $572.363.196 + Cr 21202140 $103.474.304."),
    ("2. Leasing Davivienda GECOLSA (2 motogeneradores) en el mayor desde jul-25 y omitido en el detalle 62 (tampoco validado en el PT 2025)", "=" + SXN("DAV-LGECOLSA"),
     "Mayor_Dic25: reintegro anticipo GECOLSA, doc. 2672, $213.861.900."),
    ("3. Leasings Banco de Bogotá: el mayor coincide con los certificados; el detalle 62 trae saldos de tablas teóricas", "=" + SXN("BOG-L*"),
     "Saldos_x_Obligacion: Mayor − banco ≈ 0 en los 6 contratos."),
    ("4. Leasings Davivienda 1019441, 1019448 y 1021275: detalle 62 = tabla teórica; mayor = extracto dic-25 − pagos 2026", "=" + SXN("DAV-L1*"),
     "1021275: la tabla capitaliza intereses; banco mantiene $1.876.684.264."),
    ("5. Leasings Bancolombia 362095, 336817 y 331330: detalle 62 desactualizado (362095 con saldo de dic-25)", f'={SXN("BCL-L362095")}+{SXN("BCL-L336817")}+{SXN("BCL-L331330")}',
     "Ver Rev_Tablas_Amort (error en tabla 362095)."),
]
for k, (lab, f_, sop) in enumerate(EXPL):
    rr_ = 14 + k
    put(ws, rr_, 1, lab, wrap=True)
    ws.merge_cells(start_row=rr_, start_column=1, end_row=rr_, end_column=5)
    put(ws, rr_, 6, f_, fmt=NUM)
    put(ws, rr_, 9, sop, wrap=True)
    ws.row_dimensions[rr_].height = 30
put(ws, 19, 1, "6. Diferencia residual no explicada", wrap=True)
ws.merge_cells("A19:E19")
put(ws, 19, 6, "=F7-SUM(F14:F18)", fmt=NUM, fill=fill_key)
put(ws, 20, 1, "Total diferencia leasing", bold=True, fill=fill_tot)
put(ws, 20, 6, "=SUM(F14:F19)", fmt=NUM, bold=True, fill=fill_tot)
put(ws, 20, 7, '=IF(ABS(F20-F7)<0.01,"Cuadra","Revisar")')

put(ws, 22, 1, "C. Roll-forward de capital por obligación (detalle 62 + movimiento ene–jul) → saldo inicial implícito al 1-ene-26", font=f_sec, border=False)
header(ws, 23, ["ID auditoría", "Banco", "Tipo", "Saldo detalle 62 (31-jul)", "Pagos de capital (débitos)",
                "Desembolsos / renovaciones (créditos)", "Saldo inicial implícito 1-ene", "En detalle 62?", "Comentario"])
RF_IDS = [d[2] for d in det] + ["DAV-8906", "BCL-L386735", "DAV-LGECOLSA", "N/A-COMISION", "N/A-CRUCE"]
RF_COM = {
    "DAV-8906": "Crédito Davivienda $10.000 MM cancelado el 28-ene-26 con el crédito Bogotá 1155297469 (refinanciación con otro banco: baja en cuentas NIIF 9.3.3.1).",
    "BCL-L386735": "Leasing nuevo 28-jul-26 no incluido en el detalle 62 → saldo inicial implícito negativo = omisión.",
    "DAV-LGECOLSA": "Leasing GECOLSA: sin movimiento 2026 y fuera del detalle 62; su saldo ($213,9 MM) está en el saldo inicial del mayor.",
    "N/A-COMISION": "Comisiones Valley Trading registradas y reclasificadas en 21202040 (neto 0). Partida ajena a obligaciones.",
    "N/A-CRUCE": "NI-4731 'cruce anticipo calderas' en 21202140 Davivienda (neto 0). Partida ajena a obligaciones.",
    "BOG-1159383796": "Incluye el crédito 1155297469 (desembolso 28-ene) y su prórroga al 1159383796 (29-jul).",
    "BBVA-0141": "Renovación 30-jun-26 (CE-35090) a la par. El detalle 62 lo identifica como 252059; el mayor como 9600290141/0141 (y 4476 en ene–mar).",
    "BCL-1260105867": "Mayor glosa '12660105867' (error de digitación).",
}
r = 24
RF_FIRST = r
CAP_ACCTS = ("21050100", "21050201", "21202040", "21202140")
for ident in RF_IDS:
    meta = next((d for d in det if d[2] == ident), None)
    bk = meta[0] if meta else {"DAV-8906": "BANCO DAVIVIENDA", "BCL-L386735": "BANCOLOMBIA S.A.", "DAV-LGECOLSA": "BANCO DAVIVIENDA"}.get(ident, "N/A")
    tp = "Leasing" if ("-L" in ident or ident.startswith("N/A")) else "Crédito"
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, bk, font=f_in)
    put(ws, r, 3, tp)
    put(ws, r, 4, f"=SUMIFS({DR('F')},{DR('C')},A{r})", fmt=NUM)
    put(ws, r, 5, "=" + "+".join(f'SUMIFS({R_DEB},{R_ID},$A{r},{R_CTA},"{a}")' for a in CAP_ACCTS), fmt=NUM)
    put(ws, r, 6, "=" + "+".join(f'SUMIFS({R_CRE},{R_ID},$A{r},{R_CTA},"{a}")' for a in CAP_ACCTS), fmt=NUM)
    put(ws, r, 7, f"=D{r}+E{r}-F{r}", fmt=NUM)
    put(ws, r, 8, f'=IF(COUNTIF({DR("C")},A{r})>0,"Sí","NO")')
    put(ws, r, 9, RF_COM.get(ident, ""), font=f_in, wrap=True)
    r += 1
RF_LAST = r - 1
for tp, accts_si in (("Crédito", ("21050100", "21050201")), ("Leasing", ("21202040", "21202140"))):
    put(ws, r, 1, f"Total {tp}", bold=True, fill=fill_tot)
    for c in (4, 5, 6, 7):
        L = get_column_letter(c)
        put(ws, r, c, f'=SUMIFS({L}{RF_FIRST}:{L}{RF_LAST},$C${RF_FIRST}:$C${RF_LAST},"{tp}")', fmt=NUM, bold=True, fill=fill_tot)
    r += 1
    put(ws, r, 1, f"Mayor {tp} (SI / Déb / Créd / —)")
    put(ws, r, 7, f"=-({BAL(accts_si[0],'SI')}+{BAL(accts_si[1],'SI')})", fmt=NUM)
    put(ws, r, 5, f"={BAL(accts_si[0],'D')}+{BAL(accts_si[1],'D')}", fmt=NUM)
    put(ws, r, 6, f"={BAL(accts_si[0],'C')}+{BAL(accts_si[1],'C')}", fmt=NUM)
    r += 1
    put(ws, r, 1, f"Diferencia {tp} (detalle implícito − mayor)")
    for c in (5, 6, 7):
        L = get_column_letter(c)
        put(ws, r, c, f"={L}{r - 2}-{L}{r - 1}", fmt=NUM, fill=fill_key if c == 7 else None)
    r += 2
ws.freeze_panes = "B24"

# ============================================================================ 7. RECÁLCULO DE INTERESES
ws = ws_new("Recalculo_Intereses", "Recálculo de intereses (NIIF 9 – costo amortizado, método del interés efectivo)",
            "Interés = Capital × ((1 + EA)^(días/base) − 1). Las fechas y tasas en azul provienen del certificado o del detalle 62 (ver columna 'Fuente').",
            [17, 30, 17, 9, 11, 11, 7, 16, 16, 16, 16, 13, 11, 13, 50])
put(ws, 4, 1, "Parte 1. Recálculo de intereses liquidados por el banco (prueba de exactitud de la liquidación bancaria)", font=f_sec, border=False)
header(ws, 5, ["ID auditoría", "Periodo / concepto", "Capital base", "Tasa EA", "Desde", "Hasta", "Días",
               "Recalc. base 365", "Recalc. base 360", "Interés banco", "Dif. mínima (recalc − banco)",
               "% dif.", "EA implícita (365)", "Conclusión", "Comentario"], height=40)
P1 = [
    ("DAV-830624", "Intereses 21-mar a 21-sep-26 (próximo pago)", "=Certificados!J{}".format(CERT_ROW["DAV-830624"]), 0.1551,
     dt.date(2026, 3, 21), dt.date(2026, 9, 21), 565738188.80, ""),
    ("DAV-607691", "Intereses 28-ene a 28-jul-26", "=Certificados!H{}".format(CERT_ROW["DAV-607691"]), 0.1258,
     dt.date(2026, 1, 28), dt.date(2026, 7, 28), 272299382.05,
     "La tasa 'cobrada' de 12,58% EA del extracto no reconcilia con el interés facturado; la tasa implícita (~14,4% EA) es coherente con IBR 10,69% + 3,40 pts. Confirmar con Davivienda."),
    ("DAV-569032", "Intereses 28-ene a 28-jul-26", "=Certificados!H{}".format(CERT_ROW["DAV-569032"]), 0.1258,
     dt.date(2026, 1, 28), dt.date(2026, 7, 28), 206287213.20, "Igual situación que 607691."),
    ("BOG-856670629", "Cuota 12: 21-mar a 21-jun-26", 1639499992.59, 0.1494, dt.date(2026, 3, 21), dt.date(2026, 6, 21), 59390887.23,
     "Finagro: el banco liquida en base 360."),
    ("BOG-858177259", "Próxima cuota: 25-may a 25-ago-26", "=Certificados!H{}".format(CERT_ROW["BOG-858177259"]), 0.1665,
     dt.date(2026, 5, 25), dt.date(2026, 8, 25), 45166249.86, "Finagro: base 360."),
    ("BOG-453436746", "Cuota 100: 17-jun a 17-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-453436746"]), 0.1696,
     dt.date(2026, 6, 17), dt.date(2026, 7, 17), 38894999.71, "Ordinaria comercial: base 365."),
    ("BOG-1159383796", "Crédito 1155297469 cuota 2: 14-abr a 14-jul-26", 10000000000, 0.1239, dt.date(2026, 4, 14),
     dt.date(2026, 7, 14), 299617500.00, "Finagro: base 360."),
    ("BCL-1260105867", "Intereses causados 22-jul a 11-ago-26", "=Certificados!H{}".format(CERT_ROW["BCL-1260105867"]), 0.1417,
     dt.date(2026, 7, 22), dt.date(2026, 8, 11), 24607407.00, ""),
    ("BOG-L557285005", "Canon jul: 22-abr a 22-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L557285005"]), 0.1604,
     dt.date(2026, 4, 22), dt.date(2026, 7, 22), 39578631.88, ""),
    ("BOG-L556449064", "Canon jul: 27-abr a 27-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L556449064"]), 0.1425,
     dt.date(2026, 4, 27), dt.date(2026, 7, 27), 9497274.30, ""),
    ("BOG-L556449117", "Canon jul: 27-abr a 27-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L556449117"]), 0.1425,
     dt.date(2026, 4, 27), dt.date(2026, 7, 27), 9497274.29, ""),
    ("BOG-L556449180", "Canon jul: 23-abr a 23-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L556449180"]), 0.1422,
     dt.date(2026, 4, 23), dt.date(2026, 7, 23), 14118857.99, ""),
    ("BOG-L556449224", "Canon jul: 23-abr a 23-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L556449224"]), 0.1422,
     dt.date(2026, 4, 23), dt.date(2026, 7, 23), 14118857.98, ""),
    ("BOG-L556449288", "Canon jul: 20-abr a 20-jul-26", "=Certificados!H{}".format(CERT_ROW["BOG-L556449288"]), 0.1422,
     dt.date(2026, 4, 20), dt.date(2026, 7, 20), 14121779.84, ""),
]
r = 6
P1_FIRST = r
for ident, per, cap, ea, d0, d1, ib, com in P1:
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, per, font=f_in, wrap=True)
    put(ws, r, 3, cap, font=f_link if str(cap).startswith("=") else f_in, fmt=NUM)
    put(ws, r, 4, ea, font=f_in, fmt=PCT)
    put(ws, r, 5, d0, font=f_in, fmt=DATE)
    put(ws, r, 6, d1, font=f_in, fmt=DATE)
    put(ws, r, 7, f"=F{r}-E{r}", fmt=NUM0)
    put(ws, r, 8, f"=C{r}*((1+D{r})^(G{r}/365)-1)", fmt=NUM)
    put(ws, r, 9, f"=C{r}*((1+D{r})^(G{r}/360)-1)", fmt=NUM)
    put(ws, r, 10, ib, font=f_in, fmt=NUM)
    put(ws, r, 11, f"=IF(ABS(H{r}-J{r})<=ABS(I{r}-J{r}),H{r}-J{r},I{r}-J{r})", fmt=NUM)
    put(ws, r, 12, f"=K{r}/J{r}", fmt=PCT)
    put(ws, r, 13, f"=(1+J{r}/C{r})^(365/G{r})-1", fmt=PCT)
    put(ws, r, 14, f'=IF(ABS(L{r})<=Resumen!$C$7,"Razonable","Diferencia")')
    put(ws, r, 15, com, font=f_in, wrap=True)
    ws.row_dimensions[r].height = 26
    r += 1
P1_LAST = r - 1
put(ws, r, 2, "TOTAL", bold=True, fill=fill_tot)
for c in (10, 11):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{P1_FIRST}:{L}{P1_LAST})", fmt=NUM, bold=True, fill=fill_tot)
put(ws, r, 12, f"=K{r}/J{r}", fmt=PCT, bold=True, fill=fill_tot)
P1_TOT = r
r += 2

put(ws, r, 1, "Parte 2. Causación de intereses al 31-jul-2026 – recálculo vs registrado en 2130 (prueba de corte y valuación)", font=f_sec, border=False)
r += 1
header(ws, r, ["ID auditoría", "Fuente capital / tasa", "Capital 31-jul", "Tasa EA", "Inicio causación (último pago)",
               "Fecha corte", "Días", "Base días", "Interés causado recalculado", "Causado s/mayor (2130)",
               "No registrado (recalc − mayor)", "Cuenta 2130", "Último pago s/mayor", "", "Comentario"], height=40)
r += 1
P2_FIRST = r
cap_cert = lambda i: f"=Certificados!J{CERT_ROW[i]}"
cap_det = lambda i: f"=SUMIFS({DR('F')},{DR('C')},\"{i}\")"
cap_bank = lambda i: f"=Saldos_x_Obligacion!L{SX_ROW[i]}"
ea_det = lambda i: f"=INDEX({DR('J')},MATCH(\"{i}\",{DR('C')},0))"
P2 = [
    ("DAV-830624", "Certificado", cap_cert("DAV-830624"), 0.1551, dt.date(2026, 3, 21), 365,
     "Interés semestral vencido; el cliente causa cifras fijas mensuales hasta el día ~24 → subcausación."),
    ("DAV-607691", "Cert. / tasa detalle 62", cap_cert("DAV-607691"), ea_det("DAV-607691"), dt.date(2026, 7, 28), 365,
     "3 días del nuevo semestre. Tasa del nuevo periodo no certificada: se usa la del detalle 62."),
    ("DAV-569032", "Cert. / tasa detalle 62", cap_cert("DAV-569032"), ea_det("DAV-569032"), dt.date(2026, 7, 28), 365, "Ídem."),
    ("BBVA-0141", "Detalle 62 (sin certificado)", cap_det("BBVA-0141"), ea_det("BBVA-0141"), dt.date(2026, 6, 30), 365,
     "Renovado 30-jun-26 con intereses mensuales: no hay causación de julio en 2130. Verificar si el interés de julio se pagó y llevó a 53052020; si no, debe causarse."),
    ("BOG-453436746", "Certificado", cap_cert("BOG-453436746"), 0.1696, dt.date(2026, 7, 17), 365,
     "Crédito mensual: el cliente no causa entre cuotas (gasto al pago)."),
    ("BOG-856670629", "Certificado", cap_cert("BOG-856670629"), 0.1494, dt.date(2026, 6, 21), 360, "Finagro, base 360."),
    ("BOG-858177259", "Certificado", cap_cert("BOG-858177259"), 0.1665, dt.date(2026, 5, 25), 360, "Finagro, base 360."),
    ("BOG-1159383796", "Cert. 1155297469 / tasa detalle", cap_cert("BOG-1159383796"), ea_det("BOG-1159383796"),
     dt.date(2026, 7, 22), 360, "Inicio estimado 22-jul (vencimiento 22-ene-27, 6 meses, según detalle 62). Sin certificado del nuevo crédito; confirmar si hubo intereses 14 a 22-jul."),
    ("POP-631309940", "Detalle 62 (sin certificado)", cap_det("POP-631309940"), ea_det("POP-631309940"), dt.date(2026, 6, 17), 365,
     "Inicio estimado 17-jun (cuotas trimestrales al 17; vence 17-dic-26). Pago registrado 30-jun con intereses de mora $54.242."),
    ("BCL-1260105867", "Certificado", cap_cert("BCL-1260105867"), 0.1417, dt.date(2026, 7, 22), 365,
     "El banco certifica $24.607.407 causados del 22-jul al 11-ago (ver Parte 1)."),
    ("DAV-L1021275", "Banco [E] / tasa extracto", cap_bank("DAV-L1021275"), 0.1441, dt.date(2026, 7, 23), 365, ""),
    ("DAV-L1019441", "Banco [E] / tasa extracto", cap_bank("DAV-L1019441"), 0.1603, dt.date(2026, 7, 13), 365, ""),
    ("DAV-L1019448", "Banco [E] / tasa extracto", cap_bank("DAV-L1019448"), 0.1593, dt.date(2026, 7, 6), 365, ""),
    ("BOG-L557285005", "Certificado", cap_cert("BOG-L557285005"), 0.1604, dt.date(2026, 7, 22), 365, ""),
    ("BOG-L556449064", "Certificado", cap_cert("BOG-L556449064"), 0.1425, dt.date(2026, 7, 27), 365, ""),
    ("BOG-L556449117", "Certificado", cap_cert("BOG-L556449117"), 0.1425, dt.date(2026, 7, 27), 365, ""),
    ("BOG-L556449180", "Certificado", cap_cert("BOG-L556449180"), 0.1422, dt.date(2026, 7, 23), 365, ""),
    ("BOG-L556449224", "Certificado", cap_cert("BOG-L556449224"), 0.1422, dt.date(2026, 7, 23), 365, ""),
    ("BOG-L556449288", "Certificado", cap_cert("BOG-L556449288"), 0.1422, dt.date(2026, 7, 20), 365, ""),
    ("BCL-L362095", "Banco [E] / tasa detalle 62", cap_bank("BCL-L362095"), ea_det("BCL-L362095"), dt.date(2026, 6, 30), 365,
     "Inicio según periodo 'ABR-MAY-JUN' del CL-414. Confirmar fechas de canon con Bancolombia (vencimientos al día 14)."),
    ("BCL-L336817", "Banco [E] / tasa detalle 62", cap_bank("BCL-L336817"), ea_det("BCL-L336817"), dt.date(2026, 7, 19), 365,
     "Inicio = fecha del último canon registrado (CL-421)."),
    ("BCL-L331330", "Banco [E] / tasa detalle 62", cap_bank("BCL-L331330"), ea_det("BCL-L331330"), dt.date(2026, 7, 9), 365,
     "Inicio = fecha del último canon registrado (CL-420)."),
    ("BCL-L386735", "Mayor (no está en detalle 62)", f'=SUMIFS({R_CRE},{R_ID},"BCL-L386735")',
     f"=(1+({RATE_CELL['IBR 3 MESES']}+0.0135)/4)^4-1", dt.date(2026, 7, 28), 365, "IBR TV + 1,35 según glosa NB-1014."),
]
for ident, fuente, cap, ea, d0, base, com in P2:
    acct = "21300102" if "-L" in ident else "21300101"
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, fuente, font=f_in)
    put(ws, r, 3, cap, fmt=NUM)
    put(ws, r, 4, ea, font=f_in if not str(ea).startswith("=") else None, fmt=PCT)
    put(ws, r, 5, d0, font=f_in, fmt=DATE)
    put(ws, r, 6, CORTE, font=f_in, fmt=DATE)
    put(ws, r, 7, f"=F{r}-E{r}", fmt=NUM0)
    put(ws, r, 8, base, font=f_in, fmt=NUM0)
    put(ws, r, 9, f"=C{r}*((1+D{r})^(G{r}/H{r})-1)", fmt=NUM)
    put(ws, r, 12, acct, font=f_in)
    put(ws, r, 13, f'=_xlfn.MAXIFS({R_FEC},{R_ID},A{r},{R_CTA},L{r},{R_DEB},">1000")', fmt='yyyy-mm-dd;;"sin reversión"')
    put(ws, r, 10, f'=SUMIFS({R_CRE},{R_ID},A{r},{R_CTA},L{r},{R_FEC},">"&M{r})-SUMIFS({R_DEB},{R_ID},A{r},{R_CTA},L{r},{R_FEC},">"&M{r})', fmt=NUM)
    put(ws, r, 11, f"=I{r}-J{r}", fmt=NUM, fill=fill_bad)
    put(ws, r, 15, com, font=f_in, wrap=True)
    ws.row_dimensions[r].height = 26
    r += 1
P2_LAST = r - 1
put(ws, r, 2, "TOTAL", bold=True, fill=fill_tot)
for c in (9, 10, 11):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{P2_FIRST}:{L}{P2_LAST})", fmt=NUM, bold=True, fill=fill_tot)
P2_TOT = r
r += 1
put(ws, r, 2, "  de los cuales créditos (21300101)")
for c in (9, 10, 11):
    L = get_column_letter(c)
    put(ws, r, c, f'=SUMIFS({L}{P2_FIRST}:{L}{P2_LAST},$L${P2_FIRST}:$L${P2_LAST},"21300101")', fmt=NUM)
P2_OF = r
r += 1
put(ws, r, 2, "  de los cuales leasing (21300102)")
for c in (9, 10, 11):
    L = get_column_letter(c)
    put(ws, r, c, f'=SUMIFS({L}{P2_FIRST}:{L}{P2_LAST},$L${P2_FIRST}:$L${P2_LAST},"21300102")', fmt=NUM)
P2_LS = r
r += 1
put(ws, r, 2, "Saldo 2130 según balance")
put(ws, r, 10, f"=-{BAL('2130','SF')}", fmt=NUM)
r += 1
put(ws, r, 2, "Diferencia causado por obligación vs balance 2130")
put(ws, r, 10, f"=J{P2_TOT}-J{r - 1}", fmt=NUM)
put(ws, r, 11, "Redondeos de reversiones (p.ej. 556449288 $0,16).", font=f_sub, border=False)
r += 2

put(ws, r, 1, "Parte 3. Saldo de capital implícito – leasings Davivienda sin saldo en el extracto", font=f_sec, border=False)
r += 1
header(ws, r, ["ID auditoría", "Canon", "Interés del canon", "Tasa EA", "Desde", "Hasta", "Días", "Capital implícito antes del canon",
               "Capital del canon", "Capital implícito después", "Saldo detalle 62", "Dif. detalle − implícito", "", "", "Comentario"], height=40)
r += 1
P3_FIRST = r
P3 = [
    ("DAV-L1021275", "Canon 006 (sólo intereses)", 64041851, 0.1441, dt.date(2026, 4, 23), dt.date(2026, 7, 23), 0),
    ("DAV-L1019441", "Canon 033", 5062372, 0.1603, dt.date(2026, 6, 11), dt.date(2026, 7, 13), 12327298),
    ("DAV-L1019448", "Canon 031", 4177893, 0.1593, dt.date(2026, 6, 5), dt.date(2026, 7, 6), 8867342),
]
P3_ROW = {}
for ident, can, it, ea, d0, d1, capc in P3:
    P3_ROW[ident] = r
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, can, font=f_in)
    put(ws, r, 3, it, font=f_in, fmt=NUM)
    put(ws, r, 4, ea, font=f_in, fmt=PCT)
    put(ws, r, 5, d0, font=f_in, fmt=DATE)
    put(ws, r, 6, d1, font=f_in, fmt=DATE)
    put(ws, r, 7, f"=F{r}-E{r}", fmt=NUM0)
    put(ws, r, 8, f"=C{r}/((1+D{r})^(G{r}/365)-1)", fmt=NUM)
    put(ws, r, 9, capc, font=f_in, fmt=NUM)
    put(ws, r, 10, f"=H{r}-I{r}", fmt=NUM)
    put(ws, r, 11, f"=SUMIFS({DR('F')},{DR('C')},A{r})", fmt=NUM)
    put(ws, r, 12, f"=K{r}-J{r}", fmt=NUM)
    put(ws, r, 15, "Estimación indicativa (la tasa puede variar dentro del periodo). Solicitar certificado de saldo.", font=f_in, wrap=True)
    r += 1
put(ws, r, 2, "TOTAL", bold=True, fill=fill_tot)
for c in (10, 11, 12):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{P3_FIRST}:{L}{r - 1})", fmt=NUM, bold=True, fill=fill_tot)
P3_TOT = r
for rr_, ident in DOCS_ROWS:  # comparación diferida en Certificados sección D
    p3 = P3_ROW[ident]
    put(CERT_WS, rr_, 14, f'=IF(H{rr_}=Recalculo_Intereses!C{p3},"Sí","NO")')
    put(CERT_WS, rr_, 15, f'=IF(G{rr_}=Recalculo_Intereses!I{p3},"Sí","NO")')
ws.freeze_panes = "C6"

# -*- coding: utf-8 -*-
# Bloque insertado en build_cruce_62.py después de "7. RECÁLCULO DE INTERESES".
# Construye: anexos A01..Ann (tabla del cliente vs recálculo), Rev_Tablas_Amort, Int_Glosas e Intereses_x_Obligacion.
import importlib.util as _ilu

_spec = _ilu.spec_from_file_location("pta", __file__.rsplit("/", 1)[0] + "/parse_tablas_amort.py")
_pta = _ilu.module_from_spec(_spec)
_spec.loader.exec_module(_pta)
F_TAB = f"{UP}/89656b61-64._Tabla_amortizaci_n_obligaciones_financieras.xlsm"
TAB = _pta.parse(F_TAB)

ANX_ORDER = ["DAV-L1019441", "DAV-L1019448", "DAV-L1021275", "DAV-830624", "DAV-607691", "DAV-569032",
             "BOG-453436746", "BOG-856670629", "BOG-858177259", "BOG-1155297469", "BOG-1159383796",
             "BOG-L556449224", "BOG-L556449180", "BOG-L556449117", "BOG-L556449064", "BOG-L556449288", "BOG-L557285005",
             "POP-631309940", "BBVA-9600277262", "BCL-1260105867", "BCL-L362095", "BCL-L336817", "BCL-L331330",
             "DYP-482800190700", "DAV-L1013934"]
TAB_ID_SX = {"BOG-1159383796": "BOG-1159383796"}  # ID de la tabla -> ID de saldo (mismo salvo excepciones)
OBS_T = {
    "DAV-L1019441": "Tabla teórica a tasa fija (TIR mensual constante) y cuota fija $17.681.217; el banco liquida a tasa variable (canon 033: capital $12.327.298 / interés $5.062.372). El saldo del detalle 62 es el de esta tabla, no el del banco.",
    "DAV-L1019448": "Tabla teórica a tasa fija y cuota constante; el banco liquida a tasa variable (canon 031: capital $8.867.342 / interés $4.177.893). Detalle 62 = saldo de la tabla.",
    "DAV-L1021275": "La TIR de la tabla supera el canon de sólo intereses de los primeros 12 trimestres: la tabla capitaliza intereses y el saldo crece ($1.876,7 MM → $1.883,7 MM). El banco mantiene el capital en $1.876.684.264 y cobra 14,41% EA ($64,0 MM en jul-26 vs $51,5 MM en la tabla).",
    "DAV-830624": "Interés mensual simple con IBR 6M + 3 pts / 12; el banco liquida semestre vencido a 15,51% EA. La columna 'medición al costo' (costo amortizado) resta los intereses del saldo: error de fórmula (ver columna Z).",
    "DAV-607691": "Interés mensual simple con IBR 6M + 3 pts; el contrato es IBR + 3,40 y el banco factura semestral. Columna de costo amortizado con el mismo error de fórmula (ver columna Z).",
    "DAV-569032": "Igual que 607691: puntos adicionales de la tabla (3%) distintos al contrato (3,40%); costo amortizado mal formulado.",
    "BOG-453436746": "Capital fijo $62,5 MM mensual coincide con el banco; el interés de la tabla usa el IBR proyectado.",
    "BOG-856670629": "Capital trimestral $182,2 MM coincide con el banco (diferencia de redondeo en el saldo).",
    "BOG-858177259": "Capital trimestral $125 MM coincide con el banco.",
    "BOG-1155297469": "Tabla del crédito cancelado/prorrogado el 14/29-jul-26.",
    "BOG-1159383796": "La hoja 'Hoja2' es copia de la tabla del 1155297469: fechas feb–jul 2026 anteriores al inicio (22-jul-26) y saldo cero al 14-jul. No sirve como tabla del crédito nuevo.",
    "BOG-L556449224": "Tabla a tasa fija; el banco liquida DTF variable: saldo certificado mayor que el de la tabla.",
    "BOG-L556449180": "Tabla a tasa fija; el banco liquida DTF variable: saldo certificado mayor que el de la tabla.",
    "BOG-L556449117": "Tabla a tasa fija con TIR muy baja (0,59% trimestral) y capital acelerado: saldo $57 MM por debajo del banco.",
    "BOG-L556449064": "Tabla a tasa fija con TIR muy baja (0,59% trimestral) y capital acelerado: saldo $57 MM por debajo del banco.",
    "BOG-L556449288": "Tabla a tasa fija; el banco liquida DTF variable: saldo certificado mayor que el de la tabla.",
    "BOG-L557285005": "Tabla a tasa fija; el banco liquida IBR 3M variable: saldo certificado mayor que el de la tabla.",
    "POP-631309940": "Capital trimestral $143,75 MM coincide con lo pagado; vence 17-dic-26 (todo corriente).",
    "BBVA-9600277262": "Tabla del crédito anterior (11-dic-25 a 11-jun-26). No hay tabla del crédito renovado el 30-jun-26 (glosa 'crédito 0141').",
    "BCL-1260105867": "Capital anual $1.666,7 MM; coincide con el certificado (capital vigente $3.333.333.334).",
    "BCL-L362095": "ERROR: el periodo 5 (14-jun-26) aumenta el saldo en $48,8 MM (capital negativo) y la última cuota deja saldo negativo. El detalle 62 usa el saldo de dic-25 ($497,2 MM) sin actualizar.",
    "BCL-L336817": "Tabla a tasa fija; detalle 62 = saldo de la tabla. Sin extracto 2026.",
    "BCL-L331330": "Tabla con amortización de capital constante $14.586.381, pero desde feb-26 se paga $17.114.382 (mayor): la tabla no refleja los pagos reales.",
    "DYP-482800190700": "Crédito Davivienda DYPSIS $72.300 MM (30-may-25 a 28-may-27) incluido en el archivo de tablas pero NO registrado en el mayor de Guaicaramo. Confirmar titular, vinculación y garantías otorgadas.",
    "DAV-L1013934": "Contrato INPARME terminado; sin saldo en el mayor 2026.",
}

ANX = {}
for n, ident in enumerate(ANX_ORDER, start=1):
    t = TAB[ident]
    sh = f"A{n:02d}-{t['hoja']}"[:31]
    lea = t["tipo"] == "leasing"
    p = t["param"]
    ws = ws_new(sh, f"Anexo {sh} – Tabla de amortización del cliente vs recálculo de auditoría – {ident}",
                f"Fuente: 64._Tabla_amortización_obligaciones_financieras.xlsm, hoja '{t['hoja']}'. Columnas A–G: cliente [A]; H–P: recálculo auditoría [B].",
                [7, 11, 17, 15, 15, 15, 17, 17, 15, 17, 13, 13, 16, 9, 8, 8])
    vi = p.get("Valor razonable") if lea else p.get("Capital")
    if lea:
        tasa_txt, fini, plazo = p.get("Tasa acordada"), p.get("Fecha inicio del arrendamiento"), f"{p.get('Plazo del arrendamiento')} {p.get('Plazo del arrendamiento (unidad)') or ''}"
        unidad = str(p.get("Tasa de interés (unidad)") or "").upper()
        nper = 12 if "MENS" in unidad else 4
    else:
        tasa_txt, fini, plazo = p.get("Tasa"), p.get("Fecha inicio"), f"{p.get('Plazo meses')} meses"
        unidad, nper = "MENSUAL", 12
    put(ws, 3, 1, "Banco / contrato", bold=True)
    put(ws, 3, 3, f"{t.get('banco')} – {t.get('contrato')}", font=f_in)
    put(ws, 4, 1, "Valor inicial", bold=True)
    put(ws, 4, 3, vi, font=f_in, fmt=NUM)
    put(ws, 5, 1, "Fecha inicio", bold=True)
    put(ws, 5, 3, dt.date.fromisoformat(fini) if isinstance(fini, str) else fini, font=f_in, fmt=DATE)
    put(ws, 6, 1, "Plazo / tasa pactada", bold=True)
    put(ws, 6, 3, f"{plazo} – {tasa_txt}", font=f_in)
    put(ws, 7, 1, "Tasa periódica tabla (TIR cliente)", bold=True)
    put(ws, 7, 3, t.get("tir") if lea else "Por fila (col. N)", font=f_in, fmt="0.0000%")
    put(ws, 8, 1, "TIR recalculada (flujos col. M)", bold=True)
    put(ws, 9, 1, "Tasa EA equivalente", bold=True)
    put(ws, 3, 6, "Resultados al 31-jul-2026", bold=True, border=False)
    labels = ["Fecha última cuota ≤ corte", "Saldo tabla cliente", "Saldo recalculado", "Interés tabla ene–jul 26",
              "Capital siguiente 12 meses (CP)", "Máx. dif. saldo |cliente − recalc.|"]
    for k, lab in enumerate(labels):
        put(ws, 4 + k, 6, lab)
    HR = 11
    header(ws, HR, ["Per.", "Fecha", "Saldo inicial", "Interés", "Capital", "Pago / cuota", "Saldo final",
                    "Saldo inicial recalc.", "Interés recalc.", "Saldo final recalc.", "Dif. interés", "Dif. saldo",
                    "Flujo (TIR)", "Tasa periodo", "Ene–jul 26", "CP 12m"], height=36)
    r = HR + 1
    A_FIRST = r
    filas = t["filas"]
    if lea:
        filas = [{"per": 0, "fecha": fini, "saldo_ini": None, "interes": 0, "capital": 0, "pago": 0, "saldo_fin": vi}] + filas
    for k, fl in enumerate(filas):
        fe = dt.date.fromisoformat(fl["fecha"]) if isinstance(fl["fecha"], str) else fl["fecha"]
        put(ws, r, 1, fl["per"], font=f_in)
        put(ws, r, 2, fe, font=f_in, fmt=DATE)
        put(ws, r, 3, fl["saldo_ini"], font=f_in, fmt=NUM)
        put(ws, r, 4, fl["interes"], font=f_in, fmt=NUM)
        put(ws, r, 5, fl["capital"], font=f_in, fmt=NUM)
        put(ws, r, 6, fl["pago"] if lea else fl.get("cuota"), font=f_in, fmt=NUM)
        put(ws, r, 7, fl["saldo_fin"], font=f_in, fmt=NUM)
        put(ws, r, 14, "=$C$7" if lea else fl.get("tasa"), font=f_base if lea else f_in, fmt="0.0000%")
        if k == 0:
            put(ws, r, 8, f"=G{r}" if lea else f"=C{r}", fmt=NUM)
            put(ws, r, 9, 0 if lea else f"=H{r}*N{r}", fmt=NUM)
            put(ws, r, 10, f"=H{r}" if lea else f"=H{r}-E{r}", fmt=NUM)
            put(ws, r, 13, f"=-$C$4", fmt=NUM)
        else:
            put(ws, r, 8, f"=J{r - 1}", fmt=NUM)
            put(ws, r, 9, f"=H{r}*N{r}", fmt=NUM)
            put(ws, r, 10, f"=H{r}+I{r}-F{r}" if lea else f"=H{r}-E{r}", fmt=NUM)
            put(ws, r, 13, f"=F{r}", fmt=NUM)
        put(ws, r, 11, f"=D{r}-I{r}", fmt=NUM)
        put(ws, r, 12, f"=G{r}-J{r}", fmt=NUM)
        put(ws, r, 15, f'=IF(AND(B{r}>=DATE(2026,1,1),B{r}<=Resumen!$C$5),1,0)', fmt=NUM0)
        put(ws, r, 16, f'=IF(AND(B{r}>Resumen!$C$5,B{r}<=DATE(2027,7,31)),1,0)', fmt=NUM0)
        r += 1
    A_LAST = r - 1
    rng = lambda c: f"${c}${A_FIRST}:${c}${A_LAST}"
    put(ws, 8, 3, f'=IFERROR(IRR({rng("M")}),"N/A")', fmt="0.0000%")
    put(ws, 9, 3, f'=IF(ISNUMBER(C8),(1+C8)^{nper}-1,"N/A")', fmt="0.00%")
    put(ws, 4, 9, f'=_xlfn.MAXIFS({rng("B")},{rng("B")},"<="&Resumen!$C$5)', fmt=DATE)
    put(ws, 5, 9, f'=SUMIFS({rng("G")},{rng("B")},I4)', fmt=NUM, fill=fill_key)
    put(ws, 6, 9, f'=SUMIFS({rng("J")},{rng("B")},I4)', fmt=NUM)
    put(ws, 7, 9, f'=SUMIFS({rng("D")},{rng("O")},1)', fmt=NUM)
    put(ws, 8, 9, f'=SUMIFS({rng("E")},{rng("P")},1)', fmt=NUM)
    put(ws, 9, 9, f"=SUMPRODUCT(MAX(ABS({rng('L')})))", fmt=NUM)
    ws.freeze_panes = ws.cell(row=HR + 1, column=3)
    ANX[ident] = {"sh": sh, "n": n}

# ---------------------------------------------------------------------------- Rev_Tablas_Amort
ws = ws_new("Rev_Tablas_Amort", "Revisión y recálculo de las tablas de amortización del cliente (papel 64) – corte 31-jul-2026",
            "Cada fila resume un anexo A01–A25. Saldo banco desde Saldos_x_Obligacion; interés registrado desde Intereses_x_Obligacion.",
            [17, 9, 13, 20, 20, 8, 16, 11, 26, 11, 10, 9, 11, 17, 17, 15, 17, 15, 16, 16, 15, 16, 16, 15, 15, 17, 17, 60])
header(ws, 4, ["ID auditoría", "Hoja cliente", "Anexo", "Banco", "Contrato", "Tipo", "Valor inicial", "Fecha inicio",
               "Tasa pactada", "TIR tabla (periodo)", "TIR recalc.", "Dif. TIR (pb)", "Última cuota ≤ corte",
               "Saldo tabla 31-jul", "Saldo detalle 62", "Tabla − detalle 62", "Saldo banco 31-jul", "Tabla − banco",
               "Interés tabla ene–jul 26", "Interés registrado 2026", "Registrado − tabla", "CP 12m según tabla",
               "CP detalle 62", "CP tabla − detalle", "Máx. dif. recálculo saldo", "Costo amortizado cliente (col. W) 31-jul",
               "C. amortizado − saldo tabla", "Observación de auditoría"], height=56)
r = 5
RT_FIRST = r
for ident in ANX_ORDER:
    t, a = TAB[ident], ANX[ident]
    sh = f"'{a['sh']}'"
    sxid = TAB_ID_SX.get(ident, ident)
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, t["hoja"], font=f_in)
    c = put(ws, r, 3, a["sh"], font=Font(name=FN, size=9, color="0563C1", underline="single"))
    c.hyperlink = f"#{sh}!A1"
    put(ws, r, 4, str(t.get("banco")), font=f_in)
    put(ws, r, 5, str(t.get("contrato")), font=f_in)
    put(ws, r, 6, "Leasing" if t["tipo"] == "leasing" else "Crédito")
    put(ws, r, 7, f"={sh}!C4", fmt=NUM)
    put(ws, r, 8, f"={sh}!C5", fmt=DATE)
    put(ws, r, 9, f"={sh}!C6", wrap=True)
    put(ws, r, 10, f'=IF(ISNUMBER({sh}!C7),{sh}!C7,"Por fila")', fmt="0.000%")
    put(ws, r, 11, f"=IFERROR({sh}!C8,\"N/A\")", fmt="0.000%")
    put(ws, r, 12, f'=IF(AND(ISNUMBER(J{r}),ISNUMBER(K{r})),(K{r}-J{r})*10000,"N/A")', fmt=PB)
    put(ws, r, 13, f"={sh}!I4", fmt=DATE)
    put(ws, r, 14, f"={sh}!I5", fmt=NUM)
    put(ws, r, 15, f'=IF(COUNTIF({DR("C")},A{r})>0,SUMIFS({DR("F")},{DR("C")},A{r}),"No está")', fmt=NUM)
    put(ws, r, 16, f'=IF(ISNUMBER(O{r}),N{r}-O{r},"N/A")', fmt=NUM)
    put(ws, r, 17, f'=IFERROR(INDEX(Saldos_x_Obligacion!$L:$L,MATCH("{sxid}",Saldos_x_Obligacion!$A:$A,0)),"N/D")', fmt=NUM)
    put(ws, r, 18, f'=IF(ISNUMBER(Q{r}),N{r}-Q{r},"N/D")', fmt=NUM)
    put(ws, r, 19, f"={sh}!I7", fmt=NUM)
    put(ws, r, 20, f'=IFERROR(INDEX(Intereses_x_Obligacion!$I:$I,MATCH("{sxid}",Intereses_x_Obligacion!$A:$A,0)),"N/D")', fmt=NUM)
    put(ws, r, 21, f'=IF(ISNUMBER(T{r}),T{r}-S{r},"N/D")', fmt=NUM)
    put(ws, r, 22, f"=MAX({sh}!I8,0)", fmt=NUM)
    put(ws, r, 23, f'=IF(COUNTIF({DR("C")},A{r})>0,SUMIFS({DR("M")},{DR("C")},A{r}),"No está")', fmt=NUM)
    put(ws, r, 24, f'=IF(ISNUMBER(W{r}),V{r}-W{r},"N/A")', fmt=NUM)
    put(ws, r, 25, f"={sh}!I9", fmt=NUM)
    if t["tipo"] == "credito":
        f_ = [x for x in t["filas"] if x["fecha"] and x["fecha"] <= "2026-07-31"]
        ca = f_[-1].get("ca_saldo") if f_ else None
        put(ws, r, 26, ca, font=f_in, fmt=NUM)
        put(ws, r, 27, f'=IF(ISNUMBER(Z{r}),Z{r}-N{r},"N/A")', fmt=NUM, fill=fill_bad if ca and abs(ca - (f_[-1]["saldo_fin"] or 0)) > 1000 else None)
    else:
        put(ws, r, 26, "N/A")
        put(ws, r, 27, "N/A")
    put(ws, r, 28, OBS_T.get(ident, ""), font=f_in, wrap=True)
    ws.row_dimensions[r].height = 48
    r += 1
RT_LAST = r - 1
put(ws, r, 1, "TOTAL (obligaciones en el mayor)", bold=True, fill=fill_tot)
for c in (14, 15, 17, 19, 20, 22, 23):
    L = get_column_letter(c)
    put(ws, r, c, f'=SUMIFS({L}{RT_FIRST}:{L}{RT_LAST},$A${RT_FIRST}:$A${RT_LAST},"<>DYP*",$A${RT_FIRST}:$A${RT_LAST},"<>DAV-L1013934",$A${RT_FIRST}:$A${RT_LAST},"<>BOG-1155297469",$A${RT_FIRST}:$A${RT_LAST},"<>BBVA-9600277262")',
        fmt=NUM, bold=True, fill=fill_tot)
RT_TOT = r
r += 2
put(ws, r, 1, "Conclusiones de la revisión de tablas", font=f_sec, border=False)
r += 1
for txt in [
    "1. Los saldos del detalle 62 se toman de estas tablas teóricas (tasa fija / IBR-DTF proyectado), no de los extractos: en leasings la diferencia con el banco proviene de la tasa variable efectivamente cobrada.",
    "2. El recálculo aritmético (columna Y) valida la consistencia interna de cada tabla (saldo inicial + interés − pago = saldo final). Las diferencias materiales indican fórmulas alteradas o filas pegadas como valor.",
    "3. Las tablas de créditos Davivienda tienen la columna de costo amortizado mal formulada (resta intereses del saldo): no deben usarse para medir el pasivo bajo NIIF 9 (B5.4.1).",
    "4. Faltan tablas para: leasing Bancolombia 386735, leasing GECOLSA (Davivienda), BBVA renovado (0141) y el crédito Bogotá 1159383796 ('Hoja2' es una copia del crédito anterior).",
    "5. El archivo incluye el crédito Davivienda DYPSIS ($72.300 MM) que no está en el mayor de Guaicaramo: confirmar titular y garantías (NIC 24, NIIF 9 §2.1(e) garantías financieras, NIC 37).",
]:
    put(ws, r, 1, txt, wrap=True, border=False)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
    ws.row_dimensions[r].height = 28
    r += 1
ws.freeze_panes = "B5"

# ---------------------------------------------------------------------------- Int_Glosas
DIRECTOS = {"BOG-453436746", "DAV-L1019441", "DAV-L1019448", "BCL-L331330", "BCL-L336817"}


def pint(desc):
    m = re.search(r"INTE?RE?SES?\s*[:$]?\s*([\d][\d\.,]*\d)", desc)
    if not m:
        return None
    s_ = m.group(1)
    if re.fullmatch(r"\d{1,3}(\.\d{3})+\.\d{1,2}", s_):
        a_, b_ = s_.rsplit(".", 1)
        return float(a_.replace(".", "") + "." + b_)
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", s_):
        return float(s_.replace(".", ""))
    if re.fullmatch(r"\d+", s_):
        return float(s_)
    return None


ws = ws_new("Int_Glosas", "Intereses pagados directamente al gasto (obligaciones sin causación en 2130) – según glosa del comprobante",
            "Fuente: descripción de los comprobantes CL del Mayor_Mov. Aplica a obligaciones de pago mensual que el cliente no causa en 2130.",
            [17, 13, 11, 11, 17, 12, 90])
header(ws, 4, ["ID auditoría", "Documento", "Fecha", "Cuenta", "Interés según glosa", "¿Dato?", "Descripción"])
r = 5
IG_FIRST = r
for row in mov_rows:
    (i, _a, acct, nom, _cc, terc, deb, cre, _net, fecha, nit, _t, doc, tdoc, desc, *_rest) = row
    acct = str(acct).strip()
    desc = (desc or "").strip()
    if acct[:4] not in ("2105", "2120") or acct == "21050301":
        continue
    ident = assign_id(acct, terc, desc)
    if ident not in DIRECTOS:
        continue
    v = pint(desc)
    put(ws, r, 1, ident, font=f_bold)
    put(ws, r, 2, (doc or "").strip(), font=f_in)
    put(ws, r, 3, fecha, font=f_in, fmt=DATE)
    put(ws, r, 4, acct, font=f_in)
    put(ws, r, 5, v, font=f_in, fmt=NUM, fill=None if v is not None else fill_bad)
    put(ws, r, 6, "Sí" if v is not None else "Sin dato")
    put(ws, r, 7, desc, font=f_in)
    r += 1
IG_LAST = r - 1
put(ws, r, 4, "TOTAL", bold=True, fill=fill_tot)
put(ws, r, 5, f"=SUM(E{IG_FIRST}:E{IG_LAST})", fmt=NUM, bold=True, fill=fill_tot)
put(ws, r, 6, f'=COUNTIF(F{IG_FIRST}:F{IG_LAST},"Sin dato")', fmt=NUM0, bold=True, fill=fill_tot)
ws.freeze_panes = "B5"

# ---------------------------------------------------------------------------- Intereses_x_Obligacion
ws = ws_new("Intereses_x_Obligacion", "Intereses por obligación 2026 – causación (2130), pagos, gasto y recálculo al corte",
            "Saldo 2130 por obligación = detalle 31-dic-25 + causaciones 2026 − reversiones 2026. Gasto 2026 = causaciones 2130 + intereses pagados directo al gasto (Int_Glosas).",
            [17, 15, 16, 16, 16, 16, 16, 9, 17, 17, 16, 17, 16, 45])
header(ws, 4, ["ID auditoría", "Banco", "2130 al 31-dic-25", "Causaciones 2026 (Cr 2130)", "Reversiones 2026 (Db 2130)",
               "2130 al 31-jul-26 por obligación", "Pagado directo a gasto (glosas)", "Cuotas sin dato", "Gasto intereses 2026 registrado",
               "Interés según tabla ene–jul 26", "Registrado − tabla", "Causación recalculada 31-jul", "Causación no registrada",
               "Comentario"], height=48)
r = 5
IX_FIRST = r
P2R = lambda c: f"Recalculo_Intereses!${c}${P2_FIRST}:${c}${P2_LAST}"
for bank_, ids in SX_IDS:
    for ident in ids:
        put(ws, r, 1, ident, font=f_bold)
        put(ws, r, 2, bank_, font=f_in)
        put(ws, r, 3, f'=SUMIFS({MD("G")},{MD("J")},A{r},{MD("C")},"2130*")', fmt=NUM)
        put(ws, r, 4, f'=SUMIFS({R_CRE},{R_ID},A{r},{R_CTA},"2130*")', fmt=NUM)
        put(ws, r, 5, f'=SUMIFS({R_DEB},{R_ID},A{r},{R_CTA},"2130*")', fmt=NUM)
        put(ws, r, 6, f"=C{r}+D{r}-E{r}", fmt=NUM)
        put(ws, r, 7, f"=SUMIFS(Int_Glosas!$E${IG_FIRST}:$E${IG_LAST},Int_Glosas!$A${IG_FIRST}:$A${IG_LAST},A{r})", fmt=NUM)
        put(ws, r, 8, f'=COUNTIFS(Int_Glosas!$A${IG_FIRST}:$A${IG_LAST},A{r},Int_Glosas!$F${IG_FIRST}:$F${IG_LAST},"Sin dato")', fmt=NUM0)
        put(ws, r, 9, f"=D{r}+G{r}", fmt=NUM)
        put(ws, r, 10, f'=IFERROR(INDEX(Rev_Tablas_Amort!$S:$S,MATCH(A{r},Rev_Tablas_Amort!$A:$A,0)),"Sin tabla")', fmt=NUM)
        put(ws, r, 11, f'=IF(ISNUMBER(J{r}),I{r}-J{r},"N/A")', fmt=NUM)
        put(ws, r, 12, f"=SUMIFS({P2R('I')},{P2R('A')},A{r})", fmt=NUM)
        put(ws, r, 13, f"=L{r}-F{r}", fmt=NUM, fill=fill_bad)
        put(ws, r, 14, {"DAV-8906": "Cancelado ene-26: causación de 2025 reversada con el pago.",
                        "BOG-453436746": "Paga mensual; el cliente no causa entre cuotas.",
                        "BBVA-0141": "Sin causación de julio tras la renovación (PEND-3).",
                        "DAV-LGECOLSA": "Sin tabla ni extracto: intereses no identificados en 2026 (PEND-12)."}.get(ident, ""), font=f_in, wrap=True)
        r += 1
IX_LAST = r - 1
put(ws, r, 1, "TOTAL", bold=True, fill=fill_tot)
for c in (3, 4, 5, 6, 7, 8, 9, 12, 13):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{IX_FIRST}:{L}{IX_LAST})", fmt=NUM, bold=True, fill=fill_tot)
IX_TOT = r
r += 1
put(ws, r, 1, "Balance 2130 (signo +)")
put(ws, r, 3, f"=-{BAL('2130', 'SI')}", fmt=NUM)
put(ws, r, 4, f"={BAL('2130', 'C')}", fmt=NUM)
put(ws, r, 5, f"={BAL('2130', 'D')}", fmt=NUM)
put(ws, r, 6, f"=-{BAL('2130', 'SF')}", fmt=NUM)
r += 1
put(ws, r, 1, "Diferencia")
for c in (3, 4, 5, 6):
    L = get_column_letter(c)
    put(ws, r, c, f"={L}{r - 2}-{L}{r - 1}", fmt=NUM, fill=fill_key)
r += 2
section(ws, r, "B. Cruce de cuentas de intereses del balance contra la suma por obligación", 14)
r += 1
header(ws, r, ["Código", "Cuenta contable", "Saldo / movimiento balance", "Valor según obligaciones", "Diferencia", "", "Comentario"])
r += 1
IXB_FIRST = r
lea_ids = lambda col: f'SUMIFS({col}{IX_FIRST}:{col}{IX_LAST},$A${IX_FIRST}:$A${IX_LAST},"*-L*")'
crb_ids = lambda col: f'(SUM({col}{IX_FIRST}:{col}{IX_LAST})-{lea_ids(col)})'
rowsB = [
    ("21300101", "Intereses financieros por pagar O.F. (SF)", f"=-{BAL('21300101', 'SF')}", f"={crb_ids('F')}", ""),
    ("21300102", "Intereses leasing por pagar (SF)", f"=-{BAL('21300102', 'SF')}", f"={lea_ids('F')}", ""),
    ("53052030", "Intereses sobre préstamos causados (SF)", f"={BAL('53052030', 'SF')}", f"={crb_ids('F')}", "El saldo debe igualar los intereses por pagar de créditos."),
    ("53052040", "Intereses leasing causados (SF)", f"={BAL('53052040', 'SF')}", f"={lea_ids('F')}", ""),
    ("53052020", "Intereses préstamos efectivamente pagados (débitos)", f"={BAL('53052020', 'D')}", f"={crb_ids('E')}+{crb_ids('G')}",
     "Débitos = reversiones de 2130 + intereses pagados directo. La diferencia corresponde a pagos sin dato en glosa u otros intereses."),
    ("53052010", "Intereses leasing efectivamente pagados (débitos)", f"={BAL('53052010', 'D')}", f"={lea_ids('E')}+{lea_ids('G')}", "Ídem para leasing."),
    ("5305", "Gasto de intereses neto ene–jul (10+20+30+40)", "=Gasto_Intereses!B14", f"=I{IX_TOT}", "Gasto registrado vs suma por obligación (causaciones + pagos directos)."),
]
for cod, nom_, a_, b_, com in rowsB:
    put(ws, r, 1, cod, font=f_in)
    put(ws, r, 2, nom_)
    put(ws, r, 3, a_, fmt=NUM)
    put(ws, r, 4, b_, fmt=NUM)
    put(ws, r, 5, f"=C{r}-D{r}", fmt=NUM)
    put(ws, r, 7, com, wrap=True)
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=14)
    ws.row_dimensions[r].height = 26
    r += 1
IXB_LAST = r - 1
ws.freeze_panes = "C5"

# ============================================================================ 8. GASTO DE INTERESES (prueba analítica)
ws = ws_new("Gasto_Intereses", "Prueba analítica sustantiva – gasto por intereses ene–jul 2026",
            "Expectativa = deuda promedio × ((1 + tasa promedio ponderada)^(días/365) − 1).",
            [58, 20, 60])
rows = [
    ("Deuda financiera 1-ene-26 (2105 sin TC + 2120)", f"=-({BAL('21050100','SI')}+{BAL('21050201','SI')}+{BAL('2120','SI')})", NUM, "Balance de prueba."),
    ("Deuda financiera 31-jul-26 (2105 sin TC + 2120)", f"=-({BAL('21050100','SF')}+{BAL('21050201','SF')}+{BAL('2120','SF')})", NUM, "Balance de prueba."),
    ("Deuda promedio simple", "=AVERAGE(B4:B5)", NUM, ""),
    ("Tasa EA promedio ponderada (detalle 62, tasas del cliente)", f"={D_TPOND}", PCT, "Detalle_62."),
    ("Días del periodo (1-ene a 31-jul-26)", "=DATE(2026,7,31)-DATE(2025,12,31)", NUM0, ""),
    ("Gasto de intereses esperado", "=B6*((1+B7)^(B8/365)-1)", NUM, ""),
    ("Intereses leasing pagados (53052010)", "=" + BAL("53052010", "SF"), NUM, ""),
    ("Intereses préstamos pagados (53052020)", "=" + BAL("53052020", "SF"), NUM, ""),
    ("Intereses préstamos causados (53052030)", "=" + BAL("53052030", "SF"), NUM, ""),
    ("Intereses leasing causados (53052040)", "=" + BAL("53052040", "SF"), NUM, ""),
    ("Gasto de intereses registrado", "=SUM(B10:B13)", NUM, "Excluye mora (53052002) y sobregiros (53052001)."),
    ("Causación no registrada al 31-jul (Recalculo_Intereses Parte 2)", f"=Recalculo_Intereses!K{P2_TOT}", NUM, ""),
    ("Gasto de intereses ajustado", "=B14+B15", NUM, ""),
    ("Diferencia ajustado − esperado", "=B16-B9", NUM, ""),
    ("% diferencia", "=B17/B9", PCT, ""),
    ("Umbral de expectativa (±)", 0.10, PCT, "Supuesto de auditoría: ±10% para una prueba analítica de precisión moderada (tasas variables IBR/DTF)."),
    ("Conclusión", '=IF(ABS(B18)<=B19,"Gasto razonable dentro del umbral","Fuera de umbral – investigar")', None, ""),
    ("Intereses de mora registrados (53052002)", "=" + BAL("53052002", "SF"), NUM, "Incluye $54.242 Banco Popular (CL-411)."),
]
header(ws, 3, ["Concepto", "Valor", "Fuente / comentario"])
for k, (lab, f, fmt, com) in enumerate(rows, start=4):
    put(ws, k, 1, lab, bold=lab.startswith(("Gasto de intereses", "Conclusión")))
    put(ws, k, 2, f, font=f_in if not str(f).startswith("=") else None, fmt=fmt, fill=fill_key if k in (19,) else None)
    put(ws, k, 3, com, font=f_sub)
GI_PCT, GI_CONC = "Gasto_Intereses!$B$18", "Gasto_Intereses!$B$20"

# ============================================================================ 9. CLASIFICACIÓN CP/LP
ws = ws_new("Clasif_CP_LP", "Clasificación corriente / no corriente (NIC 1 párr. 69–76) – por obligación con tablas de amortización",
            "CP auditoría = saldo total si vence ≤ 31-jul-27; si no, capital de los próximos 12 meses según la tabla (papel 64), acotado al saldo del mayor. Sin tabla: se mantiene la clasificación contable.",
            [26, 18, 18, 18, 18, 18, 18, 18, 18, 50])
header(ws, 4, ["Concepto", "CP auditoría (tablas / vencimiento)", "CP detalle 62 corregido", "Cuenta CP mayor", "Cuenta LP mayor",
               "Saldo total mayor", "", "", "", "Comentario"])
put(ws, 9, 1, "Reclasificación propuesta LP → CP (CP auditoría − cuenta CP del mayor)", font=f_sec, border=False)
header(ws, 10, ["Concepto", "Reclasificación", "", "", "", "", "", "", "", "Comentario"])
put(ws, 16, 1, "Porción corriente por obligación", font=f_sec, border=False)
header(ws, 17, ["ID auditoría", "Tipo", "Saldo mayor 31-jul-26", "Vencimiento final", "CP 12m según tabla", "CP registrado (sin tabla)",
                "CP auditoría", "CP detalle 62 corregido", "CP auditoría − detalle", "Criterio"], height=36)
VENC_OVR = {"BOG-1159383796": dt.date(2027, 1, 22), "BBVA-0141": dt.date(2026, 12, 30)}
r = 18
CL_FIRST = r
for bank_, ids in SX_IDS:
    for ident in ids:
        tp = "Leasing" if ("-L" in ident or ident.startswith("N/A")) else "Crédito"
        put(ws, r, 1, ident, font=f_bold)
        put(ws, r, 2, tp)
        put(ws, r, 3, f"=Saldos_x_Obligacion!H{SX_ROW[ident]}", fmt=NUM)
        if ident in VENC_OVR:
            put(ws, r, 4, VENC_OVR[ident], font=f_in, fmt=DATE)
        else:
            put(ws, r, 4, f'=IFERROR(INDEX({DR("P")},MATCH(A{r},{DR("C")},0)),"N/D")', fmt=DATE)
        put(ws, r, 5, f'=IFERROR(INDEX(Rev_Tablas_Amort!$V:$V,MATCH(A{r},Rev_Tablas_Amort!$A:$A,0)),"Sin tabla")', fmt=NUM)
        cpa = "21202040" if tp == "Leasing" else "21050100"
        put(ws, r, 6, f'=SUMIFS({MD("G")},{MD("J")},A{r},{MD("C")},"{cpa}")+SUMIFS({R_CRE},{R_ID},A{r},{R_CTA},"{cpa}")-SUMIFS({R_DEB},{R_ID},A{r},{R_CTA},"{cpa}")', fmt=NUM)
        put(ws, r, 7, f'=IF(C{r}<=0,0,IF(AND(ISNUMBER(D{r}),D{r}<=DATE(2027,7,31)),C{r},IF(ISNUMBER(E{r}),MIN(MAX(E{r},0),C{r}),MIN(MAX(F{r},0),C{r}))))', fmt=NUM)
        put(ws, r, 8, f"=SUMIFS({DR('W')},{DR('C')},A{r})", fmt=NUM)
        put(ws, r, 9, f"=G{r}-H{r}", fmt=NUM)
        put(ws, r, 10, f'=IF(C{r}<=0,"Sin saldo",IF(AND(ISNUMBER(D{r}),D{r}<=DATE(2027,7,31)),"Vence en 12 meses: todo corriente",IF(ISNUMBER(E{r}),"Capital próximos 12 meses según tabla","Sin tabla: se mantiene CP contable [P]")))', wrap=True)
        r += 1
CL_LAST = r - 1
put(ws, r, 1, "TOTAL", bold=True, fill=fill_tot)
for c in (3, 6, 7, 8, 9):
    L = get_column_letter(c)
    put(ws, r, c, f"=SUM({L}{CL_FIRST}:{L}{CL_LAST})", fmt=NUM, bold=True, fill=fill_tot)
CL_R = lambda col, tp: f'SUMIFS(${col}${CL_FIRST}:${col}${CL_LAST},$B${CL_FIRST}:$B${CL_LAST},"{tp}")'
for k, (tp, cp, lp) in enumerate((("Crédito", "21050100", "21050201"), ("Leasing", "21202040", "21202140")), start=5):
    put(ws, k, 1, f"{tp}s")
    put(ws, k, 2, "=" + CL_R("G", tp), fmt=NUM)
    put(ws, k, 3, "=" + CL_R("H", tp), fmt=NUM)
    put(ws, k, 4, f"=-{BAL(cp, 'SF')}", fmt=NUM)
    put(ws, k, 5, f"=-{BAL(lp, 'SF')}", fmt=NUM)
    put(ws, k, 6, f"=D{k}+E{k}", fmt=NUM)
put(ws, 7, 1, "TOTAL", bold=True, fill=fill_tot)
for c in range(2, 7):
    L = get_column_letter(c)
    put(ws, 7, c, f"={L}5+{L}6", fmt=NUM, bold=True, fill=fill_tot)
put(ws, 5, 10, "Las cuentas CP/LP del mayor se asignan por origen del crédito, no por vencimiento.", font=f_in, wrap=True)
put(ws, 6, 10, "Incluye 386735 y GECOLSA sin tabla (se mantiene su clasificación contable).", font=f_in, wrap=True)
put(ws, 11, 1, "Créditos: 21050201 → 21050100")
put(ws, 11, 2, "=B5-D5", fmt=NUM, fill=fill_key)
put(ws, 12, 1, "Leasing: 21202140 → 21202040")
put(ws, 12, 2, "=B6-D6", fmt=NUM, fill=fill_key)
put(ws, 12, 10, "Negativo = el mayor tiene más corto plazo que el calculado (reclasificar CP → LP).", font=f_in, wrap=True)
put(ws, 13, 1, "Tarjetas de crédito: agrupadas en 210502 (LP) → presentar como corriente")
put(ws, 13, 2, f"=-{BAL('21050301','SF')}", fmt=NUM, fill=fill_key)
r += 2
put(ws, r, 1, "Errores de clasificación en el detalle 62", font=f_sec, border=False)
header(ws, r + 1, ["ID auditoría", "Saldo", "CP cliente", "LP cliente", "CP+LP − saldo", "", "", "", "", "Error"])
k = r + 2
for bank, num, ident, rr in det:
    cp, lp, sal = rr[10] or 0, rr[11] or 0, rr[3]
    if abs(cp + lp - sal) > 1 or cp < 0 or lp < 0 or cp > sal:
        rowd = D_FIRST + [d[2] for d in det].index(ident)
        put(ws, k, 1, ident, font=f_bold)
        put(ws, k, 2, f"=Detalle_62!F{rowd}", fmt=NUM)
        put(ws, k, 3, f"=Detalle_62!M{rowd}", fmt=NUM)
        put(ws, k, 4, f"=Detalle_62!N{rowd}", fmt=NUM)
        put(ws, k, 5, f"=C{k}+D{k}-B{k}", fmt=NUM)
        put(ws, k, 10, "CP mayor que el saldo y LP negativo." if cp > sal else "CP + LP no suma el saldo (LP posiblemente copiado de otra fila).",
            font=f_in, wrap=True)
        k += 1

# ============================================================================ 10. HALLAZGOS
ws = ws_new("Hallazgos", "Hallazgos, referencias NIIF y ajustes propuestos",
            "Montos vinculados a las hojas de trabajo. Ajustes sujetos a la materialidad del encargo (no informada).",
            [5, 16, 70, 18, 26, 55])
header(ws, 4, ["#", "Afirmación", "Hallazgo", "Monto (COP)", "Referencia normativa", "Recomendación / acción", "Estado"])
ws.column_dimensions["G"].width = 24
H_PEND = {2: "PEND-1, PEND-2, PEND-12", 3: "PEND-2", 5: "PEND-3, PEND-5, PEND-10", 6: "PEND-6", 8: "PEND-8",
          11: "PEND-7", 13: "PEND-1 a PEND-5", 15: "PEND-11", 18: "PEND-12", 20: "PEND-14", 21: "PEND-5, PEND-14",
          22: "PEND-14", 23: "PEND-13", 24: "PEND-8"}
cdm = "Cruce_Detalle_Mayor"
H = [
    ("Integridad", "El leasing Bancolombia 386735 (camión VW Constellation, IBR TV + 1,35) se registró en el mayor el 28-jul-26 (NB-00001014) pero NO está en el detalle 62 del cliente.",
     f"={cdm}!F14", "NIIF 16 p.26; NIC 1 p.15", "Incluir el contrato en el detalle 62 con su tabla de amortización; solicitar contrato y certificado."),
    ("Exactitud / Integridad", "Diferencia de capital de leasing entre el mayor (2120) y el detalle 62, explicada obligación por obligación (Saldos_x_Obligacion): 386735 y GECOLSA omitidos; Bogotá, Davivienda y Bancolombia con saldos de tablas teóricas en el detalle.",
     f"={cdm}!F7", "NIIF 16 p.36; NIC 1 p.15", "Actualizar el detalle 62 con los saldos del banco y conciliar mensualmente por contrato."),
    ("Exactitud", "Leasings Bancolombia: el detalle 62 difiere del saldo estimado del banco (362095 conserva el saldo de dic-25; la tabla 331330 no refleja la cuota real pagada).",
     f'=SUMIFS({SX("P")},{SX("A")},"BCL-L3*")', "NIIF 16 p.36", "Obtener certificados Bancolombia al 31-jul-26 y actualizar tablas."),
    ("Exactitud", "Leasings Banco de Bogotá (pozos 1–5 y planta Jenbacher): el detalle 62 está por debajo del saldo certificado tras el canon de julio (p.ej. pozos 3 y 4 −$57 MM c/u).",
     f"=-{C_BOGL}", "NIIF 16 p.36", "Actualizar el detalle 62 con los saldos certificados."),
    ("Corte / Valuación", "Intereses causados al 31-jul-26 no registrados: el cliente causa cifras fijas hasta ~el día 24 y no causa en obligaciones de pago mensual (453436746, BBVA, leasings mensuales) ni en el periodo posterior al último pago.",
     f"=Recalculo_Intereses!K{P2_TOT}", "NIC 1 p.27-28 (base de acumulación); NIIF 9 p.5.4.1 y B5.4 (método del interés efectivo)", "Registrar AJE-1; implementar causación diaria con tabla de amortización por obligación."),
    ("Exactitud", "Davivienda 607691 y 569032: la tasa 'cobrada' informada (12,58% EA) no reconcilia con los intereses facturados (tasa implícita ~14,4% EA). El detalle 62 usa 16,02% EA.",
     f"=Recalculo_Intereses!K{P1_FIRST + 1}+Recalculo_Intereses!K{P1_FIRST + 2}", "NIIF 9 p.B5.4.5 (tasas variables); NIIF 7 p.39", "Solicitar a Davivienda la liquidación detallada (IBR aplicado y spread)."),
    ("Exactitud (detalle)", "Tasas EA del detalle 62 calculadas con periodicidad distinta a la declarada (BBVA, Bancolombia 1260105867 y 362095) y tasas del detalle distintas a las certificadas (p.ej. 453436746: 17,52% vs 16,96%; 856670629: 15,83% vs 14,94%).",
     None, "NIIF 7 p.33-40 (riesgo de tasa)", "Usar la tasa certificada por el banco y la periodicidad contractual."),
    ("Presentación", "Clasificación CP/LP: el detalle 62 tiene errores (Banco Popular CP $575 MM > saldo $287,5 MM y LP negativo; Pozo #2 CP + LP ≠ saldo) y las cuentas CP/LP del mayor no reflejan vencimientos a 12 meses.",
     "=Clasif_CP_LP!B11+Clasif_CP_LP!B12", "NIC 1 p.69-76", "Reclasificar la porción corriente (AJE-2) calculada con tablas y vencimientos; corregir el detalle 62."),
    ("Presentación", "Las tarjetas de crédito por pagar (21050301) están agrupadas dentro de 'Obligaciones financieras a largo plazo' (210502).",
     "=Clasif_CP_LP!B13", "NIC 1 p.69", "Presentar como pasivo corriente."),
    ("Exactitud", "Abonos a capital aplicados por Davivienda por excedentes de pago (127,57 + 269,47 + 627,79 + 617,95 + 786,80) no registrados: el cliente lleva todo el pago a intereses.",
     "=127.57+269.47+627.79+617.95+786.8", "NIIF 9 p.3.3.1", "Inmaterial; registrar en el siguiente pago."),
    ("Ocurrencia / Exactitud", "Leasing Davivienda 1019441: el banco reporta el canon 032 (11-jun) pagado por $1.029.562, sólo intereses; la contabilidad registró capital $11.973.677 (CL-416).",
     11973677, "NIIF 16 p.36", "Obtener el estado de cuenta del contrato y verificar la aplicación del pago."),
    ("Control interno", "Partidas ajenas registradas en cuentas de obligaciones financieras: comisiones Valley Trading en 21202040 y cruce de anticipo de calderas (NI-4731) en 21202140 (netas cero).",
     "=4068134+31360500+8261283", "NIC 1 p.32 (no compensar)", "Restringir el uso de las cuentas 2120 a contratos de leasing con tercero y número de contrato."),
    ("Integridad / Confirmación", "Obligaciones sin certificado ni extracto en el PDF: BBVA ($4.700 MM), Banco Popular, Bogotá 1159383796 (nuevo crédito) y leasings Bancolombia.",
     f"=Certificados!K{C_SINCERT}", "NIA 505", "Enviar confirmaciones bancarias."),
    ("Exactitud (detalle)", "Vencimientos del detalle 62 distintos a los certificados: pozos 3 y 4 (abr-2028 vs 27-jul-2028), pozos 1 y 2 (20-ene vs 23-ene-2030), Bancolombia 1260105867 (21 vs 22-abr-2028).",
     None, "NIIF 7 p.39 (análisis de vencimientos)", "Corregir; afecta la clasificación CP/LP y la revelación de liquidez."),
    ("Valuación / Revelación", "Renovaciones y refinanciaciones (BBVA 30-jun; Bogotá 1155297469 → 1159383796 el 29-jul; Davivienda 8906 pagado con el crédito Bogotá 1155297469): se registraron a la par sin costos de transacción.",
     None, "NIIF 9 p.3.3.2 y B3.3.6 (prueba del 10%)", "Documentar la prueba del 10% (modificación sustancial) y los costos de transacción, si los hubo."),
    ("Control interno", "Causación no sistemática de intereses de leasing (p.ej. pozos 3 y 4: feb–mar $480.943 vs abr $8,6 MM) y glosas con valores errados (CL-393, NI-4397) o números de crédito errados (BBVA 252059 vs 0141; '12660105867').",
     None, "NIC 8 p.41; NIC 1 p.27", "Usar tablas de amortización por obligación; revisar las glosas."),
    ("Integridad", "Mayor vs balance: los débitos, créditos y saldos de las cuentas 21 del movimiento cuadran con los dos balances de prueba (24-ago y 29-sep), que son idénticos.",
     f"=Cruce_Mov_Balance!H{CMB_TOT}", "—", "Sin excepción."),
    ("Integridad", "Leasing Davivienda GECOLSA (2 motogeneradores, registrado jul-25 en 21202140) incluido en el mayor pero omitido en el detalle 62 y en la validación del PT 2025; sin tabla ni extracto.",
     f'=SUMIFS({SX("H")},{SX("A")},"DAV-LGECOLSA")', "NIIF 16 p.26 y 47; NIA 505", "Obtener contrato, tabla y certificado; incluirlo en el detalle 62."),
    ("Valuación", "Las tablas de leasing del cliente son teóricas (tasa fija / índice proyectado) y no se remiden cuando cambian IBR/DTF: el saldo de la tabla difiere del saldo bancario.",
     f'=SUMIFS(Rev_Tablas_Amort!R{RT_FIRST}:R{RT_LAST},Rev_Tablas_Amort!F{RT_FIRST}:F{RT_LAST},"Leasing")', "NIIF 16 p.36, 42(b) y 43 (nueva medición por cambio en índice o tasa)", "Remedir el pasivo con la tasa vigente o usar el saldo bancario; actualizar tablas cada periodo."),
    ("Exactitud (tabla)", "Tabla Bancolombia 362095 con error: el periodo 5 (14-jun-26) aumenta el saldo (capital negativo) y la última cuota deja saldo negativo.",
     f"=ABS('{ANX['BCL-L362095']['sh']}'!E17)", "NIIF 16 p.36", "Corregir la tabla con el plan de pagos del banco."),
    ("Exactitud (tabla)", "La tabla 'Hoja2' del crédito Bogotá 1159383796 es copia de la del 1155297469 (fechas feb–jul 2026 y saldo cero al 14-jul): no representa el crédito vigente.",
     None, "NIIF 9 B5.4", "Elaborar la tabla del crédito nuevo (22-jul-26 a 22-ene-27)."),
    ("Valuación (tabla)", "Columna 'medición al costo' (costo amortizado) de las tablas de créditos Davivienda mal formulada: resta los intereses del saldo (830624: costo amortizado < capital).",
     f'=INDEX(Rev_Tablas_Amort!AA:AA,MATCH("DAV-830624",Rev_Tablas_Amort!A:A,0))', "NIIF 9 p.4.2.1, 5.4.1 y B5.4.1", "Corregir la fórmula; no usar esa columna para medir el pasivo."),
    ("Derechos y obligaciones", "El archivo de tablas incluye el crédito Davivienda DYPSIS 7000482800190700 por $72.300 MM (may-25 a may-27) que no está en el mayor de Guaicaramo.",
     72300000000, "NIC 24; NIIF 9 §2.1(e) y B2.5 (garantías financieras); NIC 37 p.86", "Confirmar titular, vinculación y si Guaicaramo es garante/codeudor; revelar si aplica."),
    ("Integridad (tablas)", "Faltan tablas de amortización para: leasing Bancolombia 386735, leasing GECOLSA, BBVA renovado (0141) y crédito Bogotá 1159383796.",
     None, "NIIF 7 p.39; NIC 1 p.69", "Solicitar las tablas o planes de pago bancarios."),
]
r = 5
H_ROW = {}
for n, (af, txt, monto, ref, rec) in enumerate(H, start=1):
    put(ws, r, 1, n)
    put(ws, r, 2, af, wrap=True)
    put(ws, r, 3, txt, wrap=True)
    put(ws, r, 4, monto if monto is not None else "Cualitativo", fmt=NUM, font=f_in if isinstance(monto, (int, float)) else None)
    put(ws, r, 5, ref, wrap=True)
    put(ws, r, 6, rec, wrap=True)
    put(ws, r, 7, f"Pendiente de soporte ({H_PEND[n]})" if n in H_PEND else "Concluido", wrap=True,
        fill=fill_bad if n in H_PEND else None, bold=True)
    ws.row_dimensions[r].height = 64
    H_ROW[n] = r
    r += 1
r += 1
section(ws, r, "Ajustes propuestos (AJE)", 6)
r += 1
header(ws, r, ["AJE", "Cuenta", "Descripción", "Débito", "Crédito", "Soporte"])
r += 1
AJE = [
    ("AJE-1", "53052030", "Intereses sobre préstamos causados", f"=Recalculo_Intereses!K{P2_OF}", None, "Causación de intereses de créditos al 31-jul-26"),
    ("AJE-1", "21300101", "Intereses financieros por pagar O.F.", None, f"=Recalculo_Intereses!K{P2_OF}", ""),
    ("AJE-1", "53052040", "Intereses leasing causados", f"=Recalculo_Intereses!K{P2_LS}", None, "Causación de intereses de leasing al 31-jul-26"),
    ("AJE-1", "21300102", "Intereses leasing por pagar", None, f"=Recalculo_Intereses!K{P2_LS}", ""),
    ("AJE-2", "21050201", "Obligación financiera LP", "=Clasif_CP_LP!B11", None, "Reclasificación de la porción corriente (presentación)"),
    ("AJE-2", "21050100", "Obligación financiera CP", None, "=Clasif_CP_LP!B11", ""),
    ("AJE-2", "21202140", "Leasing LP", "=Clasif_CP_LP!B12", None, ""),
    ("AJE-2", "21202040", "Leasing CP", None, "=Clasif_CP_LP!B12", ""),
]
A_FIRST = r
for aje, cta, des, db, cr, sop in AJE:
    put(ws, r, 1, aje)
    put(ws, r, 2, cta, font=f_in)
    put(ws, r, 3, des)
    put(ws, r, 4, db, fmt=NUM)
    put(ws, r, 5, cr, fmt=NUM)
    put(ws, r, 6, sop, wrap=True)
    r += 1
A_FIRST_H, A_LAST_H = A_FIRST, r - 1
put(ws, r, 3, "Sumas iguales", bold=True, fill=fill_tot)
put(ws, r, 4, f"=SUM(D{A_FIRST}:D{r - 1})", fmt=NUM, bold=True, fill=fill_tot)
put(ws, r, 5, f"=SUM(E{A_FIRST}:E{r - 1})", fmt=NUM, bold=True, fill=fill_tot)
r += 1
put(ws, r, 3, "AJE-2 se calcula con tablas de amortización y vencimientos (Clasif_CP_LP); para 386735 y GECOLSA (sin tabla) se mantiene la clasificación contable. El mayor de leasing cuadra con el banco en Bogotá; en Davivienda/Bancolombia se usa el extracto de dic-25 (estimado).",
    font=f_sub, border=False)

# ============================================================================ 10b. PENDIENTES
ws = ws_new("Pendientes", "Pendientes de soporte – el papel se emite sin esta información",
            "Las cifras afectadas se dejan calculadas con la información disponible y supuestos documentados; se actualizarán al recibir el soporte.",
            [9, 62, 22, 18, 16, 22, 14, 50])
header(ws, 4, ["Ref.", "Pendiente / información requerida", "Obligación (ID)", "Monto afectado (COP)", "Hallazgo relacionado",
               "Solicitar a", "Estado", "Tratamiento provisional en este papel"], height=36)
PEND = [
    ("PEND-1", "Certificados de saldo de capital al 31-jul-26 de los leasings Davivienda 1019441-2, 1019448-3 y 1021275-3 (los extractos sólo traen la factura del canon).",
     "DAV-L1019441 / L1019448 / L1021275", f"=Recalculo_Intereses!K{P3_TOT}", "2, 3, 13", "Davivienda Leasing",
     "Saldo banco estimado = extracto 31-dic-25 (PT 2025) + movimientos 2026 [E] (Saldos_x_Obligacion). Facturas LEASING_1275-3/9441-2/9448-3_DAVIVIENDA.pdf recibidas sin saldo de capital.",
     "Parcial – estimado con extracto dic-25"),
    ("PEND-2", "Certificados de saldo de los leasings Bancolombia 362095, 336817, 331330 y contrato + tabla de amortización del nuevo leasing 386735.",
     "BCL-L362095 / L336817 / L331330 / L386735", f"={cdm}!F16", "1, 2, 3, 13", "Bancolombia",
     "Saldo banco estimado con extracto dic-25 (PT 2025) [E]; 386735 se toma del mayor (NB-1014). La diferencia mayor vs detalle quedó explicada por obligación.",
     "Parcial – estimado con extracto dic-25"),
    ("PEND-3", "Confirmación BBVA del crédito renovado 30-jun-26 ($4.700 MM): número de obligación, tasa y si el interés de julio se pagó (y a qué cuenta se llevó).",
     "BBVA-0141", f'=SUMIFS(Recalculo_Intereses!K{P2_FIRST}:K{P2_LAST},Recalculo_Intereses!A{P2_FIRST}:A{P2_LAST},"BBVA-0141")',
     "5, 13", "BBVA / cliente", "Se incluye en AJE-1 la causación estimada de julio; se retira si se demuestra el pago en 53052020."),
    ("PEND-4", "Confirmación Banco Popular del crédito 631309940-2 (saldo, tasa, fechas de cuota).",
     "POP-631309940", f'=SUMIFS({DR("F")},{DR("C")},"POP-631309940")', "13", "Banco Popular",
     "Saldo del detalle 62; inicio de causación estimado 17-jun."),
    ("PEND-5", "Certificado del nuevo crédito Bogotá 1159383796 (prórroga del 1155297469): fecha de inicio, tasa y si hubo intereses entre el 14 y el 22-jul.",
     "BOG-1159383796", f'=SUMIFS(Recalculo_Intereses!K{P2_FIRST}:K{P2_LAST},Recalculo_Intereses!A{P2_FIRST}:A{P2_LAST},"BOG-1159383796")',
     "5, 13", "Banco de Bogotá", "Inicio estimado 22-jul con tasa del detalle 62 (IBR 3M + 0,8)."),
    ("PEND-6", "Liquidación detallada Davivienda créditos 607691 y 569032 (IBR aplicado y spread): la tasa declarada 12,58% EA no reconcilia con lo facturado.",
     "DAV-607691 / DAV-569032", f"=Recalculo_Intereses!K{P1_FIRST + 1}+Recalculo_Intereses!K{P1_FIRST + 2}", "6", "Davivienda",
     "Se acepta lo facturado (coincide con lo contabilizado); diferencia de tasa queda abierta."),
    ("PEND-7", "Estado de cuenta del leasing Davivienda 1019441: aplicación del pago del canon 032 (banco: sólo intereses $1.029.562; contabilidad: capital $11.973.677).",
     "DAV-L1019441", 11973677, "11", "Cliente", "Sin ajuste hasta verificar. LEASING_9441-2_DAVIVIENDA.pdf reconfirma el lado del banco; falta la explicación del cliente del registro CL-416.",
     "Parcial – banco reconfirma"),
    ("PEND-8", "Tablas de amortización por obligación para determinar la porción corriente a 12 meses (NIC 1.69).",
     "Todas", "=Clasif_CP_LP!B11+Clasif_CP_LP!B12", "8, 24", "Cliente", "Tablas recibidas (papel 64): CP calculado por obligación en Clasif_CP_LP. Faltan tablas de 386735, GECOLSA, BBVA 0141 y 1159383796.",
     "Parcial – tablas recibidas"),
    ("PEND-9", "Materialidad del encargo (global, de ejecución y umbral de errores triviales).",
     "—", None, "Todos", "Socio / gerente del encargo", "Umbral de redondeo $1.000 y tolerancia de recálculo 2% (hoja Resumen, celdas amarillas)."),
    ("PEND-10", "Fechas reales de último pago/canon de Banco Popular y leasings Bancolombia (362095, 336817, 331330) para la causación al corte.",
     "POP / BCL-L*", f'=SUMIFS(Recalculo_Intereses!K{P2_FIRST}:K{P2_LAST},Recalculo_Intereses!A{P2_FIRST}:A{P2_LAST},"BCL-L*")+SUMIFS(Recalculo_Intereses!K{P2_FIRST}:K{P2_LAST},Recalculo_Intereses!A{P2_FIRST}:A{P2_LAST},"POP-*")',
     "5", "Cliente / bancos", "Fechas estimadas (anotadas en Recalculo_Intereses Parte 2)."),
    ("PEND-11", "Documentación de la prueba del 10% (NIIF 9 B3.3.6) y costos de transacción de las renovaciones BBVA y Bogotá y de la refinanciación Davivienda 8906.",
     "BBVA-0141 / BOG-1159383796 / DAV-8906", None, "15", "Cliente", "Se asume modificación no sustancial a la par, sin costos."),
    ("PEND-12", "Contrato, tabla de amortización y certificado del leasing Davivienda GECOLSA (2 motogeneradores) registrado en 21202140 y no incluido en el detalle 62.",
     "DAV-LGECOLSA", f'=SUMIFS({SX("H")},{SX("A")},"DAV-LGECOLSA")', "18", "Davivienda Leasing / cliente", "Se toma el saldo del mayor; sin intereses identificados en 2026."),
    ("PEND-13", "Aclaración del crédito Davivienda DYPSIS 7000482800190700 ($72.300 MM) incluido en el archivo de tablas: titular, vinculación y garantías de Guaicaramo.",
     "DYP-482800190700", 72300000000, "23", "Cliente / Davivienda", "No se incluye en el pasivo de Guaicaramo; posible revelación (NIC 24 / garantías)."),
    ("PEND-14", "Tablas de amortización corregidas: 362095 (error en periodo 5), crédito 1159383796 (Hoja2 copiada) y columna de costo amortizado de créditos Davivienda.",
     "BCL-L362095 / BOG-1159383796 / DAV-*", None, "20, 21, 22", "Cliente", "Se usan los saldos del mayor y del banco; las tablas no se usan para medir."),
]
r = 5
PN_FIRST = r
for ref, txt, ident, monto, hz, quien, trat, *est in PEND:
    est = est[0] if est else "Pendiente"
    put(ws, r, 1, ref, bold=True)
    put(ws, r, 2, txt, wrap=True)
    put(ws, r, 3, ident, wrap=True)
    put(ws, r, 4, monto if monto is not None else "Cualitativo", fmt=NUM, font=f_in if isinstance(monto, (int, float)) else None)
    put(ws, r, 5, hz)
    put(ws, r, 6, quien, wrap=True)
    put(ws, r, 7, est, bold=True, fill=fill_bad, wrap=True)
    put(ws, r, 8, trat, wrap=True)
    ws.row_dimensions[r].height = 52
    r += 1
PN_LAST = r - 1
r += 1
put(ws, r, 2, "Pendientes abiertos", bold=True, fill=fill_tot)
put(ws, r, 4, f'=COUNTIF(G{PN_FIRST}:G{PN_LAST},"<>Recibido")', fmt=NUM0, bold=True, fill=fill_tot)
PN_COUNT = f"Pendientes!$D${r}"
r += 1
put(ws, r, 2, "Instrucción: al recibir un soporte suficiente, cambiar el Estado a 'Recibido' (se descuenta del contador), actualizar los datos de entrada (azul) en la hoja correspondiente y recalcular.",
    font=f_sub, border=False)
ws.freeze_panes = "B5"

# ============================================================================ 11. RESUMEN
ws = wb.create_sheet("Resumen", 0)
ws.sheet_view.showGridLines = True
for i, w in enumerate([4, 62, 20, 20, 20, 16], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws["A1"] = "GUAICARAMO S.A.S. – Auditoría de Obligaciones Financieras (cuenta 21 / papel 62) – corte 31-jul-2026"
ws["A1"].font = f_title
ws["A2"] = ("Cruce movimiento ↔ detalle 62 ↔ certificados bancarios ↔ balance de prueba; recálculo de intereses (NIIF 9 costo amortizado, NIIF 16, NIC 1). "
            "Cifras en pesos colombianos. Azul = dato de la fuente; negro = fórmula; verde = vínculo a otra hoja; amarillo = cifra clave.")
ws["A2"].font = f_sub
put(ws, 4, 2, "Parámetros", font=f_sec, border=False)
put(ws, 5, 2, "Fecha de corte")
put(ws, 5, 3, CORTE, font=f_in, fmt=DATE)
put(ws, 6, 2, "Umbral de diferencia trivial / redondeo (COP)")
put(ws, 6, 3, 1000, font=f_in, fmt=NUM, fill=fill_key)
note(ws, 6, 3, "Supuesto de auditoría: diferencias de hasta $1.000 se consideran redondeo. Ajustar según la materialidad del encargo.")
put(ws, 7, 2, "Tolerancia recálculo de intereses (%)")
put(ws, 7, 3, 0.02, font=f_in, fmt=PCT, fill=fill_key)
put(ws, 4, 5, "Estado del papel", font=f_sec, border=False)
put(ws, 5, 5, '="Emitido con " & ' + PN_COUNT + ' & " pendientes"', bold=True, fill=fill_bad)
ws.merge_cells("E5:F5")
put(ws, 6, 5, "Ver hoja 'Pendientes'", font=f_sub, border=False)
note(ws, 7, 3, "Supuesto de auditoría: ±2% por diferencias de convención (días/base, fecha de fijación del IBR/DTF).")

put(ws, 9, 2, "Resultados por prueba", font=f_sec, border=False)
header(ws, 10, ["", "Prueba", "Valor A", "Valor B", "Diferencia", "Estado"], c0=1)
TESTS = [
    ("1. Integridad: movimiento auxiliar vs balance – SF cuenta 21",
     f"=Cruce_Mov_Balance!F{CMB_TOT}", f"=Cruce_Mov_Balance!G{CMB_TOT}"),
    ("2. Integridad: balance 29-sep vs balance 24-ago – nº cuentas con diferencias", "=" + BAL_NDIFF, "=0"),
    ("3. Detalle 62: suma recalculada vs 'Total obligaciones' del cliente", f"=Detalle_62!F{D_TOT}", f"=Detalle_62!F{D_TOT + 1}"),
    ("4. Créditos: detalle 62 vs mayor (21050100 + 21050201)", f"={cdm}!D6", f"={cdm}!E6"),
    ("5. Leasing: detalle 62 vs mayor (21202040 + 21202140)", f"={cdm}!D7", f"={cdm}!E7"),
    ("6. Certificados: capital banco 31-jul vs detalle 62 (obligaciones con saldo certificado)",
     f'=SUMIFS(Certificados!J{C_FIRST}:J{C_LAST},Certificados!H{C_FIRST}:H{C_LAST},"<>")',
     f'=SUMIFS(Certificados!K{C_FIRST}:K{C_LAST},Certificados!H{C_FIRST}:H{C_LAST},"<>")'),
    ("7. Certificados: pagos/intereses certificados vs registrados (comprobantes)",
     f"=Certificados!H{C_PAGOS_TOT}", f"=Certificados!I{C_PAGOS_TOT}"),
    ("8. Recálculo de intereses liquidados por el banco (Parte 1)",
     f"=Recalculo_Intereses!J{P1_TOT}+Recalculo_Intereses!K{P1_TOT}", f"=Recalculo_Intereses!J{P1_TOT}"),
    ("9. Intereses causados al 31-jul: recalculado vs registrado en 2130",
     f"=Recalculo_Intereses!I{P2_TOT}", f"=Recalculo_Intereses!J{P2_TOT}"),
    ("10. Intereses por pagar: suma por obligación vs balance 2130",
     f"=Recalculo_Intereses!J{P2_TOT}", f"=-{BAL('2130','SF')}"),
]
r = 11
for lab, a_, b_ in TESTS:
    put(ws, r, 2, lab, wrap=True)
    put(ws, r, 3, a_, fmt=NUM)
    put(ws, r, 4, b_, fmt=NUM)
    put(ws, r, 5, f"=C{r}-D{r}", fmt=NUM)
    put(ws, r, 6, f'=IF(ABS(E{r})<=$C$6,"OK","DIFERENCIA")', bold=True)
    ws.row_dimensions[r].height = 26
    r += 1
r += 1
put(ws, r, 2, "Indicadores clave", font=f_sec, border=False)
r += 1
KPIS = [
    ("Saldo de capital según mayor 31-jul (2105 sin TC + 2120)", f"=Gasto_Intereses!B5", NUM),
    ("Saldo según detalle 62", f"=Detalle_62!F{D_TOT}", NUM),
    ("Leasing 386735 omitido en el detalle 62", f"={cdm}!F14", NUM),
    ("Leasing GECOLSA (Davivienda) omitido en el detalle 62", f"={cdm}!F15", NUM),
    ("Leasings Bogotá: detalle 62 por debajo del mayor/certificados", f"={cdm}!F16", NUM),
    ("Diferencia de leasing sin explicar (residual)", f"={cdm}!F19", NUM),
    ("% del saldo del detalle 62 cubierto con certificados", f"={C_COBERT}", PCT),
    ("Intereses no causados al 31-jul (AJE-1)", f"=Recalculo_Intereses!K{P2_TOT}", NUM),
    ("Reclasificación LP → CP propuesta (créditos + leasing)", "=Clasif_CP_LP!B11+Clasif_CP_LP!B12", NUM),
    ("Gasto de intereses: desviación vs expectativa", f"={GI_PCT}", PCT),
    ("Conclusión prueba analítica", f"={GI_CONC}", None),
]
for lab, f, fmt in KPIS:
    put(ws, r, 2, lab)
    put(ws, r, 3, f, fmt=fmt, fill=fill_key)
    r += 1
r += 1
put(ws, r, 2, "Conclusión", font=f_sec, border=False)
r += 1
concl = [
    "INTEGRIDAD – Mayor vs balance: el movimiento auxiliar de la cuenta 21 cuadra exactamente (débitos, créditos y saldos) con los dos balances de prueba, que son idénticos entre sí. Sin excepción.",
    "INTEGRIDAD – Detalle 62 vs mayor: los créditos cuadran (diferencia de redondeo $0,67). El leasing NO cuadra ($1.018 MM): leasing Bancolombia 386735 y leasing GECOLSA omitidos y saldos de tablas teóricas en el detalle; la diferencia queda explicada por obligación (Saldos_x_Obligacion).",
    "EXACTITUD – Certificados: los créditos cuadran con los extractos salvo redondeos de abonos a capital (< $5.000). Los leasings de Banco de Bogotá están subvalorados en el detalle 62. Los pagos registrados coinciden con los extractos, salvo el canon 032 del leasing 1019441.",
    "RECÁLCULO – La liquidación bancaria es razonable en base 365 (Bogotá Finagro en base 360), excepto Davivienda 607691/569032, cuya tasa declarada no reconcilia. Hay intereses causados no registrados al 31-jul (AJE-1) por la política de causación del cliente.",
    "PRESENTACIÓN – Se requiere reclasificar la porción corriente (NIC 1.69) y corregir los errores CP/LP del detalle 62. Ver la hoja 'Hallazgos'.",
    "PENDIENTES – El papel se emite sin certificados de saldo de leasings Davivienda/Bancolombia, sin confirmaciones de BBVA, Banco Popular y del crédito Bogotá 1159383796, sin tablas de amortización y sin materialidad del encargo. Las conclusiones afectadas quedan 'Pendiente de soporte' (ver hoja 'Pendientes').",
]
for t in concl:
    put(ws, r, 2, t, wrap=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 40
    r += 1
r += 1
put(ws, r, 2, "Hojas del libro", font=f_sec, border=False)
r += 1
for nm, ds in [("Mayor_Mov", "Movimiento auxiliar ene–jul con el ID de obligación asignado por auditoría"),
               ("Balance_21", "Cuentas 21/5305 de ambos balances de prueba y su comparación"),
               ("Pruebas_Auditoria", "Bloques 1–4: movimiento vs balance, certificados bancarios, saldo por obligación y recálculo de intereses (cálculos en amarillo)"),
               ("Tablas_Amortizacion", "Revisión y recálculo de las 25 tablas de amortización del cliente (recálculo en amarillo)"),
               ("Detalle_62", "Detalle del cliente con recálculo de tasas, subtotales y CP/LP"),
               ("Cruce_Detalle_Mayor", "Detalle 62 vs mayor y roll-forward de capital por obligación"),
               ("Gasto_Intereses", "Prueba analítica del gasto por intereses"),
               ("Clasif_CP_LP", "Clasificación corriente/no corriente (NIC 1)"),
               ("Hallazgos", "Hallazgos, normas y ajustes propuestos"),
               ("Pendientes", "Información pendiente de soporte y tratamiento provisional"),
               ("PT_62", "Papel de trabajo principal (formato PT): aseveraciones, procedimientos, secciones 1–5 y conclusión"),
               ("Mayor_Dic25", "Detalle del mayor al 31-dic-25 (PT 2025) por obligación"),
               ("Intereses_x_Obligacion", "Intereses 2026 por obligación y cruce de cuentas 2130/5305"),
               ("Int_Glosas", "Intereses pagados directo al gasto según glosas")]:
    put(ws, r, 2, nm, bold=True)
    put(ws, r, 3, ds)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    r += 1

# -*- coding: utf-8 -*-
# Bloque insertado al final de build_cruce_62.py (antes de guardar): hoja principal "PT_62" con formato de papel de trabajo.
ws = wb.create_sheet("PT_62", 0)
ws.sheet_view.showGridLines = True
for i, w in enumerate([4, 6, 22, 30, 17, 17, 17, 17, 17, 15, 12, 13, 13, 18, 16, 16, 16, 16], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
f_box = Font(name=FN, size=9, bold=True, color="1F3864")


def txt(r, c, v, font=None, merge_to=None, h=None, wrap=True, border=False, fill=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or f_base
    cell.alignment = Alignment(wrap_text=wrap, vertical="top")
    if fill:
        cell.fill = fill
    if border:
        cell.border = box
    if merge_to:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
    if h:
        ws.row_dimensions[r].height = h
    return cell


# Encabezado
txt(2, 3, "GUAICARAMO S.A.S.", font=f_title, merge_to=9)
for k, (lab, val) in enumerate([("CLIENTE:", "GUAICARAMO S.A.S."), ("NIT:", "860.040.584-0"),
                                ("PERIODO:", "1-ene-2026 a 31-jul-2026"), ("FECHA DE CORTE:", "=Resumen!C5"),
                                ("NOMBRE P/T:", "OBLIGACIONES FINANCIERAS (cuenta 21 – papel 62/64)"),
                                ("MARCO TÉCNICO:", "NIIF Plenas (Grupo 1): NIIF 9, NIIF 16, NIC 1, NIIF 7, NIC 23, NIC 24"),
                                ("MONEDA:", "Pesos colombianos (COP)")]):
    txt(4 + k, 3, lab, font=f_bold)
    c = txt(4 + k, 4, val, font=f_in if not str(val).startswith("=") else f_link, merge_to=9)
    if k == 3:
        c.number_format = DATE
for c_, lab in ((14, "PREPARADO:"), (15, "REVISADO:"), (16, "Ref. PT")):
    txt(4, c_, lab, font=f_hdr, fill=fill_hdr, border=True)
for c_, val in ((14, ""), (15, ""), (16, "62 / 64")):
    txt(5, c_, val, font=Font(name=FN, size=8, bold=True, color="FF0000"), border=True)
for c_ in (14, 15):
    txt(6, c_, "Fecha:", font=f_sub, border=True)
txt(7, 14, "Iniciales y fecha del preparador / revisor", font=f_sub, merge_to=16)

r = 12
SEC = lambda rr, t_: (txt(rr, 3, t_, font=f_sec, merge_to=16, fill=fill_sec))
SEC(r, "CÓMO LEER ESTE PAPEL (guía rápida)")
for gl in ["1. Esta hoja resume TODO el trabajo: qué se revisó, cómo y qué se encontró. Las hojas de soporte están enlazadas al final (clic en el nombre).",
           "2. Celdas BLANCAS = información del cliente, del balance o del banco.   Celdas AMARILLAS = cálculos y recálculos hechos por auditoría.",
           "3. 'OK' = la cifra cuadra.   'DIFERENCIA' = hay que revisar (la explicación está en la hoja 'Hallazgos').   [P] = falta un soporte (hoja 'Pendientes').",
           "4. Orden sugerido de lectura: Conclusión (al final) → secciones 1 a 5 → hoja 'Pruebas_Auditoria' → hoja 'Tablas_Amortizacion' → 'Hallazgos'."]:
    r += 1
    txt(r, 3, gl, merge_to=16, h=14)
r += 2
SEC(r, "OBJETIVO GENERAL")
r += 1
txt(r, 3, "Obtener evidencia suficiente y apropiada sobre la razonabilidad de los saldos de obligaciones financieras (créditos y pasivos por "
          "arrendamiento financiero) e intereses por pagar al 31-jul-2026, y del gasto financiero asociado, verificando integridad, existencia, "
          "exactitud, valuación a costo amortizado (método del interés efectivo – NIIF 9 §4.2.1 y B5.4; NIIF 16 §36), corte y clasificación "
          "corriente/no corriente (NIC 1 §69-76), con base en el balance de prueba, el auxiliar contable, el detalle 62, las tablas de "
          "amortización (papel 64) y los extractos/certificados bancarios.", merge_to=16, h=52)
r += 2
SEC(r, "LIMITACIONES AL ALCANCE")
r += 1
txt(r, 3, '="Se emite con "&' + PN_COUNT + '&" pendientes de soporte (hoja Pendientes): certificados de saldo al 31-jul-26 de leasings Davivienda y Bancolombia, confirmaciones BBVA y Banco Popular, '
          'certificado del crédito Bogotá 1159383796, tablas del leasing 386735/GECOLSA y materialidad del encargo. Las conclusiones afectadas se marcan [P]."',
    merge_to=16, h=40)
r += 2
SEC(r, "ASEVERACIONES")
r += 1
for c_, lab in enumerate(["Aseveración", "Procedimiento principal", "Hoja soporte", "Resultado"], start=3):
    txt(r, c_ if c_ < 5 else (5 if c_ == 5 else 9), lab, font=f_hdr, fill=fill_hdr, border=True)
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
ASEV = [
    ("Existencia / Ocurrencia", "Cotejo con certificados y extractos bancarios; pagos certificados vs comprobantes.", "Certificados",
     f'="Cobertura con extracto 2026: "&TEXT({C_COBERT},"0%")&"; diferencias de pagos: "&TEXT(Certificados!J{C_PAGOS_TOT},"#,##0")'),
    ("Integridad", "Movimiento vs balance; detalle 62 vs mayor por obligación; obligaciones del mayor no incluidas en el detalle.", "Saldos_x_Obligacion",
     f'="Mayor vs balance: "&Resumen!F11&". Leasing mayor − detalle 62: "&TEXT(Cruce_Detalle_Mayor!F7,"#,##0")&" (386735 y GECOLSA no están en el detalle)."'),
    ("Exactitud / Valuación", "Recálculo de intereses, de las tablas de amortización y de la causación al corte (NIIF 9 B5.4, NIIF 16 §36).", "Recalculo_Intereses / Rev_Tablas_Amort",
     f'="Causación no registrada al 31-jul: "&TEXT(Recalculo_Intereses!K{P2_TOT},"#,##0")&"; tablas con error: 362095, Hoja2 (1159383796), costo amortizado Davivienda."'),
    ("Corte", "Causación de intereses entre el último pago y el 31-jul; prórrogas y desembolsos cercanos al cierre.", "Recalculo_Intereses",
     f'="AJE-1 por "&TEXT(Recalculo_Intereses!K{P2_TOT},"#,##0")'),
    ("Derechos y obligaciones", "Titularidad de las obligaciones (NIT Guaicaramo) y obligaciones de terceros incluidas en el archivo de tablas.", "Rev_Tablas_Amort",
     '="Crédito DYPSIS $72.300 MM en el archivo de tablas no está en el mayor: confirmar titular/garantías [P]."'),
    ("Clasificación y presentación", "Porción corriente según tablas y vencimientos (NIC 1 §69); agrupación de tarjetas de crédito.", "Clasif_CP_LP",
     f'="Reclasificación LP→CP propuesta: "&TEXT(Clasif_CP_LP!B11+Clasif_CP_LP!B12,"#,##0")'),
    ("Revelación", "Datos para NIIF 7 (vencimientos, tasas variables IBR/DTF) y NIIF 16 (pasivos por arrendamiento).", "Detalle_62 / Rev_Tablas_Amort",
     '="Tasas y vencimientos del detalle 62 difieren de los certificados en varias obligaciones (hallazgos 7 y 14)."'),
]
r += 1
for a_, p_, h_, res_ in ASEV:
    txt(r, 3, a_, font=f_bold, border=True)
    txt(r, 4, p_, border=True)
    txt(r, 5, h_, font=f_link, border=True, merge_to=8)
    txt(r, 9, res_, font=f_base, border=True, merge_to=16)
    ws.row_dimensions[r].height = 36
    r += 1
r += 1
SEC(r, "TÉCNICAS DE AUDITORÍA APLICADAS")
r += 1
txt(r, 3, "☑ Análisis   ☑ Recálculo   ☑ Inspección documental (extractos/certificados)   ☑ Comparación (mayor ↔ detalle ↔ tabla ↔ banco)   ☑ Prueba analítica sustantiva   ☐ Confirmación externa (pendiente)",
    merge_to=16, h=18)
r += 2
SEC(r, "PROCEDIMIENTOS")
r += 1
for c_, lab in ((2, "#"), (3, "Procedimiento"), (9, "Ref. hoja"), (12, "Resultado")):
    txt(r, c_, lab, font=f_hdr, fill=fill_hdr, border=True)
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=11)
ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=16)
PROC = [
    ("Cruzar el auxiliar de la cuenta 21 contra el balance de prueba (débitos, créditos y saldos) y comparar las dos versiones del balance.", "Cruce_Mov_Balance", "=Resumen!F11"),
    ("Reconstruir el saldo del mayor por obligación (detalle 31-dic-25 + movimientos 2026) y cruzarlo con el detalle 62.", "Saldos_x_Obligacion",
     f'=IF(ABS(Saldos_x_Obligacion!H{SX_CRED[1]})+ABS(Saldos_x_Obligacion!H{SX_LEAS[1]})<=Resumen!$C$6,"Por obligación cuadra con balance","Revisar")'),
    ("Cotejar saldos, tasas, vencimientos y pagos con extractos/certificados bancarios.", "Certificados", "=Resumen!F16"),
    ("Revisar y recalcular las tablas de amortización del cliente (papel 64): aritmética, TIR, saldo al corte, porción corriente.", "Rev_Tablas_Amort / A01–A25",
     f'="Tablas con hallazgos: 362095, Hoja2 (1159383796), costo amortizado Davivienda; "&COUNTIF(Rev_Tablas_Amort!R{RT_FIRST}:R{RT_LAST},"N/D")&" sin saldo banco"'),
    ("Recalcular intereses liquidados por los bancos y la causación al 31-jul (costo amortizado).", "Recalculo_Intereses / Intereses_x_Obligacion", "=Resumen!F18"),
    ("Cruzar las cuentas de gasto financiero (5305) y de intereses por pagar (2130) contra la suma por obligación.", "Intereses_x_Obligacion",
     f'=IF(ABS(Intereses_x_Obligacion!E{IXB_FIRST})<=Resumen!$C$6,"2130 cuadra por obligación","Revisar")'),
    ("Prueba analítica sustantiva del gasto por intereses.", "Gasto_Intereses", "=Gasto_Intereses!B20"),
    ("Evaluar la clasificación corriente / no corriente y la presentación.", "Clasif_CP_LP", '="Ver reclasificación AJE-2"'),
    ("Concluir y documentar hallazgos, ajustes y pendientes.", "Hallazgos / Pendientes", '=' + PN_COUNT + '&" pendientes abiertos"'),
]
r += 1
for n, (p_, h_, res_) in enumerate(PROC, start=1):
    txt(r, 2, n, border=True)
    txt(r, 3, p_, border=True, merge_to=8)
    txt(r, 9, h_, font=f_link, border=True, merge_to=11)
    txt(r, 12, res_, font=f_bold, border=True, merge_to=16)
    ws.row_dimensions[r].height = 28
    r += 1
r += 1
SEC(r, "FUENTES DE INFORMACIÓN")
r += 1
for f_ in ["[A] Balance de prueba NIIF a jul-2026 (Balance_2026_GUUU.xlsx – 24-ago-26 – y Obl_Financieros.xlsx/Hoja1 – 29-sep-26).",
           "[A] Auxiliar cuenta 21 ene–jul 2026 (Obl_Financieros.xlsx/Sheet1).",
           "[A] Detalle de obligaciones financieras a jul-26 (62._Obligaciones_Financieras.xlsx).",
           "[A] Tablas de amortización (64._Tabla_amortización_obligaciones_financieras.xlsm).",
           "[A] PT Obligaciones Financieras al 31-dic-2025 (ejemplo.xlsx): detalle del mayor por obligación y saldos de extractos a dic-25.",
           "[C] Extractos/certificados bancarios (ilovepdf_merged_1.pdf; LEASING_1275-3/9441-2/9448-3_DAVIVIENDA.pdf)."]:
    txt(r, 3, f_, merge_to=16)
    r += 1
r += 1

# ******* 1 ******* Prueba cuenta vs detalle
SEC(r, "******* 1 *******   Prueba de auditoría – Cuentas del balance vs detalle 62 vs saldo por obligación")
r += 1
hdr1 = ["CUENTA", "DESCRIPCIÓN", "SALDO BALANCE 31/07/2026", "SALDO DETALLE 62", "SALDO POR OBLIGACIÓN (mayor)", "DIF. BALANCE − DETALLE 62",
        "DIF. BALANCE − POR OBLIGACIÓN", "MARCA"]
for k, lab in enumerate(hdr1):
    txt(r, 3 + k, lab, font=f_hdr, fill=fill_hdr, border=True)
ws.row_dimensions[r].height = 30
r += 1
P1_ROWS = [
    ("210501 + 210502", "Créditos (sin tarjetas de crédito)", f"=-({BAL('21050100', 'SF')}+{BAL('21050201', 'SF')})",
     f'=SUMIFS({DR("F")},{DR("D")},"Crédito")', f"=Saldos_x_Obligacion!H{SX_CRED[0]}"),
    ("212020 + 212021", "Leasing financiero (NIIF 16)", f"=-({BAL('21202040', 'SF')}+{BAL('21202140', 'SF')})",
     f'=SUMIFS({DR("F")},{DR("D")},"Leasing")', f"=Saldos_x_Obligacion!H{SX_LEAS[0]}"),
    ("213001", "Intereses financieros por pagar", f"=-{BAL('2130', 'SF')}", "No incluido", f"=Intereses_x_Obligacion!F{IX_TOT}"),
    ("21050301", "Tarjetas de crédito por pagar", f"=-{BAL('21050301', 'SF')}", "No incluido", "N/A"),
]
p1_first = r
for cta, des, a_, b_, c__ in P1_ROWS:
    txt(r, 3, cta, font=f_in, border=True)
    txt(r, 4, des, border=True)
    for k, v in enumerate((a_, b_, c__)):
        cell = txt(r, 5 + k, v, border=True)
        cell.number_format = NUM
        cell.font = f_link if str(v).startswith("=") and "!" in str(v) else f_base
    for k, (x, y) in enumerate((("E", "F"), ("E", "G"))):
        cell = txt(r, 8 + k, f'=IF(AND(ISNUMBER({x}{r}),ISNUMBER({y}{r})),{x}{r}-{y}{r},"N/A")', border=True)
        cell.number_format = NUM
    txt(r, 10, "[D] [A]", font=f_sub, border=True)
    r += 1
txt(r, 4, "TOTAL", font=f_bold, border=True, fill=fill_tot)
for c_ in (5, 6, 7, 8, 9):
    L = get_column_letter(c_)
    cell = txt(r, c_, f"=SUM({L}{p1_first}:{L}{r - 1})", font=f_bold, border=True, fill=fill_tot)
    cell.number_format = NUM
r += 2

# ******* 2 ******* Validación por terceros y número de obligación
SEC(r, "******* 2 *******   Validación por terceros y número de obligación financiera (saldo mayor vs detalle vs tabla vs extracto)")
r += 1
hdr2 = ["N° obligación (ID)", "Descripción", "Saldo mayor 31-jul", "Saldo detalle 62", "Saldo tabla amort.", "Saldo extracto/banco",
        "Dif. mayor − banco", "Tipo tasa", "Frecuencia", "Fecha venc.", "Tasa EA banco", "Intereses 2026 registrados",
        "Int. por pagar 2130", "Causación recalc. 31-jul", "Marca"]
for bank_, ids in SX_IDS:
    txt(r, 3, bank_.title(), font=f_box)
    r += 1
    for k, lab in enumerate(hdr2):
        txt(r, 3 + k, lab, font=f_hdr, fill=fill_hdr, border=True)
    ws.row_dimensions[r].height = 30
    r += 1
    b_first = r
    for ident in ids:
        sx = SX_ROW[ident]
        txt(r, 3, ident, font=f_bold, border=True)
        txt(r, 4, f"=Saldos_x_Obligacion!D{sx}", font=f_link, border=True)
        for k, col in enumerate(("H", "I", "J", "L", "O")):
            cell = txt(r, 5 + k, f"=Saldos_x_Obligacion!{col}{sx}", font=f_link, border=True)
            cell.number_format = NUM
        txt(r, 10, f'=IFERROR(INDEX({DR("H")},MATCH(C{r},{DR("C")},0)),"—")', font=f_link, border=True)
        txt(r, 11, f'=IFERROR(INDEX({DR("L")},MATCH(C{r},{DR("C")},0)),"—")', font=f_link, border=True)
        cell = txt(r, 12, f'=IFERROR(INDEX({DR("P")},MATCH(C{r},{DR("C")},0)),"—")', font=f_link, border=True)
        cell.number_format = DATE
        cell = txt(r, 13, f'=IFERROR(INDEX(Certificados!$M:$M,MATCH(C{r},Certificados!$E:$E,0)),"—")', font=f_link, border=True)
        cell.number_format = PCT
        for k, (sheet_, col) in enumerate((("Intereses_x_Obligacion", "I"), ("Intereses_x_Obligacion", "F"), ("Intereses_x_Obligacion", "L"))):
            cell = txt(r, 14 + k, f"=INDEX({sheet_}!${col}:${col},MATCH(C{r},{sheet_}!$A:$A,0))", font=f_link, border=True)
            cell.number_format = NUM
        txt(r, 17, f'=IF(LEFT(Saldos_x_Obligacion!M{sx},3)="[C]","[C] [B]",IF(LEFT(Saldos_x_Obligacion!M{sx},3)="[E]","[E] [B]","[P]"))', border=True)
        r += 1
    txt(r, 4, "Subtotal", font=f_bold, border=True, fill=fill_tot)
    for c_ in (5, 6, 7, 8, 9, 14, 15, 16):
        L = get_column_letter(c_)
        cell = txt(r, c_, f"=SUM({L}{b_first}:{L}{r - 1})", font=f_bold, border=True, fill=fill_tot)
        cell.number_format = NUM
    r += 2

# ******* 3 ******* Cuentas de intereses
SEC(r, "******* 3 *******   Cuentas de intereses del balance vs suma por obligación (causación, pagos y gasto)")
r += 1
for k, lab in enumerate(["Código", "Cuenta contable", "Saldo / mov. balance", "Valor según obligaciones", "Diferencia"]):
    txt(r, 3 + k, lab, font=f_hdr, fill=fill_hdr, border=True)
r += 1
for rr_ in range(IXB_FIRST, IXB_LAST + 1):
    for k, col in enumerate("ABCDE"):
        cell = txt(r, 3 + k, f"=Intereses_x_Obligacion!{col}{rr_}", font=f_link, border=True)
        if k >= 2:
            cell.number_format = NUM
    r += 1
r += 1

# ******* 4 ******* Revisión de tablas
SEC(r, "******* 4 *******   Revisión y recálculo de tablas de amortización (papel 64) – ver Rev_Tablas_Amort y anexos A01–A25")
r += 1
for k, lab in enumerate(["Concepto", "", "Tabla", "Detalle 62", "Banco", "Diferencia"]):
    txt(r, 3 + k, lab, font=f_hdr, fill=fill_hdr, border=True)
r += 1
for lab, a_, b_, c__ in [
    ("Saldo al 31-jul (obligaciones con tabla en el mayor)", f"=Rev_Tablas_Amort!N{RT_TOT}", f"=Rev_Tablas_Amort!O{RT_TOT}", f"=Rev_Tablas_Amort!Q{RT_TOT}"),
    ("Interés ene–jul 2026 (tabla vs registrado)", f"=Rev_Tablas_Amort!S{RT_TOT}", "", f"=Rev_Tablas_Amort!T{RT_TOT}"),
    ("Porción corriente 12 meses (tabla vs detalle)", f"=Rev_Tablas_Amort!V{RT_TOT}", f"=Rev_Tablas_Amort!W{RT_TOT}", ""),
]:
    txt(r, 3, lab, border=True, merge_to=4)
    for k, v in enumerate((a_, b_, c__)):
        cell = txt(r, 5 + k, v, font=f_link, border=True)
        cell.number_format = NUM
    cell = txt(r, 8, f'=IF(ISNUMBER(G{r}),E{r}-G{r},IF(ISNUMBER(F{r}),E{r}-F{r},"N/A"))', border=True)
    cell.number_format = NUM
    r += 1
r += 1

# ******* 5 ******* Ajustes
SEC(r, "******* 5 *******   Ajustes y reclasificaciones propuestos (detalle en hoja Hallazgos)")
r += 1
for k, lab in enumerate(["AJE", "Cuenta", "Descripción", "Débito", "Crédito"]):
    txt(r, 3 + k, lab, font=f_hdr, fill=fill_hdr, border=True)
r += 1
for rr_ in range(A_FIRST_H, A_LAST_H + 1):
    for k, col in enumerate("ABCDE"):
        cell = txt(r, 3 + k, f'=IF(Hallazgos!{col}{rr_}="","",Hallazgos!{col}{rr_})', font=f_link, border=True)
        if k >= 3:
            cell.number_format = NUM
    r += 1
r += 1

# Conclusión
SEC(r, "CONCLUSIÓN")
r += 1
txt(r, 3, '="Con base en los procedimientos aplicados, el auxiliar de la cuenta 21 cuadra con el balance y el saldo reconstruido por obligación explica el 100% del balance. '
          'El detalle 62 del cliente NO es confiable para leasing: proviene de tablas teóricas (tasa fija) y omite el leasing Bancolombia 386735 y el leasing GECOLSA (Davivienda). '
          'Los créditos cuadran con los extractos (redondeos). Se proponen el AJE-1 por "&TEXT(Recalculo_Intereses!K' + str(P2_TOT) + ',"$#,##0")&" de intereses no causados al corte '
          'y el AJE-2 de reclasificación a corto plazo. La conclusión queda sujeta a los "&' + PN_COUNT + '&" pendientes de soporte."',
    merge_to=16, h=70)
r += 2
SEC(r, "MARCAS DE AUDITORÍA")
r += 1
for m_, d_ in [("[A]", "Información suministrada por el cliente"), ("[B]", "Cálculos / recálculos realizados por auditoría"),
               ("[C]", "Cotejado con extracto o certificado bancario 2026"), ("[D]", "Cotejado con balance de prueba / mayor"),
               ("[E]", "Estimado: extracto al 31-dic-25 (PT 2025) + movimientos 2026 – pendiente certificado"), ("[P]", "Pendiente de soporte")]:
    txt(r, 2, m_, font=f_bold)
    txt(r, 3, d_, merge_to=10)
    r += 1
r += 1
SEC(r, "VER SOPORTES / ANEXOS (clic para abrir)")
r += 1
links = ["Resumen", "Pruebas_Auditoria", "Tablas_Amortizacion", "Hallazgos", "Pendientes", "Cruce_Detalle_Mayor",
         "Intereses_x_Obligacion", "Clasif_CP_LP", "Gasto_Intereses", "Detalle_62", "Mayor_Dic25", "Mayor_Mov", "Balance_21", "Int_Glosas"]
for k, nm in enumerate(links):
    cc = 3 + (k % 4) * 3
    rr_ = r + k // 4
    cell = ws.cell(row=rr_, column=cc, value=nm)
    cell.font = Font(name=FN, size=9, color="0563C1", underline="single")
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{nm}'!A1")
ws.freeze_panes = "A4"

# ============================================================================ 12. UNIÓN DE HOJAS (pruebas y tablas)
import importlib.util as _ilu2
_sp = _ilu2.spec_from_file_location("msh", __file__.rsplit("/", 1)[0] + "/merge_sheets.py")
MSH = _ilu2.module_from_spec(_sp)
_sp.loader.exec_module(MSH)
T_ = Font(name=FN, size=11, bold=True, color="002060")
G_ = Font(name=FN, size=8, color="000000")
INTRO_P = [("PRUEBAS DE AUDITORÍA – OBLIGACIONES FINANCIERAS – corte 31-jul-2026", T_),
           ("Cómo leer esta hoja:  celdas BLANCAS = datos tomados del cliente, del balance o del banco;   celdas AMARILLAS = cálculo o recálculo hecho por auditoría.", G_),
           ("'OK' = la cifra cuadra;   'DIFERENCIA' = revisar (explicación en la columna de comentarios o en la hoja Hallazgos).   Cifras en pesos colombianos.", G_),
           ("Contenido:  BLOQUE 1 Movimiento vs balance  ·  BLOQUE 2 Certificados bancarios  ·  BLOQUE 3 Saldo por obligación  ·  BLOQUE 4 Recálculo de intereses", G_)]
BLOCKS_P = [
    ("Cruce_Mov_Balance", "BLOQUE 1 – ¿El auxiliar contable de la cuenta 21 cuadra con el balance?",
     ["Qué se hizo: se sumaron todos los débitos y créditos del auxiliar (ene–jul 2026) y se compararon con el balance de prueba.",
      "Cómo se lee: saldo inicial + débitos − créditos = saldo final calculado (amarillo), que se compara con el saldo final del balance.",
      f'="Resultado: "&Cruce_Mov_Balance!M{CMB_TOT}&"  –  diferencia total en saldos: "&TEXT(Cruce_Mov_Balance!H{CMB_TOT},"#,##0")']),
    ("Certificados", "BLOQUE 2 – ¿Los saldos del detalle 62 coinciden con lo que dice el banco?",
     ["Qué se hizo: se tomó cada extracto o certificado bancario y se comparó capital, tasa y vencimiento con el detalle 62 del cliente.",
      "Capital banco 31-jul (amarillo) = saldo del certificado − abonos a capital pagados después de la fecha del certificado.  Sección C: cada pago del banco vs el comprobante contable.",
      f'="Resultado: diferencia detalle 62 − banco "&TEXT(Certificados!L{C_TOT},"#,##0")&";  saldo cubierto con certificados: "&TEXT({C_COBERT},"0%")']),
    ("Saldos_x_Obligacion", "BLOQUE 3 – Saldo de cada obligación: contabilidad vs detalle 62 vs tabla vs banco",
     ["Qué se hizo: saldo contable de cada crédito/leasing = saldo al 31-dic-25 (papel 2025) + nuevos préstamos − pagos 2026 (amarillo).",
      "Luego se compara con el detalle 62, con la tabla de amortización del cliente y con el banco ([C] certificado 2026; [E] estimado con el extracto de dic-25).",
      f'="Resultado: la suma por obligación cuadra con el balance (diferencia "&TEXT(Saldos_x_Obligacion!H{SX_CRED[1]}+Saldos_x_Obligacion!H{SX_LEAS[1]},"#,##0")&"); el detalle 62 está desactualizado en leasing."']),
    ("Recalculo_Intereses", "BLOQUE 4 – Recálculo de intereses",
     ["Fórmula: interés = capital × ((1 + tasa EA)^(días / base) − 1), con base 365 o 360 según el banco.",
      "Parte 1: ¿el banco cobró bien?  ·  Parte 2: intereses desde el último pago hasta el 31-jul que no están contabilizados (AJE-1)  ·  Parte 3: saldo implícito de leasings.",
      f'="Resultado: intereses no causados al 31-jul = "&TEXT(Recalculo_Intereses!K{P2_TOT},"$#,##0")']),
]
INTRO_T = [("REVISIÓN DE LAS TABLAS DE AMORTIZACIÓN DEL CLIENTE (papel 64) – 25 tablas", T_),
           ("Cómo leer esta hoja: primero el RESUMEN de las 25 tablas; debajo, cada tabla completa (usa los enlaces de la columna 'Anexo' del resumen).", G_),
           ("En cada tabla: columnas A–G BLANCAS = tabla tal como la entregó el cliente;   columnas H–P AMARILLAS = recálculo de auditoría.", G_),
           ("Si 'Dif. saldo' = 0 la tabla está bien calculada.   'Tabla − banco' muestra cuánto se aleja la tabla de lo que realmente cobra el banco.", G_)]
BLOCKS_T = [("Rev_Tablas_Amort", "RESUMEN DE LA REVISIÓN DE TABLAS",
             ["Cada fila es una tabla. La última columna ('Observación de auditoría') explica en palabras lo encontrado."])]
for ident in ANX_ORDER:
    t = TAB[ident]
    BLOCKS_T.append((ANX[ident]["sh"], f"TABLA {ANX[ident]['n']:02d} – {ident} – {t.get('banco')} {t.get('contrato')}",
                     ["Observación: " + OBS_T.get(ident, "")]))
mapP, stP, wsP = MSH.merge(wb, "Pruebas_Auditoria", BLOCKS_P, YELLOW, intro=INTRO_P)
mapT, stT, wsT = MSH.merge(wb, "Tablas_Amortizacion", BLOCKS_T, YELLOW, intro=INTRO_T)
MAP = {**mapP, **mapT}
MSH.rewrite_workbook(wb, MAP, removed=[b[0] for b in BLOCKS_P] + [b[0] for b in BLOCKS_T])
wsP.freeze_panes = None
wsT.freeze_panes = None
TXT_REP = [("Rev_Tablas_Amort / A01–A25", "Tablas_Amortizacion"), ("Rev_Tablas_Amort", "Tablas_Amortizacion (resumen)"),
           ("anexos A01–A25", "hoja Tablas_Amortizacion"), ("A01–A25", "Tablas_Amortizacion"), ("A01-A25", "Tablas_Amortizacion"),
           ("Saldos_x_Obligacion", "Pruebas_Auditoria (bloque 3)"), ("Recalculo_Intereses", "Pruebas_Auditoria (bloque 4)"),
           ("Cruce_Mov_Balance", "Pruebas_Auditoria (bloque 1)"), ("Hoja Certificados", "Pruebas_Auditoria (bloque 2)"),
           ("Certificados sección", "Pruebas_Auditoria bloque 2, sección")]
for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and not v.startswith("="):
                if v == "Certificados":
                    c.value = "Pruebas_Auditoria (bloque 2)"
                    continue
                for a_, b_ in TXT_REP:
                    v = v.replace(a_, b_)
                c.value = v
            elif isinstance(v, str) and v.startswith("=") and w.title == "PT_62" and '"' in v:
                for a_, b_ in TXT_REP:
                    v = v.replace(a_, b_)
                c.value = v
ORDER = ["PT_62", "Resumen", "Pruebas_Auditoria", "Tablas_Amortizacion", "Hallazgos", "Pendientes", "Cruce_Detalle_Mayor",
         "Intereses_x_Obligacion", "Clasif_CP_LP", "Gasto_Intereses", "Detalle_62", "Mayor_Dic25", "Mayor_Mov", "Balance_21", "Int_Glosas"]
wb._sheets = [wb[n] for n in ORDER] + [w for w in wb.worksheets if w.title not in ORDER]
wb.active = 0
for w in wb.worksheets:
    w.sheet_properties.tabColor = None
    w.page_setup.orientation = "landscape"
    w.sheet_properties.pageSetUpPr.fitToPage = True
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
wb.save(OUT)
print("OK", OUT, "mov rows", len(mov_rows), "det rows", len(det))
