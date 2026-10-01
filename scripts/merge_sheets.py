# -*- coding: utf-8 -*-
"""Une varias hojas de un libro openpyxl en una sola (bloques apilados verticalmente),
reescribiendo todas las referencias de fórmulas e hipervínculos del libro."""
import copy
import re

from openpyxl.comments import Comment
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

CELL_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
COL_RE = re.compile(r"^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$")


def _qs(name):
    return f"'{name}'" if re.search(r"[^A-Za-z0-9_]", name) else name


def _split(ref, cur):
    if "!" in ref:
        sh, r = ref.rsplit("!", 1)
        return sh.strip("'"), r, True
    return cur, ref, False


def _map_ref(ref, cur_sheet, mapping):
    sh, r, explicit = _split(ref, cur_sheet)
    if sh not in mapping:
        if explicit and sh in RENAMED:
            return f"{_qs(RENAMED[sh])}!{r}"
        return ref
    tgt, off, first, last = mapping[sh]
    m = COL_RE.match(r)
    if m:
        r2 = f"{m.group(1)}{m.group(2)}${first}:{m.group(3)}{m.group(4)}${last}"
    else:
        r2 = CELL_RE.sub(lambda mm: f"{mm.group(1)}{mm.group(2)}{mm.group(3)}{int(mm.group(4)) + off}", r)
    return r2 if tgt == NEWHOME.get(cur_sheet, cur_sheet) else f"{_qs(tgt)}!{r2}"


RENAMED = {}
NEWHOME = {}
SRCMAP = {}  # (hoja destino, fila, columna) -> hoja origen


def translate(formula, cur_sheet, mapping):
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    tok = Tokenizer(formula)
    out = []
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            out.append(_map_ref(t.value, cur_sheet, mapping))
        else:
            out.append(t.value)
    return "=" + "".join(out)


def merge(wb, target, blocks, yellow, intro=None, gap=2):
    """blocks: lista de (nombre_hoja, titulo_bloque, [lineas explicativas]).
    Devuelve dict hoja -> fila inicial en la hoja destino."""
    ws_t = wb.create_sheet(target)
    ws_t.sheet_view.showGridLines = True
    r = 1
    if intro:
        for k, (txt, fnt) in enumerate(intro):
            c = ws_t.cell(row=r, column=1, value=txt)
            c.font = fnt
            c.alignment = Alignment(wrap_text=False, vertical="top")
            r += 1
        r += 1
    mapping, starts, widths = {}, {}, {}
    plan = []
    for name, title, lines in blocks:
        ws = wb[name]
        hdr_rows = []
        c = ws_t.cell(row=r, column=1, value=title)
        c.font = Font(name="Arial Narrow", size=10, bold=True, color="FF0000")
        hdr_rows.append(r)
        r += 1
        for ln in lines:
            c = ws_t.cell(row=r, column=1, value=ln)
            c.font = Font(name="Arial Narrow", size=8, italic=not ln.startswith("="), bold=ln.startswith("="), color="002060")
            r += 1
        off = r - 1
        mapping[name] = (target, off, 1 + off, ws.max_row + off)
        starts[name] = r
        NEWHOME[name] = target
        plan.append((ws, off))
        for col, dim in ws.column_dimensions.items():
            if dim.width:
                widths[col] = max(widths.get(col, 0), dim.width)
        r = ws.max_row + off + 1 + gap
    # copiar celdas
    for ws, off in plan:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and not c.has_style:
                    continue
                d = ws_t.cell(row=c.row + off, column=c.column)
                d.value = c.value
                if c.has_style:
                    d.font = copy.copy(c.font)
                    d.fill = copy.copy(c.fill)
                    d.border = copy.copy(c.border)
                    d.alignment = copy.copy(c.alignment)
                    d.number_format = c.number_format
                if c.comment:
                    d.comment = Comment(c.comment.text, c.comment.author or "Auditoría")
                if c.hyperlink:
                    d.hyperlink = copy.copy(c.hyperlink)
                SRCMAP[(target, d.row, d.column)] = ws.title
        for rng in ws.merged_cells.ranges:
            ws_t.merge_cells(start_row=rng.min_row + off, start_column=rng.min_col, end_row=rng.max_row + off, end_column=rng.max_col)
        for rr, dim in ws.row_dimensions.items():
            if dim.height:
                ws_t.row_dimensions[rr + off].height = dim.height
    for col, w in widths.items():
        ws_t.column_dimensions[col].width = w
    ws_t.column_dimensions["A"].width = max(widths.get("A", 10), 14)
    # yellow para fórmulas (cálculo de auditoría) en la hoja destino
    for row in ws_t.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and (target, c.row, c.column) in SRCMAP:
                cur = c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None
                if cur not in ("FF002060",):
                    c.fill = PatternFill("solid", fgColor=yellow)
    return mapping, starts, ws_t


def rewrite_workbook(wb, mapping, removed):
    """Reescribe fórmulas e hipervínculos de todo el libro y elimina las hojas origen."""
    for ws in wb.worksheets:
        if ws.title in removed:
            continue
        for row in ws.iter_rows():
            for c in row:
                src = SRCMAP.get((ws.title, c.row, c.column), ws.title)
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = translate(c.value, src, mapping)
                if c.hyperlink is not None:
                    loc = c.hyperlink.location or (c.hyperlink.target[1:] if isinstance(c.hyperlink.target, str) and c.hyperlink.target.startswith("#") else None)
                    if loc:
                        new = _map_ref(loc if "!" in loc else f"{_qs(src)}!{loc}", ws.title, mapping)
                        if "!" not in new:
                            new = f"{_qs(ws.title)}!{new}"
                        c.hyperlink = Hyperlink(ref=c.coordinate, location=new)
    for name in removed:
        del wb[name]
