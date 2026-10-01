# -*- coding: utf-8 -*-
"""Extrae las tablas de amortización del archivo del cliente (64._Tabla_amortización...xlsm)
a un JSON estructurado para el papel de trabajo."""
import datetime as dt
import json
import sys

import openpyxl


# hoja -> (ID auditoría, tipo)
SHEETS = {
    "9441": ("DAV-L1019441", "leasing"), "9448": ("DAV-L1019448", "leasing"), "2753": ("DAV-L1021275", "leasing"),
    "3934": ("DAV-L1013934", "leasing"),
    "5005": ("BOG-L557285005", "leasing"), "9224": ("BOG-L556449224", "leasing"), "9180": ("BOG-L556449180", "leasing"),
    "9117": ("BOG-L556449117", "leasing"), "9064": ("BOG-L556449064", "leasing"), "9299": ("BOG-L556449288", "leasing"),
    "362095": ("BCL-L362095", "leasing"), "336817": ("BCL-L336817", "leasing"), "331330": ("BCL-L331330", "leasing"),
    "D-0624": ("DAV-830624", "credito"), "D-7691": ("DAV-607691", "credito"), "D-9032": ("DAV-569032", "credito"),
    "B-6746": ("BOG-453436746", "credito"), "B-0629": ("BOG-856670629", "credito"), "B-7259": ("BOG-858177259", "credito"),
    "B-7469": ("BOG-1155297469", "credito"), "Hoja2": ("BOG-1159383796", "credito"),
    "P-9940": ("POP-631309940", "credito"), "V-7262": ("BBVA-9600277262", "credito"), "BC-5867": ("BCL-1260105867", "credito"),
    "DD-0709": ("DYP-482800190700", "credito"),
}


def d2s(x):
    return x.strftime("%Y-%m-%d") if isinstance(x, (dt.datetime, dt.date)) else x


def parse(src):
  wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
  res = {}
  for sh, (ident, tipo) in SHEETS.items():
      rows = [list(r) for r in wb[sh].iter_rows(values_only=True, max_col=45)]
      get = lambda r, c: rows[r - 1][c - 1] if r - 1 < len(rows) and c - 1 < len(rows[r - 1]) else None
      t = {"hoja": sh, "id": ident, "tipo": tipo, "filas": []}
      if tipo == "leasing":
          t["banco"] = get(11, 2)
          t["contrato"] = str(get(12, 2))
          p = {}
          for r in range(14, 30):
              k = get(r, 2)
              if isinstance(k, str):
                  p[k.strip().rstrip(":")] = d2s(get(r, 3))
                  if k.strip().startswith("Plazo") or k.strip().startswith("Tasa de inter") or k.strip().startswith("Cuota"):
                      p[k.strip().rstrip(":") + " (unidad)"] = get(r, 4)
          t["param"] = p
          # TIR y tabla de causación (bloque 'Meses','Saldo inicial' en columnas B..G)
          hdr = None
          for r in range(1, len(rows) + 1):
              if get(r, 2) == "TIR":
                  t["tir"] = get(r, 3)
                  t["tir_unidad"] = get(r, 4)
              if get(r, 2) == "Meses" and get(r, 3) == "Saldo inicial":
                  hdr = r
          for r in range(hdr + 1, len(rows) + 1):
              per, si, it, pg, sf, fe = (get(r, c) for c in range(2, 8))
              if not isinstance(per, (int, float)) or not isinstance(si, (int, float)):
                  break
              cap = si - sf  # capital amortizado (negativo = capitalización de intereses)
              t["filas"].append({"per": per, "fecha": d2s(fe), "saldo_ini": si, "interes": it or 0, "pago": cap + (it or 0),
                                 "capital": cap, "saldo_fin": sf})
      else:
          t["banco"] = get(6, 7)
          t["contrato"] = str(get(5, 7))
          t["param"] = {"Tasa": get(8, 7), "Liquidación": get(9, 7), "Fecha inicio": d2s(get(11, 7)),
                        "Fecha terminación": d2s(get(12, 7)), "Capital": get(16, 7), "Plazo meses": get(15, 7),
                        "Tiempo de gracia": get(14, 7), "Valor cuota K": get(9, 10)}
          t["tir"] = None
          for r in range(26, len(rows) + 1):
              mes, fe = get(r, 2), get(r, 3)
              if not isinstance(mes, (int, float)) or not isinstance(fe, (dt.datetime, dt.date)):
                  if t["filas"]:
                      break
                  continue
              t["filas"].append({"per": mes, "fecha": d2s(fe), "ibr": get(r, 7), "puntos": get(r, 6), "tasa": get(r, 8),
                                 "interes": get(r, 9) or 0, "capital": get(r, 10) or 0, "cuota": get(r, 11) or 0,
                                 "saldo_fin": get(r, 12), "ca_saldo": get(r, 23)})
          for i, f in enumerate(t["filas"]):
              f["saldo_ini"] = t["filas"][i - 1]["saldo_fin"] if i else t["param"]["Capital"]
      res[ident] = t
  return res


if __name__ == "__main__":
    res = parse(sys.argv[1])
    json.dump(res, open(sys.argv[2], "w"), ensure_ascii=False, indent=1, default=str)
    CORTE = "2026-07-31"
    for ident, t in res.items():
        f = [x for x in t["filas"] if x["fecha"] and x["fecha"] <= CORTE]
        last = f[-1] if f else None
        print(ident, t["hoja"], len(t["filas"]), last["fecha"] if last else None, last["saldo_fin"] if last else None, t.get("tir"))
