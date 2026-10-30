"""Paso 2: inventario de celdas de plantillas (ST o FINREP/COREP en layout de anexo).

Para cada celda de la rejilla: plantilla, fila, columna, etiquetas de fila/columna y,
si la celda contiene una fórmula, la fórmula. Las fórmulas sencillas (sumas y restas
de celdas, ``SUM(rango)``) se traducen **sin LLM** a reglas de tipo ``plantilla``:
son controles internos ya codificados por la EBA en la propia plantilla.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

from ..loaders.excel_annex import DEFAULT_CODE_RE, _detect
from ..model import DatapointRef

_CELL_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_SUM_RE = re.compile(r"SUM\(\s*\$?([A-Z]{1,3}\$?\d+)\s*:\s*\$?([A-Z]{1,3}\$?\d+)\s*\)", re.I)


def _label(ws, row: int, upto_col: int) -> str | None:
    for c in range(1, upto_col):
        v = ws.cell(row, c).value
        if isinstance(v, str) and v.strip() and not re.match(DEFAULT_CODE_RE, v.strip()):
            return v.strip()
    return None


def _col_label(ws, col: int, header_row: int) -> str | None:
    for r in range(header_row - 1, 0, -1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip():
            return v.strip()
    return None


def build_inventory(xlsx: str | Path, framework: str, sheet_map: dict | None = None) -> dict:
    xlsx = Path(xlsx)
    wb = load_workbook(xlsx, data_only=False)
    code_re = re.compile(DEFAULT_CODE_RE)
    cells, formulas, issues = [], [], []
    for ws in wb.worksheets:
        conf = (sheet_map or {}).get(ws.title, {})
        tpl, z = conf.get("template", ws.title), conf.get("z")
        row_codes, col_codes, warns = _detect(ws, code_re)
        issues += [f"{ws.title}: {w}" for w in warns]
        if not row_codes or not col_codes:
            continue
        header_row = min(row_codes) - 1
        addr2ref = {}
        for r, rc in row_codes.items():
            for c, cc in col_codes.items():
                ref = DatapointRef(framework, tpl, rc, cc, z)
                addr = f"{get_column_letter(c)}{r}"
                addr2ref[addr] = ref
                v = ws.cell(r, c).value
                cells.append({
                    "ref": str(ref), "sheet": ws.title, "cell": addr,
                    "row_label": _label(ws, r, min(col_codes)),
                    "col_label": _col_label(ws, c, header_row),
                    "formula": v if isinstance(v, str) and v.startswith("=") else None,
                })
        for item in cells:
            if item["sheet"] == ws.title and item["formula"]:
                rule = formula_to_rule(item["ref"], item["formula"], addr2ref)
                if rule:
                    formulas.append({**rule, "_celda": f"{ws.title}!{item['cell']}"})
                else:
                    issues.append(f"{ws.title}!{item['cell']}: fórmula no traducible {item['formula']!r}")
    return {"file": str(xlsx), "framework": framework, "cells": cells,
            "template_rules": formulas, "issues": issues}


_TOK_RE = re.compile(r"\s*([+-]|\(|\)|[A-Z]{1,3}\d+)")


def formula_to_rule(target_ref: str, formula: str, addr2ref: dict) -> dict | None:
    """``=D5+D6-SUM(D7:D9)`` -> regla ``target = expresión``.

    Devuelve None si la fórmula no es una combinación lineal de celdas de la rejilla
    (productos, funciones distintas de SUM, referencias a otras hojas...). Esas quedan
    como incidencias para revisión manual o para el agente.
    """
    body = formula.lstrip("=").replace("$", "").upper()

    def expand_sum(m):
        c1, r1, c2, r2 = range_boundaries(f"{m.group(1)}:{m.group(2)}")
        return "(" + "+".join(f"{get_column_letter(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)) + ")"

    body = _SUM_RE.sub(expand_sum, body)
    toks, pos = [], 0
    while pos < len(body):
        m = _TOK_RE.match(body, pos)
        if not m:
            return None
        toks.append(m.group(1))
        pos = m.end()

    coefs: dict[str, float] = {}
    stack, sign, expect_operand = [1.0], 1.0, True
    for t in toks:
        if t in "+-":
            sign = sign * (-1.0 if t == "-" else 1.0)
            expect_operand = True
        elif t == "(":
            stack.append(stack[-1] * sign)
            sign, expect_operand = 1.0, True
        elif t == ")":
            if len(stack) == 1:
                return None
            stack.pop()
            sign, expect_operand = 1.0, False
        else:
            if not expect_operand or t not in addr2ref:
                return None
            key = str(addr2ref[t])
            coefs[key] = coefs.get(key, 0.0) + stack[-1] * sign
            sign, expect_operand = 1.0, False
    if len(stack) != 1 or not coefs:
        return None

    parts = []
    for ref, k in coefs.items():
        if k == 0:
            continue
        mag = "" if abs(k) == 1 else f"{abs(k):g}*"
        parts.append(f"{'-' if k < 0 else '+'} {mag}{{{ref}}}")
    rhs = " ".join(parts).removeprefix("+ ").strip()
    t = DatapointRef.parse(target_ref)
    return {
        "id": f"TPL-{t.template}-r{t.row}c{t.col}".replace(" ", ""),
        "descripcion": f"Control interno de la plantilla: {target_ref} {formula}",
        "tipo": "igualdad", "lhs": "{" + target_ref + "}", "rhs": rhs, "origen": "plantilla",
        "fuente": {"documento": "plantilla", "formula": formula},
    }


def save_inventory(inv: dict, out: str | Path) -> None:
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text(json.dumps(inv, ensure_ascii=False, indent=1), encoding="utf-8")


def known_refs(inventories: list[dict]) -> set[str]:
    return {c["ref"] for inv in inventories for c in inv["cells"]}
