"""Excel con layout de anexo EBA: códigos de fila en una columna y de columna en una fila.

Es el formato de las plantillas de los anexos ITS (FINREP/COREP) y de la mayoría de
exports de herramientas de reporting. El detector:

1. Busca celdas *de texto* con forma de código (``0010``, ``r0010``, ``c0010``). Los
   números no cuentan: un importe ``10`` no debe confundirse con el código ``0010``.
2. Si los códigos llevan prefijo ``r``/``c`` se usa el prefijo. Si no, la columna con más
   códigos es la de filas, y la fila con más códigos (fuera de esa columna) es la cabecera.
3. Lee los valores numéricos en las intersecciones.

Opciones (``options`` en la fuente):

- ``sheets``: ``{nombre_hoja: {template: "F 01.01", z: "ES"}}``. Sin esta opción la
  plantilla es el nombre de la hoja.
- ``ignore_sheets``: lista de hojas a omitir.
- ``code_regex``: regex alternativa para los códigos.
"""
from __future__ import annotations

import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ..model import DatapointRef, Fact, Source
from .base import CoverageEntry, LoadResult, SourceSpec

DEFAULT_CODE_RE = r"^[rRcC]?\d{3,4}$"


def _detect(ws, code_re: re.Pattern):
    codes = []  # (row, col, text)
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and code_re.match(c.value.strip()):
                codes.append((c.row, c.column, c.value.strip()))
    if not codes:
        return None, None, ["sin códigos de fila/columna detectados"]

    warnings: list[str] = []
    prefixed_r = [x for x in codes if x[2][0] in "rR"]
    prefixed_c = [x for x in codes if x[2][0] in "cC"]
    if prefixed_r and prefixed_c:
        row_codes = {r: t for r, _, t in prefixed_r}
        col_codes = {c: t for _, c, t in prefixed_c}
        return row_codes, col_codes, warnings

    # Par (columna de códigos de fila, fila de cabecera) que explica más códigos:
    # filas debajo de la cabecera en esa columna + columnas a su derecha en esa fila.
    best, best_score = None, 0
    for col in {c for _, c, _ in codes}:
        for hrow in {r for r, _, _ in codes}:
            score = sum(1 for r, c, _ in codes if (c == col and r > hrow) or (r == hrow and c > col))
            key = (score, -hrow, -col)  # desempate: cabecera más arriba, columna más a la izquierda
            if best is None or key > best_score:
                best, best_score = (col, hrow), key
    row_code_col, header_row = best
    if best_score[0] == 0:
        return None, None, ["no se encontró cabecera de columnas con códigos de fila debajo"]
    col_codes = {c: t for r, c, t in codes if r == header_row and c > row_code_col}
    row_codes = {r: t for r, c, t in codes if c == row_code_col and r > header_row}
    if not col_codes or not row_codes:
        return None, None, ["falta la cabecera de columnas o la columna de códigos de fila"]
    stray = len(codes) - len(col_codes) - len(row_codes)
    if stray:
        warnings.append(f"{stray} celdas con forma de código fuera de la cabecera/columna de filas (ignoradas)")
    if len(set(col_codes.values())) != len(col_codes):
        warnings.append("códigos de columna repetidos en la cabecera")
    if len(set(row_codes.values())) != len(row_codes):
        warnings.append("códigos de fila repetidos")
    return row_codes, col_codes, warnings


def load(spec: SourceSpec) -> LoadResult:
    opts = spec.options or {}
    code_re = re.compile(opts.get("code_regex", DEFAULT_CODE_RE))
    sheet_map = opts.get("sheets") or {}
    ignore = set(opts.get("ignore_sheets") or [])
    wb = load_workbook(spec.path, data_only=True, read_only=False)
    out = LoadResult()
    for ws in wb.worksheets:
        if ws.title in ignore or (sheet_map and ws.title not in sheet_map):
            continue
        conf = sheet_map.get(ws.title, {})
        template = conf.get("template", ws.title)
        z = conf.get("z")
        row_codes, col_codes, warnings = _detect(ws, code_re)
        n = 0
        if row_codes and col_codes:
            for r, rcode in row_codes.items():
                for c, ccode in col_codes.items():
                    v = ws.cell(row=r, column=c).value
                    if v is None or isinstance(v, str) and not v.strip():
                        continue
                    if isinstance(v, bool) or not isinstance(v, (int, float)):
                        warnings.append(f"{get_column_letter(c)}{r}: valor no numérico {v!r} (ignorado)")
                        continue
                    ref = DatapointRef(spec.framework, template, rcode, ccode, z)
                    cell = f"{get_column_letter(c)}{r}"
                    out.facts.append(Fact(
                        ref=ref,
                        value=float(v) * spec.scale,
                        raw_value=float(v),
                        scale=spec.scale,
                        rounding_eur=spec.rounding_eur,
                        source=Source(str(spec.path), ws.title, cell),
                    ))
                    n += 1
        out.coverage.append(CoverageEntry(str(spec.path), ws.title, template, z,
                                          len(row_codes or {}), len(col_codes or {}), n, warnings))
    return out
