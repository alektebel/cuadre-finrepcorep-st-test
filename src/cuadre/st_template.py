"""Plantillas oficiales del EU-wide stress test (formato EBA 2025/2027).

Layout observado en ``2027 EU-wide stress test - Draft Templates.xlsx`` (11/06/2026):

- Una columna con cabecera ``RowNum``: número de fila de la plantilla (entero).
- Encima de la cabecera, una fila con los números de columna (1, 2, 3...).
- Columna ``COREP CODE / COMPUTATION`` con la celda COREP de cada fila:
    * celda completa: ``C 01.00 (r0030,c0010)``, sumas/restas y ``* 2.5``;
    * solo fila: ``C 09.01, r0010`` (o ``C 09.02, r0010 or C 09.05, r0010``), con la
      columna COREP en una fila de cabecera encima de cada columna (``C 09.01, c0075``);
    * ``projection(...)``: fila de proyección, no punto de partida;
    * ``[A.1-A.1.13...]``, ``D.1/B.13``, ``R.2 - 8%``: cálculo interno, no COREP.
- Punto de partida: en hojas por periodos, la columna marcada ``Actual``/``Starting point``;
  en hojas de riesgo de crédito, las filas con ``Scenario = Actual``.

Referencia ST: ``ST:<hoja>!r<RowNum>c<nº columna>`` (los identificadores de la propia EBA).
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from .model import DatapointRef

_NUM_STR = re.compile(r"^\s*(\d{1,3})\s*$")
_COREP_COL_HDR = re.compile(r"C\s*\d{2}\.\d{2}[^,]*,\s*c\d{4}", re.I)
_START_WORDS = ("actual", "starting point")


def _as_int(v) -> int | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        m = _NUM_STR.match(v)
        return int(m.group(1)) if m else None
    return None


@dataclass
class StSheet:
    name: str
    header_row: int
    rownum_col: int
    colnums: dict[int, int]  # columna de la hoja -> nº de columna EBA
    rownums: dict[int, int]  # fila de la hoja -> RowNum
    map_col: int | None = None  # columna "COREP CODE / COMPUTATION"
    colmap_row: int | None = None  # fila con códigos de columna COREP
    attr_cols: dict[str, int] = field(default_factory=dict)  # 'scenario', 'geo', 'year'
    start_cols: set[int] = field(default_factory=set)  # columnas de punto de partida (hojas por periodo)
    warnings: list[str] = field(default_factory=list)

    def ref(self, sheet_row: int, sheet_col: int, framework: str = "ST") -> DatapointRef:
        return DatapointRef(framework, self.name, str(self.rownums[sheet_row]), str(self.colnums[sheet_col]))

    def is_starting_point_row(self, values, sheet_row: int) -> bool:
        col = self.attr_cols.get("scenario")
        if col is None:
            return True
        v = values.cell(sheet_row, col).value
        return isinstance(v, str) and v.strip().lower() in _START_WORDS


def parse_sheet(ws_values, ws_formulas=None, max_header_row: int = 25) -> StSheet | None:
    """Detecta el layout EBA en una hoja. ``ws_values``: libro con data_only=True."""
    wf = ws_formulas or ws_values
    header = rcol = None
    for r in range(1, max_header_row + 1):
        for c in range(1, min(ws_values.max_column, 60) + 1):
            v = ws_values.cell(r, c).value
            if isinstance(v, str) and v.strip() == "RowNum":
                header, rcol = r, c
                break
        if header:
            break
    if header is None:
        return None

    # fila con los números de columna (la que más enteros tenga a la derecha de RowNum)
    best_row, best = None, {}
    for r in range(1, header):
        nums = {}
        for c in range(rcol + 1, ws_values.max_column + 1):
            n = _as_int(ws_values.cell(r, c).value)
            if n is not None and 1 <= n <= 999:
                nums[c] = n
        if len(nums) > len(best):
            best_row, best = r, nums
    st = StSheet(ws_values.title, header, rcol, best, {})
    if not best:
        st.warnings.append("sin fila de números de columna")
    if best and sorted(best.values()) != list(range(1, len(best) + 1)):
        st.warnings.append("números de columna no consecutivos desde 1")

    for c in range(1, ws_values.max_column + 1):
        h = ws_values.cell(header, c).value
        if not isinstance(h, str):
            continue
        hl = h.strip().lower()
        if "corep code" in hl:
            st.map_col = c
        elif hl == "scenario":
            st.attr_cols["scenario"] = c
        elif hl.startswith("geographical breakdown"):
            st.attr_cols["geo"] = c
        elif hl == "year":
            st.attr_cols["year"] = c

    for r in range(1, header):
        hits = sum(1 for c in best if isinstance(wf.cell(r, c).value, str) and _COREP_COL_HDR.search(wf.cell(r, c).value))
        if hits:
            st.colmap_row = r
            break

    for c in best:
        for r in range(1, header + 1):
            v = ws_values.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() in _START_WORDS:
                st.start_cols.add(c)
                break
        else:
            if isinstance(ws_values.cell(header, c).value, (dt.datetime, dt.date)):
                st.start_cols.add(c)

    prev = None
    empty_run = 0
    for r in range(header + 1, ws_values.max_row + 1):
        n = _as_int(ws_values.cell(r, rcol).value)
        if n is None:
            f = wf.cell(r, rcol).value
            if isinstance(f, str) and re.fullmatch(rf"=\$?{get_column_letter(rcol)}\$?{r - 1}\+1", f.replace(" ", "")) and prev:
                n = prev + 1  # libro sin valores cacheados: RowNum = fila anterior + 1
        if n is None:
            empty_run += 1
            if empty_run > 50:
                break
            continue
        empty_run = 0
        st.rownums[r] = n
        prev = n
    if len(set(st.rownums.values())) != len(st.rownums):
        st.warnings.append("RowNum repetidos")
    return st


def row_label(ws_values, st: StSheet, r: int) -> str:
    first_num = min(st.colnums) if st.colnums else ws_values.max_column + 1
    parts = []
    for c in range(st.rownum_col + 1, first_num):
        h = ws_values.cell(st.header_row, c).value
        if isinstance(h, str) and h.strip().lower() in ("unique row identifier", "asset class identifier"):
            continue  # concatenaciones técnicas, no etiquetas
        v = ws_values.cell(r, c).value
        if isinstance(v, str) and v.strip() and v.strip() not in parts and v.strip() != "Pivot":
            parts.append(v.strip())
    return " | ".join(parts[-4:])


def col_label(ws_values, st: StSheet, c: int) -> str:
    parts = []
    for r in range(1, st.header_row + 1):
        v = ws_values.cell(r, c).value
        if isinstance(v, str) and v.strip() and not _COREP_COL_HDR.search(v) and v.strip() not in parts:
            parts.append(v.strip())
        elif isinstance(v, (dt.datetime, dt.date)):
            parts.append(v.date().isoformat() if isinstance(v, dt.datetime) else v.isoformat())
    return " | ".join(parts)


# ----------------------------------------------------------------- mapeo COREP
_FULL_TERM = re.compile(
    r"\s*(?P<sign>[+-])?\s*C\s*(?P<tpl>\d{2}\.\d{2}(?:\.[a-z0-9]+)?)\s*\(\s*r(?P<r>\d{4})\s*[,;]\s*c(?P<c>\d{4})\s*\)"
    r"(?:\s*\*\s*(?P<k>\d+(?:[.,]\d+)?))?\s*", re.I)
_ROWS_ALT = re.compile(r"^\s*C\s*(?P<tpl>\d{2}\.\d{2}(?:\.[a-z0-9]+)?)\s*,\s*(?P<codes>r\d{4}(?:\s*\+\s*r\d{4})*)\s*$", re.I)
_COLS_ALT = re.compile(r"^\s*C\s*(?P<tpl>\d{2}\.\d{2}(?:\.[a-z0-9]+)?)\s*,\s*(?P<codes>c\d{4}(?:\s*\+\s*c\d{4})*)\s*$", re.I)


def _tpl(code: str) -> str:
    return f"C {code.upper()}"


def parse_full_cells(text: str) -> list[tuple[float, str, str, str]] | None:
    """``C 01.00 (r0030,c0010) + C 01.00 (r0040,c0010) * 2.5`` -> [(coef, plantilla, fila, col)].

    Devuelve None si el texto contiene algo más que términos de celda (porcentajes,
    cálculos internos, MIN(...)...). Ignora una segunda línea ``[A.x+...]``.
    """
    text = text.split("\n[")[0].strip()
    if not text or "%" in text:
        return None
    pos, out = 0, []
    while pos < len(text):
        m = _FULL_TERM.match(text, pos)
        if not m or m.end() == pos or (out and not m.group("sign")):
            return None
        k = float(m.group("k").replace(",", ".")) if m.group("k") else 1.0
        out.append(((-k if m.group("sign") == "-" else k), _tpl(m.group("tpl")), m.group("r"), m.group("c")))
        pos = m.end()
    return out or None


def parse_alternatives(text: str, pattern: re.Pattern) -> dict[str, list[str]] | None:
    """``C 09.02, r0010 or C 09.05, r0010`` -> {"C 09.02": ["0010"], "C 09.05": ["0010"]}."""
    out: dict[str, list[str]] = {}
    for part in re.split(r"\s+or\s+", text.strip(), flags=re.I):
        m = pattern.match(part)
        if not m:
            return None
        out[_tpl(m.group("tpl"))] = re.findall(r"\d{4}", m.group("codes"))
    return out or None
