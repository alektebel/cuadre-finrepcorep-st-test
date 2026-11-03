"""Plantillas ST de la EBA rellenas por la entidad (layout ``RowNum`` + columnas numeradas).

Cada valor numérico de la rejilla RowNum × nº de columna se convierte en un hecho
``ST:<hoja>!r<RowNum>c<col>``. Opciones:

- ``sheets``: lista de hojas a leer (por defecto, todas las que tengan ``RowNum``).
- ``ignore_sheets``: hojas a omitir.
"""
from __future__ import annotations

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ..model import Fact, Source
from ..st_template import parse_sheet
from .base import CoverageEntry, LoadResult, SourceSpec


def load(spec: SourceSpec) -> LoadResult:
    opts = spec.options or {}
    only = set(opts.get("sheets") or [])
    ignore = set(opts.get("ignore_sheets") or [])
    wb = load_workbook(spec.path, data_only=True)
    wf = None  # libro con fórmulas, solo si faltan valores cacheados de RowNum
    out = LoadResult()
    for ws in wb.worksheets:
        if ws.title in ignore or (only and ws.title not in only):
            continue
        st = parse_sheet(ws)
        if st is None:
            continue
        if not st.rownums:
            wf = wf or load_workbook(spec.path, data_only=False)
            st = parse_sheet(ws, wf[ws.title])
        warnings, n, nonnum = list(st.warnings), 0, 0
        for r, rownum in st.rownums.items():
            for c, colnum in st.colnums.items():
                v = ws.cell(r, c).value
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                if isinstance(v, bool) or not isinstance(v, (int, float)):
                    nonnum += 1
                    continue
                out.facts.append(Fact(st.ref(r, c, spec.framework), float(v) * spec.scale,
                                      Source(str(spec.path), ws.title, f"{get_column_letter(c)}{r}"),
                                      raw_value=float(v), scale=spec.scale, rounding_eur=spec.rounding_eur))
                n += 1
        if nonnum:
            warnings.append(f"{nonnum} valores no numéricos en la rejilla (ignorados)")
        out.coverage.append(CoverageEntry(str(spec.path), ws.title, ws.title, None,
                                          len(st.rownums), len(st.colnums), n, warnings))
    return out
