"""Excel con layout propio, descrito en un YAML (``options.layout``).

Para exports que no siguen el layout de anexo EBA. Ejemplo de layout::

    hojas:
      - hoja: Balance
        plantilla: F 01.01
        z: null
        celdas:            # celda -> fila/columna EBA
          D12: r0010c0010
        rango:             # o bien una rejilla
          filas: {"0010": 12, "0020": 13}
          columnas: {"0010": D, "0020": E}
"""
from __future__ import annotations

from pathlib import Path

import yaml
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from ..model import DatapointRef, Fact, Source
from .base import CoverageEntry, LoadResult, SourceSpec


def _rc(code: str) -> tuple[str, str]:
    code = code.strip().lower()
    if not code.startswith("r") or "c" not in code:
        raise ValueError(f"Código {code!r} no válido; formato rFILAcCOL")
    r, c = code[1:].split("c", 1)
    return r, c


def load(spec: SourceSpec) -> LoadResult:
    layout_path = spec.options.get("layout")
    if not layout_path:
        raise ValueError(f"{spec.path}: el loader excel_declared necesita options.layout")
    layout = yaml.safe_load(Path(layout_path).read_text(encoding="utf-8"))
    wb = load_workbook(spec.path, data_only=True)
    out = LoadResult()
    for h in layout.get("hojas", []):
        ws = wb[h["hoja"]]
        tpl, z = h["plantilla"], h.get("z")
        cells: dict[str, tuple[str, str]] = {k: _rc(v) for k, v in (h.get("celdas") or {}).items()}
        grid = h.get("rango") or {}
        for rcode, rown in (grid.get("filas") or {}).items():
            for ccode, coll in (grid.get("columnas") or {}).items():
                col_idx = column_index_from_string(coll) if isinstance(coll, str) else int(coll)
                cells[ws.cell(row=int(rown), column=col_idx).coordinate] = (str(rcode), str(ccode))
        warnings, n = [], 0
        for addr, (r, c) in cells.items():
            v = ws[addr].value
            if v is None:
                continue
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                warnings.append(f"{addr}: valor no numérico {v!r}")
                continue
            out.facts.append(Fact(DatapointRef(spec.framework, tpl, r, c, z), float(v) * spec.scale,
                                  Source(str(spec.path), ws.title, addr), raw_value=float(v),
                                  scale=spec.scale, rounding_eur=spec.rounding_eur))
            n += 1
        missing = len(cells) - n - len(warnings)
        if missing:
            warnings.append(f"{missing} celdas declaradas vacías")
        out.coverage.append(CoverageEntry(str(spec.path), ws.title, tpl, z, 0, 0, n, warnings))
    return out
