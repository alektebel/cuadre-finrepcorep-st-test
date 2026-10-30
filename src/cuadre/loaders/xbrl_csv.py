"""xBRL-CSV (EBA, arquitectura DPM 2.0 / marcos 4.x).

AVISO: implementado a partir de la especificación general de xBRL-CSV y del formato
de las tablas EBA (``datapoint,factValue``), **sin validar contra un paquete real**
(no había acceso a eba.europa.eu cuando se escribió). Hay que probarlo con un
paquete de ejemplo de la taxonomía antes de confiar en él.

Traducir ``datapoint`` (``dp12345``) a plantilla/fila/columna requiere la taxonomía
o la base de datos DPM. En vez de incrustar la taxonomía, el loader pide un CSV de
mapeo (``options.datapoint_map``) con columnas ``datapoint,template,row,col,z``.
Ese CSV se genera una vez por versión de marco a partir del DPM.

Acepta el ZIP del report package o la carpeta descomprimida.
"""
from __future__ import annotations

import csv
import io
import zipfile
from pathlib import Path

from ..model import DatapointRef, Fact, Source
from .base import CoverageEntry, LoadResult, SourceSpec


def _read_members(path: Path) -> dict[str, str]:
    if path.is_dir():
        return {str(p.relative_to(path)): p.read_text(encoding="utf-8-sig")
                for p in path.rglob("*.csv")}
    with zipfile.ZipFile(path) as zf:
        return {n: zf.read(n).decode("utf-8-sig") for n in zf.namelist() if n.lower().endswith(".csv")}


def load_datapoint_map(path: str | Path) -> dict[str, tuple[str, str, str, str | None]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return {r["datapoint"]: (r["template"], r["row"], r["col"], r.get("z") or None)
                for r in csv.DictReader(fh)}


def load(spec: SourceSpec) -> LoadResult:
    dmap_path = spec.options.get("datapoint_map")
    if not dmap_path:
        raise ValueError(f"{spec.path}: xbrl_csv necesita options.datapoint_map (CSV datapoint->celda)")
    dmap = load_datapoint_map(dmap_path)
    members = _read_members(Path(spec.path))
    out = LoadResult()

    decimals = spec.decimals
    for name, text in members.items():
        if name.endswith("parameters.csv"):
            params = {r["name"]: r["value"] for r in csv.DictReader(io.StringIO(text))}
            if decimals is None and params.get("decimalsMonetary") not in (None, "", "INF"):
                decimals = int(params["decimalsMonetary"])
    rounding = 0.0 if decimals is None else 0.5 * spec.scale * 10 ** (-decimals)

    for name, text in members.items():
        reader = csv.DictReader(io.StringIO(text))
        if not reader.fieldnames or not {"datapoint", "factValue"} <= set(reader.fieldnames):
            continue
        keys = [k for k in reader.fieldnames if k not in ("datapoint", "factValue")]
        n, unmapped, nonnum = 0, 0, 0
        for i, row in enumerate(reader, start=2):
            loc = dmap.get(row["datapoint"])
            if loc is None:
                unmapped += 1
                continue
            try:
                raw = float(row["factValue"])
            except (TypeError, ValueError):
                nonnum += 1
                continue
            tpl, r, c, z = loc
            if keys:  # tablas abiertas: la clave forma parte del eje z
                z = "|".join(filter(None, [z] + [f"{k}={row[k]}" for k in keys]))
            out.facts.append(Fact(DatapointRef(spec.framework, tpl, r, c, z), raw * spec.scale,
                                  Source(str(spec.path), name, f"línea {i}", detail=row["datapoint"]),
                                  raw_value=raw, scale=spec.scale, rounding_eur=rounding))
            n += 1
        w = []
        if unmapped:
            w.append(f"{unmapped} datapoints sin mapeo a celda")
        if nonnum:
            w.append(f"{nonnum} valores no numéricos")
        out.coverage.append(CoverageEntry(str(spec.path), name, None, None, 0, 0, n, w))
    return out
