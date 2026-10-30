"""xBRL-XML (instancias EBA de marcos 3.x).

AVISO: lector mínimo y **sin validar contra instancias reales**. No resuelve la
taxonomía: cada hecho se identifica por su *clave* ``métrica|dim1=miembro1;dim2=...``
(dimensiones ordenadas) y un CSV de mapeo (``options.fact_map``, columnas
``key,template,row,col,z``) la traduce a celda. Para uso serio conviene sustituir este
lector por Arelle con la taxonomía cargada.
"""
from __future__ import annotations

import csv
from pathlib import Path

from lxml import etree

from ..model import DatapointRef, Fact, Source
from .base import CoverageEntry, LoadResult, SourceSpec

NS = {
    "xbrli": "http://www.xbrl.org/2003/instance",
    "xbrldi": "http://xbrl.org/2006/xbrldi",
}


def _local(qname: str) -> str:
    return qname.split(":")[-1]


def fact_key(metric: str, dims: dict[str, str]) -> str:
    d = ";".join(f"{_local(k)}={_local(v)}" for k, v in sorted(dims.items()))
    return f"{_local(metric)}|{d}" if d else _local(metric)


def load(spec: SourceSpec) -> LoadResult:
    fmap_path = spec.options.get("fact_map")
    if not fmap_path:
        raise ValueError(f"{spec.path}: xbrl_xml necesita options.fact_map (CSV key->celda)")
    with open(fmap_path, encoding="utf-8-sig", newline="") as fh:
        fmap = {r["key"]: (r["template"], r["row"], r["col"], r.get("z") or None) for r in csv.DictReader(fh)}

    tree = etree.parse(str(spec.path))
    root = tree.getroot()
    contexts: dict[str, dict[str, str]] = {}
    for ctx in root.iterfind("xbrli:context", NS):
        dims = {m.get("dimension"): (m.text or "").strip()
                for m in ctx.iterfind(".//xbrldi:explicitMember", NS)}
        for t in ctx.iterfind(".//xbrldi:typedMember", NS):
            dims[t.get("dimension")] = "".join(t.itertext()).strip()
        contexts[ctx.get("id")] = dims

    out = LoadResult()
    n, unmapped = 0, 0
    for el in root:
        ctx_id = el.get("contextRef")
        if ctx_id is None or el.get("unitRef") is None:
            continue  # solo hechos numéricos
        metric = etree.QName(el).localname
        key = fact_key(metric, contexts.get(ctx_id, {}))
        loc = fmap.get(key)
        if loc is None:
            unmapped += 1
            continue
        raw = float(el.text)
        dec = el.get("decimals")
        rounding = 0.0 if dec in (None, "INF") else 0.5 * spec.scale * 10 ** (-int(dec))
        tpl, r, c, z = loc
        out.facts.append(Fact(DatapointRef(spec.framework, tpl, r, c, z), raw * spec.scale,
                              Source(str(spec.path), None, f"contextRef={ctx_id}", detail=key),
                              raw_value=raw, scale=spec.scale, rounding_eur=rounding))
        n += 1
    w = [f"{unmapped} hechos numéricos sin mapeo a celda"] if unmapped else []
    out.coverage.append(CoverageEntry(str(spec.path), "instancia", None, None, 0, 0, n, w))
    return out
