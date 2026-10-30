"""Vistas tabulares de una ejecución, compartidas por la app y el informe Excel."""
from __future__ import annotations

import pandas as pd

from ..model import Status
from ..run import RunResult

STATUS_ORDER = [Status.KO.value, Status.NO_EVALUABLE.value, Status.REDONDEO.value, Status.OK.value]


def results_frame(run: RunResult) -> pd.DataFrame:
    rank = {lvl: i for i, lvl in enumerate(run.severity.levels)}
    rows = []
    for r in run.results:
        src = r.rule.source or {}
        rows.append({
            "regla": r.rule.id,
            "descripcion": r.rule.description,
            "origen": r.rule.origin.value,
            "tipo": r.rule.type.value,
            "estado": r.status.value,
            "gravedad": r.severity,
            "gravedad_abs": r.severity_abs,
            "gravedad_rel": r.severity_rel,
            "lhs_eur": r.lhs,
            "rhs_eur": r.rhs,
            "dif_bruta_eur": r.diff_raw,
            "ruido_redondeo_eur": r.rounding_noise,
            "dif_no_explicada_eur": r.diff_unexplained,
            "mensaje": r.message,
            "lhs_expr": r.rule.lhs,
            "rhs_expr": r.rule.rhs,
            "documento": src.get("documento"),
            "apartado": src.get("apartado"),
            "pagina": src.get("pagina"),
            "cita": src.get("cita"),
            "fichero_regla": src.get("_fichero_regla"),
            "_orden": (STATUS_ORDER.index(r.status.value), -rank.get(r.severity, -1), -(r.diff_unexplained or 0)),
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values("_orden").drop(columns="_orden").reset_index(drop=True)


def operands_frame(run: RunResult) -> pd.DataFrame:
    rows = []
    for r in run.results:
        for o in r.operands:
            f = o.fact
            rows.append({
                "regla": r.rule.id,
                "estado_regla": r.status.value,
                "lado": o.side,
                "coef": o.coef,
                "token": o.token,
                "referencia": str(o.ref) if o.ref else None,
                "marco": o.ref.framework if o.ref else None,
                "plantilla": o.ref.template if o.ref else None,
                "fila": o.ref.row if o.ref else None,
                "columna": o.ref.col if o.ref else None,
                "z": o.ref.z if o.ref else None,
                "valor_eur": f.value if f else None,
                "contribucion_eur": o.contribution if f else None,
                "valor_fichero": f.raw_value if f else None,
                "escala": f.scale if f else None,
                "redondeo_eur": f.rounding_eur if f else None,
                "fichero": f.source.file if f else None,
                "hoja": f.source.sheet if f else None,
                "celda": f.source.cell if f else None,
                "detalle_origen": f.source.detail if f else None,
                "nota": o.note,
            })
    return pd.DataFrame(rows)


def facts_frame(run: RunResult) -> pd.DataFrame:
    flagged: dict[str, str] = {}
    for r in run.results:
        if r.status in (Status.KO, Status.REDONDEO):
            for o in r.operands:
                if o.ref:
                    key = str(o.ref)
                    if flagged.get(key) != Status.KO.value:
                        flagged[key] = r.status.value
    rows = []
    for f in run.store.facts.values():
        rows.append({
            "referencia": str(f.ref),
            "marco": f.ref.framework,
            "plantilla": f.ref.template,
            "fila": f.ref.row,
            "columna": f.ref.col,
            "z": f.ref.z,
            "valor_eur": f.value,
            "valor_fichero": f.raw_value,
            "escala": f.scale,
            "redondeo_eur": f.rounding_eur,
            "fichero": f.source.file,
            "hoja": f.source.sheet,
            "celda": f.source.cell,
            "implicado_en": flagged.get(str(f.ref)),
        })
    return pd.DataFrame(rows)


def coverage_frame(run: RunResult) -> pd.DataFrame:
    return pd.DataFrame([{
        "fichero": c.file, "hoja": c.sheet, "plantilla": c.template, "z": c.z,
        "n_filas": c.n_rows, "n_columnas": c.n_cols, "n_hechos": c.n_facts,
        "avisos": " | ".join(c.warnings),
    } for c in run.coverage])


def conflicts_frame(run: RunResult) -> pd.DataFrame:
    return pd.DataFrame([{
        "referencia": str(a.ref), "valor_1": a.value, "origen_1": str(a.source),
        "valor_2": b.value, "origen_2": str(b.source),
    } for a, b in run.store.conflicts])
