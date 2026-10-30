"""Informe Excel trazable.

Recorrido de un descuadre: ``Resultados`` (fila de la regla) -> hipervínculo a ``Detalle``
(un renglón por operando con valor, escala, redondeo y origen) -> hipervínculo a la
celda del fichero de entrada (si es Excel) o referencia al CSV/contexto XBRL.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from ..run import RunResult
from .frames import conflicts_frame, coverage_frame, facts_frame, operands_frame, results_frame

FILL = {
    "Crítico": "C00000", "Grave": "FF4D4D", "Moderado": "FFA64D", "Leve": "FFE066",
    "REDONDEO": "BDD7EE", "OK": "C6EFCE", "NO_EVALUABLE": "D9D9D9", "KO": "FF4D4D",
}
EUR_FMT = '#,##0.00;[Red]-#,##0.00'


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", start_color=color, end_color=color)


def _write_df(ws, df, eur_cols=()):
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = _fill("1F3864")
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and v != v) else v for v in row])
    for i, col in enumerate(df.columns, start=1):
        letter = get_column_letter(i)
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200) if v is not None])
        ws.column_dimensions[letter].width = min(max(10, width + 2), 60)
        if col in eur_cols:
            for cell in ws[letter][1:]:
                cell.number_format = EUR_FMT
    ws.freeze_panes = "A2"
    if len(df):
        ws.auto_filter.ref = ws.dimensions


def _internal_link(cell, location: str) -> None:
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=location)
    cell.font = Font(color="0563C1", underline="single")


def _source_link(file: str, sheet: str | None, cell: str | None, out_dir: Path) -> str | None:
    if not all(isinstance(x, str) and x for x in (file, sheet, cell)) or not file.lower().endswith((".xlsx", ".xlsm")):
        return None
    rel = os.path.relpath(file, out_dir)
    return f"{rel}#'{sheet}'!{cell}"


def write_report(run: RunResult, out: str | Path) -> Path:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    res, ops, facts = results_frame(run), operands_frame(run), facts_frame(run)
    wb = Workbook()

    # Resumen
    ws = wb.active
    ws.title = "Resumen"
    ws.append(["Cuadre ST ↔ FINREP/COREP"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([f"Ejercicio {run.metadata['ejercicio']} · referencia {run.metadata['fecha_referencia']}"])
    ws.append([])
    ws.append(["Estado", "Nº reglas"])
    for st in ["KO", "NO_EVALUABLE", "REDONDEO", "OK"]:
        n = int((res["estado"] == st).sum()) if len(res) else 0
        ws.append([st, n])
        ws.cell(ws.max_row, 1).fill = _fill(FILL[st])
    ws.append([])
    ws.append(["Gravedad (KO)", "Nº reglas"])
    for lvl in reversed(run.severity.levels):
        n = int(((res["estado"] == "KO") & (res["gravedad"] == lvl)).sum()) if len(res) else 0
        ws.append([lvl, n])
        ws.cell(ws.max_row, 1).fill = _fill(FILL.get(lvl, "FFFFFF"))
    ws.append([])
    ws.append(["Origen", "KO", "Total"])
    for origin in sorted(res["origen"].unique()) if len(res) else []:
        sub = res[res["origen"] == origin]
        ws.append([origin, int((sub["estado"] == "KO").sum()), len(sub)])
    ws.column_dimensions["A"].width = 28

    # Resultados con enlace al detalle
    ws = wb.create_sheet("Resultados")
    res_x = res.copy()
    res_x.insert(0, "ver_detalle", "→ detalle")
    _write_df(ws, res_x, eur_cols={"lhs_eur", "rhs_eur", "dif_bruta_eur", "ruido_redondeo_eur", "dif_no_explicada_eur"})
    first_row = {}
    for i, rid in enumerate(ops["regla"] if len(ops) else [], start=2):
        first_row.setdefault(rid, i)
    col_estado = list(res_x.columns).index("estado") + 1
    col_grav = list(res_x.columns).index("gravedad") + 1
    for i, rid in enumerate(res_x["regla"], start=2):
        if rid in first_row:
            _internal_link(ws.cell(i, 1), f"'Detalle'!A{first_row[rid]}")
        st = ws.cell(i, col_estado)
        st.fill = _fill(FILL.get(st.value, "FFFFFF"))
        g = ws.cell(i, col_grav)
        if g.value in FILL:
            g.fill = _fill(FILL[g.value])
            if g.value == "Crítico":
                g.font = Font(color="FFFFFF", bold=True)

    # Detalle por operando con enlace al origen
    ws = wb.create_sheet("Detalle")
    ops_x = ops.copy()
    ops_x.insert(1, "volver", "← resultado")
    ops_x["abrir_origen"] = [
        _source_link(f, h, c, out.parent.resolve()) for f, h, c in zip(ops_x.get("fichero", []), ops_x.get("hoja", []), ops_x.get("celda", []))
    ] if len(ops_x) else []
    _write_df(ws, ops_x, eur_cols={"valor_eur", "contribucion_eur", "redondeo_eur"})
    res_row = {rid: i for i, rid in enumerate(res_x["regla"], start=2)} if len(res_x) else {}
    link_col = list(ops_x.columns).index("abrir_origen") + 1
    for i, (rid, st, link) in enumerate(zip(ops_x.get("regla", []), ops_x.get("estado_regla", []), ops_x.get("abrir_origen", [])), start=2):
        _internal_link(ws.cell(i, 2), f"'Resultados'!A{res_row.get(rid, 1)}")
        if isinstance(link, str):
            lc = ws.cell(i, link_col)
            lc.hyperlink = link
            lc.value = "abrir celda"
            lc.font = Font(color="0563C1", underline="single")
        ws.cell(i, 1).fill = _fill(FILL.get(st, "FFFFFF"))

    ws = wb.create_sheet("No_evaluables")
    _write_df(ws, res[res["estado"] == "NO_EVALUABLE"][["regla", "descripcion", "mensaje", "lhs_expr", "rhs_expr"]] if len(res) else res)

    # Datos de entrada por marco, con celdas implicadas resaltadas
    for fw in sorted(facts["marco"].unique()) if len(facts) else []:
        ws = wb.create_sheet(f"Datos_{fw}"[:31])
        sub = facts[facts["marco"] == fw].sort_values(["plantilla", "z", "fila", "columna"], na_position="first")
        _write_df(ws, sub, eur_cols={"valor_eur", "redondeo_eur"})
        col_imp = list(sub.columns).index("implicado_en") + 1
        for i, v in enumerate(sub["implicado_en"], start=2):
            if v == "KO":
                for c in ws[i]:
                    c.fill = _fill("FFC7CE")
            elif v == "REDONDEO":
                ws.cell(i, col_imp).fill = _fill(FILL["REDONDEO"])

    ws = wb.create_sheet("Cobertura_carga")
    _write_df(ws, coverage_frame(run))
    conf = conflicts_frame(run)
    if len(conf):
        ws = wb.create_sheet("Conflictos")
        _write_df(ws, conf)

    ws = wb.create_sheet("Metadatos")
    for k, v in run.metadata.items():
        ws.append([k, json.dumps(v, ensure_ascii=False, indent=1) if isinstance(v, (dict, list)) else v])
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 100
    for row in ws.iter_rows():
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out)
    return out
